#!/usr/bin/env python3
"""
bia_kit.py - spreadsheet-first metadata entry for BioImage Archive submissions.

Fill ONE Excel workbook (one row per author / biosample / specimen ...) and
generate the BioStudies PageTab TSV (template BioImages.v5) that the BioImage
Archive submission tool produces.

  * Long, repeated text (sample-prep protocol, acquisition, analysis) is written
    once and referenced by name from a dropdown.
  * REMBI "Associations" are generated from the Specimens sheet: one row per
    imaged specimen instead of a Specimen block + an Associations block in the form.
  * Biosample descriptions can be built from structured "desc:<Label>" columns.
  * Everything is validated before a TSV is written (required fields, REMBI
    minimum lengths, ORCID checksums, broken/unused references, ...).

Commands
  template    OUT.xlsx                    empty workbook
  import-tsv  PAGETAB.tsv OUT.xlsx        pre-fill a workbook from an existing PageTab TSV
  check       WORKBOOK.xlsx               validate only (exit 1 on errors)
  build       WORKBOOK.xlsx OUT.tsv       validate + write PageTab TSV (refuses on errors)
  diff        A.tsv B.tsv                 semantic comparison of two PageTab TSVs
  xenium-scan ROOT [ROOT ...] --out DIR   read experiment.xenium of each Xenium output
                                          bundle, write bundle table + file list, and
                                          (with --workbook) cross-check the workbook

Python >= 3.9. openpyxl is needed for workbook commands only; `diff` and
`xenium-scan` without --workbook are stdlib-only (run them on the cluster).
"""
from __future__ import annotations

import argparse
import csv
import difflib
import hashlib
import io
import json
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path

__version__ = "0.1.0"

# =============================================================== PageTab model
@dataclass
class Attr:
    name: str
    value: str = ""
    quals: list = field(default_factory=list)  # [("[", "Ontology", "NCBITaxon"), ...]
    ref: bool = False                           # <affiliation>-style reference


@dataclass
class Block:
    kind: str
    accno: str = ""
    parent: str = ""
    attrs: list = field(default_factory=list)

    def get(self, name: str) -> str:
        for a in self.attrs:
            if a.name == name:
                return a.value
        return ""

    def getall(self, name: str) -> list:
        return [a for a in self.attrs if a.name == name]

    def add(self, name, value="", quals=None, ref=False):
        self.attrs.append(Attr(name, value, list(quals or []), ref))


def parse_pagetab(path) -> list:
    """Parse a PageTab TSV made of blank-line separated blocks (form/export style)."""
    with open(path, newline="", encoding="utf-8-sig") as fh:
        return parse_pagetab_text(fh.read(), str(path))


def parse_pagetab_text(text: str, path: str = "<text>") -> list:
    rows = list(csv.reader(io.StringIO(text.lstrip("\ufeff"), newline=""), delimiter="\t"))
    blocks, cur = [], None
    for lineno, row in enumerate(rows, 1):
        while len(row) > 2 and row[-1] == "":
            row.pop()
        if not row or all(not c.strip() for c in row):
            cur = None
            continue
        key, val = row[0], (row[1] if len(row) > 1 else "")
        if cur is None:
            cur = Block(key, val, row[2] if len(row) > 2 else "")
            blocks.append(cur)
            continue
        if len(row) > 2:
            raise ValueError(f"{path}: row {lineno}: tables inside blocks are not supported: {row[:4]}")
        if len(key) > 1 and key[0] + key[-1] in ("[]", "()"):
            if not cur.attrs:
                raise ValueError(f"{path}: row {lineno}: qualifier {key} without attribute")
            cur.attrs[-1].quals.append((key[0], key[1:-1], val))
        elif len(key) > 1 and key[0] == "<" and key[-1] == ">":
            cur.add(key[1:-1], val, ref=True)
        else:
            cur.add(key, val)
    return blocks


def _fmt(v) -> str:
    v = "" if v is None else str(v)
    v = v.replace("\r\n", "\n").replace("\r", "\n").replace("\t", " ")
    if "\n" in v:  # PageTab: multi-line values are wrapped in double quotes
        return '"' + v.replace('"', '""') + '"'
    return v


def render_pagetab(blocks) -> str:
    out = []
    for i, b in enumerate(blocks):
        if i:
            out.append("")
        out.append("\t".join([b.kind, b.accno] + ([b.parent] if b.parent else [])))
        for a in b.attrs:
            out.append(f"{'<%s>' % a.name if a.ref else a.name}\t{_fmt(a.value)}")
            for q, qn, qv in a.quals:
                out.append(f"{q}{qn}{']' if q == '[' else ')'}\t{_fmt(qv)}")
    return "\n".join(out) + "\n"


def _block_key(b: Block):
    attrs = sorted(
        (("<>" if a.ref else "") + a.name, a.value.strip(), tuple(tuple(q) for q in a.quals))
        for a in b.attrs
    )
    return (b.kind, b.accno, b.parent, tuple(attrs))


def diff_blocks(a_blocks, b_blocks, a_name="A", b_name="B") -> list:
    """Semantic diff: block order matters, attribute order inside a block does not."""
    ka, kb = [_block_key(b) for b in a_blocks], [_block_key(b) for b in b_blocks]
    out = []

    def label(b):
        t = b.get("Title") or b.get("Name") or b.get("Specimen") or ""
        return f"{b.kind} {b.accno or '-'} (parent {b.parent or '-'}) {t[:50]!r}"

    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(a=ka, b=kb, autojunk=False).get_opcodes():
        if op == "equal":
            continue
        if op == "replace" and (i2 - i1) == (j2 - j1):
            for x, y in zip(a_blocks[i1:i2], b_blocks[j1:j2]):
                xa, ya = set(_block_key(x)[3]), set(_block_key(y)[3])
                hdr = [] if _block_key(x)[:3] == _block_key(y)[:3] else [f"header {x.kind}/{x.accno}/{x.parent} -> {y.kind}/{y.accno}/{y.parent}"]
                out.append(f"~ {label(x)}: " + "; ".join(
                    hdr + [f"only in {a_name}: {n}={v[:60]!r}" for n, v, _ in sorted(xa - ya)]
                    + [f"only in {b_name}: {n}={v[:60]!r}" for n, v, _ in sorted(ya - xa)]))
            continue
        for b in a_blocks[i1:i2]:
            out.append(f"- only in {a_name}: {label(b)}")
        for b in b_blocks[j1:j2]:
            out.append(f"+ only in {b_name}: {label(b)}")
    return out


# ============================================================ workbook schema
@dataclass
class Col:
    header: str
    required: bool = False
    width: int = 18
    help: str = ""
    ref: str = ""        # sheet whose first column supplies a dropdown
    wrap: bool = False


STUDY_FIELDS = [
    ("Accession", False,
     "The archive's accession number for this submission, if one has already been assigned. "
     "Leave empty for a new submission: the archive assigns one."),
    ("Template", True,
     "Name of the submission template this file follows. Keep the pre-filled value unless the archive asks otherwise."),
    ("DOI", False, "The study DOI, only if the archive has already issued one. Otherwise leave empty."),
    ("AttachTo", True,
     "The archive collection the study belongs to. Keep the pre-filled value for BioImage Archive submissions."),
    ("Title", True,
     "A descriptive title for the dataset, shown in search results. At least 25 characters; "
     "often the same as the title of the associated publication."),
    ("Release date", True,
     "Date on which the data become public, written as year-month-day (YYYY-MM-DD). "
     "The submission stays private until then."),
    ("Description", True,
     "What the dataset contains and why it was produced. At least 25 characters; "
     "the abstract of the associated publication is acceptable."),
    ("Keywords", True,
     "Words that help others find the data: organism, tissue or cell type, technique, biological topic. "
     "Separate several keywords with ';' or put each on its own line."),
    ("Acknowledgements", False, "People, facilities or groups to thank who are not authors. Optional."),
    ("License", True,
     "Licence under which the data are released; choose one of the two the archive offers: "
     "CC0 (no conditions) or CC BY 4.0 (re-users must credit you)."),
    ("License URL", False, "Leave empty: it is filled in automatically from the chosen licence."),
    ("Funding statement", False,
     "Funding acknowledgement as it should appear with the data (free text). "
     "Also list each grant separately on the Grants sheet."),
    ("_study_accno", False, "Internal identifier used by the submission tool. Leave as is; empty means it will be generated."),
]

_INTERNAL = "Internal identifier used by the submission tool. Leave as is (empty is fine for new rows)."
TABLES = {
    "Authors": [
        Col("Name", True, 24, "Full name of the person as it should be credited, or the name of a group or consortium."),
        Col("E-mail", False, 30, "Contact e-mail address. Required for the submitter and for corresponding authors."),
        Col("ORCID", False, 22,
            "The person's ORCID iD: four blocks of four characters separated by hyphens (the last character may be X). "
            "The check digit is verified."),
        Col("Roles", False, 34,
            "The person's role(s) in the study, separated by ';'. At least one author must have the role 'submitter'."),
        Col("Affiliations", True, 16,
            "Key(s) of the person's organisation(s) as defined on the Organisations sheet, separated by ';'."),
        Col("_accno", width=10, help=_INTERNAL),
    ],
    "Organisations": [
        Col("Key", True, 10,
            "A short label you choose for the organisation (letters, digits, '-' or '_'). "
            "It only links authors to the organisation and is not published as such."),
        Col("Name", True, 44, "Full official name of the organisation."),
        Col("ROR ID", False, 30,
            "The organisation's Research Organization Registry identifier as a full link (https://ror.org/ followed by "
            "the identifier). Look it up on ror.org. Recommended."),
        Col("Address", False, 30, "Postal address of the organisation. Optional."),
    ],
    "Links": [
        Col("URL", True, 44,
            "Full web address (starting with http:// or https://) of a related resource: code, data portal, related dataset."),
        Col("Description", False, 50, "One line saying what the link points to."),
    ],
    "Publications": [
        Col("Title", False, 50, "Title of the article that describes or uses these data."),
        Col("Authors", False, 40, "Author list as printed in the article."),
        Col("DOI", False, 26, "The article's DOI, starting with '10.' (without the https://doi.org/ prefix)."),
        Col("Year", False, 8, "Year of publication, four digits."),
        Col("PubMed ID", False, 12, "The article's PubMed identifier (digits), if it is indexed in PubMed. Optional."),
        Col("_accno", width=10, help=_INTERNAL),
    ],
    "Grants": [
        Col("Agency", False, 40, "Name of the funding organisation."),
        Col("Grant ID", False, 24, "The grant or award number exactly as issued by the funder."),
        Col("_accno", width=10, help=_INTERNAL),
    ],
    "Biosamples": [
        Col("Title", True, 14,
            "A short, unique name for one biological source (the identifier your lab uses for the donor, animal, "
            "culture or line). Specimens refer to it by this name."),
        Col("Organism", True, 24, "Species of the sample, written as its scientific name."),
        Col("NCBITaxon ID", False, 12,
            "Numeric NCBI Taxonomy identifier of the species (digits only). Filled in automatically for common "
            "model organisms; the web form can also look it up."),
        Col("Biological entity", True, 30,
            "What was imaged, in words: the organ, tissue, cell type or structure, and its state or condition."),
        Col("Experimental variable", False, 22,
            "What was deliberately varied between samples in this study. Optional."),
        Col("Intrinsic variable", False, 18,
            "Inherent properties of the sample relevant for interpretation, such as genetic background or disease state. Optional."),
        Col("Extrinsic variable", False, 18,
            "External factors the sample was exposed to, such as treatments or environmental conditions. Optional."),
        Col("Description", False, 30,
            "Any other information about the sample, as free text. Structured facts can instead go into "
            "descriptor columns (headers starting with 'desc:'), each written as a 'label: value' line.", wrap=True),
        Col("_accno", width=10, help=_INTERNAL),
    ],
    "Protocols": [
        Col("Key", True, 22, "A short name for this sample-preparation procedure; specimens select it by this name."),
        Col("Sample preparation protocol", True, 100,
            "How specimens were prepared for imaging: collection, fixation, embedding, sectioning, staining or "
            "labelling, mounting. Cite kit and protocol identifiers where they exist.", wrap=True),
    ],
    "Acquisitions": [
        Col("Title", True, 28, "A short name for this imaging setup; specimens select it by this name."),
        Col("Imaging instrument", True, 26, "Microscope or imaging platform used, with manufacturer and model."),
        Col("Image acquisition parameters", True, 70,
            "Settings needed to interpret or reuse the images: objective or magnification, pixel/voxel size, channels "
            "and what they label, exposure, z-sampling, tiling, acquisition software and version.", wrap=True),
        Col("Imaging method", True, 24,
            "The imaging technique, given as the label of a term from the Biological Imaging Methods Ontology (FBbi)."),
        Col("Imaging method ontology", False, 10,
            "Ontology the term comes from. Filled in automatically when an FBbi term identifier is given."),
        Col("Imaging method term ID", False, 16, "Identifier of the FBbi term: 'FBbi_' followed by eight digits."),
        Col("_accno", width=10, help=_INTERNAL),
    ],
    "Analyses": [
        Col("Title", True, 22,
            "A short name for this processing/analysis procedure; specimens select it by this name. "
            "Define one entry per software version if several versions were used."),
        Col("Image analysis overview", True, 70,
            "How the images were processed or analysed: software and versions, main steps and parameters, "
            "and which results are included.", wrap=True),
        Col("_accno", width=10, help=_INTERNAL),
    ],
    "Correlations": [
        Col("Title", True, 22, "A short name for this correlation procedure; specimens select it by this name."),
        Col("Spatial and temporal alignment", False, 40,
            "How images from different modalities or time points were aligned to each other.", wrap=True),
        Col("Fiducials used", False, 30, "Reference markers used for the alignment, if any.", wrap=True),
        Col("Transformation matrix", False, 30,
            "Where the transformation between image spaces is stored, or how it is described.", wrap=True),
        Col("_accno", width=10, help=_INTERNAL),
    ],
    "StudyComponents": [
        Col("Name", True, 26,
            "A short name for a group of data described together (same type of experiment or images); "
            "specimens select it by this name."),
        Col("Description", True, 40, "What this group of data contains.", wrap=True),
        Col("File List", True, 26,
            "File name of the file list uploaded for this group: a table whose first column header is 'Files' "
            "and that lists one data file per row."),
        Col("_accno", width=10, help=_INTERNAL),
    ],
    "Specimens": [
        Col("Specimen title", True, 28,
            "Unique name of one imaged specimen, as used in your records or on the instrument. One row per specimen."),
        Col("Biosample", True, 14, "The biological source this specimen comes from (choose from Biosamples).",
            ref="Biosamples"),
        Col("Protocol", True, 22,
            "How this specimen was prepared (choose from Protocols). May stay empty when only one protocol is defined.",
            ref="Protocols"),
        Col("Image acquisition", True, 28,
            "How this specimen was imaged (choose from Acquisitions). May stay empty when only one is defined.",
            ref="Acquisitions"),
        Col("Image analysis", False, 22, "How the images of this specimen were analysed (choose from Analyses). Optional.",
            ref="Analyses"),
        Col("Image correlation", False, 18, "Correlation procedure applied, if any (choose from Correlations). Optional.",
            ref="Correlations"),
        Col("Study component", True, 24,
            "Group of data the images belong to (choose from StudyComponents). May stay empty when only one is defined.",
            ref="StudyComponents"),
        Col("_spec_accno", width=10, help=_INTERNAL),
        Col("_asso_accno", width=10, help=_INTERNAL),
    ],
}

LICENSES = {
    "CC0": "https://creativecommons.org/publicdomain/zero/1.0/legalcode",
    "CC BY 4.0": "https://creativecommons.org/licenses/by/4.0/legalcode",
}
# NCBI Taxonomy ids for common model organisms (keyed by lower-case scientific name)
TAXA = {
    "homo sapiens": "9606", "mus musculus": "10090", "rattus norvegicus": "10116",
    "danio rerio": "7955", "drosophila melanogaster": "7227", "caenorhabditis elegans": "6239",
    "saccharomyces cerevisiae": "4932", "arabidopsis thaliana": "3702", "xenopus laevis": "8355",
    "sus scrofa": "9823", "macaca mulatta": "9544", "gallus gallus": "9031",
}
DESC_PREFIX = "desc:"
GUIDE_MARK = "ⓘ "   # row 2 of every table sheet explains its columns; rows starting with this are not data
TEMPLATE_DEFAULTS = {"Template": "BioImages.v5", "AttachTo": "BioImages"}
README_LINES = [
    ("BioImage Archive submission workbook", "title"),
    (f"bia_kit {__version__}. This workbook is converted into the submission file (PageTab TSV) for the BioImage Archive.", ""),
    ("", ""),
    ("How to fill it", "h"),
    ("Every sheet has a grey row under the headers that explains each column. Enter data from the row below it.", ""),
    ("1. Study: one value per field; the Explanation column says what each field needs.", ""),
    ("2. Organisations first, then Authors. Authors refer to organisations by their Key; several keys or roles", ""),
    ("   in one cell are separated by ';'.", ""),
    ("3. Describe each procedure once: Protocols (sample preparation), Acquisitions (imaging), Analyses", ""),
    ("   (processing) and, only if needed, Correlations. Specimens point to them by name.", ""),
    ("4. Biosamples: one row per biological source. For structured facts that every sample has, add a column", ""),
    ("   whose header is 'desc:' followed by a label; each value becomes a 'label: value' line in the description.", ""),
    ("5. Specimens: one row per imaged specimen. Choose its biosample, protocol, acquisition, analysis and study", ""),
    ("   component from the drop-down lists. Each row becomes a specimen plus its association in the submission.", ""),
    ("6. Convert:  python bia_kit.py build <this workbook> <output>.tsv   (checks everything first and refuses on errors),", ""),
    ("   or open the workbook in the web form (bia-submission-form.html) and use 'Download PageTab TSV'.", ""),
    ("", ""),
    ("Conventions", "h"),
    ("Orange headers are required, blue headers optional. Grey headers starting with '_' are internal identifiers:", ""),
    ("leave them as they are (duplicates created by copying rows are detected and replaced).", ""),
    ("All cells are formatted as text on purpose, so identifiers and dates are not altered by the spreadsheet program.", ""),
    ("Empty rows on Publications and Grants are ignored.", ""),
]


# ============================================================ helpers
def cellstr(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime):
        return v.date().isoformat() if v.time() == time(0) else v.isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return str(v).replace("\r\n", "\n").strip()


def split_multi(s: str) -> list:
    return [x.strip() for x in re.split(r"[;\n]", s or "") if x.strip()]


def orcid_ok(s: str) -> bool:
    m = re.fullmatch(r"(?:https?://orcid\.org/)?(\d{4}-\d{4}-\d{4}-\d{3}[\dX])", s.strip())
    if not m:
        return False
    digits = m.group(1).replace("-", "")
    total = 0
    for ch in digits[:-1]:
        total = (total + int(ch)) * 2
    r = (12 - total % 11) % 11
    return digits[-1] == ("X" if r == 10 else str(r))


def short_hash(seed: str, n=4) -> str:
    x = int(hashlib.sha1(seed.encode("utf-8")).hexdigest(), 16)
    alphabet, s = "0123456789abcdefghijklmnopqrstuvwxyz", ""
    while len(s) < n:
        x, r = divmod(x, 36)
        s += alphabet[r]
    return s


class Issues:
    ORDER = {"ERROR": 0, "WARN": 1, "INFO": 2}

    def __init__(self):
        self.items = []

    def add(self, level, where, msg):
        self.items.append((level, where, msg))

    def count(self, level):
        return sum(1 for i in self.items if i[0] == level)

    def render(self) -> str:
        if not self.items:
            return "No issues found."
        lines = [f"{self.count('ERROR')} error(s), {self.count('WARN')} warning(s), {self.count('INFO')} note(s)", ""]
        for lvl, where, msg in sorted(self.items, key=lambda i: self.ORDER[i[0]]):
            lines.append(f"[{lvl:5}] {where}: {msg}")
        return "\n".join(lines)


# ============================================================ PageTab -> workbook data
def import_blocks(blocks) -> tuple:
    """Return (study: dict, tables: dict[sheet -> list[dict]], notes: list[str])."""
    notes = []
    if not blocks or blocks[0].kind != "Submission":
        raise ValueError("first block must be 'Submission'")
    sub = blocks[0]
    study_b = next((b for b in blocks if b.kind == "Study"), None)
    if study_b is None:
        raise ValueError("no 'Study' block found")
    tables = {k: [] for k in TABLES}
    lic = study_b.getall("License")
    study = {
        "Accession": sub.accno, "Template": sub.get("Template"), "DOI": sub.get("DOI"),
        "AttachTo": sub.get("AttachTo"),
        "Title": study_b.get("Title") or sub.get("Title"),
        "Release date": study_b.get("ReleaseDate") or sub.get("ReleaseDate"),
        "Description": study_b.get("Description"),
        "Keywords": "; ".join(a.value for a in study_b.attrs if a.name in ("Keywords", "Keyword")),
        "Acknowledgements": study_b.get("Acknowledgements"),
        "License": lic[0].value if lic else "",
        "License URL": next((qv for q, qn, qv in (lic[0].quals if lic else []) if qn == "URL"), ""),
        "Funding statement": study_b.get("Funding statement"),
        "_study_accno": study_b.accno,
    }
    known = {
        "Submission": {"Template", "Title", "DOI", "ReleaseDate", "AttachTo"},
        "Study": {"Title", "ReleaseDate", "Description", "Keywords", "Keyword", "Acknowledgements",
                  "License", "Funding statement"},
        "Link": {"Description"},
        "Publication": {"Pubmed ID", "Title", "Authors", "DOI", "Year"},
        "Funding": {"Agency", "grant_id"},
        "Biosample": {"Title", "Organism", "Description", "Biological entity", "Experimental variable",
                      "Intrinsic variable", "Extrinsic variable"},
        "Specimen": {"Title", "Sample preparation protocol"},
        "Image acquisition": {"Title", "Imaging instrument", "Image acquisition parameters", "Imaging method"},
        "Image analysis": {"Title", "Image analysis overview"},
        "Image correlation": {"Title", "Spatial and temporal alignment", "Fiducials used", "Transformation matrix"},
        "Study Component": {"Name", "Description", "File List"},
        "Associations": {"Biosample", "Specimen", "Image acquisition", "Image analysis", "Image correlation"},
        "Author": {"Name", "E-mail", "ORCID", "affiliation", "Role"},
        "organisation": {"Name", "RORID", "Address"},
    }
    known["organization"] = known["organisation"]

    protocols, specimen_info, sc_names = {}, {}, {}
    desc_labels = []
    kv_re = re.compile(r"^([^:\n]{1,80}):\s+(.*)$")

    for b in blocks:
        kn = known.get(b.kind)
        if kn is None:
            notes.append(f"UNMAPPED block type {b.kind!r} ({b.accno}) - not imported, add it by hand")
            continue
        for a in b.attrs:
            if a.name not in kn:
                notes.append(f"UNMAPPED attribute {a.name!r} in {b.kind} {b.accno or b.get('Title')!r} - not imported")
        if b.kind == "Link":
            tables["Links"].append({"URL": b.accno, "Description": b.get("Description")})
        elif b.kind == "Publication":
            tables["Publications"].append({"Title": b.get("Title"), "Authors": b.get("Authors"),
                                           "DOI": b.get("DOI"), "Year": b.get("Year"),
                                           "PubMed ID": b.get("Pubmed ID"), "_accno": b.accno})
        elif b.kind == "Funding":
            tables["Grants"].append({"Agency": b.get("Agency"), "Grant ID": b.get("grant_id"), "_accno": b.accno})
        elif b.kind == "Biosample":
            org = b.getall("Organism")
            term = next((qv for q, qn, qv in (org[0].quals if org else []) if qn == "TermId"), "")
            row = {"Title": b.get("Title"), "Organism": org[0].value if org else "",
                   "NCBITaxon ID": term.replace("NCBITaxon_", "").replace("NCBITaxon:", ""),
                   "Biological entity": b.get("Biological entity"),
                   "Experimental variable": b.get("Experimental variable"),
                   "Intrinsic variable": b.get("Intrinsic variable"),
                   "Extrinsic variable": b.get("Extrinsic variable"),
                   "Description": "", "_accno": b.accno}
            desc = b.get("Description")
            lines = [l for l in desc.split("\n") if l.strip()]
            parsed = [kv_re.match(l.strip()) for l in lines]
            if lines and all(parsed):
                for m in parsed:
                    lab = DESC_PREFIX + m.group(1).strip()
                    if lab not in desc_labels:
                        desc_labels.append(lab)
                    row[lab] = m.group(2).strip()
            else:
                row["Description"] = desc
            tables["Biosamples"].append(row)
        elif b.kind == "Specimen":
            text = b.get("Sample preparation protocol")
            if text not in protocols:
                protocols[text] = f"Protocol {len(protocols) + 1}"
            specimen_info[b.get("Title")] = (protocols[text], b.accno)
        elif b.kind == "Image acquisition":
            im = b.getall("Imaging method")
            q = {qn: qv for _, qn, qv in (im[0].quals if im else [])}
            tables["Acquisitions"].append({
                "Title": b.get("Title"), "Imaging instrument": b.get("Imaging instrument"),
                "Image acquisition parameters": b.get("Image acquisition parameters"),
                "Imaging method": im[0].value if im else "", "Imaging method ontology": q.get("Ontology", ""),
                "Imaging method term ID": q.get("TermId", ""), "_accno": b.accno})
        elif b.kind == "Image analysis":
            tables["Analyses"].append({"Title": b.get("Title"),
                                       "Image analysis overview": b.get("Image analysis overview"),
                                       "_accno": b.accno})
        elif b.kind == "Image correlation":
            tables["Correlations"].append({"Title": b.get("Title"),
                                           "Spatial and temporal alignment": b.get("Spatial and temporal alignment"),
                                           "Fiducials used": b.get("Fiducials used"),
                                           "Transformation matrix": b.get("Transformation matrix"),
                                           "_accno": b.accno})
        elif b.kind == "Study Component":
            sc_names[b.accno] = b.get("Name")
            tables["StudyComponents"].append({"Name": b.get("Name"), "Description": b.get("Description"),
                                              "File List": b.get("File List"), "_accno": b.accno})
        elif b.kind == "Associations":
            spec = b.get("Specimen")
            pkey, saccno = specimen_info.get(spec, ("", ""))
            tables["Specimens"].append({
                "Specimen title": spec, "Biosample": b.get("Biosample"), "Protocol": pkey,
                "Image acquisition": b.get("Image acquisition"), "Image analysis": b.get("Image analysis"),
                "Image correlation": b.get("Image correlation"),
                "Study component": sc_names.get(b.parent, b.parent),
                "_spec_accno": saccno, "_asso_accno": b.accno})
        elif b.kind == "Author":
            tables["Authors"].append({
                "Name": b.get("Name"), "E-mail": b.get("E-mail"), "ORCID": b.get("ORCID"),
                "Roles": "; ".join(a.value for a in b.getall("Role")),
                "Affiliations": "; ".join(a.value for a in b.attrs if a.ref and a.name == "affiliation"),
                "_accno": b.accno})
        elif b.kind in ("organisation", "organization"):
            tables["Organisations"].append({"Key": b.accno, "Name": b.get("Name"),
                                            "ROR ID": b.get("RORID"), "Address": b.get("Address")})

    associated = {r["Specimen title"] for r in tables["Specimens"]}
    for title, (pkey, saccno) in specimen_info.items():
        if title not in associated:
            notes.append(f"Specimen {title!r} has no association - added with blank study component")
            tables["Specimens"].append({"Specimen title": title, "Protocol": pkey, "_spec_accno": saccno})
    tables["Protocols"] = [{"Key": k, "Sample preparation protocol": t} for t, k in protocols.items()]
    return study, tables, desc_labels, notes


# ============================================================ workbook I/O
def _require_openpyxl():
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        sys.exit("openpyxl is required for workbook commands:  pip install openpyxl")


def _desc_col(label: str) -> Col:
    return Col(label, False, 16,
               "Descriptor column: each value is added to the sample description as a line "
               "'%s: <value>'." % label[len(DESC_PREFIX):])


def write_workbook(path, study=None, tables=None, desc_labels=None):
    """Write a workbook. Row 1 = headers, row 2 = column explanations (guide row), data from row 3."""
    _require_openpyxl()
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    study, tables, desc_labels = study or {}, tables or {}, desc_labels or []
    REQ = PatternFill("solid", fgColor="F8CBAD")
    OPT = PatternFill("solid", fgColor="DDEBF7")
    INT = PatternFill("solid", fgColor="D9D9D9")
    GUIDE = PatternFill("solid", fgColor="F2F2F2")
    GUIDE_FONT = Font(italic=True, color="595959", size=9)
    FIRST, NROWS = 3, 300

    wb = Workbook()
    ws = wb.active
    ws.title = "README"
    for i, (text, style) in enumerate(README_LINES, 1):
        c = ws.cell(row=i, column=1, value=text)
        if style == "title":
            c.font = Font(bold=True, size=14)
        elif style == "h":
            c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 120

    # --- Study (key/value)
    ws = wb.create_sheet("Study")
    for j, (h, w) in enumerate([("Field", 22), ("Value", 90), ("Explanation", 70)], 1):
        c = ws.cell(row=1, column=j, value=h)
        c.font = Font(bold=True)
        ws.column_dimensions[c.column_letter].width = w
    for i, (name, req, help_) in enumerate(STUDY_FIELDS, 2):
        k = ws.cell(row=i, column=1, value=name)
        k.font = Font(bold=True)
        k.fill = INT if name.startswith("_") else (REQ if req else OPT)
        k.alignment = Alignment(vertical="top")
        v = ws.cell(row=i, column=2, value=study.get(name, "") or None)
        v.number_format = "@"
        v.alignment = Alignment(wrap_text=True, vertical="top")
        e = ws.cell(row=i, column=3, value=help_)
        e.alignment = Alignment(wrap_text=True, vertical="top")
        e.font = GUIDE_FONT
        if name == "License":
            dv = DataValidation(type="list", formula1='"%s"' % ",".join(LICENSES), allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(v.coordinate)
    ws.freeze_panes = "B2"

    # --- tables
    for sheet, cols in TABLES.items():
        ws = wb.create_sheet(sheet)
        cols = list(cols)
        if sheet == "Biosamples":
            at = next(i for i, c in enumerate(cols) if c.header == "_accno")
            cols[at:at] = [_desc_col(lab) for lab in desc_labels]
        for j, col in enumerate(cols, 1):
            c = ws.cell(row=1, column=j, value=col.header)
            c.font = Font(bold=True)
            c.fill = INT if col.header.startswith("_") else (REQ if col.required else OPT)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if col.help:
                c.comment = Comment(col.help, "bia_kit")
            g = ws.cell(row=2, column=j, value=GUIDE_MARK + (col.help or ("required" if col.required else "optional")))
            g.font, g.fill = GUIDE_FONT, GUIDE
            g.alignment = Alignment(wrap_text=True, vertical="top")
            ws.column_dimensions[c.column_letter].width = col.width
            if col.ref:
                # OOXML stores list formulas WITHOUT a leading '=' (as Excel itself writes them)
                dv = DataValidation(type="list", formula1=f"{col.ref}!$A${FIRST}:$A${NROWS}", allow_blank=True,
                                    showErrorMessage=True, errorTitle="Unknown value",
                                    error=f"Choose a value defined on the {col.ref} sheet.")
                ws.add_data_validation(dv)
                dv.add(f"{c.column_letter}{FIRST}:{c.column_letter}{NROWS}")
        ws.row_dimensions[2].height = 75
        rows = tables.get(sheet, [])
        for i in range(FIRST, max(NROWS, len(rows) + FIRST) + 1):
            r = rows[i - FIRST] if i - FIRST < len(rows) else {}
            for j, col in enumerate(cols, 1):
                val = r.get(col.header, "")
                c = ws.cell(row=i, column=j, value=val if val != "" else None)
                c.number_format = "@"
                if col.wrap:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = "B3"
    wb.save(path)


def _is_guide_row(values) -> bool:
    vals = [cellstr(v) for v in values if cellstr(v)]
    return bool(vals) and all(v.startswith(GUIDE_MARK.strip()) for v in vals)


def read_workbook(path) -> tuple:
    _require_openpyxl()
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True)
    if "Study" not in wb.sheetnames:
        raise ValueError(f"{path}: no 'Study' sheet - is this a bia_kit workbook?")
    study = {}
    for row in wb["Study"].iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            study[cellstr(row[0])] = cellstr(row[1] if len(row) > 1 else None)
    tables = {}
    for sheet in TABLES:
        if sheet not in wb.sheetnames:
            tables[sheet] = []
            continue
        ws = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            tables[sheet] = []
            continue
        headers = [cellstr(h) for h in rows[0]]
        out = []
        for n, r in enumerate(rows[1:], 2):
            if _is_guide_row(r):          # explanation row, not data (older workbooks have none)
                continue
            d = {h: cellstr(v) for h, v in zip(headers, r) if h}
            if any(d.values()):
                d["__row"] = n
                out.append(d)
        tables[sheet] = out
    return study, tables


# ============================================================ workbook -> PageTab
FILE_PATH_SEGMENT = re.compile(r"^[A-Za-z0-9!\-_.*'() ]+$")


def check_file_list(path: Path, issues: Issues, where: str):
    """BioImage Archive file-list rules (help-file-list page)."""
    with open(path, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh, delimiter="\t"))
    if not rows or not rows[0] or rows[0][0] != "Files":
        issues.add("ERROR", where, f"{path.name}: first header column must be 'Files'")
        return
    seen, bad = set(), []
    for n, r in enumerate(rows[1:], 2):
        if not r or not any(c.strip() for c in r):
            issues.add("ERROR", where, f"{path.name}: blank line at row {n}")
            continue
        p = r[0]
        if p in seen:
            issues.add("ERROR", where, f"{path.name}: duplicate path at row {n}: {p}")
        seen.add(p)
        if "\\" in p or p.endswith(" ") or p.endswith("/") or not all(
                FILE_PATH_SEGMENT.match(s) for s in p.split("/")):
            bad.append(f"row {n}: {p}")
    if bad:
        issues.add("ERROR", where, f"{path.name}: {len(bad)} path(s) with disallowed characters; first ones: {bad[:3]}")
    issues.add("INFO", where, f"{path.name}: {len(rows) - 1} file rows checked")


def build_blocks(study, tables, keep_empty=False, filelist_dir=None) -> tuple:
    I = Issues()
    T = {k: tables.get(k, []) for k in TABLES}

    def W(sheet, row=None):
        return f"{sheet}" + (f" row {row['__row']}" if row and "__row" in row else "")

    # ---------- accno bookkeeping (detect copy-pasted duplicates)
    taken = set()
    def accno(row, key, prefix, n, seed):
        a = row.get(key, "")
        if a and a not in taken:
            taken.add(a)
            return a
        if a:
            I.add("INFO", W("?", row), f"duplicate internal id {a!r} (copied row?) - regenerated")
        a = f"{prefix}-{n}-{short_hash(prefix + seed)}"
        while a in taken:
            a += "x"
        taken.add(a)
        return a

    # ---------- Study
    S = study
    for name, req, _ in STUDY_FIELDS:
        if req and not S.get(name):
            I.add("ERROR", "Study", f"'{name}' is required")
    for name in ("Title", "Description"):
        if S.get(name) and len(S[name]) < 25:
            I.add("ERROR", "Study", f"'{name}' must be >= 25 characters (REMBI)")
    rd = S.get("Release date", "")
    if rd:
        try:
            if date.fromisoformat(rd) <= date.today():
                I.add("INFO", "Study", f"Release date {rd} is not in the future - data will be public on release")
        except ValueError:
            I.add("ERROR", "Study", f"Release date {rd!r} is not YYYY-MM-DD")
    keywords = split_multi(S.get("Keywords", ""))
    lic = S.get("License", "")
    lic_url = S.get("License URL") or LICENSES.get(lic, "")
    if lic and lic not in LICENSES:
        I.add("WARN", "Study", f"License {lic!r} is not one of {list(LICENSES)}")
    study_acc = S.get("_study_accno") or f"study-0-{short_hash(S.get('Title', ''))}"
    taken.add(study_acc)

    blocks = []
    sub = Block("Submission", S.get("Accession", ""))
    sub.add("Template", S.get("Template", "BioImages.v5"))
    sub.add("Title", S.get("Title", ""))
    if S.get("DOI"):
        sub.add("DOI", S["DOI"])
    sub.add("ReleaseDate", rd)
    sub.add("AttachTo", S.get("AttachTo", "BioImages"))
    blocks.append(sub)

    st = Block("Study", study_acc)
    st.add("Title", S.get("Title", ""))
    st.add("ReleaseDate", rd)
    st.add("Description", S.get("Description", ""))
    for k in keywords:
        st.add("Keywords", k)
    if S.get("Acknowledgements"):
        st.add("Acknowledgements", S["Acknowledgements"])
    if lic:
        st.add("License", lic, [("[", "URL", lic_url)] if lic_url else [])
    if S.get("Funding statement"):
        st.add("Funding statement", S["Funding statement"])
    blocks.append(st)

    # ---------- Links
    for r in T["Links"]:
        if not r.get("URL"):
            I.add("ERROR", W("Links", r), "URL is required")
            continue
        if not re.match(r"^(https?|ftp)://", r["URL"]):
            I.add("WARN", W("Links", r), f"URL does not start with http(s):// : {r['URL']}")
        lb = Block("Link", r["URL"])
        lb.add("Description", r.get("Description", ""))
        blocks.append(lb)

    # ---------- Publications / Grants
    for n, r in enumerate(T["Publications"], 1):
        filled = any(r.get(k) for k in ("Title", "Authors", "DOI", "Year", "PubMed ID"))
        if not filled:
            I.add("WARN", W("Publications", r), "empty publication " + ("emitted (--keep-empty)" if keep_empty else "skipped"))
            if not keep_empty:
                continue
        else:
            if not r.get("Title"):
                I.add("ERROR", W("Publications", r), "Title is required")
            if r.get("DOI") and not re.match(r"^10\.\d{4,9}/\S+$", r["DOI"]):
                I.add("WARN", W("Publications", r), f"DOI should look like 10.xxxx/yyyy (no https://doi.org/): {r['DOI']}")
            if r.get("Year") and not re.fullmatch(r"\d{4}", r["Year"]):
                I.add("WARN", W("Publications", r), f"Year should be 4 digits: {r['Year']}")
        pb = Block("Publication", accno(r, "_accno", "publ", n, r.get("Title", "") + str(n)), study_acc)
        for name, col in (("Pubmed ID", "PubMed ID"), ("Title", "Title"), ("Authors", "Authors"),
                          ("DOI", "DOI"), ("Year", "Year")):
            pb.add(name, r.get(col, ""))
        blocks.append(pb)
    if not any(any(r.get(k) for k in ("Title", "DOI")) for r in T["Publications"]):
        I.add("INFO", "Publications", "no publication given (fine before the paper is out; add it before release)")

    grants_filled = [r for r in T["Grants"] if r.get("Agency") or r.get("Grant ID")]
    for n, r in enumerate(T["Grants"], 1):
        if r not in grants_filled:
            I.add("WARN", W("Grants", r), "empty grant row " + ("emitted (--keep-empty)" if keep_empty else "skipped"))
            if not keep_empty:
                continue
        elif not (r.get("Agency") and r.get("Grant ID")):
            I.add("WARN", W("Grants", r), "both Agency and Grant ID are expected")
        fb = Block("Funding", accno(r, "_accno", "fund", n, r.get("Grant ID", "") + str(n)), study_acc)
        fb.add("Agency", r.get("Agency", ""))
        fb.add("grant_id", r.get("Grant ID", ""))
        blocks.append(fb)
    if S.get("Funding statement") and not grants_filled:
        ids = re.findall(r"\b\d{5,}(?:/[A-Z0-9]+)*\b", S["Funding statement"])
        I.add("WARN", "Grants", "Funding statement is filled but the Grants sheet is empty"
              + (f"; candidate grant ids in the statement: {', '.join(dict.fromkeys(ids))}" if ids else ""))

    # ---------- reference tables
    def index(sheet, keycol):
        idx = {}
        for r in T[sheet]:
            k = r.get(keycol, "")
            if not k:
                I.add("ERROR", W(sheet, r), f"'{keycol}' is required")
                continue
            if k in idx:
                I.add("ERROR", W(sheet, r), f"duplicate {keycol} {k!r}")
                continue
            idx[k] = r
        return idx

    bios = index("Biosamples", "Title")
    prots = index("Protocols", "Key")
    acqs = index("Acquisitions", "Title")
    anas = index("Analyses", "Title")
    cors = index("Correlations", "Title")
    scs = index("StudyComponents", "Name")
    orgs = index("Organisations", "Key")
    used = {k: set() for k in ("Biosamples", "Protocols", "Acquisitions", "Analyses", "Correlations",
                                "StudyComponents", "Organisations")}

    for sheet, idx in (("Biosamples", bios), ("Protocols", prots), ("Acquisitions", acqs),
                       ("Analyses", anas), ("StudyComponents", scs)):
        for k, r in idx.items():
            for col in TABLES[sheet]:
                if col.required and not r.get(col.header):
                    I.add("ERROR", W(sheet, r), f"'{col.header}' is required ({k})")

    # ---------- Specimens (sample sheet) -> specimen + association plans
    def default(sheet, idx, value, required, row, col):
        if value:
            if value not in idx:
                I.add("ERROR", W("Specimens", row), f"{col} {value!r} is not defined on the {sheet} sheet")
            return value
        if required:
            if len(idx) == 1:
                return next(iter(idx))
            I.add("ERROR", W("Specimens", row), f"{col} is required" + (f" (choose one of {len(idx)})" if idx else f" - define it on the {sheet} sheet"))
        return ""

    spec_order, spec_def, assoc_by_sc, combos = [], {}, {}, set()
    for r in T["Specimens"]:
        title = r.get("Specimen title", "")
        if not title:
            I.add("ERROR", W("Specimens", r), "Specimen title is required")
            continue
        b = default("Biosamples", bios, r.get("Biosample", ""), True, r, "Biosample")
        p = default("Protocols", prots, r.get("Protocol", ""), True, r, "Protocol")
        a = default("Acquisitions", acqs, r.get("Image acquisition", ""), True, r, "Image acquisition")
        an = default("Analyses", anas, r.get("Image analysis", ""), False, r, "Image analysis")
        co = default("Correlations", cors, r.get("Image correlation", ""), False, r, "Image correlation")
        sc = default("StudyComponents", scs, r.get("Study component", ""), True, r, "Study component")
        for k, v in (("Biosamples", b), ("Protocols", p), ("Acquisitions", a), ("Analyses", an),
                     ("Correlations", co), ("StudyComponents", sc)):
            if v:
                used[k].add(v)
        if title in spec_def:
            if spec_def[title]["protocol"] != p:
                I.add("ERROR", W("Specimens", r), f"specimen {title!r} appears with different protocols")
        else:
            spec_order.append(title)
            spec_def[title] = {"protocol": p, "row": r}
        combo = (sc, b, title, a, an, co)
        if combo in combos:
            I.add("WARN", W("Specimens", r), f"duplicate row for specimen {title!r}")
        combos.add(combo)
        assoc_by_sc.setdefault(sc, []).append({"b": b, "s": title, "a": a, "an": an, "co": co, "row": r})
        # sanity: specimen title starting with ANOTHER biosample's title
        for bt in sorted(bios, key=len, reverse=True):
            if bt != b and re.match(re.escape(bt) + r"[-_ .]", title) and not re.match(re.escape(b) + r"[-_ .]", title):
                I.add("WARN", W("Specimens", r), f"specimen {title!r} looks like it belongs to biosample {bt!r}, but row says {b!r}")
                break
    # capitalisation consistency: titles that share a pattern but differ only in letter case
    shapes = {}
    for t in spec_order:
        shp = re.sub(r"\d+", "#", t)
        shapes.setdefault(shp.lower(), {}).setdefault(shp, []).append(t)
    for variants in shapes.values():
        if len(variants) > 1:
            ex = "; ".join(f"{v[0]!r} ({len(v)}x)" for v in variants.values())
            I.add("WARN", "Specimens", f"inconsistent capitalisation in specimen titles: {ex}")

    # ---------- Biosamples
    human_cols = set()
    for n, (t, r) in enumerate(bios.items(), 1):
        org = r.get("Organism", "")
        tax = re.sub(r"^NCBITaxon[_:]", "", r.get("NCBITaxon ID", ""))
        if not tax:
            sci = org.split("(")[0].strip().lower()
            tax = TAXA.get(sci, "")
            if tax:
                I.add("INFO", W("Biosamples", r), f"NCBITaxon ID {tax} filled from organism name {org!r}")
            else:
                I.add("WARN", W("Biosamples", r), f"no NCBITaxon ID for {org!r} - add it (number only)")
        elif not tax.isdigit():
            I.add("ERROR", W("Biosamples", r), f"NCBITaxon ID must be a number: {r.get('NCBITaxon ID')!r}")
        desc_parts = [r["Description"]] if r.get("Description") else []
        desc_cols = [h for h in r if h.startswith(DESC_PREFIX) and r[h]]
        desc_parts += [f"{h[len(DESC_PREFIX):]}: {r[h]}" for h in desc_cols]
        if tax == "9606":
            human_cols.update(h[len(DESC_PREFIX):] for h in desc_cols)
        bb = Block("Biosample", accno(r, "_accno", "bios", n, t), study_acc)
        bb.add("Title", t)
        bb.add("Organism", org, [("[", "Ontology", "NCBITaxon"), ("[", "TermId", f"NCBITaxon_{tax}")] if tax else [])
        bb.add("Description", "\n".join(desc_parts))
        bb.add("Biological entity", r.get("Biological entity", ""))
        for name in ("Experimental variable", "Intrinsic variable", "Extrinsic variable"):
            if r.get(name):
                bb.add(name, r[name])
        blocks.append(bb)
    if human_cols:
        I.add("WARN", "Biosamples",
              "human donors: these per-donor fields will be PUBLIC on release: " + ", ".join(sorted(human_cols))
              + ". Confirm consent / biobank agreement allow it and that the combination cannot identify a donor.")

    # ---------- Specimens
    for n, t in enumerate(spec_order, 1):
        d = spec_def[t]
        sb = Block("Specimen", accno(d["row"], "_spec_accno", "spec", n, t), study_acc)
        sb.add("Title", t)
        sb.add("Sample preparation protocol", prots.get(d["protocol"], {}).get("Sample preparation protocol", ""))
        blocks.append(sb)

    # ---------- Acquisitions / Correlations / Analyses
    for n, (t, r) in enumerate(acqs.items(), 1):
        ab = Block("Image acquisition", accno(r, "_accno", "imag", n, "acq" + t), study_acc)
        ab.add("Title", t)
        ab.add("Imaging instrument", r.get("Imaging instrument", ""))
        ab.add("Image acquisition parameters", r.get("Image acquisition parameters", ""))
        term = r.get("Imaging method term ID", "")
        if term and not re.fullmatch(r"FBbi_\d{8}", term):
            I.add("WARN", W("Acquisitions", r), f"Imaging method term ID should be 'FBbi_' followed by eight digits: {term!r}")
        if not term:
            I.add("WARN", W("Acquisitions", r), "no Imaging method term ID (FBbi) - recommended")
        ont = r.get("Imaging method ontology") or ("fbbi" if term.startswith("FBbi") else "")
        quals = ([("[", "Ontology", ont)] if ont else []) + ([("[", "TermId", term)] if term else [])
        ab.add("Imaging method", r.get("Imaging method", ""), quals)
        blocks.append(ab)
    for n, (t, r) in enumerate(cors.items(), 1):
        cb = Block("Image correlation", accno(r, "_accno", "imag", n, "cor" + t), study_acc)
        cb.add("Title", t)
        for name in ("Spatial and temporal alignment", "Fiducials used", "Transformation matrix"):
            if r.get(name):
                cb.add(name, r[name])
        blocks.append(cb)
    for n, (t, r) in enumerate(anas.items(), 1):
        nb = Block("Image analysis", accno(r, "_accno", "imag", n, "ana" + t), study_acc)
        nb.add("Title", t)
        nb.add("Image analysis overview", r.get("Image analysis overview", ""))
        blocks.append(nb)

    # ---------- Study components + associations
    fl_dir = Path(filelist_dir) if filelist_dir else None
    for n, (name, r) in enumerate(scs.items(), 1):
        cb = Block("Study Component", accno(r, "_accno", "stud", n, name), study_acc)
        cb.add("Name", name)
        cb.add("Description", r.get("Description", ""))
        cb.add("File List", r.get("File List", ""))
        blocks.append(cb)
        fl = r.get("File List", "")
        if fl and not re.search(r"\.(tsv|xlsx|json)$", fl, re.I):
            I.add("WARN", W("StudyComponents", r), f"File List {fl!r} should be .tsv / .xlsx / .json")
        if fl_dir and fl:
            p = fl_dir / fl
            if p.exists() and p.suffix.lower() == ".tsv":
                check_file_list(p, I, W("StudyComponents", r))
            elif not p.exists():
                I.add("WARN", W("StudyComponents", r), f"file list not found: {p}")
        assocs = assoc_by_sc.get(name, [])
        if not assocs:
            I.add("WARN", W("StudyComponents", r), f"study component {name!r} has no specimens")
        for x in assocs:
            ab = Block("Associations", x["row"].get("_asso_accno", ""), cb.accno)
            ab.add("Biosample", x["b"])
            ab.add("Specimen", x["s"])
            ab.add("Image acquisition", x["a"])
            ab.add("Image analysis", x["an"])
            ab.add("Image correlation", x["co"])
            blocks.append(ab)

    # ---------- Authors / organisations
    if not T["Authors"]:
        I.add("ERROR", "Authors", "at least one author is required")
    roles_all = []
    for n, r in enumerate(T["Authors"], 1):
        if not r.get("Name"):
            I.add("ERROR", W("Authors", r), "Name is required")
        roles = split_multi(r.get("Roles", ""))
        roles_all += roles
        affs = split_multi(r.get("Affiliations", ""))
        if not affs:
            I.add("ERROR", W("Authors", r), f"{r.get('Name')}: at least one affiliation (organisation Key) is required")
        for k in affs:
            if k not in orgs:
                I.add("ERROR", W("Authors", r), f"{r.get('Name')}: affiliation {k!r} is not a Key on the Organisations sheet")
            used["Organisations"].add(k)
        if r.get("ORCID") and not orcid_ok(r["ORCID"]):
            I.add("ERROR", W("Authors", r), f"{r.get('Name')}: ORCID {r['ORCID']!r} is malformed or fails the checksum")
        if r.get("E-mail") and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", r["E-mail"]):
            I.add("WARN", W("Authors", r), f"{r.get('Name')}: e-mail looks malformed: {r['E-mail']!r}")
        if any("corresponding" in x.lower() or "submitter" in x.lower() for x in roles) and not r.get("E-mail"):
            I.add("WARN", W("Authors", r), f"{r.get('Name')}: submitter/corresponding author without e-mail")
        ab = Block("Author", r.get("_accno", ""), study_acc)
        ab.add("Name", r.get("Name", ""))
        if r.get("E-mail"):
            ab.add("E-mail", r["E-mail"])
        ab.add("ORCID", r.get("ORCID", ""))
        for k in affs:
            ab.add("affiliation", k, ref=True)
        for x in roles:
            ab.add("Role", x)
        blocks.append(ab)
    if T["Authors"] and not any("submitter" in x.lower() for x in roles_all):
        I.add("WARN", "Authors", "nobody has the role 'submitter'")
    for k, r in orgs.items():
        if not re.fullmatch(r"[A-Za-z0-9_\-]+", k):
            I.add("ERROR", W("Organisations", r), f"Key {k!r}: use letters, digits, '-' or '_' only")
        if not r.get("Name"):
            I.add("ERROR", W("Organisations", r), "Name is required")
        ror = r.get("ROR ID", "")
        if ror and not re.fullmatch(r"https://ror\.org/0[a-z0-9]{6}\d{2}", ror):
            I.add("WARN", W("Organisations", r), f"ROR ID should be the full link https://ror.org/ followed by the 9-character identifier: {ror!r}")
        if not ror:
            I.add("INFO", W("Organisations", r), f"{r.get('Name')}: no ROR ID (recommended)")
        ob = Block("organisation", k, study_acc)
        ob.add("Name", r.get("Name", ""))
        if ror:
            ob.add("RORID", ror)
        if r.get("Address"):
            ob.add("Address", r["Address"])
        blocks.append(ob)

    # ---------- unused definitions
    for sheet, idx in (("Biosamples", bios), ("Protocols", prots), ("Acquisitions", acqs), ("Analyses", anas),
                       ("Correlations", cors), ("Organisations", orgs)):
        for k, r in idx.items():
            if k not in used[sheet]:
                I.add("WARN", W(sheet, r), f"{k!r} is defined but never used"
                      + (" (not referenced by any Specimens row)" if sheet != "Organisations" else ""))
    return blocks, I


# ============================================================ Xenium helper
XENIUM_FIELDS = ["run_name", "region_name", "slide_id", "preservation_method", "panel_name", "panel_organism",
                 "chemistry_version", "analysis_sw_version", "instrument_sw_version", "pixel_size", "num_cells"]


def _version(s: str) -> str:
    m = re.search(r"(\d+(?:\.\d+)+)", s or "")
    return m.group(1) if m else ""


def xenium_scan(roots, out_dir: Path, relative_to=None, workbook=None):
    out_dir.mkdir(parents=True, exist_ok=True)
    bundles = []
    for root in roots:
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames.sort()
            if "experiment.xenium" in filenames:
                with open(os.path.join(dirpath, "experiment.xenium"), encoding="utf-8") as fh:
                    meta = json.load(fh)
                bundles.append((Path(dirpath), meta))
                dirnames[:] = []  # bundles do not nest
    if not bundles:
        print("no experiment.xenium found under", ", ".join(map(str, roots)))
        return 1
    base = Path(relative_to) if relative_to else Path(os.path.commonpath([str(b[0].parent) for b in bundles]))

    # optional workbook cross-check
    spec_rows, ana_versions = {}, {}
    if workbook:
        _, tables = read_workbook(workbook)
        for r in tables.get("Specimens", []):
            spec_rows.setdefault(r.get("Specimen title", ""), r)
        for r in tables.get("Analyses", []):
            v = _version(r.get("Title", "")) or _version(r.get("Image analysis overview", ""))
            if v:
                ana_versions.setdefault(v, r.get("Title", ""))
    spec_lower = {k.lower(): k for k in spec_rows}

    summary, problems, files_rows, suggestions = [], [], [], []
    for bdir, meta in bundles:
        rel_bundle = bdir.relative_to(base).as_posix() if bdir.is_relative_to(base) else bdir.as_posix()
        row = {"bundle": rel_bundle, **{k: str(meta.get(k, "")) for k in XENIUM_FIELDS}}
        version = _version(row["analysis_sw_version"])
        region = row["region_name"]
        spec = spec_rows.get(region) and region or spec_lower.get(region.lower(), "")
        row["specimen_in_workbook"] = spec
        summary.append(row)
        if workbook:
            if not spec:
                problems.append(f"bundle {rel_bundle}: region_name {region!r} matches no Specimens row")
            else:
                r = spec_rows[spec]
                if spec != region:
                    problems.append(f"{spec!r}: workbook title differs in case from region_name {region!r}")
                wb_ana = r.get("Image analysis", "")
                if version and _version(wb_ana) != version:
                    problems.append(f"{spec!r}: workbook Image analysis = {wb_ana!r} but experiment.xenium "
                                    f"analysis_sw_version = {row['analysis_sw_version']!r}")
        suggestions.append({"Specimen title": spec or region,
                            "Biosample": spec_rows.get(spec, {}).get("Biosample", "") if spec else "",
                            "Image analysis": ana_versions.get(version, f"Xenium {version}" if version else "")})
        for dirpath, dirnames, filenames in os.walk(bdir):
            dirnames.sort()
            for fn in sorted(filenames):
                p = Path(dirpath, fn)
                rel = p.relative_to(base).as_posix() if p.is_relative_to(base) else p.as_posix()
                files_rows.append({"Files": rel, "region_name": region, "slide_id": row["slide_id"],
                                   "analysis_sw_version": row["analysis_sw_version"],
                                   "biosample": spec_rows.get(spec, {}).get("Biosample", "") if spec else "",
                                   "specimen": spec})
    if workbook:
        found = {s["specimen_in_workbook"] for s in summary}
        for t in spec_rows:
            if t and t not in found:
                problems.append(f"Specimens row {t!r}: no Xenium bundle with this region_name under the scanned roots")

    def write_tsv(path, rows, cols):
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=cols, delimiter="\t", extrasaction="ignore", lineterminator="\n")
            w.writeheader()
            w.writerows(rows)

    write_tsv(out_dir / "xenium_bundles.tsv", summary, ["bundle"] + XENIUM_FIELDS + ["specimen_in_workbook"])
    fl_cols = ["Files", "biosample", "specimen", "region_name", "slide_id", "analysis_sw_version"]
    write_tsv(out_dir / "file_list.tsv", files_rows, fl_cols)
    write_tsv(out_dir / "specimens_suggested.tsv", suggestions, ["Specimen title", "Biosample", "Image analysis"])
    fl_issues = Issues()
    check_file_list(out_dir / "file_list.tsv", fl_issues, "file_list.tsv")
    print(f"{len(bundles)} Xenium bundle(s), {len(files_rows)} files; paths relative to {base}")
    print(f"wrote {out_dir/'xenium_bundles.tsv'}, {out_dir/'file_list.tsv'}, {out_dir/'specimens_suggested.tsv'}")
    print(fl_issues.render())
    if workbook:
        print(f"\nWorkbook cross-check: {len(problems)} problem(s)")
        for p in problems:
            print("  -", p)
    return 1 if problems or fl_issues.count("ERROR") else 0


# ============================================================ JSON API (web form, via Pyodide)
# The browser form holds {"study": {...}, "tables": {sheet: [row, ...]}, "desc_labels": [...]}
# - exactly what the workbook holds - and calls these functions, so validation and PageTab
# generation have ONE implementation for CLI, Excel and web.
def api_schema() -> str:
    return json.dumps({
        "version": __version__,
        "study": [{"name": n, "required": r, "help": h} for n, r, h in STUDY_FIELDS],
        "tables": {k: [vars(c) for c in v] for k, v in TABLES.items()},
        "licenses": list(LICENSES),
        "desc_prefix": DESC_PREFIX,
        "desc_help": _desc_col(DESC_PREFIX + "<label>").help,
        "taxa": TAXA,
        "template_defaults": TEMPLATE_DEFAULTS,
    })


def _payload_tables(d: dict) -> dict:
    tables = {}
    for sheet, rows in (d.get("tables") or {}).items():
        out = []
        for i, r in enumerate(rows or [], 1):
            r = {k: cellstr(v) for k, v in r.items() if not k.startswith("__")}
            if any(r.values()):
                r["__row"] = i            # row number as shown in the web grid
                out.append(r)
        tables[sheet] = out
    return tables


def api_build(payload: str) -> str:
    d = json.loads(payload)
    study = {k: cellstr(v) for k, v in (d.get("study") or {}).items()}
    blocks, issues = build_blocks(study, _payload_tables(d), keep_empty=bool(d.get("keep_empty")))
    return json.dumps({
        "tsv": render_pagetab(blocks),
        "issues": [list(i) for i in issues.items],
        "errors": issues.count("ERROR"), "warnings": issues.count("WARN"), "notes": issues.count("INFO"),
        "blocks": len(blocks),
        "specimens": sum(1 for b in blocks if b.kind == "Specimen"),
        "associations": sum(1 for b in blocks if b.kind == "Associations"),
    })


def _strip_rows(tables: dict) -> dict:
    return {k: [{c: v for c, v in r.items() if not c.startswith("__")} for r in rows] for k, rows in tables.items()}


def api_import_tsv(text: str) -> str:
    study, tables, labels, notes = import_blocks(parse_pagetab_text(text))
    return json.dumps({"study": study, "tables": tables, "desc_labels": labels, "notes": notes})


def api_import_xlsx(path: str) -> str:
    study, tables = read_workbook(path)
    labels = []
    for r in tables.get("Biosamples", []):
        labels += [k for k in r if k.startswith(DESC_PREFIX) and k not in labels]
    return json.dumps({"study": study, "tables": _strip_rows(tables), "desc_labels": labels, "notes": []})


def api_export_xlsx(payload: str, path: str) -> str:
    d = json.loads(payload)
    write_workbook(path, d.get("study") or {}, _strip_rows(_payload_tables(d)), d.get("desc_labels") or [])
    return path


# ============================================================ CLI
def main(argv=None):
    ap = argparse.ArgumentParser(prog="bia_kit", description=__doc__.split("\n\n")[0])
    ap.add_argument("--version", action="version", version=__version__)
    sp = ap.add_subparsers(dest="cmd", required=True)
    p = sp.add_parser("template", help="write a blank workbook with explanations")
    p.add_argument("out")
    p = sp.add_parser("import-tsv", help="pre-fill a workbook from a PageTab TSV")
    p.add_argument("tsv")
    p.add_argument("out")
    for name in ("check", "build"):
        p = sp.add_parser(name, help=("validate" if name == "check" else "validate and write PageTab TSV"))
        p.add_argument("workbook")
        if name == "build":
            p.add_argument("out")
            p.add_argument("--force", action="store_true", help="write the TSV even if there are errors")
        p.add_argument("--keep-empty", action="store_true", help="emit empty Publication/Funding sections")
        p.add_argument("--filelist-dir", help="directory holding the file lists, to validate them too")
        p.add_argument("--report", help="also write the validation report to this file")
    p = sp.add_parser("diff", help="semantic diff of two PageTab TSVs")
    p.add_argument("a")
    p.add_argument("b")
    p = sp.add_parser("xenium-scan", help="read experiment.xenium files; write bundle table + file list")
    p.add_argument("roots", nargs="+")
    p.add_argument("--out", required=True)
    p.add_argument("--relative-to", help="make file-list paths relative to this directory (= your upload root)")
    p.add_argument("--workbook", help="cross-check region names / analysis versions against this workbook")
    a = ap.parse_args(argv)

    if a.cmd == "template":
        write_workbook(a.out, dict(TEMPLATE_DEFAULTS))
        print("wrote", a.out)
        return 0
    if a.cmd == "import-tsv":
        study, tables, desc_labels, notes = import_blocks(parse_pagetab(a.tsv))
        write_workbook(a.out, study, tables, desc_labels)
        print(f"wrote {a.out}: " + ", ".join(f"{k}={len(v)}" for k, v in tables.items()))
        for n in notes:
            print("  NOTE:", n)
        return 0
    if a.cmd in ("check", "build"):
        study, tables = read_workbook(a.workbook)
        blocks, issues = build_blocks(study, tables, keep_empty=a.keep_empty, filelist_dir=a.filelist_dir)
        report = issues.render()
        print(report)
        if a.report:
            Path(a.report).write_text(report + "\n", encoding="utf-8")
        if a.cmd == "build":
            if issues.count("ERROR") and not a.force:
                print("\nNOT written: fix the errors above (or use --force).")
                return 1
            Path(a.out).write_text(render_pagetab(blocks), encoding="utf-8")
            n_spec = sum(1 for b in blocks if b.kind == "Specimen")
            n_asso = sum(1 for b in blocks if b.kind == "Associations")
            print(f"\nwrote {a.out}: {len(blocks)} blocks ({n_spec} specimens, {n_asso} associations)")
        return 1 if issues.count("ERROR") else 0
    if a.cmd == "diff":
        d = diff_blocks(parse_pagetab(a.a), parse_pagetab(a.b), Path(a.a).name, Path(a.b).name)
        print("\n".join(d) if d else "semantically identical")
        return 1 if d else 0
    if a.cmd == "xenium-scan":
        return xenium_scan(a.roots, Path(a.out), a.relative_to, a.workbook)
    return 2


if __name__ == "__main__":
    sys.exit(main())
