"""Tests for bia_kit.py.   Run:  python tests/test_bia_kit.py   (exit 0 = all passed)

All data used here are SYNTHETIC: the fictional example in examples/synthetic/ and Xenium-like
folders generated on the fly.
"""
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(HERE))
import bia_kit as K  # noqa: E402

EXAMPLE = HERE / "examples" / "synthetic" / "example.pagetab.tsv"
_TMP = Path(tempfile.mkdtemp(prefix="bia_kit_tests_"))
tempfile.tempdir = str(_TMP)
results = []


def case(name):
    def deco(fn):
        try:
            fn()
            results.append((name, "PASS", ""))
        except AssertionError as e:
            results.append((name, "FAIL", str(e)))
        except Exception as e:  # noqa: BLE001
            results.append((name, "FAIL", f"{type(e).__name__}: {e}"))
        return fn
    return deco


def have_openpyxl():
    try:
        import openpyxl  # noqa: F401
        return True
    except ImportError:
        return False


def load():
    study, tables, labels, notes = K.import_blocks(K.parse_pagetab(EXAMPLE))
    for rows in tables.values():
        for i, r in enumerate(rows, 3):
            r["__row"] = i
    return study, tables, labels, notes


def msgs(issues, level=None):
    return [f"{w}: {m}" for lvl, w, m in issues.items if level is None or lvl == level]


def orcid_with_check(base15: str) -> str:
    total = 0
    for ch in base15:
        total = (total + int(ch)) * 2
    r = (12 - total % 11) % 11
    d = base15 + ("X" if r == 10 else str(r))
    return "-".join(d[i:i + 4] for i in range(0, 16, 4))


@case("ORCID check digit: public test iD passes, computed iDs pass, altered iDs fail")
def _():
    assert K.orcid_ok("0000-0002-1825-0097")            # ORCID's own fictional test record
    assert not K.orcid_ok("0000-0002-1825-0098")
    assert not K.orcid_ok("0000-0002-1825-009")
    seen_x = False
    for n in range(0, 400):
        o = orcid_with_check(f"{n:015d}")
        seen_x |= o.endswith("X")
        assert K.orcid_ok(o), o
    assert seen_x, "no X check digit exercised"


@case("import: no unmapped attributes; specimens, shared protocol and descriptor columns recovered")
def _():
    study, tables, labels, notes = load()
    assert notes == [], notes
    assert len(tables["Specimens"]) == 4 and len(tables["Protocols"]) == 1
    assert labels == ["desc:Age (weeks)", "desc:Sex"], labels


@case("round trip: import -> build (keep-empty) is semantically identical to the input")
def _():
    study, tables, *_ = load()
    blocks, issues = K.build_blocks(study, tables, keep_empty=True)
    assert issues.count("ERROR") == 0, msgs(issues, "ERROR")
    out = _TMP / "rt.tsv"
    out.write_text(K.render_pagetab(blocks), encoding="utf-8")
    d = K.diff_blocks(K.parse_pagetab(EXAMPLE), K.parse_pagetab(out))
    assert d == [], d


@case("an analysis that no specimen uses is reported")
def _():
    study, tables, *_ = load()
    tables["Analyses"].append({"Title": "Unused pipeline 2.0", "Image analysis overview": "never referenced", "__row": 9})
    _, issues = K.build_blocks(study, tables)
    assert any("Unused pipeline 2.0" in m and "never used" in m for m in msgs(issues, "WARN")), msgs(issues)


@case("errors: short title, bad ORCID, unknown affiliation, undefined reference")
def _():
    study, tables, *_ = load()
    study["Title"] = "Too short"
    tables["Authors"][0]["ORCID"] = "0000-0002-1825-0098"
    tables["Authors"][1]["Affiliations"] = "no-such-org"
    tables["Specimens"][0]["Protocol"] = "No such protocol"
    _, issues = K.build_blocks(study, tables)
    e = " | ".join(msgs(issues, "ERROR"))
    for needle in (">= 25 characters", "fails the checksum", "'no-such-org' is not a Key", "'No such protocol' is not defined"):
        assert needle in e, (needle, e)


@case("a specimen assigned to the wrong biosample is flagged")
def _():
    study, tables, *_ = load()
    row = next(r for r in tables["Specimens"] if r["Specimen title"].startswith("MOUSE-02-"))
    row["Biosample"] = "MOUSE-01"
    _, issues = K.build_blocks(study, tables)
    assert any("looks like it belongs to biosample 'MOUSE-02'" in m for m in msgs(issues, "WARN")), msgs(issues)


@case("empty Protocol / Acquisition / Study component default to the single defined option")
def _():
    study, tables, *_ = load()
    for r in tables["Specimens"]:
        r["Protocol"] = r["Image acquisition"] = r["Study component"] = ""
    blocks, issues = K.build_blocks(study, tables)
    assert issues.count("ERROR") == 0, msgs(issues, "ERROR")
    acq = tables["Acquisitions"][0]["Title"]
    a = [b for b in blocks if b.kind == "Associations"]
    assert len(a) == 4 and all(b.get("Image acquisition") == acq for b in a)


@case("copy-pasted internal ids are regenerated, not duplicated")
def _():
    study, tables, *_ = load()
    new = dict(tables["Biosamples"][0])
    new["Title"], new["__row"] = "MOUSE-03", 99
    tables["Biosamples"].append(new)
    blocks, _ = K.build_blocks(study, tables)
    acc = [b.accno for b in blocks if b.kind == "Biosample"]
    assert len(acc) == len(set(acc)) == 3, acc


@case("xlsx write/read round trip is lossless")
def _():
    if not have_openpyxl():
        return
    study, tables, labels, _ = load()
    wb = _TMP / "wb.xlsx"
    K.write_workbook(wb, study, tables, labels)
    s2, t2 = K.read_workbook(wb)
    b1, _ = K.build_blocks(study, tables, keep_empty=True)
    b2, _ = K.build_blocks(s2, t2, keep_empty=True)
    assert K.diff_blocks(b1, b2) == []


@case("workbooks in the older layout (no guide row) are still read")
def _():
    if not have_openpyxl():
        return
    import openpyxl
    study, tables, labels, _ = load()
    wb_path = _TMP / "old.xlsx"
    K.write_workbook(wb_path, study, tables, labels)
    wb = openpyxl.load_workbook(wb_path)
    for sheet in K.TABLES:
        wb[sheet].delete_rows(2)          # drop the guide row -> data start at row 2
    wb.save(wb_path)
    s2, t2 = K.read_workbook(wb_path)
    assert len(t2["Specimens"]) == 4 and t2["Specimens"][0]["__row"] == 2


@case("xenium-scan (synthetic folders): version mismatch, unknown region and missing bundle are reported")
def _():
    if not have_openpyxl():
        return
    study, tables, labels, _ = load()
    tables["Analyses"] = [{"Title": "Onboard analysis 4.0.1.0", "Image analysis overview": "x"},
                          {"Title": "Onboard analysis 3.0.0.15", "Image analysis overview": "x"}]
    for r in tables["Specimens"]:
        r["Image analysis"] = "Onboard analysis 4.0.1.0"
    wb = _TMP / "xen.xlsx"
    K.write_workbook(wb, study, tables, labels)
    root = _TMP / "upload" / "xenium"
    titles = [r["Specimen title"] for r in tables["Specimens"]]
    for t in titles[:-1]:                                   # last specimen: no bundle
        ver = "xenium-3.0.0.15" if t.startswith("MOUSE-02-") else "xenium-4.0.1.0"
        d = root / f"output__{t}"
        (d / "morphology_focus").mkdir(parents=True)
        (d / "experiment.xenium").write_text(json.dumps({"region_name": t, "slide_id": "0", "analysis_sw_version": ver}))
        (d / "morphology_focus" / "focus_0000.ome.tif").write_bytes(b"")
    d = root / "output__EXTRA"
    d.mkdir(parents=True)
    (d / "experiment.xenium").write_text(json.dumps({"region_name": "EXTRA-REGION", "analysis_sw_version": "xenium-4.0.1.0"}))
    out = _TMP / "scan"
    rc = K.xenium_scan([root], out, relative_to=_TMP / "upload", workbook=wb)
    assert rc == 1
    fl = (out / "file_list.tsv").read_text().splitlines()
    assert fl[0].startswith("Files\t") and fl[1].startswith("xenium/output__"), fl[:2]
    sug = (out / "specimens_suggested.tsv").read_text()
    assert "MOUSE-02-SL01\tMOUSE-02\tOnboard analysis 3.0.0.15" in sug, sug


@case("JSON API used by the web form: import/build identical, xlsx export/import lossless")
def _():
    d = json.loads(K.api_import_tsv(EXAMPLE.read_text(encoding="utf-8")))
    assert d["notes"] == [] and len(d["tables"]["Specimens"]) == 4
    d["tables"]["Specimens"].insert(2, {})                  # an empty grid row must be ignored
    r = json.loads(K.api_build(json.dumps({**d, "keep_empty": True})))
    out = _TMP / "api.tsv"
    out.write_text(r["tsv"], encoding="utf-8")
    assert K.diff_blocks(K.parse_pagetab(EXAMPLE), K.parse_pagetab(out)) == []
    assert r["errors"] == 0 and r["specimens"] == 4
    s = json.loads(K.api_schema())
    assert set(s["tables"]) == set(K.TABLES) and s["template_defaults"] == K.TEMPLATE_DEFAULTS
    if not have_openpyxl():
        return
    x = _TMP / "api.xlsx"
    K.api_export_xlsx(json.dumps(d), str(x))
    r2 = json.loads(K.api_build(json.dumps({**json.loads(K.api_import_xlsx(str(x))), "keep_empty": True})))
    assert r2["tsv"] == r["tsv"]


# concrete identifiers / sample values must not appear in explanations
CONCRETE = [r"e\.g\.", r"FBbi_\d{8}", r"NCBITaxon[_:]\d+", r"\b\d{4}-\d{4}-\d{4}-\d{3}[\dX]\b",
            r"ror\.org/0[a-z0-9]{6}\d{2}", r"\bS-BIA[A-Z]*\d+", r"\b10\.\d{4,9}/\S+"]


@case("template: blank, guide row explains every column, dropdowns below the guide row, explanations not examples")
def _():
    texts = [h for _, _, h in K.STUDY_FIELDS] + [c.help for cols in K.TABLES.values() for c in cols] \
        + [t for t, _ in K.README_LINES]
    web = (HERE / "web" / "form.template.html").read_text(encoding="utf-8")
    hits = [(p, t[:70]) for t in texts for p in CONCRETE if re.search(p, t)]
    hits += [(p, "web form") for p in CONCRETE if re.search(p, web)]
    assert not hits, hits
    assert all(h for _, _, h in K.STUDY_FIELDS) and all(c.help for cols in K.TABLES.values() for c in cols)
    if not have_openpyxl():
        return
    import openpyxl
    tmp = _TMP / "template.xlsx"
    assert K.main(["template", str(tmp)]) == 0
    study, tables = K.read_workbook(tmp)
    assert all(rows == [] for rows in tables.values())
    assert {k: v for k, v in study.items() if v} == K.TEMPLATE_DEFAULTS, study
    wb = openpyxl.load_workbook(tmp)
    for sheet in K.TABLES:
        assert all(c.value and c.value.startswith(K.GUIDE_MARK) for c in wb[sheet][2]), sheet
    for dv in wb["Specimens"].data_validations.dataValidation:
        assert not dv.formula1.startswith("=") and "$A$3:" in dv.formula1, dv.formula1
    _, issues = K.build_blocks(study, tables)                # must validate without crashing
    assert issues.count("ERROR") > 0


@case("site build: form, template, example and guide are produced")
def _():
    if not have_openpyxl():
        return
    site = _TMP / "site"
    p = subprocess.run([sys.executable, str(HERE / "web" / "build_web.py"), "--site", str(site)],
                       capture_output=True, text=True)
    assert p.returncode == 0, p.stderr[-800:]
    html = (site / "index.html").read_text(encoding="utf-8")
    assert "/*__BIA_KIT_PY__*/" not in html and "def api_schema" in html
    for f in ("template.xlsx", "guide.html", "example/example.pagetab.tsv", "example/example.xlsx",
              "example/file_list.tsv", ".nojekyll"):
        assert (site / f).exists(), f


if __name__ == "__main__":
    for name, status, detail in results:
        print(f"{status}  {name}" + (f"\n      {detail[:600]}" if detail else ""))
    failed = sum(1 for r in results if r[1] == "FAIL")
    print(f"\n{len(results) - failed}/{len(results)} passed")
    shutil.rmtree(_TMP, ignore_errors=True)
    sys.exit(1 if failed else 0)
