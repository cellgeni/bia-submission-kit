# bia-submission-kit

[![test-and-deploy](https://github.com/cellgeni/bia-submission-kit/actions/workflows/test-and-deploy.yml/badge.svg)](https://github.com/cellgeni/bia-submission-kit/actions/workflows/test-and-deploy.yml)

**Web form: <https://cellgeni.github.io/bia-submission-kit/>**

Templates and tools for preparing the metadata of a **BioImage Archive** submission without typing
the same information repeatedly into the online form. You describe the study once, in a web form or in
an Excel workbook. The kit checks the entries and writes the submission file the archive accepts
(BioStudies **PageTab TSV**, template `BioImages.v5`).

## What is in the kit

| Path | What it is |
|---|---|
| [Web form](https://cellgeni.github.io/bia-submission-kit/) | **The web form.** Nothing to install. The site also offers the blank Excel template and a fictional example. |
| `bia_kit.py` | Command-line tool: converts, validates, compares, scans Xenium output. |
| `web/` | Source of the web form and the site build (`build_web.py`). `web/form.template.html` is not the form itself. |
| `examples/synthetic/` | A fictional, filled-in example submission. |
| `tests/` | Automated tests (`python tests/test_bia_kit.py`), run on every push. |
| `.github/workflows/` | Runs the tests, builds the website and publishes it to GitHub Pages from `main`. |

## Three ways to fill in a submission

All three use the same rules and produce identical output, so you can switch between them at any time.

1. **Web form:** open the [web form](https://cellgeni.github.io/bia-submission-kit/), or a local copy built
   with `python web/build_web.py` (`dist/bia-submission-form.html`, a single file that also works offline
   once the browser has loaded the Python engine). Work through the pages from top to bottom; each field and
   column explains what it needs, and problems are listed as you type. Then use *Download PageTab TSV*.
   *Save draft* stores a copy to reopen later or send to a colleague; *Export Excel* / *Open Excel* switch to
   the spreadsheet. *Load example* fills in the fictional example.
2. **Excel / Google Sheets:** download the blank template from the website (or create one with
   `python bia_kit.py template blank.xlsx`). Fill it in from row 3 downwards, then convert it with
   `python bia_kit.py build my_study.xlsx my_study.tsv`, or open it in the web form and download from there.
3. **Starting from an existing submission:** export the PageTab TSV from the BioStudies submission tool, then
   use *Import PageTab TSV* in the web form, or `python bia_kit.py import-tsv export.tsv my_study.xlsx`.

## How the information is organised

The template follows the REMBI metadata model used by the archive. Information that applies to many images
is written **once** and referenced by name:

* **Study, Organisations, Authors, Links, Publications, Grants:** who, what and how it was funded.
* **Biosamples:** one entry per biological source of material.
* **Sample preparation (Protocols), Image acquisition, Image analysis, Image correlation:** each distinct
  procedure is described once.
* **Study components:** groups of data described together; each comes with one file list that you upload.
* **Specimens:** one row per imaged specimen. The row links the specimen to its biosample, preparation,
  acquisition, analysis and study component. In the submission file, each row becomes a specimen plus its
  association.

Biosamples may get extra **descriptor columns** (headers starting with `desc:`) for facts every sample has
but that have no column of their own. Each value is written into the sample's description as a
`label: value` line.

### Workbook sheets → submission file (technical)

| Sheet | PageTab block(s) |
|---|---|
| Study | `Submission` + `Study` (Title, ReleaseDate, Description, Keywords, License [URL], Funding statement) |
| Links / Publications / Grants | `Link`, `Publication`, `Funding` (empty rows skipped; `--keep-empty` to emit them) |
| Biosamples | `Biosample` (Organism + `[Ontology] NCBITaxon` / `[TermId] NCBITaxon_<id>`) |
| Protocols + Specimens | `Specimen` (one per distinct specimen title) |
| Acquisitions / Correlations / Analyses | `Image acquisition` (+ FBbi qualifiers) / `Image correlation` / `Image analysis` |
| StudyComponents + Specimens | `Study Component` + one `Associations` block per Specimens row |
| Authors / Organisations | `Author` (`<affiliation>` references, one `Role` per role) / `organisation` (`RORID`) |

Columns whose header starts with `_` hold the submission tool's internal section identifiers, so an updated
submission keeps the same identifiers. Duplicates created by copying rows are detected and replaced.

## Checks performed before a file is written

* **Errors (must fix):**
  * required fields are filled; title and description are at least 25 characters; dates are year-month-day;
  * ORCID iDs are well-formed and pass the check-digit test;
  * every reference (affiliation, biosample, protocol, …) points to something that is defined;
  * each specimen has exactly one preparation;
  * file lists start with a `Files` column, use only the archive's allowed characters, and have no blank or
    duplicate lines.
* **Warnings (should check):**
  * entries that are defined but never used;
  * a specimen whose name suggests a different biosample than the one selected;
  * inconsistent capitalisation in specimen names;
  * a funding statement without structured grants;
  * personal data of human donors that will become public;
  * missing ontology or registry identifiers.

## Command-line reference

```bash
python3 -m venv .venv && ./.venv/bin/pip install -r requirements.txt   # once (needs openpyxl)
python bia_kit.py template    blank.xlsx                 # new blank template
python bia_kit.py import-tsv  export.tsv  study.xlsx     # existing submission -> workbook
python bia_kit.py check       study.xlsx  [--filelist-dir DIR]
python bia_kit.py build       study.xlsx  study.tsv      # refuses to write while errors remain
python bia_kit.py diff        old.tsv     new.tsv        # what changed between two submission files
python bia_kit.py xenium-scan DATA_DIR --relative-to UPLOAD_ROOT --out scan/ [--workbook study.xlsx]
python web/build_web.py                                  # build the single-file form -> dist/
python web/build_web.py --site _site                     # build the whole website (needs openpyxl; markdown optional)
python tests/test_bia_kit.py                             # run the tests
```

`xenium-scan` reads the `experiment.xenium` file of every 10x Xenium output folder and writes three things:
an overview table, a file list with per-file metadata, and specimen rows whose analysis entry matches the
analysis software version recorded by the instrument. With `--workbook`, it reports specimens whose analysis
entry or name disagrees with the instrument record. Matching assumes the region name entered on the
instrument equals the specimen title. The scan itself needs only the Python standard library, so it can run
on the cluster where the data are.

## Web form: how it works and what to know

* It runs the same `bia_kit.py` inside the browser (Python via Pyodide), so all three routes give the same result.
* **Internet:** needed when the page opens, to download the Python engine from the jsDelivr CDN. For offline
  or locked-down networks, host the Pyodide files yourself.
* **Your data:** nothing you enter is uploaded. Entries are kept in the browser's local storage until you
  click *New*. The ontology search sends only the typed search text to EBI OLS4.
* **Collaboration:** no simultaneous co-editing. For that, share the Excel workbook (Excel Online / Google
  Sheets) and use the web form to check and export.
* **Browsers:** tested with Chromium and WebKit (Safari's engine).

## Development and publishing

* `main` is published automatically. Each push runs the tests on two Python versions, builds the site with
  `web/build_web.py --site`, and deploys it to GitHub Pages (workflow `test-and-deploy`). Pull requests get
  the tests and the site build, but no deployment.
* The website is generated, never committed: `dist/`, `_site/` and all `.xlsx` files are git-ignored.
* After changing `bia_kit.py`, nothing else needs updating. The form embeds the module at build time, and its
  pages and help texts are generated from the module's schema.
* **Never commit real submission data.** Unpublished metadata, donor information and contact details stay
  outside the repository; use the synthetic example for tests and demonstrations.

## Submitting the file

Upload the TSV in the BioStudies submission tool, or with the API
(`POST https://www.ebi.ac.uk/biostudies/submissions/api/submissions/async/direct`, header `X-SESSION-TOKEN`,
form field `submission=@file.tsv`). A file that carries the accession of an existing submission **updates**
that submission, so agree on this with the BioImage Archive curators first. Upload acceptance has not yet
been verified end-to-end with this kit; try a test submission first.

## Not verified / limitations

* The attribute names of *Image correlation* follow the archive's own `bia-agent` tool (older template) and
  have not been checked against the current form.
* `xenium-scan` has been tested only on synthetic folders.

## Relation to upstream

The archive's [`bia-agent`](https://github.com/BioImage-Archive/bia-agent) (Apache-2.0) converts REMBI YAML to
PageTab. At the version reviewed, it writes the older `BioImages.v4` template and allows only one association
per study component, so it cannot describe studies with many specimens. Adding that to bia-agent would be a
useful upstream contribution.
