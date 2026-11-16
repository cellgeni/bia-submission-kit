# Synthetic example

A **fictional** submission that shows what a filled-in study looks like. None of it describes real
data. The person and the article are the public test records that ORCID and Crossref provide for
exactly this purpose: Josiah Carberry, ORCID 0000-0002-1825-0097, and DOI 10.5555/12345678.

| File | Content |
|---|---|
| `example.pagetab.tsv` | The submission file (PageTab TSV) |
| `file_list.tsv` | The matching file list for the study component |
| `make_example.py` | Script that generates both files from the definitions it contains |

To try it:

* **In the web form:** click *Load example*.
* **From the command line:**

  ```bash
  python bia_kit.py import-tsv examples/synthetic/example.pagetab.tsv example.xlsx
  python bia_kit.py check example.xlsx --filelist-dir examples/synthetic
  ```

Do not add real submissions to this repository. Unpublished metadata, personal data of donors and
authors' contact details belong outside version control.
