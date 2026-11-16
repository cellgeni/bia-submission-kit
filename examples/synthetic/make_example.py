#!/usr/bin/env python3
"""Generate the SYNTHETIC example submission (fictional study, people and data).

    python examples/synthetic/make_example.py

Writes example.pagetab.tsv and file_list.tsv next to this script. Everything here is made up;
the person and the article are the public test records ORCID and Crossref provide for this
purpose (Josiah Carberry, ORCID 0000-0002-1825-0097, DOI 10.5555/12345678).
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent.parent))
import bia_kit as K  # noqa: E402

PROTOCOL = ("SYNTHETIC EXAMPLE. Brains were fixed by immersion in 4% paraformaldehyde for 24 h, cut into "
            "300 µm coronal slices on a vibratome, cleared with a hydrogel-based clearing protocol and stained "
            "with a nuclear dye. Slices were mounted in refractive-index-matched medium.")
study = {
    "Template": "BioImages.v5", "AttachTo": "BioImages",
    "Title": "SYNTHETIC EXAMPLE - Light-sheet imaging of cleared mouse brain slices (fictional study)",
    "Release date": "2099-01-01",
    "Description": ("This is a fictional study used to demonstrate the bia-submission-kit. It does not describe real "
                    "data. Two mice contribute two cleared brain slices each, imaged with one light-sheet setup and "
                    "processed with one segmentation pipeline."),
    "Keywords": "light-sheet microscopy; tissue clearing; mouse brain; example",
    "Acknowledgements": "Fictional imaging facility.",
    "License": "CC BY 4.0",
    "Funding statement": "Fictional work funded by the Example Funding Agency (grant EX-0001).",
}
tables = {
    "Organisations": [{"Key": "inst1", "Name": "Example Institute of Imaging (fictional)"}],
    "Authors": [
        {"Name": "Josiah Carberry", "E-mail": "josiah.carberry@example.org", "ORCID": "0000-0002-1825-0097",
         "Roles": "submitter; corresponding author", "Affiliations": "inst1"},
        {"Name": "Example Researcher", "E-mail": "researcher@example.org", "Roles": "first author",
         "Affiliations": "inst1"},
    ],
    "Links": [{"URL": "https://example.org/fictional-viewer", "Description": "Interactive viewer (fictional)"}],
    "Publications": [{"Title": "Toward a Unified Theory of High-Energy Metaphysics: Silly String Theory",
                      "Authors": "Josiah Carberry", "DOI": "10.5555/12345678", "Year": "2008"}],
    "Grants": [{"Agency": "Example Funding Agency", "Grant ID": "EX-0001"}],
    "Biosamples": [
        {"Title": "MOUSE-01", "Organism": "Mus musculus", "NCBITaxon ID": "10090",
         "Biological entity": "Adult mouse brain", "Experimental variable": "None (example)",
         "desc:Age (weeks)": "12", "desc:Sex": "female"},
        {"Title": "MOUSE-02", "Organism": "Mus musculus", "NCBITaxon ID": "10090",
         "Biological entity": "Adult mouse brain", "Experimental variable": "None (example)",
         "desc:Age (weeks)": "14", "desc:Sex": "male"},
    ],
    "Protocols": [{"Key": "Clearing and nuclear stain", "Sample preparation protocol": PROTOCOL}],
    "Acquisitions": [{
        "Title": "Light-sheet volume", "Imaging instrument": "Light-sheet microscope (fictional model)",
        "Image acquisition parameters": ("SYNTHETIC EXAMPLE. 10x detection objective, 0.65 µm lateral pixel size, "
                                         "2 µm z-step, one channel (nuclear stain), tiled acquisition."),
        "Imaging method": "selective plane illumination fluorescence microscopy",
        "Imaging method ontology": "fbbi", "Imaging method term ID": "FBbi_00000369"}],
    "Analyses": [{"Title": "Nuclei segmentation v1",
                  "Image analysis overview": "SYNTHETIC EXAMPLE. Tiles were stitched and nuclei segmented in 3D."}],
    "StudyComponents": [{"Name": "Light-sheet volumes", "Description": "Stitched volumes and segmentation masks.",
                         "File List": "file_list.tsv"}],
    "Specimens": [
        {"Specimen title": f"MOUSE-0{m}-SL0{s}", "Biosample": f"MOUSE-0{m}", "Image analysis": "Nuclei segmentation v1"}
        for m in (1, 2) for s in (1, 2)
    ],
}

if __name__ == "__main__":
    for name, rows in tables.items():
        for i, r in enumerate(rows, 1):
            r["__row"] = i
    blocks, issues = K.build_blocks(study, tables)
    print(issues.render())
    assert issues.count("ERROR") == 0
    (HERE / "example.pagetab.tsv").write_text(K.render_pagetab(blocks), encoding="utf-8")
    lines = ["Files\tbiosample\tspecimen\tcontent"]
    for r in tables["Specimens"]:
        t = r["Specimen title"]
        lines += [f"{t}/{t}_volume.ome.zarr.zip\t{r['Biosample']}\t{t}\tstitched volume",
                  f"{t}/{t}_nuclei_masks.ome.zarr.zip\t{r['Biosample']}\t{t}\tsegmentation masks"]
    (HERE / "file_list.tsv").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("wrote example.pagetab.tsv and file_list.tsv")
