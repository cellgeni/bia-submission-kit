#!/usr/bin/env python3
"""Build the web form, or the complete GitHub Pages site.

    python web/build_web.py                -> dist/bia-submission-form.html   (one self-contained file)
    python web/build_web.py --site _site   -> _site/  index.html (the form), template.xlsx,
                                              example/ (synthetic example), guide.html (README)

The form inlines bia_kit.py and runs it in the browser with Pyodide, loaded from the jsDelivr
CDN when the page opens. Building the site needs openpyxl; the guide uses the `markdown`
package if installed (otherwise the README is shown as plain text).
"""
import argparse
import html
import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MARK_SRC = '/*__BIA_KIT_PY__*/""'
MARK_LINKS = "/*__SITE_LINKS__*/null"
SITE_MARKER = ".nojekyll"          # marks a directory as a generated site (safe to replace)


def render_form(links=None) -> str:
    tpl = (ROOT / "web" / "form.template.html").read_text(encoding="utf-8")
    src = (ROOT / "bia_kit.py").read_text(encoding="utf-8")
    assert tpl.count(MARK_SRC) == 1 and tpl.count(MARK_LINKS) == 1, "placeholder missing in form.template.html"
    # json.dumps gives a valid JS string literal; escape '</' so the source cannot close the <script>
    out = tpl.replace(MARK_SRC, json.dumps(src).replace("</", "<\\/"))
    return out.replace(MARK_LINKS, json.dumps(links) if links else "null")


def render_guide() -> str:
    text = (ROOT / "README.md").read_text(encoding="utf-8")
    try:
        import markdown
        body = markdown.markdown(text, extensions=["tables", "fenced_code"])
    except ImportError:
        body = f"<pre>{html.escape(text)}</pre>"
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Guide - BioImage Archive submission kit</title>
<style>
 body {{ font:15px/1.55 -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        color:#1d2330; max-width:900px; margin:0 auto; padding:24px 20px 80px; }}
 a {{ color:#2f6fde; }} code, pre {{ background:#f3f5f8; border-radius:4px; }} code {{ padding:1px 4px; }}
 pre {{ padding:10px 12px; overflow:auto; }} pre code {{ padding:0; }}
 table {{ border-collapse:collapse; margin:10px 0; }} th, td {{ border:1px solid #d9dde5; padding:5px 8px; text-align:left; vertical-align:top; }}
 th {{ background:#f0f2f6; }} .back {{ display:inline-block; margin-bottom:12px; }}
</style></head><body><a class="back" href="./">&larr; Open the submission form</a>
{body}
</body></html>
"""


def build_site(out: Path):
    sys.path.insert(0, str(ROOT))
    import bia_kit as K

    if out.exists():
        if any(out.iterdir()) and not (out / SITE_MARKER).exists():
            sys.exit(f"refusing to overwrite {out}: not empty and not a generated site")
        shutil.rmtree(out)
    (out / "example").mkdir(parents=True)
    links = {"guide": "guide.html", "template": "template.xlsx", "example": "example/example.pagetab.tsv"}
    (out / "index.html").write_text(render_form(links), encoding="utf-8")
    (out / "guide.html").write_text(render_guide(), encoding="utf-8")
    K.write_workbook(out / "template.xlsx", dict(K.TEMPLATE_DEFAULTS))
    ex = ROOT / "examples" / "synthetic"
    for name in ("example.pagetab.tsv", "file_list.tsv"):
        shutil.copy(ex / name, out / "example" / name)
    study, tables, labels, _ = K.import_blocks(K.parse_pagetab(ex / "example.pagetab.tsv"))
    K.write_workbook(out / "example" / "example.xlsx", study, tables, labels)
    (out / SITE_MARKER).write_text("", encoding="utf-8")
    for p in sorted(out.rglob("*")):
        if p.is_file():
            print(f"  {p.relative_to(out)}  ({p.stat().st_size / 1024:.0f} KiB)")


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--site", metavar="DIR", help="build the complete static site into DIR")
    a = ap.parse_args()
    if a.site:
        build_site(Path(a.site))
        print(f"site written to {a.site}")
    else:
        out = ROOT / "dist" / "bia-submission-form.html"
        out.parent.mkdir(exist_ok=True)
        out.write_text(render_form(), encoding="utf-8")
        print(f"wrote {out} ({out.stat().st_size / 1024:.0f} KiB)")


if __name__ == "__main__":
    main()
