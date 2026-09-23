# Shared helpers for Wayground sets with pictures.
# Images live in this public repo; Wayground's "Image Link" column points at the raw GitHub URL.
import csv, os
import fitz, openpyxl
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = "https://raw.githubusercontent.com/ferhodes/wayground-images/main/"
HEAD = ["Question Text", "Question Type", "Option 1", "Option 2", "Option 3", "Option 4", "Option 5",
        "Correct Answer", "Time in seconds", "Image Link", "Answer explanation"]

def svg_to_png(svg_text, out_rel, width_px=1200):
    """Render an SVG string to a PNG at ROOT/out_rel. Returns the public URL."""
    out = os.path.join(ROOT, out_rel)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    doc = fitz.open(stream=svg_text.encode("utf-8"), filetype="svg")
    page = doc[0]
    zoom = width_px / page.rect.width
    page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False).save(out)
    return url(out_rel)

def url(rel):
    return RAW + rel.replace("\\", "/")

def write_set(rows, xlsx_path):
    """rows: [question, type, [options], correct(1-5 or ''), seconds, image_url or '', explanation].
    Writes the .xlsx and a .csv beside it, in Wayground's template columns."""
    out = []
    for q, typ, opts, ans, secs, img, why in rows:
        opts = list(opts) + [""] * (5 - len(opts))
        assert len(opts) == 5 and (typ != "Multiple Choice" or 1 <= int(ans) <= 5), q
        out.append([q, typ, *opts, str(ans), secs, img, why])
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(HEAD)
    for r in out: ws.append(r)
    wb.save(xlsx_path)
    with open(os.path.splitext(xlsx_path)[0] + ".csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f); w.writerow(HEAD); w.writerows(out)
    return len(out)
