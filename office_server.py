"""
Office Suite - PYTHON BACKEND  (office_server.py)

Formats:
  Word   : open/save  .docx  .txt  .html   (open: .md bhi)
  Sheet  : open/save  .xlsx  .csv  .txt    (open: .tsv bhi)
  Slides : open/save  .pptx  .txt

Chalane ka tareeka:
  1. office_server.py aur OfficeSuite.html ek hi folder mein rakhein
  2. python office_server.py
     (pehli baar libraries khud install hoti hain; internet chahiye)
  3. Browser khud khulega: http://127.0.0.1:8000
Har save ki copy "exports" folder mein bhi banti hai.
"""
import csv
import datetime
import html as H
import importlib.util
import io
import json
import os
import re
import subprocess
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs, urlparse

for _mod, _pkg in (("docx", "python-docx"), ("openpyxl", "openpyxl"), ("pptx", "python-pptx")):
    if importlib.util.find_spec(_mod) is None:
        print(f"Installing {_pkg} ...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", _pkg])
        except Exception:
            print(f"{_pkg} install nahi hui. Manual: pip install {_pkg}")

HOST, PORT = "127.0.0.1", 8000
BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "exports")
os.makedirs(OUT, exist_ok=True)
MAX_UPLOAD = 50 * 1024 * 1024
BLANK = {"t": "", "b": "", "bg": "#ffffff", "fg": "#1c1e21"}


def decode(data):
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252", "replace")


# ============================ WORD ============================
def build_docx(d):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH as A
    from docx.shared import Pt, RGBColor

    align = {"center": A.CENTER, "right": A.RIGHT, "justify": A.JUSTIFY, "left": A.LEFT}

    def add_runs(p, rs):
        for r in rs:
            run = p.add_run(r["t"])
            run.bold, run.italic, run.underline = r.get("b"), r.get("i"), r.get("u")
            if r.get("size"):
                run.font.size = Pt(round(r["size"] * 0.75))
            if r.get("font"):
                run.font.name = r["font"]
            if r.get("color"):
                run.font.color.rgb = RGBColor(*r["color"])

    doc = Document()
    for b in d.get("blocks", []):
        k = b["type"]
        if k == "table":
            rows = b["rows"]
            if not rows:
                continue
            t = doc.add_table(rows=len(rows), cols=max(len(r) for r in rows))
            t.style = "Table Grid"
            for i, row in enumerate(rows):
                for j, cell in enumerate(row):
                    add_runs(t.cell(i, j).paragraphs[0], cell)
            doc.add_paragraph()
            continue
        if k in ("h1", "h2", "h3"):
            p = doc.add_heading(level=int(k[1]))
        elif k == "ul":
            p = doc.add_paragraph(style="List Bullet")
        elif k == "ol":
            p = doc.add_paragraph(style="List Number")
        else:
            p = doc.add_paragraph()
        add_runs(p, b.get("runs", []))
        if b.get("align") in align:
            p.alignment = align[b["align"]]
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def word_txt(d):
    lines, n = [], 0
    for b in d.get("blocks", []):
        if b["type"] == "table":
            for row in b["rows"]:
                lines.append("\t".join("".join(r["t"] for r in c).replace("\n", " ") for c in row))
            lines.append("")
            continue
        n = n + 1 if b["type"] == "ol" else 0
        t = "".join(r["t"] for r in b.get("runs", []))
        lines.append(("- " if b["type"] == "ul" else f"{n}. " if b["type"] == "ol" else "") + t)
    return "\n".join(lines).encode("utf-8")


def word_html(d, title):
    return (
        '<!DOCTYPE html>\n<html><head><meta charset="utf-8"><title>%s</title><style>'
        "body{margin:0;background:#f4f5f7}#ed{max-width:850px;margin:24px auto;background:#fff;color:#111;"
        "padding:40px 48px;font-family:Georgia,serif;font-size:18px;line-height:1.6}"
        "#ed table{border-collapse:collapse;width:100%%}#ed td{border:1px solid #999;padding:6px 8px}"
        '</style></head><body><div id="ed">%s</div></body></html>' % (H.escape(title), d.get("html", ""))
    ).encode("utf-8")


def docx_to_html(data):
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(io.BytesIO(data))

    def runs(p):
        out = []
        for r in p.runs:
            t = H.escape(r.text).replace("\n", "<br>").replace("\t", "&emsp;")
            if not t:
                continue
            f, st = r.font, []
            if f.size:
                st.append("font-size:%dpx" % round(f.size.pt * 4 / 3))
            if f.name:
                st.append("font-family:'%s'" % H.escape(f.name))
            try:
                if f.color is not None and f.color.rgb is not None:
                    st.append("color:#" + str(f.color.rgb))
            except Exception:
                pass
            if f.bold:
                t = "<b>%s</b>" % t
            if f.italic:
                t = "<i>%s</i>" % t
            if f.underline:
                t = "<u>%s</u>" % t
            if f.strike:
                t = "<s>%s</s>" % t
            out.append('<span style="%s">%s</span>' % (";".join(st), t) if st else t)
        return "".join(out)

    out, open_tag = [], None
    for el in doc.element.body.iterchildren():
        tag = el.tag.rsplit("}", 1)[-1]
        p, sname = None, ""
        if tag == "p":
            p = Paragraph(el, doc)
            sname = (p.style.name if p.style is not None else "") or ""
            want = "ul" if sname.startswith("List Bullet") else "ol" if sname.startswith("List Number") else None
            if want:
                if open_tag != want:
                    if open_tag:
                        out.append("</%s>" % open_tag)
                    out.append("<%s>" % want)
                    open_tag = want
                out.append("<li>%s</li>" % (runs(p) or "<br>"))
                continue
        if open_tag:
            out.append("</%s>" % open_tag)
            open_tag = None
        if tag == "p":
            m = re.match(r"(Heading|Title)\s*(\d?)", sname)
            t = "h%d" % (1 if m.group(1) == "Title" else min(int(m.group(2) or 1), 3)) if m else "p"
            a = str(p.alignment).upper() if p.alignment is not None else ""
            al = "center" if "CENTER" in a else "right" if "RIGHT" in a else "justify" if "JUSTIFY" in a else ""
            out.append("<%s%s>%s</%s>" % (t, ' style="text-align:%s"' % al if al else "", runs(p) or "<br>", t))
        elif tag == "tbl":
            tb = Table(el, doc)
            rows = []
            for row in tb.rows:
                cells = "".join("<td>%s</td>" % ("<br>".join(runs(c) for c in cell.paragraphs) or "<br>") for cell in row.cells)
                rows.append("<tr>%s</tr>" % cells)
            out.append("<table>%s</table><p><br></p>" % "".join(rows))
    if open_tag:
        out.append("</%s>" % open_tag)
    return "".join(out)


def text_to_html(txt):
    return "".join("<p>%s</p>" % (H.escape(l) or "<br>") for l in txt.splitlines()) or "<p><br></p>"


def html_body(txt):
    m = re.search(r'<div id="ed"[^>]*>(.*)</div>\s*</body>', txt, re.S | re.I) or re.search(r"<body[^>]*>(.*)</body>", txt, re.S | re.I)
    return m.group(1) if m else txt


# ============================ SHEET ============================
def build_xlsx(d):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    for r, row in enumerate(d.get("cells", []), 1):
        for c, v in enumerate(row, 1):
            if v == "":
                continue
            if not v.startswith("="):
                try:
                    v = float(v)
                    v = int(v) if v.is_integer() else v
                except ValueError:
                    pass
            ws.cell(row=r, column=c, value=v)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def sheet_text(d, delim):
    buf = io.StringIO()
    csv.writer(buf, delimiter=delim, lineterminator="\r\n").writerows(d.get("values", []))
    return buf.getvalue().encode("utf-8-sig" if delim == "," else "utf-8")


def cell_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return str(getattr(v, "text", v))


def xlsx_to_cells(data):
    from openpyxl import load_workbook

    ws = load_workbook(io.BytesIO(data)).active
    rows = [[cell_str(c.value) for c in row]
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 300), max_col=min(ws.max_column or 1, 40))]
    while rows and not any(rows[-1]):
        rows.pop()
    return rows


def csv_to_cells(txt):
    try:
        dialect = csv.Sniffer().sniff(txt[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    return [r[:40] for r in csv.reader(io.StringIO(txt), dialect)][:300]


# ============================ SLIDES ============================
def build_pptx(d):
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    hx = lambda h: RGBColor.from_string(h.lstrip("#").upper())
    for s in d.get("slides", []):
        sl = prs.slides.add_slide(prs.slide_layouts[1])
        sl.background.fill.solid()
        sl.background.fill.fore_color.rgb = hx(s["bg"])
        title, body = sl.shapes.title, sl.placeholders[1]
        title.left, title.top, title.width, title.height = Inches(0.8), Inches(0.5), Inches(11.7), Inches(1.4)
        body.left, body.top, body.width, body.height = Inches(0.8), Inches(2.1), Inches(11.7), Inches(4.8)
        title.text = s["t"].strip()
        lines = [x.strip() for x in s["b"].split("\n") if x.strip()] or [""]
        tf = body.text_frame
        tf.text = lines[0]
        for x in lines[1:]:
            tf.add_paragraph().text = x
        for shape, size in ((title, 40), (body, 26)):
            for para in shape.text_frame.paragraphs:
                for r in para.runs:
                    r.font.size, r.font.color.rgb = Pt(size), hx(s["fg"])
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def slides_txt(d):
    parts = []
    for s in d.get("slides", []):
        lines = [x.strip() for x in s["b"].split("\n") if x.strip()]
        parts.append(s["t"].strip() + "\n" + "\n".join("- " + x for x in lines))
    return "\n---\n".join(parts).encode("utf-8")


def pptx_to_slides(data):
    from pptx import Presentation
    from pptx.enum.dml import MSO_FILL

    res = []
    prs = Presentation(io.BytesIO(data))  # prs ko zinda rakhna zaroori hai
    for sl in prs.slides:
        ttl = sl.shapes.title
        title, lines = "", []
        if ttl is not None and ttl.has_text_frame:
            title = ttl.text_frame.text.strip()
        for sh in sl.shapes:
            if ttl is not None and sh.shape_id == ttl.shape_id:
                continue
            if sh.has_text_frame:
                lines += [p.text.strip() for p in sh.text_frame.paragraphs if p.text.strip()]
        if not title and lines:
            title = lines.pop(0)
        bg = "#ffffff"
        try:
            f = sl.background.fill
            if f.type == MSO_FILL.SOLID:
                bg = "#" + str(f.fore_color.rgb)
        except Exception:
            pass
        lum = int(bg[1:3], 16) * 0.299 + int(bg[3:5], 16) * 0.587 + int(bg[5:7], 16) * 0.114
        res.append({"t": title, "b": "\n".join(lines), "bg": bg, "fg": "#ffffff" if lum < 128 else "#1c1e21"})
    return res or [dict(BLANK)]


def txt_to_slides(txt):
    res = []
    for chunk in re.split(r"^\s*---\s*$", txt, flags=re.M):
        lines = [l.strip() for l in chunk.splitlines() if l.strip()]
        if lines:
            res.append({"t": lines[0].lstrip("# ").strip(),
                        "b": "\n".join(re.sub(r"^[-*\u2022]\s*", "", l) for l in lines[1:]),
                        "bg": "#ffffff", "fg": "#1c1e21"})
    return res or [dict(BLANK)]


# ============================ ROUTES ============================
def do_save(tab, fmt, name, d):
    if tab == "word":
        if fmt == "docx":
            return build_docx(d)
        if fmt == "txt":
            return word_txt(d)
        if fmt == "html":
            return word_html(d, name)
    elif tab == "sheet":
        if fmt == "xlsx":
            return build_xlsx(d)
        if fmt in ("csv", "txt"):
            return sheet_text(d, "," if fmt == "csv" else "\t")
    elif tab == "slides":
        if fmt == "pptx":
            return build_pptx(d)
        if fmt == "txt":
            return slides_txt(d)
    raise ValueError("Ye format is tab mein support nahi hai: ." + fmt)


def do_open(name, tab, data):
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    if ext in ("doc", "xls", "ppt"):
        raise ValueError("Purani .%s file support nahi. Usko Office mein .%sx save karke kholein." % (ext, ext))
    if ext == "docx":
        return {"kind": "word", "html": docx_to_html(data)}
    if ext == "xlsx":
        return {"kind": "sheet", "cells": xlsx_to_cells(data)}
    if ext == "pptx":
        return {"kind": "slides", "slides": pptx_to_slides(data)}
    if ext in ("csv", "tsv"):
        return {"kind": "sheet", "cells": csv_to_cells(decode(data))}
    if ext in ("html", "htm"):
        return {"kind": "word", "html": html_body(decode(data))}
    if ext in ("txt", "md", ""):
        txt = decode(data)
        if tab == "sheet":
            return {"kind": "sheet", "cells": csv_to_cells(txt)}
        if tab == "slides":
            return {"kind": "slides", "slides": txt_to_slides(txt)}
        return {"kind": "word", "html": text_to_html(txt)}
    raise ValueError("Ye file type support nahi: ." + ext)


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(data)

    def _json(self, code, obj):
        self._send(code, json.dumps(obj, ensure_ascii=False))

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        if self.path.split("?")[0] in ("/", "/index.html"):
            path = os.path.join(BASE, "OfficeSuite.html")
            if not os.path.exists(path):
                return self._send(404, "OfficeSuite.html isi folder mein rakhein.", "text/plain; charset=utf-8")
            with open(path, "rb") as f:
                return self._send(200, f.read(), "text/html; charset=utf-8")
        self._send(404, "Not found", "text/plain")

    def do_POST(self):
        url = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(url.query).items()}
        try:
            n = int(self.headers.get("Content-Length", 0))
            if n > MAX_UPLOAD:
                return self._json(413, {"error": "File bohat badi hai (50 MB se zyada)."})
            raw = self.rfile.read(n)
            if url.path == "/api/open":
                return self._json(200, do_open(q.get("name", ""), q.get("tab", "word"), raw))
            if url.path == "/api/save":
                fmt = re.sub(r"\W", "", q.get("fmt", ""))
                name = re.sub(r"[^\w\- ]", "_", os.path.basename(q.get("name", "")).rsplit(".", 1)[0]).strip() or "file"
                blob = do_save(q.get("tab", "word"), fmt, name, json.loads(raw.decode("utf-8")))
                with open(os.path.join(OUT, f"{name}.{fmt}"), "wb") as f:
                    f.write(blob)
                print(f"Saved: exports/{name}.{fmt}")
                return self._send(200, blob, "application/octet-stream")
            self._json(404, {"error": "Not found"})
        except Exception as e:
            self._json(500, {"error": "%s: %s" % (type(e).__name__, e)})

    def log_message(self, *a):
        pass


if __name__ == "__main__":
    srv = None
    for port in range(PORT, PORT + 20):
        try:
            srv = HTTPServer((HOST, port), Handler)
            break
        except OSError:
            continue
    if srv is None:
        sys.exit("Koi free port nahi mila.")
    url = f"http://{HOST}:{srv.server_port}"
    print(f"Office Suite chal raha hai: {url}  (band karne ke liye Ctrl+C)")
    try:
        webbrowser.open(url)
    except Exception:
        pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        srv.server_close()
