"""Probe thesis docx + PPT template, write report to UTF-8 file."""
from docx import Document
from pptx import Presentation
from pathlib import Path

OUT = Path(r"C:/Users/99375/Desktop/demo/_probe_out.txt")
lines = []
def w(s=""):
    lines.append(str(s))

# ---------- docx ----------
docx_path = r"D:\file\组会\开题报告\2026.5答辩\专硕1组-1-陈未艾.docx"
d = Document(docx_path)

w("=" * 60)
w(f"[DOCX] total paragraphs: {len(d.paragraphs)}")
w(f"[DOCX] total tables    : {len(d.tables)}")

# Heading/title styled paragraphs (Chinese style names common)
w("\n[DOCX] heading-styled paragraphs:")
hdr_n = 0
for p in d.paragraphs:
    sty = p.style.name or ""
    t = p.text.strip()
    if t and any(k in sty.lower() for k in ("heading", "title")) or (t and ("标题" in sty or "Title" in sty)):
        w(f"  [{sty:14s}] {t[:100]}")
        hdr_n += 1
if hdr_n == 0:
    w("  (no heading-styled paragraphs found)")

# All paragraphs that look like section headers (short, bolded, numbered)
w("\n[DOCX] paragraphs that look like section headers (numbered like '1.', '1.1', or Chinese '第X章'):")
import re
pat = re.compile(r"^\s*(\d+(?:\.\d+)*[.\s]|第[一二三四五六七八九十]+[章节]|摘\s*要|Abstract|目\s*录|参考文献|致\s*谢|附\s*录)")
cnt = 0
for p in d.paragraphs:
    t = p.text.strip()
    if pat.match(t) and len(t) < 60:
        w(f"  {t}")
        cnt += 1
        if cnt > 60:
            w("  ... truncated")
            break

# Abstract (zhaiyao) + first ~1500 chars of body
w("\n[DOCX] first 2000 chars of body text (skipping empty paragraphs):")
buf = []
for p in d.paragraphs:
    t = p.text.strip()
    if t:
        buf.append(t)
    if sum(len(x) for x in buf) > 2500:
        break
w("\n".join(buf)[:2500])

# ---------- pptx ----------
w("\n" + "=" * 60)
pptx_path = r"E:\中期\中期答辩-陈未艾.pptx"
p = Presentation(pptx_path)
w(f"[PPTX] slide_size: {p.slide_width/914400:.2f}\" x {p.slide_height/914400:.2f}\"")
w(f"[PPTX] total slides: {len(p.slides)}  layouts: {len(p.slide_layouts)}")

w("\n[PPTX] slide_layouts:")
for i, lay in enumerate(p.slide_layouts):
    w(f"  layout[{i}]: {lay.name}")

w("\n[PPTX] per-slide content:")
for i, slide in enumerate(p.slides):
    layout_name = slide.slide_layout.name if slide.slide_layout else "?"
    n_pics = sum(1 for s in slide.shapes if s.shape_type == 13)
    n_shapes = len(slide.shapes)
    w(f"\n--- slide[{i:02d}]  layout={layout_name}  shapes={n_shapes}  pics={n_pics} ---")
    for sj, shp in enumerate(slide.shapes):
        if shp.has_text_frame:
            for para in shp.text_frame.paragraphs:
                t = "".join(r.text for r in para.runs).strip()
                if t:
                    w(f"  [{sj}] {t[:150]}")

# Font usage on cover
w("\n[PPTX] cover slide runs (text/font/size):")
slide0 = p.slides[0]
for shp in slide0.shapes:
    if shp.has_text_frame:
        for para in shp.text_frame.paragraphs:
            for r in para.runs:
                if r.text.strip():
                    f = r.font
                    fs = f.size.pt if f.size else None
                    color = None
                    try:
                        if f.color and f.color.rgb:
                            color = str(f.color.rgb)
                    except Exception:
                        pass
                    w(f"   '{r.text[:30]}'  name={f.name}  size={fs}  bold={f.bold}  color={color}")

OUT.write_text("\n".join(lines), encoding="utf-8")
print("wrote", OUT, len(lines), "lines")
