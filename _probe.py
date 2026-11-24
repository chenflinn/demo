"""Probe thesis docx + PPT template structure."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# ---------- docx ----------
from docx import Document

docx_path = r"D:\file\组会\开题报告\2026.5答辩\专硕1组-1-陈未艾.docx"
d = Document(docx_path)

print("=" * 60)
print("[DOCX] total paragraphs:", len(d.paragraphs))
print("[DOCX] total tables :", len(d.tables))

# Headings list
print("\n[DOCX] heading outline:")
seen_n = 0
for i, p in enumerate(d.paragraphs):
    sty = (p.style.name or "")
    if sty.lower().startswith("heading") and p.text.strip():
        level = sty.split(" ")[-1] if " " in sty else "?"
        print(f"  {sty:14s}  {p.text.strip()[:80]}")
        seen_n += 1
        if seen_n > 80:
            print("  ... truncated")
            break

# First 2k chars of body text (skip headings to get abstract / intro flavor)
print("\n[DOCX] first ~1000 chars of body:")
buf = []
for p in d.paragraphs:
    t = p.text.strip()
    if t:
        buf.append(t)
    if sum(len(x) for x in buf) > 1500:
        break
print("\n".join(buf)[:1500])

# ---------- pptx ----------
from pptx import Presentation
from pptx.util import Emu

pptx_path = r"E:\中期\中期答辩-陈未艾.pptx"
p = Presentation(pptx_path)

print("\n" + "=" * 60)
print(f"[PPTX] slide_width  : {p.slide_width / 914400:.2f}\"  ({p.slide_width} EMU)")
print(f"[PPTX] slide_height : {p.slide_height / 914400:.2f}\"  ({p.slide_height} EMU)")
print(f"[PPTX] total slides : {len(p.slides)}")
print(f"[PPTX] slide_layouts: {len(p.slide_layouts)}")
for i, lay in enumerate(p.slide_layouts):
    print(f"   layout[{i}]: {lay.name}")

print("\n[PPTX] per-slide summary:")
for i, slide in enumerate(p.slides):
    txts = []
    for shp in slide.shapes:
        if shp.has_text_frame:
            for para in shp.text_frame.paragraphs:
                t = "".join(r.text for r in para.runs).strip()
                if t:
                    txts.append(t)
    layout_name = slide.slide_layout.name if slide.slide_layout else "?"
    head = " | ".join(txts[:3])[:120]
    n_pics = sum(1 for s in slide.shapes if s.shape_type == 13)  # PICTURE=13
    print(f"  slide[{i:02d}]  layout={layout_name[:24]:24s}  pics={n_pics}  text='{head}'")

# Capture font from first non-title text in slide 0
print("\n[PPTX] cover slide fonts:")
slide0 = p.slides[0]
for shp in slide0.shapes:
    if shp.has_text_frame:
        for para in shp.text_frame.paragraphs:
            for r in para.runs:
                if r.text.strip():
                    f = r.font
                    fs = f.size.pt if f.size else None
                    print(f"   '{r.text[:40]}'  name={f.name}  size={fs}  bold={f.bold}")
                    break
            else:
                continue
            break
