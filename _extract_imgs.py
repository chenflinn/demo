"""Extract all images from thesis docx + map each to its chapter."""
import os, json, shutil, re
from pathlib import Path
from zipfile import ZipFile
from docx import Document
from docx.oxml.ns import qn

DOCX = r"D:\file\组会\开题报告\2026.5答辩\专硕1组-1-陈未艾.docx"
OUT_DIR = Path(r"C:/Users/99375/Desktop/demo/thesis_images")
OUT_DIR.mkdir(parents=True, exist_ok=True)
for f in OUT_DIR.iterdir():
    if f.is_file():
        f.unlink()

# 1) Unzip word/media/ → OUT_DIR
with ZipFile(DOCX) as z:
    medias = [n for n in z.namelist() if n.startswith("word/media/")]
    for n in medias:
        target = OUT_DIR / Path(n).name
        with z.open(n) as src, open(target, "wb") as dst:
            shutil.copyfileobj(src, dst)
print(f"extracted {len(medias)} media files to {OUT_DIR}")

# 2) Iterate paragraphs in body order, find inline blips → which rId references which media
d = Document(DOCX)
# rId → target filename map
rels = {r.rId: Path(r.target_ref).name for r in d.part.rels.values() if r.target_ref.startswith("media/")}

# Build chapter mapping by walking paragraphs in order, tracking current chapter
chapter_keywords = [
    ("cover",     re.compile(r"^(硕\s*士|摘\s*要|Abstract|关键词|目\s*录|原创|授权)")),
    ("chap1_intro",     re.compile(r"^(绪\s*论|1\.|研究背景|国内外|本文主要研究|组织结构)")),
    ("chap2_basics",    re.compile(r"^(相关理论|2\.|生成对抗网络|流模型|编码器|Transformer|卷积|池化|激活|自适应实例|对数颜色|海灵格)")),
    ("chap3_fqgan",     re.compile(r"^(3\.|Fq-?GAN|高低频分离|双分支颜色)")),
    ("chap4_sortflow",  re.compile(r"^(4\.|SortFlow|流的颜色)")),
    ("chap5_ghnet",     re.compile(r"^(5\.|GHNet|基于排序的颜色调制|梯度引导|两阶段)")),
    ("chap6_conclusion",re.compile(r"^(6\.|总结|展望|参考文献|致\s*谢)")),
]
current = "cover"
img_records = []
img_idx = 0
total_inline = 0

# We walk paragraphs in document order
for pi, p in enumerate(d.paragraphs):
    text = p.text.strip()
    if text:
        for chap, pat in chapter_keywords:
            if pat.match(text):
                current = chap
                break
    # find inline blip references inside this paragraph
    blips = p._p.findall(".//" + qn("a:blip"))
    for blip in blips:
        embed = blip.get(qn("r:embed"))
        if embed and embed in rels:
            total_inline += 1
            fname = rels[embed]
            img_records.append({
                "order": img_idx,
                "para_idx": pi,
                "image": fname,
                "chapter": current,
                "neighbor_text": text[:80],
            })
            img_idx += 1

print(f"found {total_inline} inline image refs across {len(d.paragraphs)} paragraphs")

# Also check tables (figures sometimes live in table cells)
for ti, tbl in enumerate(d.tables):
    for ri, row in enumerate(tbl.rows):
        for ci, cell in enumerate(row.cells):
            for p in cell.paragraphs:
                blips = p._p.findall(".//" + qn("a:blip"))
                for blip in blips:
                    embed = blip.get(qn("r:embed"))
                    if embed and embed in rels:
                        fname = rels[embed]
                        img_records.append({
                            "order": img_idx,
                            "para_idx": -1,
                            "table_idx": ti,
                            "image": fname,
                            "chapter": current,
                            "neighbor_text": f"(table {ti} r{ri} c{ci}) {p.text.strip()[:60]}",
                        })
                        img_idx += 1

print(f"total image records (incl table cells): {len(img_records)}")
print(f"breakdown by chapter:")
from collections import Counter
cnt = Counter(r["chapter"] for r in img_records)
for k, v in cnt.items():
    print(f"  {k}: {v}")

# Save mapping
(OUT_DIR / "_mapping.json").write_text(
    json.dumps(img_records, ensure_ascii=False, indent=2),
    encoding="utf-8",
)
print(f"saved mapping to {OUT_DIR / '_mapping.json'}")
