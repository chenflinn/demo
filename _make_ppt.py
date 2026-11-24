"""
Build the graduation defense PPT.
- base on template E:\\中期\\中期答辩-陈未艾.pptx (inherits master/theme)
- wipe existing slides, add 40 new slides per outline
- embed images from thesis docx
"""
import copy
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

TEMPLATE = r"E:\中期\中期答辩-陈未艾.pptx"
OUT      = r"D:\file\组会\开题报告\2026.5答辩\毕业答辩-陈未艾.pptx"
IMG_DIR  = Path(r"C:/Users/99375/Desktop/demo/thesis_images")

# ----- Visual constants -----
FONT        = "微软雅黑"
ACCENT      = RGBColor(0x00, 0x5C, 0xA1)   # template main blue
GRAY        = RGBColor(0x44, 0x44, 0x44)
LIGHT       = RGBColor(0x88, 0x88, 0x88)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# ----- Slide outline (40 slides) -----
SLIDES = [
    # 1
    {"kind": "cover",
     "title": "基于频率分离的高保真工业数据的颜色迁移",
     "subtitle": ["人工智能与计算机学院", "2023 级 计算机技术",
                  "汇报人:陈未艾", "指导老师:李伟"]},
    # 2
    {"kind": "toc", "title": "汇报内容",
     "items": [
         "一、研究背景与意义",
         "二、国内外研究现状",
         "三、本文主要研究内容",
         "四、方法一  Fq-GAN(基于生成对抗网络)",
         "五、方法二  SortFlow(基于流模型)",
         "六、方法三  GHNet(基于特征匹配)",
         "七、工作总结与未来展望",
         "八、学术成果",
     ]},
    # 3
    {"kind": "content", "title": "一、研究背景与意义", "subtitle": "工业图像光照变化的挑战",
     "bullets": [
         "工业产品检测系统中,照明对检测精度起到至关重要的作用",
         "工业环境复杂多变:污染侵蚀、设备震动、温度变化",
         "照明长时间运行后存在性能衰减,导致光照不稳定",
         "缺陷检测模型对颜色高度敏感,光照变化 → 色调漂移 → 检测性能下降",
     ]},
    # 4
    {"kind": "content", "title": "一、研究背景与意义", "subtitle": "颜色迁移在工业场景中的价值",
     "bullets": [
         "数据增强:利用已有缺陷数据,生成不同光照条件下的等价样本",
         "降低成本:避免重复采集多光源现场数据带来的人力与时间开销",
         "提升泛化:让检测模型在多变现场环境下保持稳定精度",
         "核心需求:颜色匹配准确  +  结构细节保留  +  训练 / 推理高效",
     ]},
    # 5
    {"kind": "content", "title": "二、国内外研究现状", "subtitle": "三类主流颜色迁移方法",
     "bullets": [
         "① 基于条件生成模型:GAN / Diffusion 等,把目标颜色作为条件控制生成 ",
         "② 基于可逆映射:利用流模型在源、目标分布之间建立可逆变换",
         "③ 基于特征匹配:利用预训练编码器在特征空间对齐源、目标颜色统计",
         "三类方法各有侧重,但在工业高保真场景下都存在明显短板",
     ]},
    # 6
    {"kind": "content", "title": "二、国内外研究现状", "subtitle": "三类方法各自的局限",
     "bullets": [
         "条件生成模型:颜色与结构耦合,简单替换样式易导致结构失真、颜色偏差",
         "可逆映射:高维空间建模复杂,易出现局部不一致或映射不稳定",
         "特征匹配:严重依赖源、目标特征相似度,差异较大时易出现颜色失真",
         "→ 工业场景需要同时具备  高保真  + 强泛化  +  低复杂度  的颜色迁移方法",
     ]},
    # 7
    {"kind": "content", "title": "三、本文主要研究内容", "subtitle": "围绕「频率分离」思想提出三种互补方法",
     "bullets": [
         "① Fq-GAN —— 基于频率分离的生成对抗网络,解耦颜色与结构信息",
         "② SortFlow —— 基于流模型,排序后的低频信息降低高维建模难度",
         "③ GHNet —— 基于特征匹配,梯度引导高频增强 + 排序颜色调制",
         "共同主线:在频率层面把颜色与结构解耦,分支建模、各自优化",
     ]},

    # ---------------- Fq-GAN ----------------
    # 8
    {"kind": "content", "title": "四、方法一:Fq-GAN", "subtitle": "研究动机",
     "bullets": [
         "条件生成 GAN 通过替换样式向量注入颜色,颜色与结构在特征空间高度耦合",
         "造成生成图像结构失真、颜色偏差、纹理不一致",
         "→ 需要在颜色与结构之间引入显式的解耦机制",
         "想法:在 频率层面 分离图像 —— 低频管颜色,高频管结构",
     ]},
    # 9 architecture image
    {"kind": "image", "title": "Fq-GAN:总体架构", "image": "image17.png",
     "caption": "双分支生成器 + 频率分离器,颜色与结构各走一路"},
    # 10
    {"kind": "image", "title": "Fq-GAN:高低频分离器", "image": "image18.png",
     "caption": "可学习的低通滤波器 → 低频 (全局色彩),残差 → 高频 (局部结构)"},
    # 11
    {"kind": "image", "title": "Fq-GAN:双分支颜色迁移模块", "image": "image19.png",
     "caption": "低频分支注入颜色直方图特征,高频分支保留结构,最终融合输出"},
    # 12
    {"kind": "content", "title": "Fq-GAN:损失函数 & 全局最优性",
     "bullets": [
         "颜色损失:在低频分支对齐 RGBuv 颜色直方图,保证全局色彩一致",
         "内容损失:在高频分支对齐预训练 VGG 特征,保证结构纹理保留",
         "对抗损失:让生成结果落在真实工业图像的分布上",
         "证明在频率分离条件下生成器的全局最优解仍可达,理论与训练自洽",
     ]},
    # 13
    {"kind": "content", "title": "Fq-GAN:实验设置",
     "bullets": [
         "数据集:自建 金属铁板 数据集 + 自建 织物 数据集 + 公共 Flowers 数据集",
         "对比方法:HistoGAN(CVPR'21)、DW-GAN(TNNLS'23)、Z*(CVPR'24)等",
         "评价指标:DISTS(结构)、Color Hellinger(颜色)、用户主观打分",
         "训练:Adam / lr=2e-4 / 1500K iter / RTX 4090 单卡",
     ]},
    # 14 quantitative table
    {"kind": "image", "title": "Fq-GAN:定量对比与视觉效果", "image": "image20.png",
     "caption": "Fq-GAN 在 DISTS / 颜色距离两个指标上均优于现有方法"},
    # 15 ablation
    {"kind": "image", "title": "Fq-GAN:消融实验", "image": "image22.png",
     "caption": "去掉频率分离 / 去掉颜色损失 / 去掉内容损失分别造成的退化"},

    # ---------------- SortFlow ----------------
    # 16
    {"kind": "content", "title": "五、方法二:SortFlow", "subtitle": "研究动机",
     "bullets": [
         "流模型可逆性强,但在高维 RGB 空间直接建模分布映射复杂、训练慢",
         "颜色分布问题本质是 1-D 的(亮度排序后的分布)",
         "→ 把高维颜色映射 转换成 低维有序分布上的传输过程,降低建模难度",
         "想法:利用 排序 + 流模型 在中间分布上建立可逆映射",
     ]},
    # 17 architecture
    {"kind": "image", "title": "SortFlow:总体架构", "image": "image26.png",
     "caption": "Encoder 把图像编成 NeuralODE 参数,在排序后的低频空间双向变换"},
    # 18 detail
    {"kind": "image", "title": "SortFlow:核心设计", "image": "image27.png",
     "caption": "按亮度排序 → base flow 推到隐空间 → target flow 反推 → scatter 复位"},
    # 19
    {"kind": "content", "title": "SortFlow:训练策略",
     "bullets": [
         "EfficientNet-B6 主干 + NeuralODE 流头",
         "输入图像通过引导滤波分离 低频 / 高频,高频直接残差叠回",
         "训练损失:重建损失 + 流模型可逆性正则",
         "Encoder 训练 750K iter 收敛,推理时只需 8 步流即可完成颜色迁移",
     ]},
    # 20
    {"kind": "content", "title": "SortFlow:实验设置",
     "bullets": [
         "数据集:Unsplash Lite 公共数据集 + 工业铁板 / 织物 数据集",
         "对比方法:HistoGAN、AdaIN、PhotoWCT2、CT 等",
         "指标:DISTS、Color Hellinger、PSNR、推理耗时",
         "硬件:RTX 4090 / 单图 < 2 秒",
     ]},
    # 21 results visual
    {"kind": "image", "title": "SortFlow:定量与定性结果", "image": "image28.png",
     "caption": "颜色对齐稳定,结构未被破坏,显著快于基于扩散的方法"},
    # 22
    {"kind": "image", "title": "SortFlow:消融实验", "image": "image30.png",
     "caption": "无排序 / 无高频残差分支 / 不同流步数 STEPS 的退化对比"},
    # 23
    {"kind": "image", "title": "SortFlow:更多可视化", "image": "image31.png",
     "caption": "不同 target 风格下的 content↔output 对照,色调匹配自然"},

    # ---------------- GHNet ----------------
    # 24
    {"kind": "content", "title": "六、方法三:GHNet", "subtitle": "研究动机",
     "bullets": [
         "传统特征匹配方法严重依赖源 / 目标图像的相似度",
         "源、目标差异大时,颜色失真严重、纹理被破坏",
         "→ 在特征空间中,把 颜色调制 和 高频增强 解耦成独立模块",
         "想法:用 排序 摆脱空间依赖,用 梯度引导 强化结构高频",
     ]},
    # 25 overall arch
    {"kind": "image", "title": "GHNet:总体架构", "image": "image32.png",
     "caption": "VGG Encoder → 多尺度特征 → GHEM + RCMM → Decoder → 输出"},
    # 26 RCMM
    {"kind": "image", "title": "GHNet:RCMM 排序颜色调制模块", "image": "image33.png",
     "caption": "对特征通道按统计量排序后做跨注意力,摆脱空间对齐依赖"},
    # 27 GHEM
    {"kind": "image", "title": "GHNet:GHEM 梯度引导高频增强模块", "image": "image34.png",
     "caption": "用低频引导的梯度作为 attention,在 skip 路径上保留并强化结构"},
    # 28 training
    {"kind": "content", "title": "GHNet:两阶段训练策略",
     "bullets": [
         "阶段一:训练 Decoder + GHEM,固定 Encoder,以重建损失收敛结构通路",
         "阶段二:加入 RCMM,引入颜色直方图损失,联合训练整网",
         "好处:避免颜色与结构同时优化时的冲突,加快收敛、提高保真度",
         "两阶段总训练时间约为传统单阶段方法的 60%",
     ]},
    # 29 exp setup
    {"kind": "content", "title": "GHNet:实验设置",
     "bullets": [
         "数据集:Unsplash Lite + 工业铁板 + 织物",
         "对比方法:WCT2、PhotoNAS、PCA-Net、CT 等经典与最新方法",
         "指标:DISTS、Color Hellinger、推理耗时、用户主观打分",
         "硬件:RTX 4090 单图 < 1 秒",
     ]},
    # 30 quant
    {"kind": "image", "title": "GHNet:定量对比与视觉效果", "image": "image36.png",
     "caption": "GHNet 在结构 / 颜色两个维度上同时达到最优"},
    # 31 ablation
    {"kind": "image", "title": "GHNet:消融实验", "image": "image38.png",
     "caption": "RCMM / GHEM / 两阶段训练 各自贡献的指标提升"},

    # ---------------- 总结 / 展望 ----------------
    # 32
    {"kind": "content", "title": "七、工作总结", "subtitle": "三种方法的共同主线与差异",
     "bullets": [
         "共同主线:在「频率层面」把颜色与结构解耦,分支独立建模",
         "Fq-GAN:GAN 路线,双分支生成器在频率层面解耦",
         "SortFlow:流模型路线,排序后的低频空间做可逆传输",
         "GHNet:特征匹配路线,RCMM 排序颜色调制 + GHEM 梯度引导高频增强",
     ]},
    # 33
    {"kind": "content", "title": "七、工作总结", "subtitle": "主要创新点",
     "bullets": [
         "① 提出 频率分离 + 双分支 颜色迁移范式,缓解颜色/结构耦合问题",
         "② 提出 基于排序的流模型 SortFlow,降低高维分布映射难度",
         "③ 提出 RCMM + GHEM 模块化设计,使特征匹配方法摆脱相似度依赖",
         "④ 三套方法均在工业 + 公共数据集上验证,DISTS / 颜色距离指标均达 SOTA",
     ]},
    # 34
    {"kind": "content", "title": "七、未来研究展望",
     "bullets": [
         "扩展到 视频 颜色迁移:在时间维度上保持帧间一致性",
         "扩展到 多模态 (红外、深度) 工业数据的统一颜色对齐",
         "探索 轻量化部署:模型压缩 / 蒸馏 / 边缘端实时推理",
         "结合 自监督预训练,降低对成对训练数据的依赖",
     ]},

    # ---------------- 学术成果 ----------------
    # 35
    {"kind": "content", "title": "八、学术成果", "subtitle": "已发表 / 投稿论文",
     "bullets": [
         "[论文 1] 第一作者,基于频率分离的颜色迁移方法,投稿中",
         "[论文 2] 第一作者,基于流模型的工业颜色迁移,投稿中",
         "[论文 3] 第一作者,基于特征匹配的工业颜色迁移,投稿中",
         "(根据实际情况填写期刊 / 会议名称与状态)",
     ]},
    # 36
    {"kind": "content", "title": "八、学术成果", "subtitle": "申请发明专利",
     "bullets": [
         "[专利 1] 一种基于频率分离的工业图像颜色迁移方法",
         "[专利 2] 一种基于流模型的颜色风格迁移装置与方法",
         "(根据实际申请情况填写专利号 / 状态)",
     ]},
    # 37
    {"kind": "content", "title": "八、学术成果", "subtitle": "项目参与",
     "bullets": [
         "国家自然科学基金面上项目 / 工业图像智能检测相关课题",
         "校企合作项目:工业现场缺陷图像增强系统研发",
         "开源代码:GitHub 已发布 SortFlow / GHNet 推理示例",
     ]},

    # ---------------- 致谢 / Q&A ----------------
    # 38
    {"kind": "content", "title": "致谢",
     "bullets": [
         "感谢导师 李伟 教授三年来在科研、写作、答辩各环节的悉心指导",
         "感谢实验室同门、师兄师姐在课题讨论、数据采集、实验调试中的帮助",
         "感谢家人在求学路上的理解与支持",
         "感谢答辩委员会各位老师在百忙之中评审本论文",
     ]},
    # 39
    {"kind": "section", "title": "感谢聆听", "subtitle": "THANKS  &  Q  A"},
    # 40 backup
    {"kind": "content", "title": "备用页", "subtitle": "可能问题的预备回答",
     "bullets": [
         "Q: 为什么选频率分离作为统一主线?    A: ...",
         "Q: SortFlow 与基于扩散的最近工作比较如何?    A: ...",
         "Q: GHNet 的 RCMM 与已有的注意力机制有何区别?    A: ...",
         "Q: 工业数据集规模 / 评测协议如何确定的?    A: ...",
     ]},
]

assert len(SLIDES) == 40, f"slide count={len(SLIDES)}"


# ----------------- helpers -----------------
def clear_slides(prs):
    """Remove every existing slide from a Presentation in-place."""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    for sId in slides:
        rId = sId.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        prs.part.drop_rel(rId)
        xml_slides.remove(sId)


def set_run(run, text, size=None, bold=None, color=None, font=FONT):
    run.text = text
    run.font.name = font
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def add_textbox(slide, x, y, w, h, lines, size=20, bold=False, color=GRAY,
                align=PP_ALIGN.LEFT, line_space=1.2):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        para.line_spacing = line_space
        r = para.add_run()
        set_run(r, line, size=size, bold=bold, color=color)
    return tb


def add_accent_bar(slide):
    """Top accent bar in template's blue."""
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 Inches(0), Inches(0),
                                 SLIDE_W, Inches(0.08))
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT
    bar.line.fill.background()
    return bar


def add_page_number(slide, n, total):
    tb = slide.shapes.add_textbox(Inches(11.8), Inches(7.05),
                                  Inches(1.4), Inches(0.35))
    tf = tb.text_frame
    tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    set_run(r, f"{n} / {total}", size=10, color=LIGHT)


def add_picture_fit(slide, image_path, max_w=Inches(11), max_h=Inches(5),
                    top=Inches(1.8)):
    """Insert centered, scaled to fit (max_w, max_h)."""
    from PIL import Image as PILImage
    im = PILImage.open(image_path)
    w_in, h_in = im.size
    ratio = min(max_w / Emu(w_in * 9525), max_h / Emu(h_in * 9525))
    new_w = Emu(int(w_in * 9525 * ratio))
    new_h = Emu(int(h_in * 9525 * ratio))
    x = (SLIDE_W - new_w) // 2
    return slide.shapes.add_picture(str(image_path), x, top, width=new_w, height=new_h)


# ----------------- slide builders -----------------
def build_cover(prs, item):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    add_accent_bar(slide)
    # title big
    add_textbox(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(2.0),
                [item["title"]], size=44, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER, line_space=1.15)
    # subtitle lines
    add_textbox(slide, Inches(0.8), Inches(4.8), Inches(11.7), Inches(2.0),
                item["subtitle"], size=22, color=GRAY,
                align=PP_ALIGN.CENTER, line_space=1.6)
    return slide


def build_toc(prs, item):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_accent_bar(slide)
    add_textbox(slide, Inches(0.7), Inches(0.4), Inches(12), Inches(1.0),
                [item["title"]], size=32, bold=True, color=ACCENT)
    add_textbox(slide, Inches(2.5), Inches(1.8), Inches(9), Inches(5.5),
                item["items"], size=22, color=GRAY, line_space=1.8)
    return slide


def build_content(prs, item):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_accent_bar(slide)
    add_textbox(slide, Inches(0.7), Inches(0.4), Inches(12), Inches(0.9),
                [item["title"]], size=30, bold=True, color=ACCENT)
    y = 1.4
    if item.get("subtitle"):
        add_textbox(slide, Inches(0.9), Inches(y), Inches(11.5), Inches(0.7),
                    [item["subtitle"]], size=22, bold=True, color=GRAY)
        y += 0.9
    bullets = [f"●  {b}" for b in item.get("bullets", [])]
    add_textbox(slide, Inches(1.1), Inches(y), Inches(11.2), Inches(7.5 - y - 0.5),
                bullets, size=20, color=GRAY, line_space=1.6)
    return slide


def build_image(prs, item):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_accent_bar(slide)
    add_textbox(slide, Inches(0.7), Inches(0.4), Inches(12), Inches(0.9),
                [item["title"]], size=30, bold=True, color=ACCENT)
    img_path = IMG_DIR / item["image"]
    if img_path.exists():
        add_picture_fit(slide, img_path, max_w=Inches(11.5), max_h=Inches(5),
                        top=Inches(1.5))
    else:
        add_textbox(slide, Inches(2), Inches(3), Inches(9), Inches(1),
                    [f"[图片缺失: {item['image']}]"], size=16, color=LIGHT,
                    align=PP_ALIGN.CENTER)
    if item.get("caption"):
        add_textbox(slide, Inches(0.7), Inches(6.7), Inches(12), Inches(0.5),
                    [item["caption"]], size=15, color=LIGHT,
                    align=PP_ALIGN.CENTER)
    return slide


def build_section(prs, item):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_accent_bar(slide)
    add_textbox(slide, Inches(0), Inches(2.8), SLIDE_W, Inches(1.5),
                [item["title"]], size=60, bold=True, color=ACCENT,
                align=PP_ALIGN.CENTER)
    if item.get("subtitle"):
        add_textbox(slide, Inches(0), Inches(4.6), SLIDE_W, Inches(1.0),
                    [item["subtitle"]], size=26, color=GRAY,
                    align=PP_ALIGN.CENTER)
    return slide


BUILDERS = {
    "cover":   build_cover,
    "toc":     build_toc,
    "content": build_content,
    "image":   build_image,
    "section": build_section,
}


# ----------------- main -----------------
def main():
    prs = Presentation(TEMPLATE)
    clear_slides(prs)
    print(f"template loaded; slide size {prs.slide_width/914400:.1f}x{prs.slide_height/914400:.1f}")
    total = len(SLIDES)
    for i, item in enumerate(SLIDES, start=1):
        builder = BUILDERS[item["kind"]]
        slide = builder(prs, item)
        # don't add page number on cover (slide 1)
        if i != 1:
            add_page_number(slide, i, total)
        print(f"  [{i:02d}/{total}] {item['kind']:8s}  {item.get('title','')[:40]}")
    out_path = Path(OUT)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    print(f"\nsaved → {out_path}")
    print(f"file size: {out_path.stat().st_size/1024:.1f} KB")


if __name__ == "__main__":
    main()
