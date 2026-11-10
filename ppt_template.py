"""
组会论文汇报 PPT 生成模板
基于 group-meeting-ppt skill 的配色、字体、布局规范

用法：
    python ppt_template.py

依赖：
    pip install python-pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy


# ============================================================
# 颜色常量（来自 skill 第 8.2 节）
# ============================================================
DEEP_BLUE     = RGBColor(0x12, 0x2E, 0x66)
MEDIUM_BLUE   = RGBColor(0x0E, 0x41, 0x9C)
RED           = RGBColor(0xFF, 0x00, 0x00)
ORANGE        = RGBColor(0xE6, 0x7E, 0x22)
DARK_GRAY     = RGBColor(0x33, 0x33, 0x33)
MID_GRAY      = RGBColor(0x7F, 0x7F, 0x7F)
LIGHT_GRAY    = RGBColor(0xD9, 0xD9, 0xD9)
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
BG_LIGHT_BLUE = RGBColor(0xEE, 0xF2, 0xF8)
BG_LIGHT_ORANGE = RGBColor(0xFD, 0xF0, 0xE3)
BG_LIGHT_GREEN  = RGBColor(0xE8, 0xF5, 0xE9)
BG_LIGHT_GRAY   = RGBColor(0xF2, 0xF2, 0xF2)

# 字体
FONT_CN_BODY  = "Noto Sans SC"
FONT_CN_TITLE = "微软雅黑"
FONT_EN_BODY  = "Arial"
FONT_EN_SERIF = "Times New Roman"
FONT_IMPACT   = "Impact"
FONT_CODE     = "Consolas"


# ============================================================
# 初始化演示文稿
# ============================================================
import os

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "模板.pptx")

def create_presentation(template_path=None):
    """传入路径时打开模板，否则新建空白演示文稿。模板原有页面和备注会保留。"""
    if template_path is not None:
        return Presentation(template_path)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


def clone_slide(prs, source_slide):
    """复制无外部关系的形状。含图片、图表、超链接或 OLE 的页面不可用。"""
    if any(not rel.reltype.endswith("/slideLayout")
           for rel in source_slide.part.rels.values()):
        raise ValueError("无法安全复制带关系的页面，请使用支持完整页面复制的工具")
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    for shape in list(new_slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(new_el)
    return new_slide


def add_picture_contained(slide, path, left, top, width, height):
    """按原图比例居中放入指定区域，不裁剪或拉伸。"""
    from pptx.parts.image import Image
    image_width, image_height = Image.from_file(path).size
    scale = min(width / image_width, height / image_height)
    placed_width = int(image_width * scale)
    placed_height = int(image_height * scale)
    return slide.shapes.add_picture(
        path,
        int(left + (width - placed_width) / 2),
        int(top + (height - placed_height) / 2),
        width=placed_width,
        height=placed_height,
    )


def _clear_slides(prs):
    """删除演示文稿中的所有页面"""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    for slide in slides:
        xml_slides.remove(slide)


# ============================================================
# 页面装饰：顶部蓝色条 + 水印
# ============================================================
def add_page_decorations(slide):
    """给空白页添加近似模板的装饰，不复制模板母版或资源。"""
    # 1. 顶部蓝色窄条（#0E419C）
    top_strip = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Emu(-12700), Emu(-12700),
        Emu(12217400), Emu(622300)
    )
    top_strip.fill.solid()
    top_strip.fill.fore_color.rgb = MEDIUM_BLUE  # #0E419C
    top_strip.line.fill.background()

    # 2. 上方蓝色装饰带（#122E66）
    blue_band = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Emu(0), Emu(0),
        Emu(12192000), Emu(3441700)
    )
    blue_band.fill.solid()
    blue_band.fill.fore_color.rgb = DEEP_BLUE  # #122E66
    blue_band.line.fill.background()

    # 3. 白色主体底板
    content_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Emu(292100), Emu(355600),
        Emu(11607800), Emu(6337300)
    )
    content_bg.fill.solid()
    content_bg.fill.fore_color.rgb = WHITE
    content_bg.line.fill.background()

    # 4. 左侧装饰线
    deco_line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Emu(425341), Emu(768237),
        Emu(355827), Emu(12700)
    )
    deco_line.fill.solid()
    deco_line.fill.fore_color.rgb = DEEP_BLUE
    deco_line.line.fill.background()

    # 5. 右下角水印
    watermark = slide.shapes.add_textbox(
        Emu(0), Emu(-12700),
        Emu(279400), Emu(279400)
    )
    tf = watermark.text_frame
    p = tf.paragraphs[0]
    p.text = "木卫林"
    p.font.name = FONT_CN_BODY
    p.font.size = Pt(4)
    p.font.color.rgb = WHITE


# ============================================================
# 封面页
# ============================================================
def add_cover_page(prs, cn_title, en_title, university, presenter, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # BLANK
    add_page_decorations(slide)
    _add_corner_decoration(slide)

    # 中文标题
    title_box = slide.shapes.add_textbox(
        Inches(1.14), Inches(1.84),
        Inches(10.15), Inches(1.27)
    )
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = cn_title
    p.font.name = FONT_CN_TITLE
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    p.alignment = PP_ALIGN.CENTER

    # 英文标题
    en_box = slide.shapes.add_textbox(
        Inches(1.14), Inches(3.66),
        Inches(10.15), Inches(0.76)
    )
    tf = en_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = en_title
    p.font.name = FONT_EN_BODY
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    p.alignment = PP_ALIGN.CENTER

    # 学校 + 汇报人
    info_box = slide.shapes.add_textbox(
        Inches(4.19), Inches(4.56),
        Inches(3.81), Inches(1.2)
    )
    tf = info_box.text_frame
    tf.word_wrap = True

    p1 = tf.paragraphs[0]
    p1.text = university
    p1.font.name = FONT_CN_BODY
    p1.font.size = Pt(18)
    p1.font.bold = True
    p1.font.color.rgb = DEEP_BLUE
    p1.alignment = PP_ALIGN.CENTER

    p2 = tf.add_paragraph()
    p2.text = f"汇报人：{presenter}"
    p2.font.name = FONT_CN_BODY
    p2.font.size = Pt(16)
    p2.font.color.rgb = DARK_GRAY
    p2.alignment = PP_ALIGN.CENTER

    p3 = tf.add_paragraph()
    p3.text = date_str
    p3.font.name = FONT_CN_BODY
    p3.font.size = Pt(14)
    p3.font.color.rgb = MID_GRAY
    p3.alignment = PP_ALIGN.CENTER


# ============================================================
# 章节分隔页
# ============================================================
def add_section_divider(prs, number, title):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)

    # 左侧装饰线
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Emu(539641), Emu(1022237),
        Emu(355827), Emu(12700)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = DEEP_BLUE
    line.line.fill.background()

    # 大号灰色数字
    num_box = slide.shapes.add_textbox(
        Inches(1.96), Inches(2.08),
        Inches(2.86), Inches(2.64)
    )
    tf = num_box.text_frame
    p = tf.paragraphs[0]
    p.text = f"{number:02d}"
    p.font.name = FONT_IMPACT
    p.font.size = Pt(166)
    p.font.color.rgb = LIGHT_GRAY
    p.alignment = PP_ALIGN.CENTER

    # 章节名
    title_box = slide.shapes.add_textbox(
        Inches(4.45), Inches(2.84),
        Inches(7.62), Inches(1.02)
    )
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = FONT_CN_BODY
    p.font.size = Pt(60)
    p.font.bold = True
    p.font.color.rgb = DEEP_BLUE


# ============================================================
# 内容页：标题条 + 正文区域
# ============================================================
def add_content_page(prs, title_text, body_text="", notes_text=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)

    # 左侧装饰线
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Emu(425341), Emu(768237),
        Emu(355827), Emu(12700)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = DEEP_BLUE
    line.line.fill.background()

    # 标题
    header = slide.shapes.add_textbox(
        Inches(0.72), Inches(0.56),
        Inches(3.70), Inches(0.43)
    )
    tf = header.text_frame
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = FONT_CN_TITLE
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x00, 0x15, 0x6C)

    # 正文区域
    if body_text:
        body = slide.shapes.add_textbox(
            Inches(0.61), Inches(1.52),
            Inches(9.48), Inches(5.0)
        )
        tf = body.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = body_text
        p.font.name = FONT_CN_BODY
        p.font.size = Pt(18)
        p.font.color.rgb = DARK_GRAY
        p.alignment = PP_ALIGN.LEFT

    # 备注
    if notes_text:
        add_notes(slide, notes_text)

    return slide


# ============================================================
# 标题条（页面顶部的深蓝色横条）
# ============================================================
def add_title_bar(slide, text, top=Inches(0.67), left=Inches(0.52),
                  width=Inches(11.6), height=Inches(0.55)):
    bar = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left, top, width, height
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = DEEP_BLUE
    bar.line.fill.background()

    tf = bar.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.15)
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = FONT_CN_BODY
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    return bar


# ============================================================
# 信息块（浅底圆角框）
# ============================================================
def add_info_block(slide, text, left, top, width, height,
                   fill_color=BG_LIGHT_BLUE, font_size=Pt(16),
                   font_color=DARK_GRAY, bold=False):
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left, top, width, height
    )
    box.fill.solid()
    box.fill.fore_color.rgb = fill_color
    box.line.fill.background()

    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.12)
    tf.margin_right = Inches(0.12)
    tf.margin_top = Inches(0.08)
    tf.margin_bottom = Inches(0.08)

    p = tf.paragraphs[0]
    p.text = text
    p.font.name = FONT_CN_BODY
    p.font.size = font_size
    p.font.color.rgb = font_color
    p.font.bold = bold
    p.alignment = PP_ALIGN.CENTER
    return box


# ============================================================
# 标签按钮（深蓝底白字小方块）
# ============================================================
def add_label_button(slide, text, left, top, width=Inches(2.1), height=Inches(0.5)):
    btn = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left, top, width, height
    )
    btn.fill.solid()
    btn.fill.fore_color.rgb = MEDIUM_BLUE
    btn.line.fill.background()

    tf = btn.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = FONT_CN_BODY
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    return btn


# ============================================================
# 文献综述页
# ============================================================
def add_literature_page(prs, ref_number, ref_title, content, notes_text=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)

    # 文献标题
    title_box = slide.shapes.add_textbox(
        Inches(0.61), Inches(0.67),
        Inches(12.0), Inches(1.0)
    )
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = f"文献[{ref_number}] "
    run.font.name = FONT_CN_BODY
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = DARK_GRAY

    run2 = p.add_run()
    run2.text = ref_title
    run2.font.name = FONT_CN_BODY
    run2.font.size = Pt(18)
    run2.font.color.rgb = DARK_GRAY

    # 正文
    body_box = slide.shapes.add_textbox(
        Inches(0.71), Inches(1.52),
        Inches(10.74), Inches(5.0)
    )
    tf = body_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = content
    p.font.name = FONT_CN_BODY
    p.font.size = Pt(16)
    p.font.color.rgb = DARK_GRAY
    p.line_spacing = Pt(24)

    if notes_text:
        add_notes(slide, notes_text)

    return slide


# ============================================================
# 备注
# ============================================================
def add_notes(slide, notes_text):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text


# ============================================================
# 对比表格页
# ============================================================
def add_comparison_table(prs, title, headers, rows, notes_text=""):
    """
    headers: list of str
    rows: list of list of str
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)

    # 标题条
    add_title_bar(slide, title)

    n_cols = len(headers)
    n_rows = len(rows) + 1  # +1 for header
    col_width = Inches(11.0 / n_cols)
    row_height = Inches(0.55)
    table_left = Inches(1.1)
    table_top = Inches(1.6)

    table = slide.shapes.add_table(
        n_rows, n_cols,
        table_left, table_top,
        Inches(11.0), row_height * n_rows
    ).table

    # 设置表头
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = DEEP_BLUE
        for p in cell.text_frame.paragraphs:
            p.font.name = FONT_CN_BODY
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

    # 设置数据行
    for i, row in enumerate(rows):
        bg = BG_LIGHT_ORANGE if i % 2 == 0 else WHITE
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            for p in cell.text_frame.paragraphs:
                p.font.name = FONT_CN_BODY
                p.font.size = Pt(13)
                p.font.color.rgb = DARK_GRAY
                p.alignment = PP_ALIGN.CENTER

    if notes_text:
        add_notes(slide, notes_text)

    return slide


# ============================================================
# 相关工作分类综述页（每页一类，含多篇文献）
# ============================================================
def add_related_work_page(prs, category_name, papers, notes_text=""):
    """
    相关工作分类综述页：一个类别 + 2~3 篇文献

    category_name: 类别名称，如 "数据驱动模型架构迭代"
    papers: list of dict，每篇含 ref_number, title, content
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)

    # 类别标签（页面顶部）
    cat_label = slide.shapes.add_textbox(
        Inches(0.61), Inches(0.55),
        Inches(11.0), Inches(0.45)
    )
    tf = cat_label.text_frame
    p = tf.paragraphs[0]
    p.text = category_name
    p.font.name = FONT_CN_TITLE
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = DEEP_BLUE

    # 逐篇文献
    y_offset = Inches(1.2)
    for paper in papers:
        # 文献标题
        title_box = slide.shapes.add_textbox(
            Inches(0.61), y_offset,
            Inches(12.0), Inches(0.6)
        )
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = f"文献[{paper['ref_number']}] "
        run.font.name = FONT_CN_BODY
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = DARK_GRAY

        run2 = p.add_run()
        run2.text = paper['title']
        run2.font.name = FONT_CN_BODY
        run2.font.size = Pt(16)
        run2.font.color.rgb = DARK_GRAY

        y_offset += Inches(0.55)

        # 文献正文
        body_box = slide.shapes.add_textbox(
            Inches(0.71), y_offset,
            Inches(11.5), Inches(1.5)
        )
        tf = body_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = paper['content']
        p.font.name = FONT_CN_BODY
        p.font.size = Pt(14)
        p.font.color.rgb = DARK_GRAY
        p.line_spacing = 1.5

        y_offset += Inches(1.6)

    if notes_text:
        add_notes(slide, notes_text)
    return slide


# ============================================================
# 相关工作总结表格页
# ============================================================
def add_related_work_summary(prs, title, categories, conclusion, notes_text=""):
    """
    相关工作总结表格页：分类汇总 + 底部本文定位

    title: 标题条文字，如 "相关工作总结"
    categories: list of dict，每类含 name, refs, strength, weakness
    conclusion: 底部一句话总结本文定位（红色加粗）
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)
    add_title_bar(slide, title)

    n_rows = len(categories) + 1  # +1 for header
    headers = ["类别", "代表文献", "优点", "不足"]
    col_widths = [Inches(2.2), Inches(2.0), Inches(3.8), Inches(3.8)]
    row_height = Inches(0.65)
    table_left = Inches(0.52)
    table_top = Inches(1.5)

    total_width = sum(w for w in col_widths)
    table = slide.shapes.add_table(
        n_rows, 4,
        table_left, table_top,
        total_width, row_height * n_rows
    ).table

    # 设置列宽
    for j, w in enumerate(col_widths):
        for i in range(n_rows):
            table.cell(i, j).width = w

    # 表头
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = DEEP_BLUE
        for p in cell.text_frame.paragraphs:
            p.font.name = FONT_CN_BODY
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

    # 数据行
    for i, cat in enumerate(categories):
        bg = WHITE if i % 2 == 0 else BG_LIGHT_BLUE
        vals = [cat['name'], cat['refs'], cat['strength'], cat['weakness']]
        for j, val in enumerate(vals):
            cell = table.cell(i + 1, j)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            for p in cell.text_frame.paragraphs:
                p.font.name = FONT_CN_BODY
                p.font.size = Pt(12)
                p.font.color.rgb = DARK_GRAY
                p.alignment = PP_ALIGN.LEFT if j >= 2 else PP_ALIGN.CENTER
                # 「不足」列标红
                if j == 3:
                    p.font.color.rgb = RED

    # 底部本文定位
    concl_top = table_top + row_height * n_rows + Inches(0.2)
    concl = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        table_left, concl_top,
        total_width, Inches(0.55)
    )
    concl.fill.solid()
    concl.fill.fore_color.rgb = BG_LIGHT_ORANGE
    concl.line.fill.background()
    tf = concl.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.15)
    p = tf.paragraphs[0]
    p.text = f"本文：{conclusion}"
    p.font.name = FONT_CN_BODY
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = RED
    p.alignment = PP_ALIGN.LEFT

    if notes_text:
        add_notes(slide, notes_text)
    return slide


# ============================================================
# 结尾页
# ============================================================
def add_ending_page(prs, text="谢谢大家", en_text="THANKS FOR LISTENING"):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)
    _add_corner_decoration(slide)

    # 中文
    box = slide.shapes.add_textbox(
        Inches(2.4), Inches(2.6),
        Inches(8.7), Inches(1.1)
    )
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = FONT_CN_BODY
    p.font.size = Pt(60)
    p.font.bold = True
    p.font.color.rgb = DEEP_BLUE
    p.alignment = PP_ALIGN.CENTER

    # 英文
    en_box = slide.shapes.add_textbox(
        Inches(2.5), Inches(4.1),
        Inches(8.3), Inches(0.6)
    )
    tf = en_box.text_frame
    p = tf.paragraphs[0]
    p.text = en_text
    p.font.name = FONT_EN_BODY
    p.font.size = Pt(14)
    p.font.color.rgb = DEEP_BLUE
    p.alignment = PP_ALIGN.CENTER


def _add_corner_decoration(slide):
    """封面页和结尾页的右上角装饰方块 + 白色自由形状"""
    # 深蓝方块
    corner = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(6.0), Inches(0),
        Inches(1.3), Inches(1.3)
    )
    corner.fill.solid()
    corner.fill.fore_color.rgb = DEEP_BLUE
    corner.line.fill.background()


# ============================================================
# 问题背景页（带渐变标签按钮）
# ============================================================
def add_background_page(prs, background_text, innovation_text, notes_text=""):
    """P2 问题背景页：左侧渐变标签 + 右侧文本"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)

    # 「问题背景」标签
    _add_gradient_label(slide, "问题背景",
                        left=Inches(0.61), top=Inches(1.52),
                        width=Inches(3.31), height=Inches(0.5))
    # 「创新点」标签
    _add_gradient_label(slide, "创新点",
                        left=Inches(8.59), top=Inches(4.13),
                        width=Inches(3.31), height=Inches(0.5))

    # 背景文本
    bg_box = slide.shapes.add_textbox(
        Inches(2.48), Inches(0.95),
        Inches(8.74), Inches(2.13)
    )
    tf = bg_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = background_text
    p.font.name = FONT_CN_BODY
    p.font.size = Pt(18)
    p.font.color.rgb = DARK_GRAY
    p.line_spacing = 1.5

    # 创新点文本
    inno_box = slide.shapes.add_textbox(
        Inches(0.66), Inches(3.43),
        Inches(9.48), Inches(2.83)
    )
    tf = inno_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = innovation_text
    p.font.name = FONT_CN_BODY
    p.font.size = Pt(18)
    p.font.color.rgb = DARK_GRAY
    p.line_spacing = 1.5

    if notes_text:
        add_notes(slide, notes_text)
    return slide


def _add_gradient_label(slide, text, left, top, width, height):
    """中蓝渐变圆角标签（问题背景/创新点 按钮样式）"""
    from pptx.oxml.ns import qn as _qn
    btn = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    btn.fill.solid()
    btn.fill.fore_color.rgb = MEDIUM_BLUE
    btn.line.fill.background()

    # 添加渐变效果（XML 级操作）
    spPr = btn._element.find(_qn('p:spPr'))
    if spPr is not None:
        solidFill = spPr.find(_qn('a:solidFill'))
        if solidFill is not None:
            spPr.remove(solidFill)
        gradFill = spPr.makeelement(_qn('a:gradFill'), {'flip': 'none', 'rotWithShape': '1'})
        gsLst = spPr.makeelement(_qn('a:gsLst'), {})
        for pos in ['0', '100000']:
            gs = spPr.makeelement(_qn('a:gs'), {'pos': pos})
            srgb = spPr.makeelement(_qn('a:srgbClr'), {'val': '0E419C'})
            gs.append(srgb)
            gsLst.append(gs)
        gradFill.append(gsLst)
        lin = spPr.makeelement(_qn('a:lin'), {'ang': '5400000', 'scaled': '1'})
        gradFill.append(lin)
        # 插入到 prstGeom 之后
        prstGeom = spPr.find(_qn('a:prstGeom'))
        if prstGeom is not None:
            prstGeom.addnext(gradFill)
        else:
            spPr.append(gradFill)

    tf = btn.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = FONT_CN_BODY
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    return btn


# ============================================================
# 模块导览页（模块标题 + 子节目录）
# ============================================================
def add_module_overview_page(prs, module_label, module_title,
                             sub_sections, framework_image=None, notes_text=""):
    """
    模块导览页：标题条 + 子节列表 + 可选框架图回指

    module_label: 如 "模块 ①"
    module_title: 如 "神经网络给出每个候选的概率"
    sub_sections: list of str，如 ["1.1 学会下一步的概率", "1.2 神经网络结构"]
    framework_image: 可选，框架图路径（用红框标出当前模块）
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)
    add_title_bar(slide, f"{module_label}：{module_title}")

    # 框架图回指（如果有）
    if framework_image:
        add_picture_contained(
            slide,
            framework_image,
            Inches(2.5), Inches(1.4),
            Inches(8.4), Inches(3.7)
        )
        sub_top = Inches(5.3)
    else:
        sub_top = Inches(2.0)

    # 引导语
    guide = slide.shapes.add_textbox(
        Inches(0.52), sub_top,
        Inches(11.6), Inches(0.5)
    )
    tf = guide.text_frame
    p = tf.paragraphs[0]
    p.text = f"{module_label}能给出什么？" if '什么' not in module_title else module_title
    p.font.name = FONT_CN_BODY
    p.font.size = Pt(16)
    p.font.color.rgb = DARK_GRAY
    p.alignment = PP_ALIGN.CENTER

    # 子节标签（横向排列）
    n = len(sub_sections)
    if n > 0:
        total_w = min(n * 2.9, 11.0)
        start_x = (13.333 - total_w) / 2
        for idx, sec in enumerate(sub_sections):
            add_label_button(
                slide, sec,
                left=Inches(start_x + idx * (total_w / n)),
                top=sub_top + Inches(0.6),
                width=Inches(total_w / n - 0.15),
                height=Inches(0.5)
            )

    if notes_text:
        add_notes(slide, notes_text)
    return slide


# ============================================================
# 子节内容页（带 N.M 编号标题）
# ============================================================
def add_subsection_page(prs, section_number, title, body_text="",
                        image_path=None, notes_text=""):
    """
    子节内容页：标题条写 N.M 标题 + 正文/图片

    section_number: 如 "2.3"
    title: 如 "约束怎么写成公式：以"关闭后终将接受"为例"
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)
    add_title_bar(slide, f"{section_number} {title}")

    if image_path:
        add_picture_contained(
            slide,
            image_path,
            Inches(0.8), Inches(1.5),
            Inches(8.4), Inches(3.7)
        )
        if body_text:
            body = slide.shapes.add_textbox(
                Inches(0.52), Inches(5.5),
                Inches(11.6), Inches(1.5)
            )
            tf = body.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = body_text
            p.font.name = FONT_CN_BODY
            p.font.size = Pt(16)
            p.font.color.rgb = DARK_GRAY
    elif body_text:
        body = slide.shapes.add_textbox(
            Inches(0.61), Inches(1.52),
            Inches(11.6), Inches(5.0)
        )
        tf = body.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = body_text
        p.font.name = FONT_CN_BODY
        p.font.size = Pt(18)
        p.font.color.rgb = DARK_GRAY
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.5

    if notes_text:
        add_notes(slide, notes_text)
    return slide


# ============================================================
# 图片内容页（标题条 + 居中图片 + 底部说明）
# ============================================================
def add_image_content_page(prs, title, image_path, caption="",
                           img_width=Inches(8.4), img_height=Inches(3.7),
                           notes_text=""):
    """
    标题条 + 大图 + 底部一行说明
    适用于：框架图、实验结果图、算法伪代码截图
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)
    add_title_bar(slide, title)

    # 居中图片
    img_left = (13.333 - img_width / 914400) / 2
    add_picture_contained(
        slide,
        image_path,
        Inches(img_left), Inches(1.5),
        img_width, img_height
    )

    # 底部说明
    if caption:
        cap = slide.shapes.add_textbox(
            Inches(0.52), Inches(5.5),
            Inches(11.6), Inches(1.0)
        )
        tf = cap.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = caption
        p.font.name = FONT_CN_BODY
        p.font.size = Pt(16)
        p.font.color.rgb = DARK_GRAY
        p.alignment = PP_ALIGN.CENTER

    if notes_text:
        add_notes(slide, notes_text)
    return slide


# ============================================================
# 连接符 / 箭头
# ============================================================
def add_connector(slide, x1, y1, x2, y2, color=DEEP_BLUE,
                  width_pt=2.0, arrow_end=True):
    """绘制连接符（带箭头）"""
    from pptx.oxml.ns import qn as _qn
    conn = slide.shapes.add_connector(
        1,  # straight connector
        x1, y1, x2, y2
    )
    conn.line.color.rgb = color
    conn.line.width = Pt(width_pt)

    if arrow_end:
        ln = conn._element.find(_qn('p:spPr'))
        if ln is None:
            ln = conn._element.find(_qn('p:style'))
        # 在 XML 中添加箭头
        spPr = conn._element.find('.//' + _qn('a:ln'))
        if spPr is not None:
            tail = spPr.makeelement(_qn('a:tailEnd'), {
                'type': 'triangle', 'w': 'med', 'len': 'med'
            })
            spPr.append(tail)
    return conn


# ============================================================
# 公式内容页（Unicode 数学符号）
# ============================================================
def add_formula_page(prs, title, formula_lines, symbol_table=None, notes_text=""):
    """
    公式 + 符号解释页

    formula_lines: list of str，每行一个公式（Unicode 数学符号）
    symbol_table: list of (symbol, meaning) 元组
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)
    add_title_bar(slide, title)

    # 公式区域
    formula_box = slide.shapes.add_textbox(
        Inches(1.5), Inches(1.6),
        Inches(10.0), Inches(2.5)
    )
    tf = formula_box.text_frame
    tf.word_wrap = True
    for idx, line in enumerate(formula_lines):
        if idx == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.name = FONT_EN_BODY
        p.font.size = Pt(20)
        p.font.color.rgb = DARK_GRAY
        p.alignment = PP_ALIGN.CENTER
        p.space_after = Pt(12)

    # 符号速查表
    if symbol_table:
        tbl_top = Inches(4.5)
        tbl = slide.shapes.add_table(
            len(symbol_table) + 1, 2,
            Inches(3.0), tbl_top,
            Inches(7.0), Inches(0.4) * (len(symbol_table) + 1)
        ).table

        tbl.cell(0, 0).text = "符号"
        tbl.cell(0, 1).text = "含义"
        for j in range(2):
            cell = tbl.cell(0, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = DEEP_BLUE
            for p in cell.text_frame.paragraphs:
                p.font.name = FONT_CN_BODY
                p.font.size = Pt(14)
                p.font.bold = True
                p.font.color.rgb = WHITE
                p.alignment = PP_ALIGN.CENTER

        for i, (sym, meaning) in enumerate(symbol_table):
            tbl.cell(i + 1, 0).text = sym
            tbl.cell(i + 1, 1).text = meaning
            for j in range(2):
                cell = tbl.cell(i + 1, j)
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if i % 2 == 0 else BG_LIGHT_BLUE
                for p in cell.text_frame.paragraphs:
                    p.font.name = FONT_CN_BODY if j == 1 else FONT_EN_BODY
                    p.font.size = Pt(14)
                    p.font.color.rgb = DARK_GRAY
                    p.alignment = PP_ALIGN.CENTER

    if notes_text:
        add_notes(slide, notes_text)
    return slide


# ============================================================
# 串联对比页（两篇论文左右对比）
# ============================================================
def add_paper_comparison_page(prs, title, left_paper, right_paper, notes_text=""):
    """
    两篇论文对比页

    left_paper / right_paper: dict with keys:
        name, features (list of str), image_path (optional)
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)
    add_title_bar(slide, title)

    for side, paper, x_offset in [('left', left_paper, 0.52), ('right', right_paper, 6.8)]:
        # 论文名
        name_box = slide.shapes.add_textbox(
            Inches(x_offset), Inches(1.5),
            Inches(5.5), Inches(0.5)
        )
        tf = name_box.text_frame
        p = tf.paragraphs[0]
        p.text = paper['name']
        p.font.name = FONT_CN_BODY
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = DEEP_BLUE
        p.alignment = PP_ALIGN.CENTER

        # 特征列表
        feat_box = slide.shapes.add_textbox(
            Inches(x_offset), Inches(2.2),
            Inches(5.5), Inches(1.5)
        )
        tf = feat_box.text_frame
        tf.word_wrap = True
        for idx, feat in enumerate(paper.get('features', [])):
            if idx == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = f"• {feat}"
            p.font.name = FONT_CN_BODY
            p.font.size = Pt(14)
            p.font.color.rgb = DARK_GRAY

        # 缩略图
        if paper.get('image_path'):
            add_picture_contained(
                slide,
                paper['image_path'],
                Inches(x_offset + 0.5), Inches(4.0),
                Inches(4.7), Inches(2.0)
            )

    if notes_text:
        add_notes(slide, notes_text)
    return slide


# ============================================================
# 演示：生成一份样例 PPT
# ============================================================
def demo():
    prs = create_presentation()

    # 1. 封面
    add_cover_page(
        prs,
        cn_title="基于概率声明式知识的神经符号预测过程监控",
        en_title="Neuro-Symbolic Predictive Process Monitoring\nwith Probabilistic Declarative Knowledge",
        university="安徽理工大学",
        presenter="XXX",
        date_str="2026-10-30"
    )

    # 2. 问题背景
    slide = add_content_page(
        prs,
        title_text="问题背景",
        body_text=(
            "神经符号预测过程监控旨在利用关于过程的符号知识，"
            "来改进从事件日志中生成的预测模型。"
        ),
        notes_text=(
            "本文发表于 ACM Transactions on Intelligent Systems and Technology。"
            "论文围绕业务流程管理中的后缀预测问题，"
            "将有限迹线性时序逻辑与深度自回归模型结合。"
        )
    )

    # 3. 章节分隔：01 引言
    add_section_divider(prs, 1, "引言")

    # 4. 贡献列表
    add_content_page(
        prs,
        title_text="本文主要贡献",
        body_text=(
            "1. 提出首个面向 PPM 的概率 BK NeSy 方案\n"
            "2. 将 ProbDeclare 模型融入束搜索\n"
            "3. 不重训神经网络，在预测时用知识修正概率"
        ),
        notes_text="KB-Contextualization 是论文给算法 1 取的名字。"
    )

    # 5. 章节分隔：02 相关工作
    add_section_divider(prs, 2, "相关工作")

    # 6. 相关工作分类综述
    add_related_work_page(
        prs,
        category_name="一、数据驱动模型架构迭代",
        papers=[
            {
                'ref_number': 1,
                'title': '基于LSTM的预测业务过程监控',
                'content': '该文研究如何仅依赖事件日志实现运行中下一活动预测，核心是学习流程执行的时序依赖与序列中的行为模式。方法上将轨迹编码为序列输入，训练LSTM输出下一活动的概率分布。'
            },
            {
                'ref_number': 2,
                'title': 'GCNN与KVP注意力用于下一事件预测',
                'content': '该文关注下一事件预测中局部模式与关键上下文的建模问题，提出卷积与注意力相结合的替代路线。'
            },
            {
                'ref_number': 3,
                'title': 'SuTraN：编码器—解码器Transformer后缀预测',
                'content': '该文面向后缀预测任务，采用编码器—解码器Transformer，以注意力机制构建全上下文表示。'
            },
        ],
        notes_text="这一类都是纯数据驱动，优点是预测准，但知识变化时跟不上。"
    )

    # 7. 相关工作总结表格
    add_related_work_summary(
        prs,
        title="相关工作总结",
        categories=[
            {'name': '数据驱动模型', 'refs': '[1][2][3]', 'strength': '预测精度高，基线扎实', 'weakness': '仅依赖历史日志，无法利用领域知识'},
            {'name': '知识注入方法', 'refs': '[4][5][6]', 'strength': '能引入Declare约束', 'weakness': '约束是硬性的，不允许例外'},
            {'name': '声明式知识与合规', 'refs': '[7][8][9]', 'strength': '知识可计算可验证', 'weakness': '主要用于事后判定，不能直接生成预测'},
            {'name': '概率声明式知识', 'refs': '[10][11][12]', 'strength': '能表达"约80%满足"', 'weakness': '与在线后缀生成结合不足'},
        ],
        conclusion='把概率声明式知识和后缀预测在生成阶段深度耦合，实现从"事后过滤"到"过程引导"的转变。',
        notes_text=(
            "以往研究围绕四个方向：数据驱动模型预测准但只靠日志；"
            "知识注入能用约束但是硬性的；声明式知识可验证但只能事后判定；"
            "概率知识能表达比例但和在线生成结合不足。"
            "本文的缺口就是把概率知识和后缀预测在生成阶段耦合。"
        )
    )

    # 7. 章节分隔：03 方法
    add_section_divider(prs, 3, "方法")

    # 8. 方法内容页（带标题条和信息块）
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_decorations(slide)
    add_title_bar(slide, "总体框架：两个离线模块 + 两个在线模块")

    # 四个模块标签
    modules = [
        ("① 神经网络（离线）", BG_LIGHT_BLUE),
        ("② 概率背景知识（离线）", BG_LIGHT_ORANGE),
        ("③ 修正打分（在线）", BG_LIGHT_GREEN),
        ("④ 束搜索（在线总控）", BG_LIGHT_BLUE),
    ]
    for idx, (label, color) in enumerate(modules):
        add_info_block(
            slide, label,
            left=Inches(1.0 + idx * 2.9),
            top=Inches(2.5),
            width=Inches(2.6),
            height=Inches(1.0),
            fill_color=color,
            font_size=Pt(14),
            bold=True
        )

    add_notes(slide, (
        "底部这行：离线是预测前只做一次，在线是每预测一步都要做。"
        "讲解顺序是 ① → ② → ③ → ④。"
    ))

    # 9. 对比表格
    add_comparison_table(
        prs,
        title="本文与以往方法对比",
        headers=["维度", "以往", "本文"],
        rows=[
            ["知识形式", "硬性约束", "概率约束 ⟨φ, p⟩"],
            ["使用时机", "生成后过滤", "束搜索每步修正"],
            ["判断结果", "对 / 错", "0~1 合规度"],
        ],
        notes_text="概率全取 1，本文方法退化为以往做法。"
    )

    # 10. 模块导览页
    add_module_overview_page(
        prs,
        module_label="模块 ②",
        module_title="把约束变成可计算的形式",
        sub_sections=[
            "2.1–2.3 约束与公式",
            "2.4–2.5 场景",
            "2.6–2.7 场景概率",
            "2.8–2.10 自动机",
        ],
        notes_text="这一部分分四段：2.1 到 2.3 讲约束和公式；2.4、2.5 讲场景。"
    )

    # 11. 子节内容页
    add_subsection_page(
        prs,
        section_number="2.2",
        title="约束怎么写成公式：以「关闭后终将接受」为例",
        body_text=(
            "每条约束都写成一条时序逻辑公式，叫 LTLf。\n"
            "举一条为例：响应(a, b)，公式是 G(a → F b)。\n"
            "G 是「从现在起一直」；a → 是「只要出现 a」；F b 是「之后某一步出现 b」。"
        ),
        notes_text="公式就是判断满足还是违反的标准，后面还能自动转成自动机。"
    )

    # 12. 公式页
    add_formula_page(
        prs,
        title="3.3 融合公式里的每个符号是什么",
        formula_lines=[
            "score' = score × NN^(1−w) × C^w",
            "score'(b) = 1 × 0.30^0.5 × 0.70^0.5 = 0.458",
            "score'(c) = 1 × 0.55^0.5 × 0.20^0.5 = 0.332",
        ],
        symbol_table=[
            ("NN", "神经网络给的原始概率"),
            ("C(a|σ)", "合规度（知识给的分数）"),
            ("w", "知识权重，越大越听知识"),
            ("score'", "融合后的分数，非归一化概率"),
        ],
        notes_text="教学示意：w = 0.5 时取几何平均，结果为分数而非归一化概率。"
    )

    # 13. 串联对比页
    add_paper_comparison_page(
        prs,
        title="本篇论文 vs 拓展论文",
        left_paper={
            'name': 'ProbDeclare NeSy（主讲）',
            'features': ['仅活动 a', 'ProbDeclare（硬约束 + 概率约束）'],
        },
        right_paper={
            'name': 'Multi-attribute NeSy（拓展）',
            'features': ['活动 a + 资源 r', 'MP-Declare（确定性约束）'],
        },
        notes_text="拓展论文把知识从单属性扩展到多属性。"
    )

    # 14. 章节分隔：04 实验
    add_section_divider(prs, 4, "实验")

    # 15. 章节分隔：05 结论
    add_section_divider(prs, 5, "结论")

    # 16. 结尾
    add_ending_page(prs)

    output_path = "demo_output.pptx"
    prs.save(output_path)
    print(f"已生成演示 PPT: {output_path}")
    print(f"共 {len(prs.slides)} 页")


if __name__ == "__main__":
    demo()
