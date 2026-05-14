"""将 config.py 中的样式配置应用到 python-docx 文档对象。"""

from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from config import PAGE_CONFIG, FONT_CONFIG, TABLE_CONFIG


def configure_page(document):
    """根据 PAGE_CONFIG 设置文档页面尺寸和边距。"""
    section = document.sections[0]
    section.page_width = Cm(PAGE_CONFIG['width'])
    section.page_height = Cm(PAGE_CONFIG['height'])
    section.left_margin = Cm(PAGE_CONFIG['margin_left'])
    section.right_margin = Cm(PAGE_CONFIG['margin_right'])
    section.top_margin = Cm(PAGE_CONFIG['margin_top'])
    section.bottom_margin = Cm(PAGE_CONFIG['margin_bottom'])


def _set_style_font(style, font_name, font_size, bold=None):
    """设置文档样式的字体，包括东亚字体。"""
    font = style.font
    font.name = font_name
    font.size = Pt(font_size)
    if bold is not None:
        font.bold = bold
    # 设置东亚字体
    rPr = style.element.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        style.element.append(rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)


def configure_styles(document):
    """根据 FONT_CONFIG 创建/修改文档内建样式。"""
    # 标题样式映射：config key -> Word style name
    style_map = {
        'title': 'Title',
        'heading1': 'Heading 1',
        'heading2': 'Heading 2',
        'heading3': 'Heading 3',
        'heading4': 'Heading 4',
    }

    for config_key, style_name in style_map.items():
        cfg = FONT_CONFIG[config_key]
        try:
            style = document.styles[style_name]
        except KeyError:
            style = document.styles.add_style(style_name, 1)  # 1 = WD_STYLE_TYPE.PARAGRAPH
        _set_style_font(style, cfg['name'], cfg['size'])
        if config_key == 'title':
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 正文样式
    body_cfg = FONT_CONFIG['body']
    normal_style = document.styles['Normal']
    _set_style_font(normal_style, body_cfg['name'], body_cfg['size'])
    pf = normal_style.paragraph_format
    pf.line_spacing = Pt(body_cfg['line_spacing'])
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.first_line_indent = Pt(body_cfg['first_line_indent'])


def set_run_font(run, font_name, font_size):
    """设置 run 的字体属性（含东亚字体）。"""
    run.font.name = font_name
    run.font.size = Pt(font_size)
    # 设置东亚字体
    r = run._element
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        r.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)


def apply_body_style(paragraph):
    """为段落应用正文样式：字体、行距、首行缩进。"""
    body_cfg = FONT_CONFIG['body']
    pf = paragraph.paragraph_format
    pf.line_spacing = Pt(body_cfg['line_spacing'])
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.first_line_indent = Pt(body_cfg['first_line_indent'])


def _create_border_element(tag, color, border_style='single', size='4'):
    """创建单个边框 XML 元素。"""
    el = OxmlElement(tag)
    el.set(qn('w:val'), border_style)
    el.set(qn('w:sz'), size)
    el.set(qn('w:color'), '{:02X}{:02X}{:02X}'.format(*color))
    el.set(qn('w:space'), '0')
    return el


def configure_table(table):
    """根据 TABLE_CONFIG 设置表格样式。"""
    # 设置表格居中
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 设置边框
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    borders = OxmlElement('w:tblBorders')
    color = TABLE_CONFIG['border_color']
    border_style = TABLE_CONFIG['border_style']

    for tag in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        borders.append(_create_border_element(f'w:{tag}', color, border_style))

    # 移除已有边框设置
    existing = tblPr.find(qn('w:tblBorders'))
    if existing is not None:
        tblPr.remove(existing)
    tblPr.append(borders)


def get_heading_config(level):
    """根据标题级别返回对应的 FONT_CONFIG 配置。"""
    level_map = {
        1: 'title',
        2: 'heading1',
        3: 'heading2',
        4: 'heading3',
        5: 'heading4',
    }
    key = level_map.get(level, 'heading4')
    return FONT_CONFIG[key]


def get_body_config():
    """返回正文的 FONT_CONFIG 配置。"""
    return FONT_CONFIG['body']
