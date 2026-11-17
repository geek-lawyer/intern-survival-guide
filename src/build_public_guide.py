"""Build the Chinese public reading edition of the Intern Survival Guide (DOCX).

Run from anywhere: python3 src/build_public_guide.py
Output: dist/实习生生存指南（公开阅读版）.docx
"""

from datetime import datetime
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Mm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import RELATIONSHIP_TYPE as RT

PROJECT_DIR = Path(__file__).resolve().parent.parent
OUT = PROJECT_DIR / "dist" / "实习生生存指南（公开阅读版）.docx"
OFFICIAL_ACCOUNT_QR = PROJECT_DIR / "assets" / "official_account_qr.jpg"
OUT.parent.mkdir(parents=True, exist_ok=True)

NAVY = "17324D"
BLUE = "2E6687"
TEAL = "2D7B78"
GOLD = "B58A3B"
INK = "222222"
MUTED = "6B7280"
LIGHT = "EEF4F6"
PALE = "F7F8FA"
WHITE = "FFFFFF"
RED = "9B2C2C"

doc = Document()
# python-docx 的基础模板默认写入 Word 2010（14）兼容级别。现代 Word 会因此
# 显示“Compatibility Mode”，并使用旧版分页规则；明确升级为 Word 2013+（15）。
compat = doc.settings._element.find(qn("w:compat"))
if compat is not None:
    for compat_setting in compat.findall(qn("w:compatSetting")):
        if compat_setting.get(qn("w:name")) == "compatibilityMode":
            compat_setting.set(qn("w:val"), "15")
sec = doc.sections[0]
sec.page_width = Mm(210)
sec.page_height = Mm(297)
sec.top_margin = Inches(0.8)
sec.bottom_margin = Inches(0.75)
sec.left_margin = Inches(0.88)
sec.right_margin = Inches(0.88)
sec.header_distance = Inches(0.35)
sec.footer_distance = Inches(0.35)

# A4 handbook geometry. Tables use exact OOXML widths instead of renderer defaults.
CONTENT_WIDTH_DXA = int(round(sec.page_width.twips - sec.left_margin.twips - sec.right_margin.twips))
TABLE_INDENT_DXA = 120
TABLE_WIDTH_DXA = CONTENT_WIDTH_DXA - TABLE_INDENT_DXA
TABLE_CELL_MARGINS_DXA = {"top": 100, "bottom": 100, "start": 120, "end": 120}

styles = doc.styles

def font(run, size=None, bold=None, color=INK, name="Arial Unicode MS"):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Arial Unicode MS")
    run._element.rPr.rFonts.set(qn("w:ascii"), name)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), name)
    if size is not None:
        run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.bold = bold

normal = styles["Normal"]
normal.font.name = "Arial Unicode MS"
normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial Unicode MS")
normal.font.size = Pt(10.5)
normal.font.color.rgb = RGBColor.from_string(INK)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.32

for name, size, color, before, after in [
    ("Heading 1", 19, NAVY, 18, 9),
    ("Heading 2", 14.5, BLUE, 14, 6),
    ("Heading 3", 11.5, TEAL, 9, 4),
]:
    st = styles[name]
    st.font.name = "Arial Unicode MS"
    st._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial Unicode MS")
    for theme_attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        st._element.rPr.rFonts.attrib.pop(qn(f"w:{theme_attr}"), None)
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = RGBColor.from_string(color)
    st.paragraph_format.space_before = Pt(before)
    st.paragraph_format.space_after = Pt(after)
    st.paragraph_format.keep_with_next = True

# Compact-reference list tokens, encoded on the real Word list styles.
for name in ("List Bullet", "List Number"):
    st = styles[name]
    st.font.name = "Arial Unicode MS"
    st._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial Unicode MS")
    st.font.size = Pt(10.5)
    st.font.color.rgb = RGBColor.from_string(INK)
    st.paragraph_format.left_indent = Inches(0.375)
    st.paragraph_format.first_line_indent = Inches(-0.188)
    st.paragraph_format.space_after = Pt(4)
    st.paragraph_format.line_spacing = 1.25

if "Lead" not in styles:
    st = styles.add_style("Lead", WD_STYLE_TYPE.PARAGRAPH)
    st.base_style = normal
    st.font.name = "Arial Unicode MS"
    st._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial Unicode MS")
    st.font.size = Pt(10.5)
    st.font.color.rgb = RGBColor.from_string(NAVY)
    st.paragraph_format.line_spacing = 1.38
    st.paragraph_format.space_after = Pt(10)

def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:fill"), fill)

def set_cell_margins(cell, top=100, start=120, bottom=100, end=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcMar.find(qn("w:" + m))
        if node is None:
            node = OxmlElement("w:" + m)
            tcMar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")

def ensure_child(parent, tag):
    child = parent.find(qn(tag))
    if child is None:
        child = OxmlElement(tag)
        parent.append(child)
    return child

def set_dxa_width(parent, tag, width_dxa):
    node = ensure_child(parent, tag)
    node.set(qn("w:type"), "dxa")
    node.set(qn("w:w"), str(int(width_dxa)))

def column_widths_from_weights(weights, total_width_dxa=TABLE_WIDTH_DXA):
    """Allocate exact integer widths while preserving the author's proportions."""
    total = float(sum(weights))
    result = [int(round(total_width_dxa * (weight / total))) for weight in weights]
    result[-1] += total_width_dxa - sum(result)
    return result

def apply_table_geometry(table, column_widths_dxa):
    """Synchronize tblW, tblInd, tblGrid and every tcW for stable rendering."""
    if sum(column_widths_dxa) != TABLE_WIDTH_DXA:
        raise ValueError("Table column widths must sum to TABLE_WIDTH_DXA")
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl_pr = table._tbl.tblPr
    set_dxa_width(tbl_pr, "w:tblW", TABLE_WIDTH_DXA)
    table_indent = ensure_child(tbl_pr, "w:tblInd")
    table_indent.set(qn("w:type"), "dxa")
    table_indent.set(qn("w:w"), str(TABLE_INDENT_DXA))
    layout = ensure_child(tbl_pr, "w:tblLayout")
    layout.set(qn("w:type"), "fixed")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in column_widths_dxa:
        grid_col = OxmlElement("w:gridCol")
        grid_col.set(qn("w:w"), str(width))
        grid.append(grid_col)

    for col_idx, width in enumerate(column_widths_dxa):
        table.columns[col_idx].width = Twips(width)
    for row in table.rows:
        row.height = None
        for col_idx, cell in enumerate(row.cells):
            width = column_widths_dxa[col_idx]
            cell.width = Twips(width)
            set_dxa_width(cell._tc.get_or_add_tcPr(), "w:tcW", width)
            set_cell_margins(cell, **TABLE_CELL_MARGINS_DXA)

def set_paragraph_shading(paragraph, fill):
    shd = ensure_child(paragraph._p.get_or_add_pPr(), "w:shd")
    shd.set(qn("w:fill"), fill)

def set_paragraph_left_rule(paragraph, color=TEAL):
    p_bdr = ensure_child(paragraph._p.get_or_add_pPr(), "w:pBdr")
    left = ensure_child(p_bdr, "w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "18")
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), color)

def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)

def set_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    cant.set(qn("w:val"), "true")
    trPr.append(cant)

def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = paragraph.add_run("第 ")
    font(r, 8.5, color=MUTED)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    paragraph._p.append(fld)
    r2 = paragraph.add_run(" 页")
    font(r2, 8.5, color=MUTED)

header = sec.header.paragraphs[0]
header.alignment = WD_ALIGN_PARAGRAPH.LEFT
hr = header.add_run("实习生生存指南  ·  公开阅读版")
font(hr, 8.5, True, MUTED)
add_page_number(sec.footer.paragraphs[0])

def p(text="", style=None, bold_prefix=None, align=None, italic=False, after=None):
    para = doc.add_paragraph(style=style)
    if align is not None:
        para.alignment = align
    if bold_prefix and text.startswith(bold_prefix):
        r1 = para.add_run(bold_prefix)
        font(r1, bold=True)
        r2 = para.add_run(text[len(bold_prefix):])
        font(r2)
    else:
        r = para.add_run(text)
        font(r)
        r.italic = italic
    if after is not None:
        para.paragraph_format.space_after = Pt(after)
    return para

def add_hyperlink(paragraph, text, url):
    """Add a styled external hyperlink without exposing the raw URL in the text."""
    rid = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rid)
    run = paragraph.add_run(text)
    font(run, 10.5, color=BLUE)
    run.font.underline = True
    hyperlink.append(run._r)
    paragraph._p.append(hyperlink)
    return hyperlink

_bookmark_counter = 0

def add_bookmark(paragraph, name):
    """Bookmark a paragraph so the static contents page can jump to it."""
    global _bookmark_counter
    _bookmark_counter += 1
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(_bookmark_counter))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(_bookmark_counter))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)

def add_internal_hyperlink(paragraph, text, anchor, size=10.5, bold=False, color=BLUE):
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    hyperlink.set(qn("w:history"), "1")
    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    r_fonts = OxmlElement("w:rFonts")
    for attr in ("ascii", "hAnsi", "eastAsia"):
        r_fonts.set(qn(f"w:{attr}"), "Arial Unicode MS")
    r_pr.append(r_fonts)
    r_color = OxmlElement("w:color")
    r_color.set(qn("w:val"), color)
    r_pr.append(r_color)
    r_size = OxmlElement("w:sz")
    r_size.set(qn("w:val"), str(int(size * 2)))
    r_pr.append(r_size)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(underline)
    if bold:
        r_pr.append(OxmlElement("w:b"))
    run.append(r_pr)
    text_node = OxmlElement("w:t")
    text_node.text = text
    run.append(text_node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    return hyperlink

def linked_bullet(title, url, note):
    para = doc.add_paragraph(style="List Bullet")
    add_hyperlink(para, title, url)
    if note:
        r = para.add_run("｜" + note)
        font(r, 10.5, color=INK)
    return para

def quote(text, label="请记住"):
    para = doc.add_paragraph()
    para.paragraph_format.left_indent = Twips(TABLE_INDENT_DXA)
    para.paragraph_format.right_indent = Twips(TABLE_INDENT_DXA)
    para.paragraph_format.space_before = Pt(5)
    para.paragraph_format.space_after = Pt(9)
    para.paragraph_format.line_spacing = 1.28
    para.paragraph_format.keep_together = True
    set_paragraph_shading(para, LIGHT)
    set_paragraph_left_rule(para)
    r = para.add_run(label + "｜")
    font(r, 10.5, True, TEAL)
    r = para.add_run(text)
    font(r, 10.5, True, NAVY)

def bullets(items):
    for item in items:
        para = doc.add_paragraph(style="List Bullet")
        r = para.add_run(item)
        font(r)

def new_list_number_id():
    """Create a fresh numbering instance so an independent list restarts at 1."""
    numbering = doc.part.numbering_part.element
    style_num_pr = styles["List Number"]._element.pPr.find(qn("w:numPr"))
    base_num_id = style_num_pr.find(qn("w:numId")).get(qn("w:val"))
    base_num = next(
        node for node in numbering.findall(qn("w:num"))
        if node.get(qn("w:numId")) == base_num_id
    )
    abstract_num_id = base_num.find(qn("w:abstractNumId")).get(qn("w:val"))
    new_num_id = max(int(node.get(qn("w:numId"))) for node in numbering.findall(qn("w:num"))) + 1

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(new_num_id))
    abstract = OxmlElement("w:abstractNumId")
    abstract.set(qn("w:val"), abstract_num_id)
    num.append(abstract)
    level_override = OxmlElement("w:lvlOverride")
    level_override.set(qn("w:ilvl"), "0")
    start_override = OxmlElement("w:startOverride")
    start_override.set(qn("w:val"), "1")
    level_override.append(start_override)
    num.append(level_override)
    numbering.append(num)
    return new_num_id

def numbers(items):
    num_id = new_list_number_id()
    for item in items:
        para = doc.add_paragraph(style="List Number")
        num_pr = ensure_child(para._p.get_or_add_pPr(), "w:numPr")
        ilvl = ensure_child(num_pr, "w:ilvl")
        ilvl.set(qn("w:val"), "0")
        num_id_node = ensure_child(num_pr, "w:numId")
        num_id_node.set(qn("w:val"), str(num_id))
        r = para.add_run(item)
        font(r)

def section_break(title, subtitle, bookmark):
    spacer = doc.add_paragraph()
    spacer.paragraph_format.page_break_before = True
    spacer.paragraph_format.space_after = Pt(70)
    k = doc.add_paragraph()
    k.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = k.add_run("INTERN SURVIVAL GUIDE")
    font(r, 9, True, GOLD)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(14)
    r = h.add_run(title)
    font(r, 27, True, NAVY)
    add_bookmark(h, bookmark)
    s = doc.add_paragraph()
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = s.add_run(subtitle)
    font(r, 12, color=MUTED)
    back = doc.add_paragraph()
    back.alignment = WD_ALIGN_PARAGRAPH.CENTER
    back.paragraph_format.space_before = Pt(10)
    back.paragraph_format.space_after = Pt(0)
    add_internal_hyperlink(back, "返回目录", "toc", size=8.5, color=MUTED)

def add_table(headers, rows, widths):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    column_widths = column_widths_from_weights(widths)
    hdr = table.rows[0]
    set_repeat_table_header(hdr)
    set_row_cant_split(hdr)
    for i, text in enumerate(headers):
        cell = hdr.cells[i]
        cell.width = Twips(column_widths[i])
        set_cell_shading(cell, NAVY)
        set_cell_margins(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_after = Pt(0)
        r = para.add_run(text)
        font(r, 9.5, True, WHITE)
    for ridx, row in enumerate(rows):
        added_row = table.add_row()
        set_row_cant_split(added_row)
        cells = added_row.cells
        for i, text in enumerate(row):
            cells[i].width = Twips(column_widths[i])
            set_cell_margins(cells[i])
            cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if ridx % 2 == 1:
                set_cell_shading(cells[i], PALE)
            para = cells[i].paragraphs[0]
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.line_spacing = 1.15
            r = para.add_run(text)
            font(r, 9)
    apply_table_geometry(table, column_widths)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return table

# Cover
for _ in range(4):
    doc.add_paragraph()
k = doc.add_paragraph()
k.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = k.add_run("INTERN SURVIVAL GUIDE")
font(r, 10, True, GOLD)
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_before = Pt(12)
title.paragraph_format.space_after = Pt(8)
r = title.add_run("实习生生存指南")
font(r, 30, True, NAVY)
sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sub.add_run("从大学生，到可靠的专业协作者")
font(r, 15, color=TEAL)
byline = doc.add_paragraph()
byline.alignment = WD_ALIGN_PARAGRAPH.CENTER
byline.paragraph_format.space_before = Pt(12)
byline.paragraph_format.space_after = Pt(2)
r = byline.add_run("作者、主编与主要撰写｜胡泽贤（George Hu）")
font(r, 10.5, True, NAVY)
contributors = doc.add_paragraph()
contributors.alignment = WD_ALIGN_PARAGRAPH.CENTER
contributors.paragraph_format.space_after = Pt(22)
r = contributors.add_run("经验共建｜历届实习生与带教成员")
font(r, 9.5, color=MUTED)
quote("这本手册不帮你熬时间。它只回答三件事：方向怎么看，能力怎么长，信任怎么赢。", "本手册的目的")
doc.add_paragraph().paragraph_format.space_after = Pt(60)
meta = doc.add_paragraph()
meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = meta.add_run("公开阅读版  ·  2026年7月版")
font(r, 9.5, color=MUTED)

doc.add_page_break()

doc.add_heading("版权与版本说明", level=1)
p("作者、版本和许可说不清，后来每一次传播都会把责任弄模糊。", style="Lead")
add_table(["项目", "信息"], [
    ("书名", "《实习生生存指南（公开阅读版）》"),
    ("作者、主编与主要撰写", "胡泽贤（George Hu）"),
    ("经验共建", "历届实习生与带教成员的公开心得、工作反馈与制度沉淀"),
    ("首版时间", "2026年7月"),
    ("当前版本", "2026年7月公开阅读版；后续修订以作者发布的最新版本为准"),
], [1.7, 4.85])
p("Copyright © 2026 胡泽贤（George Hu）. 文字内容以 CC BY-NC-SA 4.0 许可发布。", italic=True, after=4)
p("本书原创结构、文字表达与汇编成果的著作权归作者所有。团队成员与历届实习生对其原有文章、照片及其他独立作品依法享有相应权利；本书所链接或概述的第三方文章，其权利归原作者及相关权利人。", after=4)
p("本公开阅读版的文字以 CC BY-NC-SA 4.0 许可发布：注明书名、作者、版本与出处后可以复制、转载、翻译和改编，改编作品须沿用相同许可；商业使用须另经作者书面许可。团队实际采用前仍应结合自身制度、工具与适用法律进行调整。", after=4)
p("本书属于团队培养与管理指南。若其中内容与现行法律、有效制度或经确认的个别安排不一致，以法律和最新有效文件为准。", after=5)

doc.add_page_break()
doc.add_heading("联系作者、内容入口与勘误", level=1)
p("手册一旦停止修正，就开始过时。发现错漏或更好的办法，请直接反馈。", style="Lead")
add_table(["渠道", "信息"], [
    ("电子邮箱", "hu.zexian@qq.com"),
    ("微信公众号", "没事多吃核桃"),
    ("抖音", "胡泽贤 律师"),
], [1.7, 4.85])
qr_para = doc.add_paragraph()
qr_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
qr_para.paragraph_format.space_before = Pt(2)
qr_shape = qr_para.add_run().add_picture(str(OFFICIAL_ACCOUNT_QR), width=Inches(2.8))
qr_shape._inline.docPr.set("title", "微信公众号“没事多吃核桃”")
qr_shape._inline.docPr.set("descr", "扫码关注微信公众号“没事多吃核桃”的二维码")
qr_caption = doc.add_paragraph()
qr_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
qr_caption.paragraph_format.space_after = Pt(0)
r = qr_caption.add_run("扫码关注公众号“没事多吃核桃”｜欢迎反馈勘误与改进建议")
font(r, 9.5, color=MUTED)

doc.add_page_break()
doc.add_heading("使用说明", level=1)
p("一本手册若只规定实习生做什么，却不约束管理者提供什么，就只是命令清单。这里同时写工作质量、保密与协作的底线，以及实习生应得的学习机会、具体反馈和成长空间。", style="Lead")
p("这本手册也不为眼前某一批人而写。团队已经带过约七至八批实习生，专业、性格和基础各不相同；困惑、错误和成长规律却经常重复。经验若只留在某个人脑子里，下一批就只能重新交学费。")
p("第一天通读；首周结合设备操作和保密训练使用；遇到具体问题，优先查第三篇和附录。每周挑一个问题复盘。新经验经带教负责人审核后进入知识库，过时做法也要及时删掉。")

doc.add_heading("先看清我们怎样工作", level=2)
p("团队不大，战线很长：案件、客户协作、普法内容、数字流程和影像制作经常交叉，法律、财会、外语、技术和传播也会在同一项任务里相遇。资源多不等于学得多。沿一条主线从问题走到交付，做出会被使用的成果，才不至于把“接触很多”变成打杂。")
add_table(["团队特质与条件", "可能形成的学习机会", "团队必须承担的管理责任"], [
    ("真实案件、客户与法律程序", "接触文书、证据、案卷、检索、庭审观察、客户沟通和企业制度材料。", "按照保密、权限和风险分级开放；说明任务在案件链条中的位置，并及时复盘。"),
    ("多位带教者及复合专业背景", "观察不同经验和专长怎样处理问题，把法律、合规、财会、技术和传播连接起来。", "把专业履历转成新人能参与的任务、讲解和反馈；统一标准。"),
    ("线下协作、靠近业务现场与相对可控的生活成本", "在通勤和生活压力可控的环境中，观察本地法律服务的真实生态。", "如实说明实际地点、作息、补贴和食宿；不以生活成本低替代待遇与培养。"),
    ("电脑、共享存储、打印扫描、AI等数字化条件", "建立文件、检索、自动化、知识管理和信息安全能力。", "提供零基础训练、权限和SOP；不把“设备已经提供”当作“新人自然会用”。"),
    ("相机、灯光、设计剪辑软件或内容平台等可选条件", "参与海报、普法文章和短视频，把法律转成公众能理解的表达。", "建立拍摄放行、事实核验、版权、肖像和审稿机制；准确与保密优先。"),
    ("相对明确的常规作息、舒适办公环境和较高协作密度", "降低初入职场的适应成本，在真实协作中学习沟通与责任。", "例如咖啡、硬件等条件本身不是培养；紧急工作要说明原因和期限，不把长期加班当态度测验。"),
    ("七至八批实习生留下的心得、模板和手册", "后来者能够从前人的经验和错误出发，不必每批从零摸索。", "保留真实意见，更新制度、删除失效做法，让经验进入知识库。"),
], [1.35, 2.55, 2.65])

p("团队不找抽象的“最优秀”，只看工作方式是否匹配：案件要事实意识和保密，小团队要沟通与边界，AI要学习速度和人工复核，短期线下实习还要看连续投入。合适的人能在这套工作方式里尽快产生价值，团队也愿意认真培养他。")
quote("我们怎样工作，就寻找并培养怎样的能力。团队介绍和人才标准，必须是同一幅自画像。", "团队自画像")

doc.add_heading("要求必须是双向的", level=2)
bullets([
    "实习生对经手事项负责：认真、诚实、守密，发现问题不隐瞒，不把自己当临时旁观者。",
    "团队交办时说明任务目的、完成标准、期限和负责人，验收时给出具体反馈。",
    "基础工作稳定后，逐步提供检索、分析、文书、案件讨论和内容创作机会。",
    "团队从招聘阶段起如实说明地点、任务、作息、培养方式和尚未确定的条件，不用模糊承诺换取到岗。",
    "严格要求应当指向具体行为和工作质量，不以羞辱、威胁或模糊否定代替培养。",
    "双方都可以判断这段实习是否适合自己；珍惜机会不等于放弃独立思考和合理边界。",
])

doc.add_page_break()
toc_heading = doc.add_heading("先找问题，再找答案", level=1)
add_bookmark(toc_heading, "toc")
p("电子版可点击以下条目跳转；纸质阅读可按篇名和章号查找。", italic=True, after=5)
for anchor, title, scope in [
    ("part1", "第一篇  立志：为什么值得认真对待这段实习", "第一至七章"),
    ("part2", "第二篇  共同成长：实习生与管理者各自承担什么", "第八至十二章"),
    ("part3", "第三篇  上手训练：电脑、共享存储、PDF、扫描打印与办公环境", "第十三至十六章"),
    ("part4", "第四篇  办案与交付：文书、证据、案卷、委托、AI与放行机制", "第十七至二十三章"),
    ("part5", "第五篇  方向：把实习经历沉淀为可以带走的能力", "第二十四至二十六章"),
    ("appendix", "附录  可打印、可填写、可复用的工作工具", "附录一至十一"),
]:
    para = doc.add_paragraph(style="List Bullet")
    add_internal_hyperlink(para, title, anchor, bold=True)
    r = para.add_run("｜" + scope)
    font(r, 9.5, color=MUTED)

doc.add_heading("按你现在的问题查找", level=2)
for anchor, lead, target in [
    ("part3", "刚入职、不熟悉工作电脑或共享存储", "先看第三篇，再用附录一、二现场验收"),
    ("part2", "刚收到任务、不知从哪下手", "先看第八章，再填附录三《任务交付卡》"),
    ("part4", "准备对外发送、提交或流转材料", "先看第四篇，不满足红线时先暂停"),
    ("ai", "准备让AI参与检索、写作、分析或内容制作", "先看第二十三章，先划输入边界，再设计人工验收"),
    ("part2", "正在带教或管理实习生", "先看第二篇，再用附录八、十一"),
    ("part5", "准备复盘、写实习报告或判断职业方向", "先看第五篇，再参照附录九、十"),
]:
    para = doc.add_paragraph(style="List Bullet")
    r = para.add_run(lead + "：")
    font(r, bold=True)
    add_internal_hyperlink(para, target, anchor, size=10.5)

section_break("第一篇  立志", "为什么值得认真对待这段实习", "part1")

doc.add_heading("第一章  从标准答案走向真实后果", level=1)
p("你从学校来，但这里没有标准答案。", style="Lead")
p("大学常把题目、期限和评分标准一并给你。真实工作常只给零散事实、具体的人和一个不能轻率对待的期限。")
p("一页材料可能牵动财产、名誉、家庭甚至自由。实习生暂时不作最终判断，却要对经手的每个字负责：不知道就问，做不到就报，做完要查。")
p("专业并不神秘。它只是要求一个人在面对别人的重要事务时，仍然认真、诚实、克制、可靠。")
quote("实习从写下第一个字以前开始：先知道这个字会影响谁。")

doc.add_heading("第二章  先问交付给谁用", level=1)
p("完成任务是最低要求，解决问题才是工作的目的。", style="Lead")
p("整理证据，是为了看清案件事实；检索法律，是为了找到支撑判断的依据；扫描文件，是为了让下一环节准确、完整、快速地使用材料。动作只是外壳，交付用途才是任务。")
p("每次接受任务，先问自己：")
numbers([
    "这项工作的最终目的是什么？",
    "谁会使用我的成果？",
    "对方使用时最需要什么？",
    "哪些错误可能产生严重后果？",
])
p("只照表面指令行动，最容易交出“做完了但不能用”的东西。交付前至少能说明四件事：目的是什么，谁会使用，我按什么标准完成，哪些风险已经检查。")

doc.add_heading("第三章  不要用最低标准设计自己的人生", level=1)
p("“给我多少补贴，我就做多少事情”看似公平，却可能让一个人用很低的价格出售自己最宝贵的成长时间。", style="Lead")
p("补贴最容易计算，也最容易遮住其余回报。把真实任务、具体反馈、独立交付和职业判断一起算进账：便宜的实习如果只有消耗，其实很贵；补贴不高但能持续长能力，也未必吃亏。")
p("所谓长远机会，必须落到今天能做的任务、谁会给反馈、做好以后会增加什么责任。只让你先忍耐、以后再说，是空头支票。")
p("长期主义不替不合理安排辩护。团队也要交付学习机会、具体反馈和合理边界；这些长期缺席，实习生应重新判断是否值得继续投入。")
quote("不要计较一时是否吃亏，但要判断这段经历是否让你变得更有能力、更有判断、更值得信任。")

doc.add_heading("第四章  小事决定别人是否敢把大事交给你", level=1)
p("想做复杂案件、重要文书，却把整理、校对、扫描和接待看成杂事，很难赢得信任。大事从不交给只愿意谈大事的人。", style="Lead")
p("热情不能兑换信任。文件名是否方便后来者，页码是否真的核过，发现问题是否立即报告，会议室用完是否恢复——这些小事会把一个人是否可靠写得很清楚。")
p("基础工作已经稳定、准确，团队就应逐步增加任务难度。实习生也可以主动提出：下一次我能否先做一版分析或文书初稿，再请律师修改？做好小事，本来就是为了赢得大事的机会。")

doc.add_heading("第五章  别用忙碌冒充成长", level=1)
p("判断一件苦活值不值得，不看它累不累，看下一次是否仍要从头来。", style="Lead")
p("有些工作必然重复、繁琐，第一稿也可能被大量修改。但同一件事做了很多次，仍没留下方法、模板和经验，下一次还得从头开始，留下的就只有重复劳动，没有成长。")
p("每完成一项新任务，都问一遍：")
bullets([
    "我为什么耗时这么长，卡点在哪里？",
    "哪一步最容易出错，怎样增加检查？",
    "能否形成一份清单、模板或操作说明？",
    "下一个人怎样才能少走一次弯路？",
])
quote("第一次不会，可以学习；第二次不会，可以查记录；第三次再遇到同样问题，应当已有自己的方法。")

doc.add_heading("第六章  实习不是表演，而是一次双向选择", level=1)
p("不要为了表现积极而假装忙碌，也不要为了得到认可而迎合所有意见。", style="Lead")
p("实习是一场双向面试。团队看你怎样工作；你看自己能否长期面对琐碎事实、责任、期限、不确定性和冲突，也看自己是否仍愿意读、写、改。")
p("实习后发现自己不适合律师，不算失败。可惜的是从未认真试过，也没认真想过，只是顺着环境往前走。")

doc.add_heading("第七章  不要相信记忆：清单是对人性的尊重", level=1)
p("熟练让人更快，也让人自信到跳步；紧急让人集中，也让人只顾眼前。清单不替人思考，它只在人疲劳、分心和过度自信时，替专业守住底线。", style="Lead")
p("团队的《短视频拍摄放行检查暨工作日志》把开工、收工、红线、最小样本、存储位置和异常记录分开。红线没过就暂停。这比一句“注意设备”有用得多。")

doc.add_heading("SOP不是一张步骤表", level=2)
add_table(["层级", "解决的问题", "表现形式"], [
    ("操作说明", "第一次做的人，怎样完成每一个动作？", "图文步骤、设备说明、示例"),
    ("检查清单", "熟悉的人，怎样避免漏掉关键动作？", "短句、勾选、责任人"),
    ("放行门槛", "什么条件不满足，就必须暂停？", "红线项目、共同确认、暂缓选项"),
    ("工作日志", "发生了什么，怎样复盘和改进？", "时间、位置、数量、异常、签字"),
], [1.05, 3.0, 2.5])
p("四层缺一，恰好对应四种失败：新人不会做，熟手忘了做，风险来时没人敢停，出错以后没人知道为何再犯。")

doc.add_heading("把一次检查，变成团队的防线", level=2)
p("凡是有期限、不可逆、对外发送、原件流转、客户影响或保密风险的工作，都值得有放行检查。表格长短随风险，不随职位和偏好。")
add_table(["工作场景", "放行前至少确认", "典型红线"], [
    ("立案或线上提交", "法院、主体、代理地位、文书版本、附件、份数、期限、上传结果", "主体错误、期限不明、签章缺失、上传文件打不开"),
    ("文书对外发送", "收件人、发送版本、附件、内部批注、敏感信息、负责律师确认", "误发对象、带修订痕迹、未经审核"),
    ("证据扫描装订", "顺序、页码、清晰度、目录对应、原件状态、抽样检查", "漏页、倒页、原件损坏、目录不一致"),
    ("材料邮寄", "收件信息、材料清单、份数、邮寄方式、留痕、物流查询", "地址未核、材料未留底、无寄件凭证"),
    ("客户接待", "时间地点、参与人、会议室、材料、保密环境、记录授权", "客户身份或授权不清、擅自拍摄录音"),
    ("AI辅助成果", "允许输入、充分脱敏、来源核查、事实法条复核、人工负责", "输入客户秘密、虚构依据、未经复核直接使用"),
    ("收工或离场", "文件归档、任务交接、设备关机、原件归位、门窗与安全", "原件遗留、素材未备份、设备通电、门窗未锁"),
], [1.35, 3.25, 2.05])
p("低风险工作用三至五项短清单；高风险、多人协作或不可逆事项，增加红线、操作人、复核人和明确的“暂缓”。")

doc.add_heading("用反向思考设计清单", level=2)
p("先问失败：这件事最可能漏在哪一步？疲劳、分心、从众和过度自信，就是清单要拦住的东西。")
numbers([
    "目标：这项工作最终必须产生什么可靠结果？",
    "反向：如果它失败，最可能漏在哪一步、错在哪一处？",
    "偏误：熟悉、赶时间、服从权威、从众或过度自信，会不会让人跳过检查？",
    "激励：现有制度究竟奖励正确完成，还是无意中只奖励表面速度？",
    "防线：哪几项必须成为红线，谁有权宣布暂停？",
])
p("一张好清单只收三类项目：容易漏、后果重、能客观确认。异常、返工和差点出错，就是它的修订记录。先堵住失败，再追求速度。")

section_break("第二篇  共同成长", "实习生与管理者各自承担什么", "part2")

doc.add_heading("第八章  主动，不等于擅自作主", level=1)
p("成熟的主动，是先理解、再检索，形成初步判断后带着方案提问；涉及案件方向、期限、客户利益和保密风险时，则立即报告。", style="Lead")
add_table(["层级", "处理方式", "典型事项"], [
    ("第一层", "先查实习手册、团队模板库、既有工作记录和可靠公开资料。", "设备使用、文件命名、常见流程、基础格式"),
    ("第二层", "向同期实习生、带教人或事项直接交办人询问，并说明自己已经查过什么。", "模板位置、具体操作、任务理解中的局部疑问"),
    ("第三层", "立即向负责律师报告，不等待、不猜测。", "期限、案件策略、客户利益、材料缺失、信息冲突、误发或泄密风险"),
], [0.8, 3.1, 2.65])
p("别只问“这个怎么弄”。把问题升级成：“我查了手册和模板，看到两种做法。第一种依据是……，第二种风险是……。我倾向第一种，请确认。”")
p("怕问题太简单，也不能隐瞒错误、拖延报告或擅自决定。成熟不靠少提问，而靠知道什么必须立刻上报。")

doc.add_heading("第九章  用心、用力、用情：先回答三个问题", level=1)
p("口号最容易记，也最容易不用。接到重要任务、遇到困难或准备交付时，停下来问三件事。", style="Lead")
p("三问留给自己，不交日报，不做打卡，也不比谁回答漂亮。它们只负责校准方向。")

doc.add_heading("第一问｜用心：我真的理解自己在做什么吗？", level=2)
p("用心，是让大脑先进入任务：目的、事实、风险和标准，一个都不糊涂。")
p("实习生自问：", bold_prefix="实习生自问：")
bullets([
    "这项工作的目的是什么，最后由谁使用？",
    "交付什么才算完成，模板、格式、期限和存储位置是否明确？",
    "我掌握的是事实，还是未经核实的猜测？",
    "最容易出错、又可能产生严重后果的地方在哪里？",
    "提交以前，我是否逐项检查了主体、案号、金额、日期、页码和附件？",
    "这次方法能否沉淀成下一次可以复用的清单或模板？",
])
p("管理者自问：我在交办前是否想清楚了目的、标准和风险？实习生的偏差究竟来自不用心，还是来自我没有说清楚？", bold_prefix="管理者自问：")

doc.add_heading("第二问｜用力：我真的为结果采取了足够行动吗？", level=2)
p("用力，是把理解变成动作：先做样本，遇阻留痕，不能按时就提前报告。")
p("实习生自问：", bold_prefix="实习生自问：")
bullets([
    "我已经开始行动，还是仍在等待别人把每一步都安排好？",
    "我是否先做了一个最小样本，让方向尽早得到确认？",
    "遇到障碍以后，我查过什么、尝试过什么、准备了什么方案？",
    "我投入的时间产生了成果，还是只产生了重复和返工？",
    "发现无法按时完成时，我是否在截止前主动报告？",
    "收到反馈以后，我是否真正修改并避免同类问题再次发生？",
])
p("管理者自问：我是否设置了检查点、及时验收并协调必要资源？还是把任务交出去以后不闻不问，到截止时才批评结果？", bold_prefix="管理者自问：")

doc.add_heading("第三问｜用情：我真的把这件事和相关的人放在心上了吗？", level=2)
p("用情，是记得文件背后有人，也记得自己的成果会被同事接着使用。")
p("实习生自问：", bold_prefix="实习生自问：")
bullets([
    "这份材料、这个日期和这个错误，对客户意味着什么？",
    "我是否把经手事项当作共同成果，而不是一句“这不归我管”？",
    "同事接手我的成果时，能否迅速找到、理解并继续工作？",
    "我是否尊重客户、同事、共同空间和别人给予我的反馈？",
    "如果这是我自己的案件、文件或重要机会，我还会这样处理吗？",
])
p("管理者自问：我看到的是一个可以临时使用的人手，还是一个需要被培养的人？我是否听见了他的困惑和期待，并给予具体、不羞辱人格的反馈？", bold_prefix="管理者自问：")

doc.add_heading("从开始到结束，都问一遍", level=2)
add_table(["时点", "用心之问", "用力之问", "用情之问"], [
    ("开始前", "我是否理解目的和标准？", "下一步具体行动是什么？", "这件事影响谁？"),
    ("进行中", "方向和事实是否仍然正确？", "阻碍是否已处理或报告？", "协作对象需要知道什么？"),
    ("交付前", "是否完成实质与形式检查？", "是否真正达到可用状态？", "我愿意让别人依赖这份成果吗？"),
    ("结束后", "我学到了什么方法？", "怎样减少下一次成本？", "我给团队留下了什么？"),
], [0.85, 2.0, 2.0, 1.75])
p("遇到明确紧急期限，交办人说明轻重和时点；实习生判断无法按时，就在期限前报告。投入的边界，是双方为同一个承诺负责。")
quote("不给清楚目标、必要工具和及时反馈，只要求实习生投入，不叫高标准，叫单方面要求。", "双向投入")

doc.add_heading("第十章  不要害怕暴露自己的无知", level=1)
p("团队允许你暂时缺乏经验，但不能接受不诚实。", style="Lead")
p("不知道可以补，隐瞒会让错误继续。没有完成却说完成，发现错误盼着没人看见，不理解却按猜测行动——这些才会消耗信任。你可以坦率地说：“这部分我还没有完全理解，我先复述一下，请您确认。”")
p("发现错误，按“事实—影响—补救—防复发”汇报：错在哪里，可能影响什么，已经怎样补救，如何排查同类问题。汇报不是认错仪式，而是把风险重新交回可控状态。")

doc.add_heading("第十一章  把批评翻译成下一步", level=1)
p("批评最刺耳的时候，人最容易只听见语气，漏掉问题。委屈、紧张和防御都正常；情绪过去后，把批评翻译成下一步。", style="Lead")
numbers([
    "对方指出的具体事实是什么？",
    "哪些确实是自己的问题？",
    "问题产生在哪一个环节？",
    "下次要增加哪一个检查步骤？",
])
p("别把一次错误扩大成“我不适合做律师”，也别把所有批评都解释成“对方针对我”。反馈要落到事实、影响和下一步；离开这三件事，严格很容易退化成地位表演。")

doc.add_heading("带教专章  培养实习生，也是管理者的考试", level=1)
p("筛选一个现成好用的人不难；把新人教成可靠的协作者，才检验管理能力。", style="Lead")
p("新人不熟设备、流程、文件位置和完成标准，是入行常态。管理者熟悉，不等于讲解可以省略。")
p("培养不等于纵容。任务、资料和反馈给清以后，仍然敷衍、拖延或故意绕开要求，就要直接纠正。先分清不会、没听懂、不愿意；三种问题，三种教法。")

doc.add_heading("谁交办，谁带教", level=2)
p("谁交办、谁审核，谁就负责带教这一件事。不能把新人的所有问题，一律转给那个“负责实习生的人”。")
add_table(["管理方针", "对管理者的要求"], [
    ("统一标准", "同类任务使用同一模板、命名和验收口径；标准变化时及时更新手册。"),
    ("单一负责", "一项任务明确一名最终交办人；多人提出意见时，由负责人统一冲突要求。"),
    ("分级授权", "根据实习生已证明的能力逐步增加难度，不突然把高风险事项完全交出。"),
    ("样本先行", "批量工作先验收一个最小样本，确认方法正确后再扩大。"),
    ("过程可见", "设置中间检查点，关注风险和阻碍，不到截止时才第一次看成果。"),
    ("结果验收", "明确是否通过、哪里要改、谁复核；不能只说“差不多”或让任务自然消失。"),
    ("事后沉淀", "把重复问题转为清单、模板或知识库更新，减少下一次对个人记忆的依赖。"),
], [1.25, 5.3])

doc.add_heading("管理者首先管理自己", level=2)
bullets([
    "自己遵守保密、守时、文件归档和检查要求，不要求实习生执行自己长期忽视的规则。",
    "交办前想清楚目的、成果、权限、资源、期限、风险和验收人，不把自己的模糊转嫁给实习生。",
    "不以职位替代理由，不以情绪替代反馈，不以反复改变口径制造所谓抗压训练。",
    "实习生提出风险或承认错误时，先控制损失、查明事实，再讨论责任，避免逼出隐瞒文化。",
    "定期检查自己的培养结果：实习生是否获得了更强的独立能力，而不只是承担了更多杂务。",
])

doc.add_heading("严、爱、德：管理者也要接受三问", level=2)
p("只讲严格，会变成服从；只讲关怀，会变成亏欠；自己不受同一规则约束，两者都不可信。")
add_table(["管理维度", "保留什么，防止什么", "管理者要问自己的问题"], [
    ("威权：专权与尚严", "对事实、程序、期限、保密和质量可以严；对人的思想、情绪和私人边界不能专断。", "我在控制风险，还是控制人？"),
    ("仁慈：生活照顾与工作照顾", "耐心示范、具体反馈、合理照顾；不以关心侵入隐私、制造亏欠或索取忠诚。", "这份帮助是否让他更独立？"),
    ("德行：操守与公正示范", "自己先守密、守时、复核；同一标准不因亲疏和顺从改变。", "这条规则也约束我吗？"),
], [1.55, 3.25, 1.75])
p("紧急事项可以集中指挥；不能把临时例外变成日常秩序。事后解释理由，恢复提问和复盘。")
bullets([
    "要求可以随风险和能力变化；共同底线、授权理由和评价标准不能随亲疏变化。",
    "关怀只用于理解起点和提供帮助；不收集无关隐私，也不换取额外服从。",
    "“听话”不是培养结果。敢报错、能异议、少依赖、可独立交付，才是。",
])
quote("管理者可以像长者一样负责，但不能把成年人当孩子；可以对事严厉，不能靠控制人格证明权威。", "领导边界")

doc.add_heading("带教好不好，别只看感谢信", level=2)
p("历届心得最有价值的，不是热情语气，而是相同问题反复出现：设备陌生、小事返工、任务目的不清、反馈后逐步独立，以及跨专业能力进入法律现场。")
add_table(["历届心得中反复出现的反馈", "对带教者意味着什么"], [
    ("刚到岗时对设备、流程和职场表达陌生；无用功常来自没理解目的，或独自消耗太久。", "第一天教会入口、样本和提问边界；允许限时自查，风险事项或久查无果时及时升级。"),
    ("打印、页码、装订、归档等小事最容易出错，也最能训练秩序和耐心。", "不要只交代动作；要解释它在案件流程中的位置、正确顺序和错误后果。"),
    ("具体反馈和重复练习，能够让返工逐渐减少，让新人从不会做到独立。", "先验收样本、指出可观察的问题，再逐步退出；培养的结果应是减少依赖，而不是永久盯人。"),
    ("非法律专业也能在研究、设计、技术、传播和组织方面产生独特价值。", "先识别可迁移能力，以优势任务建立信心，再安排相邻的陌生任务扩展边界。"),
    ("律师和助理的工作状态、学习习惯和反馈方式，本身就是新人观察的课程。", "管理者每天怎样工作，比会议上说了什么更能塑造团队标准。"),
], [3.15, 3.4])
p("公开心得经过本人选择和团队编辑，适合看见新人学到了什么，不足以证明带教没有问题。还要看匿名意见、返工、重复错误、独立交付和离岗建议。后来者无需模仿前人的热情，更不必把感谢写成标准答案。")
quote("心得用来宣传，只留下赞美；用来改进培养，才留下组织能力。", "经验沉淀")

doc.add_heading("第二天校准会议：先修复共同理解", level=2)
p("当多人同时觉得新人“不好带”，先别急着把问题归因于态度。一次校准会把宗旨、任务、标准、带教关系和双方需求摊开；共同理解修复后，配合和主动性才有基础。")
p("校准会不训话，只校准五件事：做什么、谁指导、怎样算达标、新人卡在哪里、团队能提供什么。")

doc.add_heading("新人带着什么期待来到这里", level=3)
p("新人常带着三种问题：课堂知识能否进入真实场景；自己能否适应律师节奏；能否从小事开始，换取更大任务。管理者若只听见“想学习”，就听漏了他们也在判断职业和团队。")
quote("考公还是做律师，迟早要交卷。只有真正上手跟过案子，才能判断自己适不适合这种状态。愿意做，能够做，熟练做。", "某次校准会议中的真实声音（匿名整理）")
add_table(["实习生的期待", "团队应当怎样回应"], [
    ("让理论进入实践", "安排文书、案卷、证据和检索等真实任务；说明它在案件中的位置，及时反馈。"),
    ("理解律师的真实日常", "在保密和权限内，让新人看见完整任务链，也看见期限、客户责任和职业代价。"),
    ("训练严谨、表达与沟通", "先做最小样本；能稳定交付，再增加难度与授权。"),
    ("为团队产生实际价值", "交付真正会被使用的任务，让新人知道帮助了谁；别把实习压成无反馈的零碎杂务。"),
], [2.0, 4.55])
p("“不想给团队添麻烦”是责任感，不是禁言令。该问就问，遇险就报。若团队只记得新人要贡献，却忘了他们也为学习、体验和成长而来，双向选择就只剩单向取用。")

p("具体议程见附录九。会后只留一张行动清单：任务、负责人、样本、期限、检查点；相关成员确认。会议若没有下一步，只是一次集体表达。")

doc.add_heading("招聘只决定入口，培养决定成长", level=2)
p("BOSS直聘只是入口，平台迟早会变。团队真正需要保存的是自己的判断：什么能力在这里有用，我们愿意为哪种成长花时间。团队介绍和人才筛选都围绕这两个问题，否则前者像广告，后者像表格。")

add_table(["团队的真实特质", "由此产生的筛选维度", "应寻找的当前证据"], [
    ("真实案件与程序责任", "事实整理、证据意识、谨慎、保密和发现错误后的处理。", "小任务能区分事实与推断、指出材料缺口；面对错误先报告、核验和修正。"),
    ("小团队、高协作密度、多角色交叉", "沟通清楚、反馈及时、主动而有边界，能与不同专业成员协作。", "能复述任务、说明已尝试方法、及时报告风险；接受具体反馈后能完成第二版。"),
    ("电脑、共享存储、AI与SOP化工作", "学习敏捷、结构化习惯、工具迁移和人工校验能力。", "短时间学会一段新规则并应用；能解释AI如何使用、哪里可能错、自己怎样核对。"),
    ("法律、新媒体与视觉表达并行", "跨专业理解、信息转化、作品意识，以及准确性与传播效果之间的取舍。", "有可说明本人贡献的项目或作品；能把复杂材料转成清楚表达而不夸大结论。"),
    ("线下到岗与短期集中培养", "到岗稳定性、连续投入时间、现实预期和较快形成基础产出的可能。", "能说清实习起止、返校安排、通勤食宿和硬性边界，而不是只表示“都可以”。"),
    ("七至八批实习经验与知识沉淀", "复盘、改进和把个人经验转为团队资产的意愿。", "能讲清一次从不会到会的过程；愿意按模板、清单和反馈留下可复用成果。"),
], [1.75, 2.45, 2.35])
p("筛选不办简历竞赛，也不找最便宜、最听话的人，只找离岗位最近的证据。学历、成绩、证书和经历只能估计起点；工作样本、学习速度、反馈后的修订、稳定交付和现实条件才接近真实工作。短期实习可以有高门槛，但每道门槛都得说出理由。学校名气不是能力结论。")
quote("先验信号只决定看哪里；是否录用看当前证据；入职以后看行为、质量与成长。", "证据优先")

doc.add_heading("让候选人先看见真实条件", level=3)
p("入职前，候选人应看见真实任务、工作方式、地点、作息、补贴和食宿，也能读过历届心得后自行判断。只收集岗位证据；工作样本说明用途，不碰案件秘密，不把免费测试当生产。记录方式见附录十一。")
p("录用以后，管理者要从筛选者切换为带教者。招聘只决定谁暂时进门，不能证明团队会带人；更不能让新人不断接受临时测试，来证明自己值得被教。")

doc.add_heading("复杂任务拆开教，重复任务先做对", level=2)
p("批量开始得越早，返工可能越大。复杂任务先拆成可检查的小步；重复任务先验一个样本，再决定是否放大。")
numbers([
    "带教者先说明目的、样例、标准和截止时间。",
    "实习生先完成一个最小样本，例如先打印一份、先整理一组证据、先命名一个文件夹。",
    "带教者检查方法和结果，及时校正共性问题。",
    "确认无误后再批量完成，由实习生自行按照清单复核。",
    "任务结束后记录方法，让下一次不再依赖逐项盯人。",
])
p("效率的顺序是正确、稳定、再快。管理者不替新人做完最后检查，而用样本、检查点和清单逐步交出质量责任。")

doc.add_heading("交办要坚定，反馈要具体", level=2)
p("管理者不熟某个软件或设备，可以一起查；不能因此含糊交办。仍要明确谁负责、何时完成、以什么为准。坚定来自标准，不来自声量。")
p("别用“你怎么这么笨”“这谁都会”评价智力。把反馈说成四件事：哪一步没完成，造成什么影响，下次增加什么检查，何时重新提交。")

doc.add_heading("让协作可见，但不做机械日报", level=2)
p("普通短任务用交付说话。跨日、多人或有期限风险的任务，才更新完成项、下一步、阻碍和需要谁协助。")
p("进度可见是为了尽早发现风险，不是为了表演忙碌。需要帮助时应说明背景和已经尝试的方法；带教者也应及时回应关键阻碍。")

doc.add_heading("第十二章  保密是进入这个职业的第一道门", level=1)
p("专业能力可以逐步培养，保密意识必须从第一天建立。", style="Lead")
bullets([
    "未经允许，不得拍摄、复制、转发或带离案件材料。",
    "工作文件只存放在团队指定电脑、共享存储或协作账号，不上传未经许可的个人存储空间。",
    "工作微信与个人社交内容做好隔离，不在朋友圈、群聊或与亲友谈论案件。",
    "离开工位时锁屏；纸质材料不随意摆放、丢弃或夹带。",
    "未获明确许可，不把客户或案件材料输入任何AI工具。",
    "发错文件、材料遗失或疑似泄密时立即报告，不自行隐瞒处理。",
])
quote("保密不是因为某个案件看起来重要，而是因为客户把重要的事情交给了我们。", "底线")

section_break("第三篇  上手训练", "先掌握工作环境，再承担专业任务", "part3")

doc.add_heading("第一周训练地图", level=1)
p("第一周不比产量，先清除四种基础返工：找不到文件、转不对格式、不会用设备、不清楚保密边界。不会不可耻；学过、练过，仍只等别人代劳，才需要纠正。", style="Lead")
add_table(["阶段", "训练内容", "验收成果"], [
    ("第1天", "保密、工作账号、电脑系统、文件管理、锁屏、截图、应用切换", "独立创建并找到规定文件夹；完成截图标注并锁屏"),
    ("第2天", "文件命名、复制移动、版本另存、共享存储登录或挂载、路径确认、PDF基础", "将一组练习文件按规则处理并归档到指定共享目录"),
    ("第3天", "扫描、打印、身份证复印、OCR、合并拆分PDF、耗材与简单排障", "完成一套模拟材料的扫描、合并、页码和打印"),
    ("第4天", "任务复述、样本先行、任务更新、共享日历、工作手册与提问顺序", "完成一张任务交付卡和一次简明进度更新"),
    ("第5天", "综合练习与首周评估，补练未掌握项目", "从接收材料到共享存储归档完整独立完成一次"),
], [0.9, 3.55, 2.15])
p("第一次不会很正常；只看不做、跳过练习，才会把不会带进下一项任务。每项基础能力都要亲手完成并现场验收，没掌握就补练，不用一句“以后注意”假装问题已经解决。")
p("第一至第二天开新人校准会：说清宗旨、阶段任务、标准、带教关系和双方需求。议程见附录九。")

doc.add_heading("第十三章  接受任务的五步法", level=1)
numbers([
    "听目的：这项工作要解决什么问题，最终由谁使用？",
    "确认成果：需要文档、表格、PDF、纸质材料还是口头汇报？",
    "确认标准：格式、模板、命名、存储位置和参考样例是什么？",
    "确认期限：最终截止时间和中间检查点是什么？",
    "复述任务：用自己的话复述一遍，确认没有理解偏差。",
])
quote("不要只问“什么时候交”，还要确认“交什么才算完成”。")

doc.add_heading("第十四章  每一天，都从让环境可工作开始", level=1)
p("办公室不会自动进入工作状态。看见现场，先让它能工作。", style="Lead")
p("到岗后按当日安排检查门窗、灯光、通风、桌面、接待区和耗材；离岗前清理垃圾、归档文件、关闭设备并确认门锁。值日表分配动作，不转移其他人看见异常就处理或报告的责任。")
p("共同空间本身就是工作条件。值日可以轮换，责任不能跟着值日表一起下班。")

doc.add_heading("第十五章  先学会驾驭电脑，再谈提高效率", level=1)
p("第一次不会用团队指定的电脑系统不丢人，团队也不把系统差异当能力差。但系统差异不能长期变成返工；已有经验不会自动迁移，主动完成转换是基本责任。", style="Lead")

doc.add_heading("以Windows迁移到Mac为例", level=2)
add_table(["常见任务", "Windows习惯", "Mac工作方式"], [
    ("复制／粘贴／保存／打印", "Ctrl+C／V／S／P", "Command+C／V／S／P"),
    ("应用切换", "Alt+Tab", "Command+Tab"),
    ("关闭窗口／退出程序", "Alt+F4", "Command+W关闭窗口；Command+Q退出程序"),
    ("锁定电脑", "Windows+L", "Control+Command+Q"),
    ("搜索", "Windows搜索", "Command+Space打开聚焦搜索"),
    ("截图", "Windows+Shift+S", "Command+Shift+4选区；Command+Shift+5更多选项"),
    ("新建文件夹", "Ctrl+Shift+N", "Command+Shift+N"),
    ("重命名", "F2", "选中文件后按Return"),
    ("移动文件", "剪切后粘贴", "Command+C，再用Command+Option+V移动"),
    ("快速预览", "打开文件", "选中文件按Space，不必启动应用"),
    ("右键", "鼠标右键", "鼠标右键、双指轻点或Control+单击"),
], [1.45, 2.25, 2.85])

doc.add_heading("文件管理器不是桌面背景，而是文件工作的中心", level=2)
bullets([
    "能够切换列表、分栏和图标视图，按照名称、日期、标签和类型排序。",
    "能够从边栏进入共享存储、下载、桌面和指定工作目录，并确认当前文件的真实位置。",
    "能够创建文件夹、批量选择、复制、移动、重命名和查看文件简介。",
    "知道“最近使用”只是入口，不是文件实际存储位置；归档前必须确认路径。",
    "知道关闭窗口不等于退出应用，拔出存储卡或移动硬盘前必须安全推出。",
])

doc.add_heading("PDF与图片是必须独立完成的基本功（以下以Mac为例）", level=2)
bullets([
    "图片转PDF：按文件顺序选中图片，使用系统快速操作创建PDF，并检查方向与顺序。",
    "合并PDF：在“预览”的缩略图栏拖入其他PDF，确认顺序后另存新文件。",
    "拆分PDF：选中需要的页面拖出生成新文件，检查页数后按照规范命名。",
    "旋转、裁剪与格式转换：使用“预览”的标记、裁剪和导出功能，始终保留原始文件。",
    "页码标注、压缩、OCR、加密和电子签名按照团队指定软件及具体任务要求操作。",
])

doc.add_heading("工作电脑，第一原则是保密", level=2)
bullets([
    "离开工位立即锁屏，不把工作密码告诉他人。",
    "使用Mac时，未经允许不登录个人iCloud，不开启个人照片、桌面或文稿同步。",
    "工作材料不通过个人微信、个人邮箱、AirDrop等系统传输功能、个人网盘或私人移动硬盘传输。",
    "不知道文件会被保存到哪里时先确认，不以“系统自动同步”作为解释。",
])
p("任何处理都遵守一个朴素原则：原文件不覆盖，工作版另存；每次转换后都打开检查。软件提示“成功”，不等于材料真的完整。")

doc.add_heading("工作电脑会不会，用现场任务说话", level=2)
numbers([
    "在指定目录创建文件夹，按照规则命名，并把路径告诉带教者。",
    "复制一份模板、重命名为工作版，移动到指定共享目录且不改变原件。",
    "把若干图片生成PDF，再与另一PDF合并、旋转、拆分并核对页数。",
    "完成一次选区截图、标注和保存；找到截图实际位置。",
    "使用聚焦搜索和Finder分别找到同一文件，并查看文件简介确认路径。",
    "完成一次存储卡或移动设备的安全推出，并在离开工位前锁屏。",
])

doc.add_heading("第十六章  打印、扫描、邮寄：程序事实需要留下痕迹", level=1)
p("打印机和扫描仪不是后勤设备，它们经常处在证据、送达和程序期限的交界处。", style="Lead")
bullets([
    "打印前确认黑白或彩色、单双面、纸张规格、份数、页码和装订方式；打印后逐页检查。",
    "身份证复印尽量居中，确保正反面清晰完整，不遮挡边角和有效信息。",
    "扫描前去除订书钉并确认原件顺序；扫描后核对页数、方向、清晰度和是否漏页。",
    "喷墨打印机按团队安排定期使用；出现颜色异常时先查手册，不擅自拆机。",
    "诉讼材料按负责律师要求邮寄；留存快递单、寄件联、材料副本和物流轨迹。",
    "签收重要文件时记录签收人和时间；有疑问立即联系交办人。",
    "使用碎纸机前确认材料确属可销毁副本；原件、签章件及无法判断的文件不得擅自销毁。",
])

section_break("第四篇  办案与交付", "让每一项成果经得起使用和复核", "part4")

doc.add_heading("第十七章  文书提交前，先把错误拦下来", level=1)
p("法律文书里的错误不按大小计价。一个字，也可能改变主体、金额、期限或请求。提交前，至少完成以下检查。")
add_table(["检查类别", "检查内容"], [
    ("主体", "当事人姓名、单位全称、身份证号、统一社会信用代码、地址、联系方式"),
    ("案件", "案号、案由、法院、承办法官、程序阶段、送达信息"),
    ("数字", "日期、金额、利率、计算结果、期限起止、大小写一致"),
    ("证据", "目录、编号、名称、页码、证明目的、原件与复印件、附件齐全"),
    ("法律", "引用法条准确，确认现行有效，结论与事实相对应"),
    ("格式", "标题、字体、字号、行距、页码、落款、签章、打印效果"),
    ("文件", "文件名、版本号、存储位置正确；无内部批注、修订痕迹或不应外发信息"),
], [1.15, 5.4])
p("形成交叉检查习惯：先检查事实和逻辑，再检查格式；重要材料尽量由另一人复核。不要在同一次快速浏览中幻想自己能发现所有问题。")

doc.add_heading("第十八章  证据整理不是编号工作", level=1)
p("整理证据的目标，是帮助团队回答三个问题：发生了什么，怎样证明，对方可能怎样反驳。", style="Lead")
bullets([
    "建立案件时间线：日期、事件、参与人、对应证据、待核实问题。",
    "建立证据目录：编号、名称、形成时间、来源、页码、证明目的。",
    "标记缺口与冲突：哪些事实缺少证据，哪些材料相互矛盾。",
    "保留原始文件：不得覆盖原件；加工版、OCR版和工作版应区分。",
    "统一命名和归档：确保后来的人能够找到、识别并继续使用。",
])

doc.add_heading("第十九章  案卷整理：缺什么，就诚实地标明缺什么", level=1)
p("案卷整理只看三个结果：后来者能还原程序、找到材料、识别缺口。纸塞进档案袋，目录做得再漂亮，也不算完成。", style="Lead")
p("刑事案件应按照侦查、审查起诉和审判等实际阶段整理；民事、行政、经济与仲裁案件按照委托材料、当事人材料、诉辩文书、证据、调查材料、庭审材料、裁判文书和办案小结等类别归档。不同阶段本来就可能不存在某类文书，不要为了形式完整而制造材料。")
bullets([
    "先确认案件类型、代理地位和程序阶段，再使用对应目录模板。",
    "委托合同、授权材料、收费凭证、身份材料等基础文件缺失时，及时报告并追补。",
    "未找到材料时，用便签或缺件表标明；继续查找既有档案和经批准的工作记录。",
    "不得随便编造页码、日期、文件名称或材料内容。",
    "没有真实庭审笔录时，不得根据判决书或AI补写成仿佛真实存在的笔录；应标记缺失并请负责律师决定如何说明。",
    "原件、正式提交版、对方材料和内部工作版应明确区分。",
])
quote("档案可以不完美，但记录必须真实。", "案卷底线")

doc.add_heading("第二十章  接受委托与提交材料：先分组，再复核", level=1)
p("客户交来的是材料，不是秩序。实习生先让团队看得清，再把判断留在授权边界内。")
bullets([
    "区分身份与委托材料、授权材料、证据、公检法材料及其他资料，使用便签清楚标识。",
    "签字盖章文件完成后及时扫描归档；扫描件与纸质原件建立对应关系。",
    "客户提交的原件先清点、扫描并记录，未经允许不在原件上书写、装订或改变状态。",
    "提交法院前，根据案件、代理地位、当事人人数和法院要求，由负责律师确认材料清单与份数。",
    "线上立案上传前检查主体信息、代理人信息、起诉材料、授权材料、身份证明、证据及送达信息。",
    "任何不确定的份数、签章、期限和提交渠道都应询问负责律师，不以旧经验替代本案确认。",
])

doc.add_heading("第二十一章  信息不进系统，就等于没有留下", level=1)
p("同一个问题出现第二次，就该考虑把它变成团队资产。完成重复性文书后，团队共享资料中若没有模板，先整理一份候选模板，交业务负责人和知识库维护人审核；通过后才能标记为标准模板。")
p("所有工作文件应保存到团队指定共享存储或协作系统的对应目录。文件名采用“内容主题＋数字日期＋对象或案件标识”的统一规则；红色等标签的含义以团队最新约定为准。另存文件时不得影响原件，不随意移动、覆盖或删除。")
p("信息若只在一个人脑中，团队就等于没有它。开庭、会见、提交和回访都进共享日历，写清时间、地点与地图、人员和事项；跨日、多人或有期限风险的任务，再记进度、下一步和待协调事项。")
p("操作说明至少包括：目的、适用场景、前置条件、具体步骤、常见错误、文件位置、示例截图、更新时间、编写人和审核人。例如扫描仪、打印机、会议设备、茶水设备及物品存放位置，也可以进入知识库。")
doc.add_heading("第二十二章  接待、公共空间与设备：专业感藏在现场里", level=1)
p("共同空间不是实习生承包区。每个人都要让它保持可工作、可接待、可安全交接的状态。", style="Lead")
bullets([
    "保持个人工位和公共区域整洁，工作材料及时归位。",
    "来访者到访时主动引导、倒水并协助准备会议材料，不等待提醒。",
    "会议结束后共同恢复会议室，清理杯具、纸张和设备。",
    "下班前检查垃圾、地面、文件、设备和门窗，按照值日安排处理。",
    "发现耗材不足或设备异常，先按手册排查，再向设备或行政负责人反馈。",
])
p("设备不会用，先查手册；查过仍不明白，再问。例如茶水设备使用后应清理，拍摄前确认电量、存储、画面、收音、授权和归档位置。未经许可，不拍、不录、不发。")
p("例如，正式拍摄可以按团队最新的拍摄放行清单逐项检查。先录最小样本并回放；红线未清零，不开拍。收工以素材可打开、已完整归档、设备安全复位为准。")

ai_heading = doc.add_heading("第二十三章  AI越强，越要有人负责", level=1)
add_bookmark(ai_heading, "ai")
p("AI让信息搬运、标准答案和普通第一稿变得廉价，问题定义、专业判断、品味与责任因此更贵。会聊天不算会工作；能否接进可核验、可复用、有人负责的流程，才是分水岭。", style="Lead")
p("团队可以用Codex、腾讯WorkBuddy或其他获准工具演练。工具很快会换，下面四个判断要留下。")
add_table(["判断", "落实到行动"], [
    ("AI是能力系统，不是工具清单", "先问任务需要哪些能力，再选工具；不为追新而追新。"),
    ("AI放大已有认知", "专业底座和问题定义越清楚，输出越有价值；流畅不等于正确。"),
    ("一次回答不等于团队能力", "把有效方法沉淀成任务结构、检查点、模板或SOP，不从空白聊天框重新开始。"),
    ("人的主体性不能外包", "人定目标、给语境、作判断、担责任；AI加速探索、生成、比较和重复执行。"),
], [2.15, 4.5])

doc.add_heading("从一次提问到可复用的闭环", level=2)
numbers([
    "定义：说清目的、已知事实、未知问题、限制条件和验收标准。",
    "划界：确认哪些材料允许输入；必要时最小化并充分脱敏。",
    "试做：先用公开或虚构材料完成最小样本，确认方向后再扩大。",
    "核验：回到原件检查来源、事实、法条、案例、日期、金额和逻辑。",
    "沉淀：记录有效方法、失败原因和人工修改，经审核后进入知识库。",
])
p("“用AI做”不是交办。目标、权限、材料边界、人工检查点和最终负责人，少一个都只会把模糊管理变成更快的不可靠。")

doc.add_heading("法律团队的AI红线", level=2)
add_table(["可以做", "须经许可并逐项复核", "禁止做"], [
    ("用公开或虚构材料学习；草拟空白模板、检查清单、检索词和表达备选。", "对允许且充分脱敏的材料做结构整理、版本比较、文书框架或反方模拟。", "输入客户秘密、身份信息、证据、联系方式、内部策略或其他未获许可的信息。"),
    ("比较无敏感信息的版本；辅助公开内容、设备说明和知识库目录。", "上传文件、联网检索、批量自动化或让智能体执行可能改变外部状态的动作。", "把输出当作法律依据；虚构来源；未经人工审核对外发送、提交或作不可逆操作。"),
], [2.15, 2.25, 2.25])

doc.add_heading("把判断和责任留给人", level=2)
bullets([
    "问题是我定义的，还是我把思考本身也外包了？",
    "即将输入的信息，有没有一个字不该离开团队控制范围？",
    "哪些来源、数字、法条和结论，必须回到原件核验？",
    "如果结果出错，我能否说明过程，并愿意承担责任？",
    "这次有效的方法，能否让下一位同事少走一次弯路？",
])
quote("AI可以替你加速，但不能替你决定什么值得做、什么才算正确。", "AI时代的主体性")

section_break("第五篇  方向", "把实习经历沉淀为可以带走的能力", "part5")

doc.add_heading("第二十四章  先站稳共同底座，再长出自己的方向", level=1)
p("不同专业不必被磨成同一种人。保密、守时、文件、沟通和自检是地板；站稳以后，用长处进入真实任务，再跨出熟悉领域，用独立成果回答自己愿意走哪条路。", style="Lead")
add_table(["培养环节", "带教时要回答的问题", "应看到的阶段证据"], [
    ("建立共同底座", "保密、守时、设备、文件、沟通和检查是否达到最低可靠标准？", "通过基础评估；能独立完成一项低风险任务。"),
    ("识别入职起点", "他已经学过什么、做过什么、会使用什么工具，哪些能力可以迁移？", "形成简短能力画像，不收集与培养无关的隐私。"),
    ("从优势处进入", "哪一个真实任务既能帮助团队，又能让新人较快建立第一次胜任感？", "先完成一个可使用的最小样本，并得到具体反馈。"),
    ("跨出熟悉领域", "在已有能力旁边增加哪一种新责任，既有挑战又不至于失控？", "能说明任务目的，减少依赖，并主动报告风险。"),
    ("形成独立成果", "哪一项成果可以由他从接受任务做到自检交付，并为后来者复用？", "完成文书、目录、内容作品、清单、模板或操作说明。"),
    ("获得职业判断", "这段经历让他更了解什么工作、什么责任以及自己是否适合？", "实习报告使用具体经历作出判断，也允许得出“不适合”。"),
], [1.25, 3.55, 1.75])
p("履历只说明起点。招聘样本只用来设计第一周，不贴永久标签；培养手册只记与任务有关的信息，不复制联系方式和家庭情况。方向也无需入职当天决定，它会从任务、反馈和复盘里长出来。")
quote("不要先问一个新人像不像律师，先问他已经拥有的能力，怎样迁移到真实问题上。", "识别潜能")

doc.add_heading("第二十五章  不同专业，不同入口，同一套培养逻辑", level=1)
doc.add_heading("法学及法律相关专业", level=2)
p("法学生最容易高估法条记忆，低估事实整理。真实工作先问材料证明了什么、缺了什么、彼此哪里冲突，再谈规则怎样适用。")
add_table(["阶段", "重点能力", "建议成果"], [
    ("第一周", "保密、设备、共享存储、文件命名、扫描归档、任务复述", "通过首周基础评估；完成一份设备或流程说明"),
    ("第二至三周", "案件时间线、证据目录、基础检索、格式核对", "独立整理一组材料；形成检查清单"),
    ("第四周以后", "争点提炼、案例检索、文书初稿、案件复盘", "提交一份可讨论的分析或文书初稿；更新标准模板"),
], [1.1, 3.0, 2.45])
p("进步不看接触了多少案子，看能否从“我看过”走到“我整理清楚、核对过，并能说明尚缺什么”。")

doc.add_heading("视觉、传媒与内容相关专业", level=2)
p("法律内容做得好看却让人误解，比难看更糟。视觉与传媒实习生的价值，是不牺牲事实、保密和边界，让复杂内容更容易看懂、更值得信任。")
bullets([
    "学习保密、案件流程和法律信息核验，避免为了传播效果牺牲准确性。",
    "建立公众号文章的选题、采访、审稿、排版、发布和归档流程。",
    "将普法内容转化为图解、时间线、流程图和短视频脚本。",
    "整理团队视觉规范、图片版权记录、素材库和文件命名规则。",
    "与法律方向实习生交叉协作：一方负责事实与法律审核，另一方负责信息结构和表达。",
])

doc.add_heading("别用专业名称替人分配命运", level=2)
p("下面三种起点只说明一种方法：先从已经证明的能力进入，再把它接到法律现场。专业相同，路线也不必相同。")
add_table(["入职起点", "可以怎样从优势处进入", "随后应扩展什么"], [
    ("园林背景，做过科研辅助、创新项目和宣传组织。", "从信息收集、公众号与视频内容、活动支持等任务建立胜任感。", "进入庭审观察、法律流程和事实核验，训练专业表达与法律意识。"),
    ("生物医学工程背景，具有电子技术、社团组织和较强动手能力。", "从设备、文件、结构化信息和检查清单等任务切入。", "增加文字编辑、电话沟通、制度材料和复盘任务，发展流程与管理意识。"),
    ("数字媒体背景，学习视觉、交互、广告、摄影，掌握多种设计和剪辑工具。", "从海报、名片、宣传册、摄影和视频制作等专业任务切入。", "加入证据分类、法律事实核验、客户需求沟通和团队审稿，把“好看”升级为准确、清楚和可信。"),
], [2.0, 2.35, 2.2])
p("跨专业任务必须连得上成长主线。让设计背景的实习生整理证据，是为了理解案件逻辑和客户需求，从而让后续表达更准确；如果带教者说不清这层连接，就该重做任务安排。")

doc.add_heading("其他专业或复合背景", level=2)
p("遇到外语、财会、技术、管理等新背景，先确认共同底座，再设计一条任务主线：它能解决什么真实问题，边界在哪里，从哪个最小样本开始，用什么成果证明能力。不要把人硬塞进旧路线。")

doc.add_heading("第二十六章  让成果留下来，也让能力带得走", level=1)
p("实习报告和公众号文章只是成果的外壳。实习结束时，更重要的交付是以下问题的答案。", style="Lead")
bullets([
    "我是否比刚来时更可靠？",
    "我能否独立完成一项基础任务？",
    "我是否形成了检查、汇报和复盘习惯？",
    "我是否理解律师工作的责任与代价？",
    "我是否更加清楚自己未来要走的路？",
    "我给团队和后来者留下了什么可复用的知识？",
])
p("实习总结不用写成感恩信。收获可以写，困惑、批评和未解决的问题也可以写，甚至可以坦白自己不适合律师行业。交稿是为了训练复盘与表达；是否公开、怎样署名、使用哪些照片，由本人决定，并经过保密和事实审核。")
quote("总结不必证明这段实习值得，只要诚实说清：我做过什么、学会什么、仍不懂什么。", "实习复盘")
section_break("附录", "把思想落实为每天可以执行的动作", "appendix")

doc.add_heading("附录一  入职清单", level=1)
p("领到电脑和工位，只说明人到了；进入团队的信息、责任和保密系统，才算完成入职。具体材料以团队共享知识库中的最新入职清单为准。")
add_table(["时间", "事项", "确认"], [
    ("入职前", "按通知购买实习保险，填写承诺书及团队要求的基础表格。", "□"),
    ("入职时", "签署保密协议及必要承诺；提交身份证明、学生证明、学籍验证及团队要求的其他材料。", "□"),
    ("账号设备", "绑定企业微信或指定协作账号；领取工作手机、电脑、钥匙并登记。", "□"),
    ("第一天", "学习保密边界、文件系统、提问顺序、紧急事项报告方式和办公区域。", "□"),
    ("离岗时", "完成交接单，归还设备和钥匙，解绑账号、删除个人指纹，清理个人设备中的工作信息。", "□"),
], [1.0, 4.95, 0.6])

doc.add_heading("附录二  首周基础评估", level=1)
p("首周既练基础，也彼此判断。不会可以补；保密、诚信问题，或多次提醒仍不改的基础错误，会直接影响是否继续。")
add_table(["项目", "达标表现", "结果"], [
    ("手册学习", "知道主要制度、信息查询位置和提问顺序", "□ 已掌握  □ 补练"),
    ("工作电脑与共享存储", "能登录、查找、命名、保存、移动和归档文件", "□ 已掌握  □ 补练"),
    ("扫描打印", "能完成单双面扫描、合并、OCR、打印及简单排障", "□ 已掌握  □ 补练"),
    ("办公设备", "能使用团队常见设施，例如打印、会议或茶水设备，并完成复位或清洁", "□ 已掌握  □ 补练"),
    ("保密规范", "知道材料存储、传输、微信和AI使用边界", "□ 已掌握  □ 补练"),
    ("工作沟通", "能够复述任务、确认期限并报告进度和风险", "□ 已掌握  □ 补练"),
    ("基础质量", "文件命名、格式、页码、附件和存储位置基本准确", "□ 已掌握  □ 补练"),
], [1.15, 4.15, 1.25])

doc.add_heading("工作电脑基础操作验收明细（以Mac为例）", level=2)
add_table(["验收任务", "通过标准", "结果"], [
    ("键盘与窗口", "能使用Command、Option、Shift；会切换应用、关闭窗口、退出程序和锁屏。", "□通过 □补练"),
    ("Finder", "会创建文件夹、排序、搜索、批量选择、重命名、复制、移动并确认真实路径。", "□通过 □补练"),
    ("截图与预览", "会选区截图、标注、快速预览并找到保存位置。", "□通过 □补练"),
    ("PDF", "能完成图片转PDF、合并、拆分、旋转、裁剪、页码、OCR和另存检查。", "□通过 □补练"),
    ("共享存储归档", "能连接指定目录、按规则命名、保留原件、保存工作版并复查文件可打开。", "□通过 □补练"),
    ("设备推出", "能安全推出存储卡或移动设备，不在读写过程中直接拔出。", "□通过 □补练"),
    ("保密设置", "不使用个人同步或传输渠道；离开工位锁屏。", "□通过 □补练"),
], [1.2, 4.25, 1.1])

doc.add_heading("附录三  任务交付卡", level=1)
for label in ["任务名称", "交办人", "任务目的", "成果形式", "参考模板/文件位置", "截止时间", "中间检查点", "主要风险", "完成情况与存储位置"]:
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(7)
    r = para.add_run(label + "：")
    font(r, 10.5, True, NAVY)
    r = para.add_run("____________________________________________________________")
    font(r, 10.5, color=MUTED)

doc.add_heading("附录四  每周复盘", level=1)
for q in [
    "本周我完成得最好的一件事是什么？为什么？",
    "本周出现的一个错误或低效点是什么？根因在哪里？",
    "我学到的一项可以复用的方法是什么？",
    "我向团队留下了什么模板、清单或操作说明？",
    "下周我希望尝试哪一项更复杂的工作？",
    "这周的经历让我对未来职业方向有什么新的认识？",
]:
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(5)
    r = para.add_run(q)
    font(r, 10.5, True, NAVY)
    for _ in range(2):
        p("________________________________________________________________________________", after=2)

doc.add_heading("附录五  知识库条目模板", level=1)
add_table(["字段", "填写内容"], [
    ("条目名称", "简洁说明要解决的问题"),
    ("适用场景", "什么时候使用；不适用于什么情况"),
    ("前置条件", "需要的账号、材料、权限或设备"),
    ("操作步骤", "按顺序写明动作，必要时配截图"),
    ("常见错误", "容易遗漏、失败或产生风险的地方"),
    ("文件位置", "共享路径、模板名称或关联资料"),
    ("维护信息", "编写人、审核人、最后更新时间"),
], [1.3, 5.25])

doc.add_heading("附录六  实习成果要求", level=1)
bullets([
    "完成个人实习报告：记录参与事项、能力变化、典型错误、职业认识和后续计划。",
    "参与完成一项可公开或内部发布的内容成果：选题、事实和法律内容须经团队审核；涉及案件时严格脱敏并取得许可。交稿用于训练复盘和表达，是否公开、怎样署名及是否使用个人照片应尊重本人意愿。",
    "至少完善一项内部工作手册内容：设备操作、办公流程、文件规则、检查清单或知识库条目。",
    "发现共享资料中无适用模板时，整理候选模板并提交业务负责人和知识库维护人审核，不自行标记为标准范文。",
    "进行一次离岗交接：说明未完事项、文件位置、账号和设备归还、保密材料清理情况。",
])

doc.add_heading("附录七  通用放行型SOP模板", level=1)
p("适用于立案、邮寄、对外发送、证据装订、拍摄、客户接待、AI成果使用及其他具有期限、不可逆或对外影响的事项。具体业务应根据风险删减或增加项目。")
add_table(["模块", "填写内容"], [
    ("事项信息", "事项名称、日期、操作人、复核人、最终使用者"),
    ("目标结果", "完成后必须达到的状态，以及怎样客观验证"),
    ("开工条件", "权限、材料、设备、版本、时间和人员是否齐备"),
    ("★红线", "任何一项未完成即不得继续的条件"),
    ("最小样本", "先完成并验收的一份、一页、一组或一次试运行"),
    ("执行检查", "关键步骤、完成勾选、异常及处理"),
    ("对外放行", "收件对象、最终版本、附件、签章、保密、负责人确认"),
    ("收工条件", "归档、备份、原件、设备、现场、交接是否完成"),
    ("工作日志", "开始结束时间、数量、存储位置、异常、补救、下次改进"),
    ("最终状态", "□ 放行  □ 暂缓整改  □ 取消；操作人及复核人签字"),
], [1.3, 5.25])

doc.add_heading("附录八  管理者交办与带教自检", level=1)
p("本表适用于所有向实习生交办任务的管理者、带教者和团队成员。实习生未达标时，应同时检查管理环节是否存在缺口。")
add_table(["阶段", "管理者自检", "确认"], [
    ("交办前", "目的、成果形式、标准、模板、权限、期限、风险和验收人已经明确。", "□"),
    ("理解确认", "实习生已用自己的话复述任务；疑问和冲突要求已经澄清。", "□"),
    ("样本检查", "批量或高风险工作已先验收最小样本。", "□"),
    ("过程管理", "已设置必要检查点；实习生知道何种情况必须立即报告。", "□"),
    ("结果验收", "已明确通过或退回，并指出事实、影响和具体修改要求。", "□"),
    ("培养升级", "基础工作稳定后，已安排难度更高且有学习价值的任务。", "□"),
    ("知识沉淀", "本次新经验、错误或例外已进入清单、模板或知识库。", "□"),
    ("权威边界", "严格针对事实、程序与风险，没有靠地位、畏惧或人格控制换取服从。", "□"),
    ("关怀边界", "帮助指向成长与独立，没有以照顾交换忠诚、加班、隐私或沉默。", "□"),
    ("德行示范", "同一规则也约束自己；不同人员的差异安排有与任务相关、可以说明的理由。", "□"),
    ("自我复盘", "是否因交办模糊、多头指令、反馈延迟或标准变化造成返工？", "□"),
], [1.15, 4.85, 0.55])

doc.add_heading("附录九  新人校准会议议程", level=1)
p("建议在入职第一至第二天召开；出现多人都觉得“难带”、任务理解分歧或配合明显下降时，也应及时召开。参会人员包括实习生、主要交办人与带教者，建议控制在35至45分钟。")
add_table(["时间", "议程", "主持提示"], [
    ("3分钟", "说明会议宗旨和当前需要解决的问题", "针对协作事实，不给个人贴标签"),
    ("6分钟", "带教者、协作成员和实习生简要自我介绍", "每人约30至60秒，说明职责与专长"),
    ("8分钟", "说明实习宗旨、首周考核和阶段任务", "明确什么时间达到什么标准"),
    ("9分钟", "各带教者提出核心要求", "每人只讲最关键的三点，现场消除冲突"),
    ("6分钟", "实习生表达需求、困惑和职业期待", "每人均获得发言机会，不因问题简单而打断"),
    ("8分钟", "确认下一步任务、负责人、样本和期限", "形成书面行动清单及下次复盘时间"),
], [0.95, 3.1, 2.45])
bullets([
    "会前：收集团队认为的具体问题，不使用“态度差”“不好带”等无法行动的概括代替事实。",
    "会中：主持人负责控制时间、归纳分歧，并确保实习生也能表达。",
    "会后：发布简要纪要，列出责任人、期限、检查点和需要更新的手册内容。",
    "复盘：三至五个工作日后检查配合度、完成度、主动性和具体错误是否发生变化。",
])

doc.add_heading("附录十  历届实习心得阅读索引", level=1)
p("以下文章均发布于公众号“没事多吃核桃”，标题可点击打开。建议把它们当作不同起点的经验样本，阅读时关注具体任务、错误、反馈和能力变化，不把任何一篇的语气或结论当作未来实习生必须模仿的标准答案。")

doc.add_heading("法律现场与职业认识", level=2)
linked_bullet("《初次旁听庭审，好激动呀》", "https://mp.weixin.qq.com/s/9cRjuC24GyRrhRYt2_wHug", "观察仲裁流程、证据与争议焦点，比较真实庭审和影视想象。")
linked_bullet("《律所实习初体验》", "https://mp.weixin.qq.com/s/70rtliJ2Ntfo2tlskVeUeA", "从专业、协作、时间管理和职业道德理解律师工作。")
linked_bullet("《律所实习心得：学无止境，始于足下》", "https://mp.weixin.qq.com/s/6D6WGX1Hk5_wAzrWVeDb4A", "从基础事务看见律师真实工作量，以及提问、复盘和执行力。")
linked_bullet("《律所实习体验心得：行是知之始，知是行之成》", "https://mp.weixin.qq.com/s/ocqJtVvjPOtdGo0KvmsFmQ", "记录从页码错误到独立整理卷宗，以及设备问题应何时求助。")

doc.add_heading("从小事到可靠", level=2)
linked_bullet("《在不限零食的律所里工作是种什么体验？》", "https://mp.weixin.qq.com/s/G2QTuqmkEUf8RD7LjdhPsA", "观察工作记录、熟练度、返工减少和日常技能怎样累积成信心。")
linked_bullet("《张如意寒假实习总结》", "https://mp.weixin.qq.com/s/DW2od7LKX8N0CPhWabgZvw", "思考一次问清、可靠习惯、反复实践和向有经验者学习。")
linked_bullet("《寒假实习生陈莹实习总结》", "https://mp.weixin.qq.com/s/AAk6JgKd9kl8OCzx07Z1uQ", "思考个人能为团队带来什么，以及平台、尊重和成长时间。")
linked_bullet("《廖心怡暑假实习总结》", "https://mp.weixin.qq.com/s/J4bnka9zLIa--hRJDy9Niw", "反思任务目的、准确表达问题、主动请教和观察真实需求。")
linked_bullet("《职场菜鸟在律所实习是种什么体验？》", "https://mp.weixin.qq.com/s/Ovb6YVpRdTCHUuJAbZJx-Q", "从页码、顺序和具体反馈理解严谨、阅读写作与行动。")
linked_bullet("《记暑假律所实习心得》", "https://mp.weixin.qq.com/s/k9i8GH2XOOw7yP6s7VuaWA", "从非法律专业视角理解基础事务、办公软件、主动提问和职业适应。")

doc.add_heading("跨专业与内容创作", level=2)
linked_bullet("《在律所剪片子、做设计：我的新鲜体验与成长》", "https://mp.weixin.qq.com/s/OHRfieulR5-LFk8dtkFJTA", "看证据整理、名片、宣传册和视频如何共同训练信息表达与专业严谨。")
linked_bullet("《一场名为“成长”的小鲤鱼历险记》", "https://mp.weixin.qq.com/s/x_q8x0s_fh9pSj81h1fJyQ", "看新闻传播背景如何通过自学、设备、内容核验和SOP形成跨界能力。")

quote("读前人的心得，不是抄答案，而是提前发现问题。", "给后来者")

appendix_eleven_heading = doc.add_heading("附录十一  招聘—入职衔接卡（管理者使用）", level=1)
appendix_eleven_heading.paragraph_format.page_break_before = True
p("这张卡不负责给候选人贴标签，只把招聘中已经确认的事实翻译成首周安排。招聘负责人整理，首周带教者接收；涉及本人安排和期待的内容，由候选人在入职时确认或更正。BOSS直聘只是其中一个入口。")
add_table(["字段", "只记录什么", "怎样转成培养动作"], [
    ("岗位与团队需要", "本次为什么招人，希望解决哪一类真实问题；法律、内容、视觉或复合方向。", "确定一条主要任务线，避免所有人都被当作无差别杂务人手。"),
    ("现实条件", "是否在校、实习起止、连续天数、返校安排、实际线下地点、常规作息，以及补贴食宿的已确认边界。", "据此安排周期与工作量；尚未确认的条件标为“待确认”，不得写成承诺。"),
    ("相关起点证据", "与岗位有关的课程、项目、实习、作品、工具和可迁移能力；注明证据来源。", "选择一个能发挥已有优势的首个最小任务，帮助新人建立真实胜任感。"),
    ("短时工作样本", "任务题目、预计用时、是否使用AI、事实与结构表现、主要错误及第二版变化。", "把错误转成首周检查点；表现良好的环节可减少重复测试，尽快进入真实训练。"),
    ("学习敏捷", "学习一段新规则或新工具后，能否复述、应用、发现不确定点并验证。", "决定电脑、共享存储、AI和SOP训练的示范深度与复查频率。"),
    ("反馈与协作", "收到具体意见后的反应和修订；任务复述、进度反馈、提问及风险报告方式。", "为其确定合适的沟通频率、样本检查点和安全提问渠道。"),
    ("候选人期待", "希望学习的内容、想验证的职业方向、最担心的现实问题。", "首周校准会议主动回应；无法满足的特殊期待尽早说明。"),
    ("已确认与未定", "只区分已确认事实、当前推断和仍待验证事项，不写人格结论。", "未定维度用任务验证，不靠重复追问或团队成员相互猜测。"),
    ("首周培养建议", "共同底座、优势入口、需要补练的能力、首个可交付成果及负责人。", "直接进入首周训练地图和任务交付卡，明确样本、期限和验收人。"),
    ("风险与边界", "保密、利益冲突、设备权限、工作时间和需要立即报告的情形。", "在开放资料或权限前完成说明与确认；风险越高，授权越应分级。"),
], [1.15, 3.1, 2.4])
bullets([
    "工作样本应控制在约10至20分钟，不含真实客户秘密，不把候选人的免费测试成果直接用于商业或案件交付。",
    "记录使用“做了什么、产生什么结果、怎样修订”的行为语言，不使用“聪明”“不行”“性格怪”等无法验证的标签。",
    "学校、专业和成绩只作为理解训练起点的一部分；当前工作证据与入职后的实际进步应当拥有更高权重。",
    "招聘记录应注明负责人、接收带教者、保存位置、访问权限和清理时间；候选人未录用或不再推进后，限制留存、避免扩散，并清理没有继续保存必要的副本。",
])
quote("招聘记录不替未来下结论，只帮助带教者少走弯路，也给新人推翻旧判断的机会。", "招聘与培养衔接")

doc.add_heading("结语  不要碌碌无为", level=1)
p("年轻时最容易被浪费的，不是某一天，而是逐渐习惯于只完成最低要求。", style="Lead")
p("你不必一开始就知道自己会成为什么，也不必事事完美。先把眼前小事做出信用，再追问它把你带向哪里；接受经验和规则，也保留独立判断。实习只有几周，做事的方法会跟你很多年。")
quote("实习证明只说明你来过；做事的方法，才会跟你走。", "送给每一位实习生")

doc.add_heading("思想来源与延伸阅读", level=2)
p("本手册在文风与成长主题上借鉴《上海交通大学生存手册》；清单、反向思考、多学科模型、激励与持续学习参考《穷查理宝典》；“用心、用情、用力”和新人校准参考工作投入、组织社会化研究；管理者的尚严、仁慈、德行及边界参考家长式领导研究；AI篇参考邱懿武的课程与公开知识体系。以上均结合团队实践转述，不构成原文摘录。")
p("延伸阅读：survivesjtu.gitbook.io/survivesjtumanual ｜ press.stripe.com/poor-charlies-almanack ｜ 林姿葶、郑伯埙、周丽芳：《家长式领导之回顾与前瞻：再一次思考》，《管理学季刊》2017年第4期 ｜ Kahn, Academy of Management Journal 33(4), 1990 ｜ Bauer et al., Journal of Applied Psychology 92(3), 2007", italic=True)
source_links = doc.add_paragraph()
r = source_links.add_run("课程资料：")
font(r, 10.5, bold=True)
add_hyperlink(source_links, "邱懿武《商业中的人工智能》", "https://qiuyiwu.com/zhongda/")
r = source_links.add_run(" ｜ ")
font(r, 10.5, color=MUTED)
add_hyperlink(source_links, "《AI思维与认知》", "https://qiuyiwu.com/ai-cognition/")

# Document metadata
doc.core_properties.title = "实习生生存指南（公开阅读版）"
doc.core_properties.subject = "实习生培养、工作规范与成长指南"
doc.core_properties.author = "胡泽贤（George Hu）"
doc.core_properties.keywords = "胡泽贤, George Hu, 律师团队, 实习生, 生存指南, 招聘, 入职衔接, 培养, 保密, 工作手册, SOP, 清单, 放行机制, 家长式领导, 穷查理宝典, 组织行为学, 工作投入, 新人校准会议"
doc.core_properties.comments = "面向多种专业团队的实习生培养、工作规范与成长指南公开阅读版。"
doc.core_properties.category = "团队培养与管理"
doc.core_properties.language = "zh-CN"
doc.core_properties.last_modified_by = "胡泽贤（George Hu）"
doc.core_properties.created = datetime(2026, 7, 14, 19, 0, 0)
doc.core_properties.modified = datetime(2026, 7, 14, 19, 0, 0)
doc.core_properties.version = "2026.07"

# 发布前的少量硬性守门项。文风不能完全靠程序判断，但这些已经确认过的
# 命名错误和版本噪音不应在后续修订中反复出现。
visible_text = "\n".join(
    [paragraph.text for paragraph in doc.paragraphs]
    + [cell.text for table in doc.tables for row in table.rows for cell in row.cells]
)
for forbidden_text in ("HU ZE XIAN", "GEORGE HU LEGAL TEAM", "Workbench", "融合版"):
    if forbidden_text in visible_text:
        raise ValueError(f"发布检查失败：仍包含 {forbidden_text!r}")
if "INTERN SURVIVAL GUIDE" not in visible_text:
    raise ValueError("发布检查失败：首页或篇章页缺少统一英文眉题")

doc.save(str(OUT))
print(OUT)


# Sanity checks on the generated document.
def _visible_text(path):
    document = Document(path)
    return "\n".join(
        [paragraph.text for paragraph in document.paragraphs]
        + [cell.text for table in document.tables for row in table.rows for cell in row.cells]
    )


_public_text = _visible_text(OUT)
if "NAS" in _public_text:
    raise ValueError("Public Chinese edition still contains NAS")
if "内部学习资料" in _public_text:
    raise ValueError("Public Chinese edition still carries the old edition label")
