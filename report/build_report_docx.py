from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image


ROOT = Path(__file__).resolve().parent.parent
REPORT_MD_PATH = ROOT / "report" / "report.md"
ASSETS_DIR = ROOT / "report" / "assets"
DOCX_PATH = ROOT / "report" / "股指期货套利策略研究报告.docx"
DOCX_FALLBACK_PATH = ROOT / "report" / "股指期货套利策略研究报告_v2.docx"

BODY_COLOR = RGBColor(0x22, 0x22, 0x22)
MUTED_COLOR = RGBColor(0x55, 0x55, 0x55)
ACCENT_COLOR = RGBColor(0x00, 0x00, 0x00)
HEADER_FILL = "EDEDED"
ROW_FILL = "F7F7F7"
INFO_LEFT_FILL = "EEEEEE"
INFO_RIGHT_FILL = "FAFAFA"
BODY_LINE_SPACING_PT = 20
MAX_IMAGE_WIDTH_CM = 15.2
MAX_IMAGE_HEIGHT_CM = 18.8

FIGURE_CAPTIONS = {
    "fig_roll_yield.png": (
        "展期收益分布",
        "展示每段主力持有周期的累计收益分布。IC 和 IM 的均值都为正，说明样本期内展期机制确实带来补偿，但正收益占比仅略高于一半，单次换仓仍有很大不确定性。",
    ),
    "fig_nav.png": (
        "策略净值曲线",
        "比较吃贴水策略与现货指数的净值走势。Always 策略在 IC 和 IM 上都明显高于对应指数基准，但不是平滑套利曲线，而是跟随市场波动的大幅起伏。",
    ),
    "fig_annual.png": (
        "年度收益拆分",
        "按自然年展示策略和基准的收益对比。年度收益波动很大，基差阈值策略在某些下跌阶段有帮助，但空仓会降低回撤也会错过部分上涨。",
    ),
    "fig_drawdown.png": (
        "回撤曲线",
        "比较策略与指数的回撤深度和修复过程。IC Always 最大回撤 -61.38%，IM Always -50.81%，超过多数绝对收益资金可接受范围，更适合作为方向性配置工具。",
    ),
    "fig_heatmap.png": (
        "参数敏感性热力图",
        "展示基差阈值与展期窗口对 Sharpe 和 MaxDD 的影响。展期窗口从 1 天到 5 天影响不大，基差阈值影响更明显，但 IM 样本期短，不能将阈值当作稳定最优参数。",
    ),
    "fig_oos_validation.png": (
        "样本外验证",
        "IC 按 60/40 时间划分的训练期和测试期净值对比。测试期表现好于训练期，但测试期最大回撤仍达 -54.12%，不应解读为策略已通过样本外检验。IM 仅 872 个交易日，样本过短无法形成稳定滚动窗口，不做样本外拆分。",
    ),
    "fig_regime_segmentation.png": (
        "市场状态分段分析",
        "按市场方向、贴水深度、波动率水平三个维度将交易日分段，展示各子集 Always 策略绩效。策略首先是权益方向暴露，贴水深度也有一定信息但不是独立的择时开关。",
    ),
    "fig_risk_control.png": (
        "风险控制策略净值",
        "比较原始策略、波动率目标 15% 和回撤控制 20% 三个版本的净值走势。风控降低了回撤但收益同步被削弱，策略的贴水补偿本身不够厚，经不起降杠杆和频繁调仓的摩擦。",
    ),
    "fig_spread.png": (
        "跨期价差与布林带",
        "展示近月-次近月价差序列及其滚动均值和布林带边界。价差突破布林带后并不总是快速回归，日频下价差回归幅度较小，扣除双边成本后很难留下足够利润。",
    ),
    "fig_spread_failure.png": (
        "跨期套利信号拆解",
        "展示代表性布林带参数下的信号分布和 T+1 执行后的实际收益。日频收盘价下可捕捉偏离太少，T+1 执行进一步错过短暂回归机会。",
    ),
    "fig_bootstrap.png": (
        "Bootstrap 显著性检验",
        "展示策略相对指数 Sharpe 差异的 Block Bootstrap 重采样分布。两个品种的置信区间都覆盖 0，不能认为策略在统计意义上显著优于指数。",
    ),
    "fig_attribution.png": (
        "收益归因分解（无截距）",
        "将策略收益拆解为 beta 驱动的指数收益和残差两部分。Beta 很低，R² 接近 0，日收益中包含大量换仓跳跃、贴水收敛和期货自身价格路径的影响。",
    ),
    "fig_capm_attribution.png": (
        "有截距 CAPM 归因",
        "加入截距项后的 CAPM 归因结果。日频下 α 为正但 t 值均低于 2，不能认为有稳定独立 alpha。月频回归 Beta 约 0.67-0.76，更符合金融直觉。",
    ),
    "fig_cost_stress.png": (
        "成本压力测试",
        "将交易成本从 1 倍逐级放大到 10 倍，观察策略年化收益的变化。策略换手低，对交易成本不敏感，成本不是主要矛盾。",
    ),
    "fig_correlation.png": (
        "IC-IM 日收益相关性",
        "IC 与 IM 吃贴水策略日收益的散点图和回归线。相关性约 0.95，二者高度同涨同跌，简单组合不能形成有效分散。",
    ),
}

TABLE_CAPTIONS = [
    ("研究口径与核心设定", "概括报告的研究标的、数据区间、交易假设和评价重点。"),
    ("数据质量检查", "对比 IC 与 IM 的原始样本量、主力合约样本量、指数匹配率和合约切换统计。"),
    ("回测假设与交易口径", "统一所有策略的信号时点、成交时点、合约选择、展期规则和成本设定。"),
    ("基差率统计特征与预测能力", "ADF 均值回复检验、半衰期、基差对期货和指数前瞻收益的预测 IC。"),
    ("展期收益分布统计", "IC 与 IM 各展期段的累计收益分布，包含均值、中位数和正收益占比。"),
    ("吃贴水策略主要绩效", "比较 Always、B<-1.5% 和指数基准在 IC 与 IM 上的收益、风险与交易统计。"),
    ("样本外验证结果", "IC 按时间划分的训练期与测试期绩效对比。"),
    ("市场状态分段绩效", "按市场方向、贴水深度和波动率水平三个维度分段的 Always 策略绩效。"),
    ("风险控制版本绩效对比", "比较原始策略、波动率目标和回撤控制版本的收益与回撤。"),
    ("跨期套利信号统计", "展示跨期套利在代表性布林带参数下的信号触发率与方向正确率。"),
    ("Bootstrap 显著性检验", "Block Bootstrap 重采样下的 Sharpe 差异分布、置信区间和 p 值。"),
    ("无截距收益归因", "将策略日收益按无截距模型拆解为指数驱动和残差两部分。"),
    ("有截距 CAPM 归因（日频）", "加入截距项后的日频 CAPM 归因，α 衡量扣除指数 beta 后的平均日超额。"),
    ("有截距 CAPM 归因（月频）", "用月度收益替代日收益进行有截距 CAPM 回归，作为日频结果的稳健性检查。"),
    ("图表清单", "报告中全部 15 张图表的编号、文件名和内容说明。"),
]


def set_run_font(
    run,
    size_pt: float = 12,
    *,
    bold: bool = False,
    italic: bool = False,
    color: RGBColor | None = None,
    latin_font: str = "Times New Roman",
    east_asia_font: str = "宋体",
):
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size_pt)
    run.font.name = latin_font
    run.font.color.rgb = color or BODY_COLOR

    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), latin_font)
    rfonts.set(qn("w:hAnsi"), latin_font)
    rfonts.set(qn("w:eastAsia"), east_asia_font)


def clean_inline(text: str) -> str:
    text = text.replace("<br>", " ").replace("<br/>", " ").replace("<br />", " ")
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    return text


def add_inline_runs(
    paragraph,
    text: str,
    *,
    size_pt: float = 12,
    base_bold: bool = False,
    italic: bool = False,
    color: RGBColor | None = None,
    east_asia_font: str = "宋体",
):
    text = clean_inline(text)
    token_re = re.compile(r"(\*\*.+?\*\*|`.+?`)")
    pos = 0
    for match in token_re.finditer(text):
        if match.start() > pos:
            run = paragraph.add_run(text[pos : match.start()])
            set_run_font(
                run, size_pt=size_pt, bold=base_bold, italic=italic,
                color=color, east_asia_font=east_asia_font,
            )

        token = match.group(0)
        if token.startswith("**"):
            run = paragraph.add_run(token[2:-2].replace("`", ""))
            set_run_font(
                run, size_pt=size_pt, bold=True, italic=italic,
                color=color, east_asia_font=east_asia_font,
            )
        else:
            run = paragraph.add_run(token[1:-1])
            set_run_font(
                run, size_pt=size_pt - 0.5, color=MUTED_COLOR,
                latin_font="Consolas", east_asia_font="宋体",
            )
        pos = match.end()

    if pos < len(text):
        run = paragraph.add_run(text[pos:])
        set_run_font(
            run, size_pt=size_pt, bold=base_bold, italic=italic,
            color=color, east_asia_font=east_asia_font,
        )


def set_paragraph_bottom_border(paragraph, color: str = "808080", size: str = "8"):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = p_bdr.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        p_bdr.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)


def set_cell_shading(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top: int = 80, bottom: int = 80, left: int = 80, right: int = 80):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in {"top": top, "bottom": bottom, "left": left, "right": right}.items():
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_compact_spacing(paragraph, *, before: float = 0, after: float = 0):
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    paragraph.paragraph_format.line_spacing = 1
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(0)

    head = paragraph.add_run("第 ")
    set_run_font(head, size_pt=9.5, color=MUTED_COLOR)

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    field_run = paragraph.add_run()
    set_run_font(field_run, size_pt=9.5, color=MUTED_COLOR)
    field_run._r.append(begin)
    field_run._r.append(instr)
    field_run._r.append(end)

    tail = paragraph.add_run(" 页")
    set_run_font(tail, size_pt=9.5, color=MUTED_COLOR)


def add_header(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run("股指期货套利策略研究")
    set_run_font(run, size_pt=9.5, color=MUTED_COLOR)


def configure_document(doc: Document):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.font.color.rgb = BODY_COLOR
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.4)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.55)
    section.right_margin = Cm(2.55)

    add_header(section.header.paragraphs[0])
    add_page_number(section.footer.paragraphs[0])


def add_horizontal_rule(doc: Document, color: str = "000000", size: str = "10"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(10)
    set_paragraph_bottom_border(p, color=color, size=size)


def add_cover_page(doc: Document):
    for _ in range(5):
        doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("股指期货套利策略研究报告")
    set_run_font(run, size_pt=24, bold=True, color=ACCENT_COLOR, east_asia_font="黑体")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run("吃贴水与跨期套利的策略评估与风险拆解")
    set_run_font(run, size_pt=15, color=MUTED_COLOR, east_asia_font="黑体")

    add_horizontal_rule(doc)

    items = [
        ("研究标的", "IC（中证500股指期货）与 IM（中证1000股指期货）"),
        ("数据区间", "IC：2015-04-16 至 2026-03-02；IM：2022-07-22 至 2026-03-02"),
        ("策略口径", "吃贴水多头 + 基差阈值择时 + 跨期价差套利"),
        ("评估维度", "收益、风险、统计显著性、交易可执行性、可复现性"),
    ]
    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for left, right in items:
        row = table.add_row().cells
        set_cell_shading(row[0], INFO_LEFT_FILL)
        set_cell_shading(row[1], INFO_RIGHT_FILL)
        for cell in row:
            set_cell_margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p0 = row[0].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(left)
        set_run_font(r0, size_pt=11, bold=True, east_asia_font="黑体")
        p1 = row[1].paragraphs[0]
        add_inline_runs(p1, right, size_pt=11)

    for _ in range(7):
        doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("课程项目报告")
    set_run_font(run, size_pt=12, color=MUTED_COLOR)

    doc.add_page_break()


def add_navigation_page(doc: Document):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run("报告结构")
    set_run_font(run, size_pt=18, bold=True, east_asia_font="黑体")

    sections = [
        "摘要：核心结论与收益口径",
        "1. 数据与回测设定",
        "2. 基差因子",
        "3. 吃贴水策略",
        "4. 跨期套利",
        "5. IC-IM 组合策略",
        "6. 统计检验与收益归因",
        "7. 补充分析",
        "8. 风险提示",
        "9. 结论",
        "附录：图表清单",
    ]
    for item in sections:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(2.0)
        p.paragraph_format.space_after = Pt(4)
        add_inline_runs(p, item, size_pt=12.5, base_bold=True)

    doc.add_page_break()


def add_heading(doc: Document, text: str, level: int):
    if level == 2 and re.match(r"^\d+\.\s", text) and not text.startswith("1. "):
        doc.add_section(WD_SECTION_START.NEW_PAGE)

    p = doc.add_paragraph()
    if level == 1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(16)
        add_inline_runs(p, text, size_pt=20, base_bold=True, east_asia_font="黑体")
        return

    p.paragraph_format.space_before = Pt(12 if level == 2 else 8)
    p.paragraph_format.space_after = Pt(6 if level in (2, 3) else 3)
    if level == 2:
        add_inline_runs(p, text, size_pt=14.5, base_bold=True, east_asia_font="黑体")
        set_paragraph_bottom_border(p)
    elif level == 3:
        add_inline_runs(p, text, size_pt=12.5, base_bold=True, east_asia_font="黑体")
    else:
        add_inline_runs(p, text, size_pt=11.5, base_bold=True, color=MUTED_COLOR, east_asia_font="黑体")


def add_text_paragraph(doc: Document, text: str, *, center: bool = False, size_pt: float = 12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(BODY_LINE_SPACING_PT)
    p.paragraph_format.space_after = Pt(6)
    if not center:
        p.paragraph_format.first_line_indent = Cm(0.74)
    add_inline_runs(p, text, size_pt=size_pt)


def add_lead_line(doc: Document, text: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    add_inline_runs(p, text, size_pt=11.5, italic=True, color=MUTED_COLOR)


def add_bullet(doc: Document, text: str):
    p = doc.add_paragraph(style="List Bullet")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(BODY_LINE_SPACING_PT)
    p.paragraph_format.space_after = Pt(3)
    add_inline_runs(p, text, size_pt=12)


def add_numbered_item(doc: Document, number: str, text: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(0.74)
    p.paragraph_format.first_line_indent = Cm(-0.44)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(BODY_LINE_SPACING_PT)
    p.paragraph_format.space_after = Pt(3)
    add_inline_runs(p, f"{number}. {text}", size_pt=12)


def add_caption(doc: Document, text: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_compact_spacing(p, after=6)
    add_inline_runs(p, text, size_pt=10.5, base_bold=True, color=MUTED_COLOR)


def add_figure_caption(doc: Document, figure_no: int, title: str, note: str):
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.keep_together = True
    title_p.paragraph_format.keep_with_next = True
    set_compact_spacing(title_p, after=1)
    add_inline_runs(title_p, f"图 {figure_no} {title}", size_pt=10.5, base_bold=True, color=ACCENT_COLOR)

    note_p = doc.add_paragraph()
    note_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    note_p.paragraph_format.left_indent = Cm(0.75)
    note_p.paragraph_format.right_indent = Cm(0.75)
    note_p.paragraph_format.keep_together = True
    set_compact_spacing(note_p, after=8)
    add_inline_runs(note_p, f"说明：{note}", size_pt=10, color=MUTED_COLOR)


def add_table_caption(doc: Document, table_no: int):
    if table_no <= len(TABLE_CAPTIONS):
        title, note = TABLE_CAPTIONS[table_no - 1]
    else:
        title, note = "补充表格", "汇总报告正文中的补充数据。"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    set_compact_spacing(p, before=4, after=4)
    add_inline_runs(p, f"表 {table_no} {title}：{note}", size_pt=10.5, base_bold=True, color=MUTED_COLOR)


def add_source_line(doc: Document, text: str = "资料来源：IC/IM 期货与指数数据，本报告计算"):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_compact_spacing(p, before=0, after=6)
    add_inline_runs(p, text, size_pt=8.5, color=MUTED_COLOR)


def resolve_image(rel_path: str) -> Path:
    raw = rel_path.strip().replace("\\", "/")
    path = REPORT_MD_PATH.parent / raw
    if path.exists():
        return path
    path = ASSETS_DIR / Path(raw).name
    if path.exists():
        return path
    raise FileNotFoundError(f"Image referenced in report.md was not found: {rel_path}")


def image_display_size_cm(img_path: Path) -> tuple[float, float | None]:
    with Image.open(img_path) as img:
        width_px, height_px = img.size
    height_at_max_width = MAX_IMAGE_WIDTH_CM * height_px / width_px
    if height_at_max_width <= MAX_IMAGE_HEIGHT_CM:
        return MAX_IMAGE_WIDTH_CM, None
    width_at_max_height = MAX_IMAGE_HEIGHT_CM * width_px / height_px
    return width_at_max_height, MAX_IMAGE_HEIGHT_CM


def add_image(doc: Document, img_path: Path, figure_no: int, alt_text: str = ""):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_together = True
    p.paragraph_format.keep_with_next = True
    set_compact_spacing(p, before=4, after=2)
    run = p.add_run()
    width_cm, height_cm = image_display_size_cm(img_path)
    if height_cm is None:
        run.add_picture(str(img_path), width=Cm(width_cm))
    else:
        run.add_picture(str(img_path), height=Cm(height_cm))

    title, note = FIGURE_CAPTIONS.get(
        img_path.name,
        (alt_text or img_path.stem, "展示报告正文对应分析结果。"),
    )
    add_figure_caption(doc, figure_no, title, note)
    add_source_line(doc)


def is_table_block(lines: list[str], idx: int) -> bool:
    if idx + 1 >= len(lines):
        return False
    current = lines[idx].strip()
    nxt = lines[idx + 1].strip()
    return current.startswith("|") and nxt.startswith("|") and set(nxt.replace("|", "").strip()) <= {"-", ":", " "}


def split_table_row(line: str) -> list[str]:
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def parse_markdown_table(lines: list[str], idx: int):
    header = split_table_row(lines[idx])
    rows = []
    i = idx + 2
    while i < len(lines):
        line = lines[i].strip()
        if not line.startswith("|"):
            break
        row = split_table_row(line)
        if len(row) < len(header):
            row += [""] * (len(header) - len(row))
        rows.append(row[: len(header)])
        i += 1
    return header, rows, i


def sanitize_table_text(text: str) -> str:
    text = clean_inline(text)
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    return text.replace("`", "")


def add_table(doc: Document, header: list[str], rows: list[list[str]], table_no: int):
    add_table_caption(doc, table_no)
    table = doc.add_table(rows=1, cols=len(header))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    cols = len(header)
    font_size = 10.5 if cols <= 3 else 9.5 if cols <= 6 else 8.5
    is_info_table = [sanitize_table_text(x) for x in header] == ["项目", "内容"] or \
                    [sanitize_table_text(x) for x in header] == ["项目", "口径"]

    for j, text in enumerate(header):
        cell = table.rows[0].cells[j]
        set_cell_shading(cell, HEADER_FILL)
        set_cell_margins(cell, top=70, bottom=70, left=60, right=60)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = cell.paragraphs[0]
        set_compact_spacing(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_inline_runs(
            p,
            sanitize_table_text(text),
            size_pt=font_size,
            base_bold=True,
            east_asia_font="黑体",
        )

    for row_idx, row in enumerate(rows):
        cells = table.add_row().cells
        for j, text in enumerate(row):
            cell = cells[j]
            set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if is_info_table:
                set_cell_shading(cell, INFO_LEFT_FILL if j == 0 else INFO_RIGHT_FILL)
            elif row_idx % 2 == 1:
                set_cell_shading(cell, ROW_FILL)
            p = cell.paragraphs[0]
            set_compact_spacing(p)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            add_inline_runs(p, sanitize_table_text(text), size_pt=font_size)

    spacer = doc.add_paragraph()
    set_compact_spacing(spacer, after=2)
    add_source_line(doc)


def build_docx() -> Path:
    if not REPORT_MD_PATH.exists():
        raise FileNotFoundError(f"Cannot find {REPORT_MD_PATH}")

    lines = REPORT_MD_PATH.read_text(encoding="utf-8").splitlines()
    doc = Document()
    configure_document(doc)
    add_cover_page(doc)
    add_navigation_page(doc)

    paragraph_buffer: list[str] = []
    figure_no = 0
    table_no = 0

    def flush_paragraph():
        nonlocal paragraph_buffer
        if paragraph_buffer:
            text = " ".join(x.strip() for x in paragraph_buffer if x.strip())
            if text:
                add_text_paragraph(doc, text)
            paragraph_buffer = []

    i = 0
    while i < len(lines):
        stripped = lines[i].strip()

        if not stripped:
            flush_paragraph()
            i += 1
            continue

        if stripped in {"---", "***", "___"}:
            flush_paragraph()
            i += 1
            continue

        if stripped.startswith("# "):
            flush_paragraph()
            add_heading(doc, stripped[2:].strip(), 1)
            i += 1
            continue

        if stripped.startswith("## "):
            flush_paragraph()
            add_heading(doc, stripped[3:].strip(), 2)
            i += 1
            continue

        if stripped.startswith("### "):
            flush_paragraph()
            add_heading(doc, stripped[4:].strip(), 3)
            i += 1
            continue

        if stripped.startswith("#### "):
            flush_paragraph()
            add_heading(doc, stripped[5:].strip(), 4)
            i += 1
            continue

        if stripped.startswith("> "):
            flush_paragraph()
            add_lead_line(doc, stripped[2:].strip())
            i += 1
            continue

        if is_table_block(lines, i):
            flush_paragraph()
            header, rows, next_idx = parse_markdown_table(lines, i)
            table_no += 1
            add_table(doc, header, rows, table_no)
            i = next_idx
            continue

        image_match = re.match(r"!\[(.*?)\]\((.*?)\)", stripped)
        if image_match:
            flush_paragraph()
            alt_text = image_match.group(1).strip()
            img_path = resolve_image(image_match.group(2))
            figure_no += 1
            add_image(doc, img_path, figure_no, alt_text=alt_text)
            i += 1
            continue

        if stripped.startswith("- "):
            flush_paragraph()
            add_bullet(doc, stripped[2:].strip())
            i += 1
            continue

        number_match = re.match(r"^(\d+)\.\s+(.*)", stripped)
        if number_match:
            flush_paragraph()
            add_numbered_item(doc, number_match.group(1), number_match.group(2).strip())
            i += 1
            continue

        paragraph_buffer.append(stripped)
        i += 1

    flush_paragraph()

    try:
        doc.save(DOCX_PATH)
        return DOCX_PATH
    except PermissionError:
        doc.save(DOCX_FALLBACK_PATH)
        return DOCX_FALLBACK_PATH


if __name__ == "__main__":
    output = build_docx()
    print(f"Generated {output.relative_to(ROOT)}")
