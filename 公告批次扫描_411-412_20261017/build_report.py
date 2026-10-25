# -*- coding: utf-8 -*-
"""Build the batch 411/412 NEV passenger scan report."""

from __future__ import annotations

import json
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
PAYLOAD = json.loads((ROOT / "data" / "delivery_payload.json").read_text(encoding="utf-8"))
ASSETS = ROOT / "assets"
OUTPUT_DIR = ROOT / "outputs" / "01a147a8-938a-7ff3-a15c-94b7fcac7a5e"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT = OUTPUT_DIR / "第411批正式公告与第412批公示新能源乘用车车型扫描分析.docx"

FONT_NAME = "Microsoft YaHei"
HEADER_FILL = "1F4E78"
ALT_FILL = "F2F7FC"
BORDER = "D9D9D9"
MUTED = RGBColor(0x6B, 0x72, 0x80)
BLACK = RGBColor(0, 0, 0)


def set_run_font(run, size=10.5, bold=False, italic=False, color=BLACK):
    run.font.name = FONT_NAME
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")


def set_cell_borders(cell, color=BORDER):
    tc_pr = cell._tc.get_or_add_tcPr()
    old = tc_pr.find(qn("w:tcBorders"))
    if old is not None:
        tc_pr.remove(old)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)
        borders.append(element)
    tc_pr.append(borders)


def set_cell_margins(cell, top=80, start=90, bottom=80, end=90):
    tc_pr = cell._tc.get_or_add_tcPr()
    old = tc_pr.find(qn("w:tcMar"))
    if old is not None:
        tc_pr.remove(old)
    margins = OxmlElement("w:tcMar")
    for tag, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = OxmlElement(f"w:{tag}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        margins.append(node)
    tc_pr.append(margins)


def set_repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_row_cant_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    cant_split.set(qn("w:val"), "true")
    tr_pr.append(cant_split)


def set_table_cell_text(cell, text, bold=False, size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.12
    run = paragraph.add_run("" if text is None else str(text))
    set_run_font(run, size=size, bold=bold)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_borders(cell)
    set_cell_margins(cell)


def add_table(doc, headers, rows, widths, alignments=None, font_size=9.3):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    header = table.rows[0]
    set_repeat_header(header)
    set_row_cant_split(header)
    for index, text in enumerate(headers):
        set_cell_shading(header.cells[index], HEADER_FILL)
        set_table_cell_text(
            header.cells[index],
            text,
            bold=True,
            size=font_size,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )
        header.cells[index].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        header.cells[index].width = Inches(widths[index])
    for row_index, row in enumerate(rows):
        cells = table.add_row().cells
        set_row_cant_split(table.rows[-1])
        for col_index, value in enumerate(row):
            if row_index % 2 == 1:
                set_cell_shading(cells[col_index], ALT_FILL)
            alignment = (alignments or ["left"] * len(headers))[col_index]
            set_table_cell_text(
                cells[col_index],
                value,
                size=font_size,
                align={
                    "left": WD_ALIGN_PARAGRAPH.LEFT,
                    "center": WD_ALIGN_PARAGRAPH.CENTER,
                    "right": WD_ALIGN_PARAGRAPH.RIGHT,
                }[alignment],
            )
            cells[col_index].width = Inches(widths[col_index])
    return table


def add_caption(doc, text, above=False):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(4 if not above else 0)
    paragraph.paragraph_format.space_after = Pt(4 if above else 8)
    paragraph.paragraph_format.keep_with_next = above
    run = paragraph.add_run(text)
    set_run_font(run, size=9, italic=True, color=MUTED)
    return paragraph


def add_body(doc, text):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(7)
    paragraph.paragraph_format.line_spacing = 1.22
    run = paragraph.add_run(text)
    set_run_font(run, size=10.5)
    return paragraph


def add_numbered(doc, text):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.left_indent = Inches(0.22)
    paragraph.paragraph_format.first_line_indent = Inches(-0.22)
    paragraph.paragraph_format.space_after = Pt(4)
    paragraph.paragraph_format.line_spacing = 1.18
    run = paragraph.add_run(text)
    set_run_font(run, size=10.3)
    return paragraph


def configure_document(doc):
    section = doc.sections[0]
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = FONT_NAME
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = BLACK
    normal.paragraph_format.space_after = Pt(7)
    normal.paragraph_format.line_spacing = 1.2

    for style_name, size in (("Title", 24), ("Heading 1", 16), ("Heading 2", 13), ("Heading 3", 11)):
        style = styles[style_name]
        style.font.name = FONT_NAME
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)
        style.font.size = Pt(size)
        style.font.color.rgb = BLACK
        style.font.bold = True
        style.paragraph_format.space_before = Pt(0 if style_name == "Title" else 12)
        style.paragraph_format.space_after = Pt(8)
        style.paragraph_format.keep_with_next = True
    title_style = styles["Title"]
    title_p_pr = title_style._element.get_or_add_pPr()
    old_border = title_p_pr.find(qn("w:pBdr"))
    if old_border is not None:
        title_p_pr.remove(old_border)


def route_row(batch, route):
    return next(
        (row for row in PAYLOAD["batch_summary"][batch] if row["技术路线"] == route),
        None,
    )


def fmt(value, suffix="", decimals=1):
    if value is None:
        return "n.a."
    if decimals == 0:
        return f"{value:.0f}{suffix}"
    return f"{value:.{decimals}f}{suffix}"


def top_models(batch, limit):
    records = [
        record
        for record in PAYLOAD["records"][batch]
        if record.get("电机总功率(kW)") is not None
    ]
    records.sort(key=lambda record: record["电机总功率(kW)"], reverse=True)
    return records[:limit]


def build_report():
    payload = PAYLOAD
    doc = Document()
    configure_document(doc)

    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = title.add_run("第411批正式公告与第412批公示新能源乘用车车型扫描分析")
    set_run_font(run, size=23, bold=True, color=BLACK)
    subtitle = doc.add_paragraph()
    subtitle.paragraph_format.space_before = Pt(8)
    subtitle.paragraph_format.space_after = Pt(4)
    run = subtitle.add_run("数据截至 2026-10-17")
    set_run_font(run, size=12, color=MUTED)
    meta = doc.add_paragraph()
    run = meta.add_run("正式公告口径：第411批；公示候选口径：第412批")
    set_run_font(run, size=10.5, color=MUTED)
    add_body(
        doc,
        "本报告用于识别新批次新能源乘用车的产品结构、主要参数与驱动系统方向。"
        "第411批已经形成正式公告，作为本轮结论的主证据；第412批仍处于公示期，"
        "仅使用官方公示清单和详情页中已经披露的数据，续航、容量和部分供应商字段将在正式公告发布后复核。",
    )
    doc.add_page_break()

    doc.add_heading("结论摘要", level=1)
    add_body(
        doc,
        "411正式公告识别到47个新能源乘用车型号，其中纯电23个、插电混动21个、增程3个。"
        "412公示清单识别到58个新能源乘用车型号，其中纯电37个、插电混动15个、增程6个。"
        "412公示整体更偏向纯电，但公示详情页尚未统一披露纯电续航和电池容量，因此不能把公示样本与正式公告样本直接等同。",
    )
    add_numbered(
        doc,
        "1. 411正式公告的动力结构更均衡。纯电占48.9%，插混与增程合计51.1%；412公示中纯电占63.8%，插混与增程合计36.2%。",
    )
    add_numbered(
        doc,
        "2. 411插混车型明显大电量化。411插混平均纯电续航248.5km、平均电池容量55.5kWh，较410基准的224.9km和43.4kWh继续上移；样本中最高容量达到71.24kWh。",
    )
    add_numbered(
        doc,
        "3. 412纯电车型继续向高功率集中。在37个纯电样本中，平均电机总功率270.5kW，200kW及以上占有效功率样本的70.3%，多电机车型占有效驱动形式样本的21.6%。",
    )
    add_numbered(
        doc,
        "4. 412增程车型集中在50kWh以上电池和250km以上纯电续航。公示附件中3个增程样本的纯电续航为251至265km，电池容量为52.01至56.036kWh。",
    )
    add_numbered(
        doc,
        "5. 后续开发判断需要同时看“驱动功率”和“电池带电量”。411的插混验证了大电池长续航路线，412的纯电高功率和多电机样本则提示中大型车电驱系统需要预留更高的功率上限和平台化扩展空间。",
    )

    doc.add_heading("数据范围与口径", level=1)
    add_body(
        doc,
        "历史基准来自341至410批权威底表，411采用工业和信息化部正式公告附件与官方车型公示详情页，"
        "412采用2026-10-16发布的拟发布内容公示、车辆新产品公示清单和官方详情页。"
        "车型范围限定为新能源乘用车，排除燃油车、普通混合动力车、客车、货车、专用车和摩托车。",
    )
    add_body(
        doc,
        "批次数量按产品型号统计，型号级新增表示该型号未在341至410权威底表中出现，不等同于全新车系或全新品牌。"
        "均值只按公告已经披露的有值样本计算，空白和n.a.不按0处理。电池容量优先取正式公告附件，"
        "公示车型仅保留官方详情页已经公开的字段。",
    )

    batch_rows = [
        [
            "410正式基准",
            "2026-09-18",
            "70",
            "47",
            "16",
            "7",
            "n.a.",
            fmt(payload["batch_summary"]["overall_410"]["平均续航(km)"], " km"),
        ],
        [
            "411正式公告",
            "2026-10-16",
            "47",
            "23",
            "21",
            "3",
            "26",
            "593.5 km（纯电）",
        ],
        [
            "412公示",
            "2026-10-16",
            "58",
            "37",
            "15",
            "6",
            "58",
            "待正式公告",
        ],
    ]
    add_caption(doc, "表1 各批次新能源乘用车扫描范围", above=True)
    add_table(
        doc,
        ["批次", "状态", "NEV乘用车", "纯电", "插混", "增程", "型号级新增", "平均续航口径"],
        batch_rows,
        [0.92, 1.02, 0.86, 0.62, 0.62, 0.62, 0.86, 1.36],
        ["center", "center", "right", "right", "right", "right", "right", "left"],
    )

    doc.add_heading("按批次扫描结果", level=1)
    add_body(
        doc,
        "411正式公告的47个型号中，纯电和插混数量接近，增程数量较少。车型结构呈现两头分布："
        "一端是紧凑型和中型纯电产品，另一端是5.1米以上的大型插混和增程产品。"
        "本批次纯电平均车长4810mm、平均整备质量2020kg；插混平均车长5163.7mm、平均整备质量2688.8kg。",
    )
    add_body(
        doc,
        "412公示的58个型号中，纯电占比达到63.8%。纯电样本平均车长4986.5mm、平均整备质量2189.1kg，"
        "比411纯电样本更长、更重；插混和增程样本仍集中在5.0米左右的中大型产品。"
        "该批次公示清单覆盖1074个新产品，其中新能源349个型号，本次扫描保留58个新能源乘用车型号，"
        "其余为商用车、专用车、摩托车或非乘用车产品。",
    )
    doc.add_picture(str(ASSETS / "chart_batch_structure.png"), width=Inches(6.75))
    add_caption(doc, "图1 各批次新能源乘用车结构。412为公示候选口径，正式公告发布后应复核。")

    doc.add_heading("主要参数与趋势", level=1)
    add_body(
        doc,
        "411纯电样本的平均续航为593.5km、平均电池容量72.4kWh、平均电机总功率258.1kW。"
        "相比410纯电基准，续航和容量略低，但电机功率明显提高；这更像是车型级别和样本构成变化，"
        "不宜直接解释为技术退步。411插混样本的平均纯电续航248.5km、平均容量55.5kWh，"
        "相对410插混基准提升明显，说明长续航插混继续扩大。",
    )
    add_body(
        doc,
        "412公示的纯电样本仅有功率、尺寸和质量可用，续航和容量不能从当前公示详情页统一提取，"
        "因此表中保留n.a.。插混样本平均纯电续航225.1km、平均容量48.3kWh；增程样本平均纯电续航257.0km、"
        "平均容量54.7kWh。该结果说明增程新车正在向50kWh以上的带电量集中。",
    )
    parameter_rows = []
    for route in ("纯电(BEV)", "插电混动(PHEV)", "增程(EREV)"):
        row = [
            route,
            str(route_row("410", route)["数量"]),
            str(route_row("411", route)["数量"]),
            str(route_row("412", route)["数量"]),
            f'{fmt(route_row("410", route)["平均续航(km)"])} / {fmt(route_row("411", route)["平均续航(km)"])} / {fmt(route_row("412", route)["平均续航(km)"])}',
            f'{fmt(route_row("410", route)["平均电池容量(kWh)"])} / {fmt(route_row("411", route)["平均电池容量(kWh)"])} / {fmt(route_row("412", route)["平均电池容量(kWh)"])}',
            f'{fmt(route_row("410", route)["平均电机功率(kW)"])} / {fmt(route_row("411", route)["平均电机功率(kW)"])} / {fmt(route_row("412", route)["平均电机功率(kW)"])}',
        ]
        parameter_rows.append(row)
    add_caption(doc, "表2 路线参数对比。数值顺序为410 / 411 / 412；功率均值只按公开有值样本计算。", above=True)
    add_table(
        doc,
        ["技术路线", "410", "411", "412", "平均续航 km", "平均容量 kWh", "平均功率 kW"],
        parameter_rows,
        [1.15, 0.42, 0.42, 0.42, 1.52, 1.52, 1.52],
        ["left", "right", "right", "right", "center", "center", "center"],
    )
    doc.add_picture(str(ASSETS / "chart_power.png"), width=Inches(6.75))
    add_caption(
        doc,
        "图2 批次平均电机总功率。411插混仅6个型号公开电机功率，图中均值不代表全部21个插混样本。",
    )

    doc.add_heading("驱动系统扫描", level=1)
    add_body(
        doc,
        "本报告把电机峰值功率序列求和作为电机总功率，并按公开峰值功率数列推断单电机、双电机、三电机和四电机。"
        "当官方详情页只披露发动机最大净功率、没有披露电机功率时，该行不参与电机功率均值。",
    )
    drive_rows = []
    for batch_label, batch in (("411正式", "411"), ("412公示", "412")):
        for route in ("纯电(BEV)", "插电混动(PHEV)", "增程(EREV)"):
            row = route_row(batch, route)
            if not row:
                continue
            drive_rows.append(
                [
                    batch_label,
                    route,
                    f'{row["平均电机功率(kW)"]:.1f}' if row["平均电机功率(kW)"] is not None else "n.a.",
                    f'{row["200kW及以上占比"]:.1%}' if row["200kW及以上占比"] is not None else "n.a.",
                    f'{row["多电机占比"]:.1%}' if row["多电机占比"] is not None else "n.a.",
                    str(row["功率有效样本"]),
                    str(row["多电机有效样本"]),
                ]
            )
    add_caption(doc, "表3 驱动系统结构。百分比只在有效样本内计算，样本数列单独展示。", above=True)
    add_table(
        doc,
        ["批次", "技术路线", "平均电机总功率 kW", "200kW及以上占比", "多电机占比", "功率样本", "驱动样本"],
        drive_rows,
        [0.82, 1.22, 1.18, 1.12, 0.98, 0.72, 0.72],
        ["center", "left", "right", "right", "right", "right", "right"],
    )
    add_body(
        doc,
        "411纯电样本中，200kW及以上占60.9%，多电机占39.1%；412纯电样本中，"
        "200kW及以上占70.3%，多电机占21.6%。412的单电机占比更高，但高功率单电机和"
        "双电机高性能车型同时存在，说明平台策略正在分化为效率型和高性能型两条路线。",
    )
    add_body(
        doc,
        "把功率和电池放在一起看，411插混的均值功率388.3kW建立在6个有效样本上，不能直接外推；"
        "但六款样本均达到200kW以上，且多电机占比50%。该现象足以支持一个判断："
        "中大型插混和增程产品的电驱功率门槛已经上移，开发规划不能继续只按早期100至150kW的电驱平台设定上限。",
    )

    doc.add_heading("重点车型观察", level=1)
    add_body(
        doc,
        "下表按电机总功率排序，仅列公开披露电机功率的车型。续航和容量为空表示官方公示暂未提供，"
        "不代表产品没有相关参数。车型通用名称只在公告或目录原文给出时展示。",
    )
    top_rows = []
    for record in top_models("411", 8):
        top_rows.append(
            [
                "411",
                record["企业名称"],
                record["产品商标"],
                record["产品型号"],
                f'{record["电机总功率(kW)"]:.0f} / {record["驱动形式"] or "n.a."}',
                fmt(record["纯电续航(km)"]) + " / " + fmt(record["电池容量(kWh)"]),
            ]
        )
    for record in top_models("412", 8):
        top_rows.append(
            [
                "412",
                record["企业名称"],
                record["产品商标"],
                record["产品型号"],
                f'{record["电机总功率(kW)"]:.0f} / {record["驱动形式"] or "n.a."}',
                fmt(record["纯电续航(km)"]) + " / " + fmt(record["电池容量(kWh)"]),
            ]
        )
    add_caption(doc, "表4 高功率重点车型，功率/驱动形式与续航/容量。", above=True)
    add_table(
        doc,
        ["批次", "企业", "商标", "型号", "功率 / 驱动形式", "续航 / 容量"],
        top_rows,
        [0.45, 1.18, 0.82, 1.18, 1.32, 1.55],
        ["center", "left", "left", "left", "center", "center"],
        font_size=8.8,
    )

    doc.add_heading("对驱动系统和产品规划的建议", level=1)
    add_body(
        doc,
        "先确定产品级别，再确定电驱平台，不要把全部车型压到同一个功率段。"
        "对C级和中型主流纯电产品，150至250kW单电机仍可作为效率基线；"
        "对D级、MPV和旗舰SUV，至少需要预留400至600kW的双电机或三电机扩展能力。"
        "412公示中的多电机样本已经覆盖双电机、三电机和四电机，平台规划必须支持前后轴功率解耦和不同电机数量组合。",
    )
    add_body(
        doc,
        "插混和增程应按“大电池纯电体验”定义系统边界。411插混的平均容量已经达到55.5kWh，"
        "412增程样本集中在52至56kWh、纯电续航251至265km。建议把200至300km纯电续航、"
        "45至70kWh电池包和200至400kW电驱作为中大型插混、增程产品的规划区间，"
        "并对旗舰车型预留更高的功率和电池容量上限。",
    )
    add_body(
        doc,
        "下一步的重点不是继续增加单一峰值参数，而是把电驱效率和整车质量一起纳入门禁。"
        "建议在后续车型定义中同时设置电机功率、整备质量、续航和电池容量四项设计目标，"
        "用公告批次滚动验证实现情况。若只有功率上升而质量、能耗同步恶化，不应把它解释为技术进步。",
    )

    doc.add_heading("持续扫描流程", level=1)
    add_numbered(
        doc,
        "1. 每月跟踪工信部公告发布、拟发布内容公示、车船税目录和车辆购置税目录。",
    )
    add_numbered(
        doc,
        "2. 公示阶段先抓取新产品清单和官方详情页，标记纯电、插混、增程和乘用车范围；续航、容量等待正式附件。",
    )
    add_numbered(
        doc,
        "3. 正式公告发布后，用附件替换公示候选值，按产品型号与341批以来的历史底表比对，区分型号级新增和历史延续。",
    )
    add_numbered(
        doc,
        "4. 每轮更新扫描底表、驱动系统对比和报告结论，缺失字段保留空白并记录原因。",
    )
    add_numbered(
        doc,
        "5. 每季度比较主要车企的功率、电池容量、续航和驱动形式变化，形成进入下一轮产品定义的技术门槛。",
    )

    doc.add_heading("来源与限制", level=1)
    source_rows = [
        ["第411批正式公告", "2026-10-16", "https://www.miit.gov.cn/jgsj/zbys/wjfb/art/2026/art_ef051926ce924a4da93e1677c7e18b34.html"],
        ["第412批拟发布内容公示", "2026-10-16", "https://www.miit.gov.cn/jgsj/zbys/qcgy/art/2026/art_ec785d818a604ccd9618b73d4f13f062.html"],
        ["第412批新产品公示清单", "2026-10-12", "https://www.miit.gov.cn/datainfo/cpgg/art/2026/art_43a4b4be0de2483d85f28e9423201494.html"],
        ["341至410批权威底表", "2026-10-05", "本地项目底表"],
    ]
    add_table(
        doc,
        ["来源", "日期", "链接/说明"],
        source_rows,
        [1.25, 1.0, 4.25],
        ["left", "center", "left"],
        font_size=8.4,
    )
    add_body(
        doc,
        "限制一：412仍是公示候选数据，续航、容量和供应商字段待正式公告补齐。"
        "限制二：411插混的电机功率有效样本仅6个，均值不外推。"
        "限制三：型号级新增仅代表型号未见于341至410底表。"
        "限制四：驱动形式由峰值功率数列推断，正式开发决策仍需官方技术文件复核。",
    )

    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build_report()
