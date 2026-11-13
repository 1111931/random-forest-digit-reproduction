#!/usr/bin/env python
"""Build the complete Chinese manuscript as a styled DOCX."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path
from typing import Iterable, Sequence

import numpy as np
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor, Twips


ROOT = Path(__file__).resolve().parents[1]
RESULT_DIR = ROOT / "results"
FIGURE_DIR = ROOT / "figures"
CHECKPOINT_DIR = ROOT / "checkpoints"
PAPER_DIR = ROOT / "paper"

DOCS_SKILL_SCRIPTS = Path(
    r"C:\Users\24062\.codex\plugins\cache\openai-primary-runtime"
    r"\documents\26.802.11031\skills\documents\scripts"
)
sys.path.insert(0, str(DOCS_SKILL_SCRIPTS))
from table_geometry import apply_table_geometry, column_widths_from_weights


NAVY = RGBColor(0x0B, 0x25, 0x45)
HEADING_BLUE = RGBColor(0x1F, 0x3A, 0x5F)
DARK_BLUE = RGBColor(0x1F, 0x4D, 0x78)
MUTED = RGBColor(0x60, 0x6B, 0x76)
BLACK = RGBColor(0, 0, 0)
GOLD = RGBColor(0x7A, 0x5A, 0x00)

BODY_FONT_LATIN = "Times New Roman"
BODY_FONT_CJK = "宋体"
HEAD_FONT_LATIN = "Arial"
HEAD_FONT_CJK = "微软雅黑"

GITHUB_URL = (
    "https://github.com/1111931/random-forest-digit-reproduction"
)


def set_run_font(
    run,
    *,
    latin: str = BODY_FONT_LATIN,
    cjk: str = BODY_FONT_CJK,
    size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: RGBColor | None = None,
) -> None:
    run.font.name = latin
    if run._element.rPr is None:
        run._element.get_or_add_rPr()
    r_fonts = run._element.rPr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        run._element.rPr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), latin)
    r_fonts.set(qn("w:hAnsi"), latin)
    r_fonts.set(qn("w:eastAsia"), cjk)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def set_style_font(style, latin: str, cjk: str, size: float, color: RGBColor) -> None:
    style.font.name = latin
    style.font.size = Pt(size)
    style.font.color.rgb = color
    r_pr = style.element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), latin)
    r_fonts.set(qn("w:hAnsi"), latin)
    r_fonts.set(qn("w:eastAsia"), cjk)


def configure_document(doc: Document) -> None:
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)
    section.different_first_page_header_footer = True

    normal = doc.styles["Normal"]
    set_style_font(normal, BODY_FONT_LATIN, BODY_FONT_CJK, 11, BLACK)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.333
    normal.paragraph_format.widow_control = True

    h1 = doc.styles["Heading 1"]
    set_style_font(h1, HEAD_FONT_LATIN, HEAD_FONT_CJK, 16, HEADING_BLUE)
    h1.font.bold = True
    h1.paragraph_format.space_before = Pt(18)
    h1.paragraph_format.space_after = Pt(10)
    h1.paragraph_format.keep_with_next = True
    h1.paragraph_format.line_spacing = 1.1

    h2 = doc.styles["Heading 2"]
    set_style_font(h2, HEAD_FONT_LATIN, HEAD_FONT_CJK, 13, HEADING_BLUE)
    h2.font.bold = True
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(6)
    h2.paragraph_format.keep_with_next = True
    h2.paragraph_format.line_spacing = 1.1

    h3 = doc.styles["Heading 3"]
    set_style_font(h3, HEAD_FONT_LATIN, HEAD_FONT_CJK, 12, DARK_BLUE)
    h3.font.bold = True
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    h3.paragraph_format.keep_with_next = True
    h3.paragraph_format.line_spacing = 1.1


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    prefix = paragraph.add_run("第 ")
    set_run_font(prefix, size=8.5, color=MUTED)
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text, end])
    set_run_font(run, size=8.5, color=MUTED)
    suffix = paragraph.add_run(" 页")
    set_run_font(suffix, size=8.5, color=MUTED)


def configure_furniture(doc: Document) -> None:
    section = doc.sections[0]
    header = section.header
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run("手写数字视觉识别中的随机森林复现")
    set_run_font(
        run,
        latin=HEAD_FONT_LATIN,
        cjk=HEAD_FONT_CJK,
        size=8.5,
        color=MUTED,
    )
    add_page_number(section.footer.paragraphs[0])
    add_page_number(section.first_page_footer.paragraphs[0])


def add_para(
    doc: Document,
    text: str = "",
    *,
    bold_prefix: str | None = None,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    size: float = 11,
    bold: bool = False,
    italic: bool = False,
    color: RGBColor = BLACK,
    before: float = 0,
    after: float = 8,
    line: float = 1.333,
    indent_cm: float = 0,
    keep_with_next: bool = False,
) -> object:
    paragraph = doc.add_paragraph()
    paragraph.alignment = align
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line
    paragraph.paragraph_format.keep_with_next = keep_with_next
    if indent_cm:
        paragraph.paragraph_format.first_line_indent = Inches(indent_cm)
    if bold_prefix and text.startswith(bold_prefix):
        prefix_run = paragraph.add_run(bold_prefix)
        set_run_font(
            prefix_run,
            size=size,
            bold=True,
            italic=italic,
            color=color,
        )
        body_run = paragraph.add_run(text[len(bold_prefix) :])
        set_run_font(
            body_run,
            size=size,
            bold=bold,
            italic=italic,
            color=color,
        )
    else:
        run = paragraph.add_run(text)
        set_run_font(
            run,
            size=size,
            bold=bold,
            italic=italic,
            color=color,
        )
    return paragraph


def add_heading(doc: Document, text: str, level: int = 1) -> object:
    paragraph = doc.add_paragraph(style=f"Heading {level}")
    run = paragraph.add_run(text)
    latin = HEAD_FONT_LATIN
    cjk = HEAD_FONT_CJK
    color = HEADING_BLUE if level < 3 else DARK_BLUE
    size = 16 if level == 1 else 13 if level == 2 else 12
    set_run_font(run, latin=latin, cjk=cjk, size=size, bold=True, color=color)
    return paragraph


def add_hyperlink(paragraph, text: str, url: str) -> None:
    part = paragraph.part
    relationship_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)
    new_run = OxmlElement("w:r")
    run_properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "1F4D78")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), BODY_FONT_LATIN)
    fonts.set(qn("w:hAnsi"), BODY_FONT_LATIN)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "21")
    run_properties.extend([fonts, color, underline, size])
    text_element = OxmlElement("w:t")
    text_element.text = text
    new_run.extend([run_properties, text_element])
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    shading.set(qn("w:val"), "clear")
    tc_pr.append(shading)


def set_table_borders(table, color: str = "C9D2DC", size: int = 5) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), str(size))
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)
        borders.append(element)


def set_cell_text(
    cell,
    text: str,
    *,
    bold: bool = False,
    align=WD_ALIGN_PARAGRAPH.CENTER,
    size: float = 9.5,
    color: RGBColor = BLACK,
) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    run = paragraph.add_run(str(text))
    set_run_font(run, size=size, bold=bold, color=color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def add_caption(
    doc: Document,
    text: str,
    *,
    keep_with_next: bool = False,
) -> None:
    add_para(
        doc,
        text,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        size=9,
        color=MUTED,
        before=3,
        after=8,
        line=1.0,
        keep_with_next=keep_with_next,
    )


def add_figure(
    doc: Document,
    figure_name: str,
    caption: str,
    *,
    width_inches: float = 6.25,
) -> None:
    path = FIGURE_DIR / figure_name
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.keep_together = True
    run = paragraph.add_run()
    shape = run.add_picture(str(path), width=Inches(width_inches))
    doc_pr = shape._inline.docPr
    doc_pr.set("descr", caption)
    doc_pr.set("title", caption)
    add_caption(doc, caption)


def add_table(
    doc: Document,
    caption: str,
    headers: Sequence[str],
    rows: Sequence[Sequence[object]],
    weights: Sequence[float],
    *,
    font_size: float = 9.0,
) -> None:
    add_caption(doc, caption, keep_with_next=True)
    table = doc.add_table(rows=1, cols=len(headers))
    table.autofit = False
    set_table_borders(table)
    header_row_properties = table.rows[0]._tr.get_or_add_trPr()
    repeat_header = OxmlElement("w:tblHeader")
    repeat_header.set(qn("w:val"), "true")
    header_row_properties.append(repeat_header)
    for index, header in enumerate(headers):
        cell = table.rows[0].cells[index]
        set_cell_text(cell, header, bold=True, size=font_size, color=NAVY)
        shade_cell(cell, "F4F6F9")
    for values in rows:
        cells = table.add_row().cells
        for index, value in enumerate(values):
            align = (
                WD_ALIGN_PARAGRAPH.LEFT
                if index == 0 and isinstance(value, str) and len(value) > 18
                else WD_ALIGN_PARAGRAPH.CENTER
            )
            set_cell_text(cells[index], value, size=font_size, align=align)
    widths = column_widths_from_weights(weights, 9360)
    apply_table_geometry(
        table,
        widths,
        table_width_dxa=9360,
        indent_dxa=120,
        cell_margins_dxa={"top": 80, "bottom": 80, "start": 120, "end": 120},
    )
    for row in table.rows:
        tr_pr = row._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        tr_pr.append(cant_split)
    add_para(doc, "", after=2, line=1.0)


def add_equation(
    doc: Document,
    expression: str,
    number: str,
) -> None:
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(4)
    paragraph.paragraph_format.space_after = Pt(8)
    run = paragraph.add_run(f"{expression}    ({number})")
    set_run_font(
        run,
        latin="Cambria Math",
        cjk=BODY_FONT_CJK,
        size=10.5,
        italic=False,
        color=BLACK,
    )


def pct(value: float, digits: int = 2) -> str:
    return f"{100.0 * value:.{digits}f}%"


def signed_pp(value: float, digits: int = 2) -> str:
    return f"{100.0 * value:+.{digits}f} 个百分点"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def load_metrics() -> dict:
    with (RESULT_DIR / "metrics.json").open(encoding="utf-8") as handle:
        return json.load(handle)


def build_paper(output_path: Path) -> None:
    data = load_metrics()
    if data["is_quick_mode"]:
        raise RuntimeError("Quick-mode results cannot be used to build the paper.")
    split = data["dataset"]
    summary = {row["method"]: row for row in data["main_summary"]}
    checkpoint = data["checkpoint_metrics"]
    ci = data["bootstrap_95_ci"]
    mcnemar_baseline = data["mcnemar"]["baseline_vs_proposed"]
    mcnemar_bagging = data["mcnemar"]["bagging_vs_proposed"]
    oob_values = [item["oob_score"] for item in data["oob_per_seed"]]
    oob_mean = float(np.mean(oob_values))
    oob_std = float(np.std(oob_values, ddof=1))
    acc_tree = summary["baseline"]["accuracy_mean"]
    acc_bag = summary["ablation_bagging"]["accuracy_mean"]
    acc_rf = summary["proposed"]["accuracy_mean"]
    gain_bag = acc_bag - acc_tree
    gain_feature = acc_rf - acc_bag
    gain_total = acc_rf - acc_tree

    doc = Document()
    configure_document(doc)
    configure_furniture(doc)
    core_properties = doc.core_properties
    core_properties.title = "从单棵决策树到随机森林：手写数字视觉识别的历史方法复现与实证分析"
    core_properties.subject = "随机森林在手写数字识别中的可复现实验"
    core_properties.author = "独立复现研究组"
    core_properties.keywords = "random forest; handwritten digit recognition; ensemble learning; reproducibility"
    core_properties.comments = "All numerical results were generated from the included reproducibility package."

    # Title block
    add_para(
        doc,
        "计算机视觉历史方法复现研究",
        align=WD_ALIGN_PARAGRAPH.CENTER,
        size=9.5,
        bold=True,
        color=GOLD,
        after=8,
        line=1.0,
    )
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(8)
    title.paragraph_format.line_spacing = 1.18
    title_run = title.add_run(
        "从单棵决策树到随机森林：\n手写数字视觉识别的历史方法复现与实证分析"
    )
    set_run_font(
        title_run,
        latin=HEAD_FONT_LATIN,
        cjk=HEAD_FONT_CJK,
        size=21,
        bold=True,
        color=NAVY,
    )
    add_para(
        doc,
        "作者：独立复现研究组",
        align=WD_ALIGN_PARAGRAPH.CENTER,
        size=11,
        bold=True,
        after=3,
        line=1.0,
    )
    add_para(
        doc,
        "单位：人工智能与计算机视觉实验室",
        align=WD_ALIGN_PARAGRAPH.CENTER,
        size=9.5,
        color=MUTED,
        after=10,
        line=1.0,
    )
    add_para(
        doc,
        "From a Single Decision Tree to a Random Forest:\n"
        "A Reproducible Study of Historical Handwritten Digit Recognition",
        align=WD_ALIGN_PARAGRAPH.CENTER,
        size=12.5,
        italic=True,
        color=MUTED,
        after=4,
        line=1.15,
    )
    add_para(
        doc,
        "Independent Reproduction Group",
        align=WD_ALIGN_PARAGRAPH.CENTER,
        size=10.5,
        bold=True,
        after=2,
        line=1.0,
    )
    add_para(
        doc,
        "Artificial Intelligence and Computer Vision Laboratory",
        align=WD_ALIGN_PARAGRAPH.CENTER,
        size=9.2,
        italic=True,
        color=MUTED,
        after=10,
        line=1.0,
    )
    abstract_text = (
        "摘要：随机森林是机器学习中被广泛使用的经典集成方法。为检验其核心历史经验，"
        "本文在一个完全可复现的小规模计算机视觉任务上比较单棵 CART 决策树、"
        "64 维全特征 Bagging 与随机像素子空间随机森林。实验使用 scikit-learn 自带的 "
        "1,797 幅 8×8 手写数字图像，采用固定分层划分得到 1,437 个训练样本和 360 个"
        f"测试样本。在 5 个模型随机种子上，单树、Bagging 和随机森林的平均准确率分别为 "
        f"{pct(acc_tree)}、{pct(acc_bag)} 和 {pct(acc_rf)}。随机森林相对单树提高 "
        f"{100.0 * gain_total:.2f} 个百分点，相对全特征 Bagging 提高 "
        f"{100.0 * gain_feature:.2f} 个百分点；"
        "在像素噪声和中心遮挡条件下，其退化幅度也明显更小。消融结果表明，"
        "Bootstrap 集成主要降低单树的高方差，而随机像素子空间进一步增加了基学习器的"
        "多样性，是随机森林优于 Bagging 的关键因素。本文同时发现随机森林的概率校准"
        "弱于 Bagging，说明准确率提升不等价于概率质量提升。所有实验代码、原始结果、"
        "图表与基线和本文方法的模型权重均已整理进复现包，数值可由同一命令完整复现。"
    )
    add_para(doc, abstract_text, size=10.3, line=1.25, after=8)
    add_para(
        doc,
        "关键词：随机森林；手写数字识别；集成学习；特征随机化；可复现实验",
        size=10.3,
        bold_prefix="关键词：",
        after=12,
        line=1.15,
    )

    add_para(
        doc,
        "Abstract: Random forests are a classical ensemble method in machine learning. "
        "This study tests their central historical claims on a fully reproducible "
        "small-scale computer-vision task by comparing a single CART tree, full-feature "
        "bagging, and a random forest with random pixel subspaces. We use 1,797 8 x 8 "
        "handwritten digit images bundled with scikit-learn and a fixed stratified split "
        "of 1,437 training and 360 test samples. Across five model seeds, the mean test "
        f"accuracies of the tree, bagging, and random forest are {pct(acc_tree)}, "
        f"{pct(acc_bag)}, and {pct(acc_rf)}, respectively. The random forest improves "
        f"over the tree by {100.0 * gain_total:.2f} percentage points and over "
        f"full-feature bagging by {100.0 * gain_feature:.2f} percentage points. "
        "It also degrades substantially less under Gaussian pixel "
        "noise and central occlusion. The ablation indicates that bootstrap aggregation "
        "mainly reduces the variance of a single tree, whereas random feature subspaces "
        "add the diversity responsible for the additional gain. We also find that the "
        "random forest is less calibrated than bagging, showing that higher accuracy does "
        "not imply better probability estimates. All code, raw results, figures, and "
        "model checkpoints are released with the reproduction package and reproduce "
        "the reported numbers.",
        size=9.8,
        italic=False,
        color=BLACK,
        line=1.18,
        after=8,
    )
    add_para(
        doc,
        "Keywords: random forest; handwritten digit recognition; ensemble learning; "
        "feature randomization; reproducible research",
        size=9.8,
        bold_prefix="Keywords:",
        line=1.12,
        after=14,
    )

    add_heading(doc, "1 引言", 1)
    add_para(
        doc,
        "手写数字识别是计算机视觉与模式识别中最早被系统研究的任务之一。它既包含"
        "笔画粗细、书写倾斜和局部形变等自然变化，又具有类别数少、输入尺寸小的特点，"
        "因此常被用于验证分类器对视觉表征、决策边界和泛化误差的基本处理能力。"
        "从特征工程与传统分类器，到 LeNet-5 [6]、AlexNet [7] 和残差网络 [8]，"
        "该方向的技术路线经历了显著变化。深度网络在大规模数据上已成为主流，但在"
        "样本规模有限、计算资源受限、解释性和可复现性要求较高的场景中，经典集成方法"
        "仍具有研究和教学价值。",
    )
    add_para(
        doc,
        "Breiman 在 2001 年系统总结了随机森林 [1]。该方法在 Bootstrap 重采样得到的"
        "训练子集上训练多棵决策树，并在每次节点分裂时只考虑随机选取的特征子集，"
        "最后对各树的类别概率取平均。与单棵决策树相比，随机森林通过样本扰动降低模型"
        "方差；与只采用 Bootstrap 的 Bagging [3] 相比，随机特征子空间进一步削弱了"
        "树与树之间的相关性。Ho 的随机子空间方法 [2] 以及 Amit 和 Geman 的随机树"
        "工作 [5] 也为这一思想提供了历史基础。",
    )
    add_para(
        doc,
        "本文不主张新的算法创新，而是直接复现并检验上述历史结论。核心研究假设可以"
        "分为三点。H1：对高方差单树进行 Bootstrap 集成，应显著提升测试准确率与"
        "宏平均 F1。H2：在像素特征维数高于有效判别维数时，随机特征子空间应通过"
        "去相关进一步改善集成性能。H3：若 H1 和 H2 所描述的方差降低与多样性提升"
        "确实发生，那么随机森林在像素噪声和局部遮挡下也应比单树及全特征 Bagging "
        "更稳定。",
    )
    add_para(
        doc,
        "为检验这些假设，本文使用 scikit-learn [14] 自带的手写数字数据 [13, 19]，"
        "预先固定训练/测试划分，在同一测试集上比较单棵 CART 树、全特征 Bagging 和"
        "随机森林，并用 5 个模型随机种子评估波动。除主结果外，本文还报告学习曲线、"
        "集成树数量曲线、逐类 F1、混淆矩阵、特征重要性、高斯像素噪声、中心遮挡、"
        "Bootstrap 置信区间和 McNemar 检验。本文的主要贡献是：第一，给出一个不需要"
        "外部下载、可在普通 CPU 上完整重跑的历史方法复现；第二，通过三层对照把"
        "“集成收益”和“特征随机化收益”分开；第三，用多种扰动实验检验引言中的三个"
        "假设，并同时报告概率校准这一不利结果。",
    )
    add_para(
        doc,
        "需要强调，本文的目标不是与 ResNet 或 Vision Transformer 竞争当前最优准确率。"
        "8×8 的数据分辨率、360 个测试样本和单一的固定划分都会限制结论外推。本文希望"
        "回答的是一个更窄但可验证的问题：在透明、可解释且可重复的设置中，随机森林"
        "相对于其直接历史前身是否仍然表现出预期优势，以及这种优势来自哪里。",
    )

    add_heading(doc, "2 相关工作", 1)
    add_heading(doc, "2.1 手写数字识别", 2)
    add_para(
        doc,
        "传统手写数字识别通常依赖人工设计的形状、梯度或局部特征，再结合最近邻、"
        "支持向量机或决策树等分类器。LeCun 等提出的 LeNet-5 [6] 将卷积、池化和"
        "反向传播结合，使特征学习与分类器训练统一到端到端框架中。大规模图像数据与"
        "GPU 训练进一步推动了深度卷积网络的发展 [7-9]。在这一背景下，像素空间"
        "随机森林并不是最先进的识别器，但它提供了可直接检查的局部阈值决策和特征"
        "重要性，适合作为历史方法的受控复现。",
    )
    add_heading(doc, "2.2 集成学习与随机森林", 2)
    add_para(
        doc,
        "Bagging [3] 通过对训练集进行有放回采样构造多个基学习器，并以投票或平均降低"
        "不稳定学习器的方差。随机子空间 [2] 从特征维度引入随机性，使不同基学习器"
        "关注不同的输入子集。随机森林 [1] 同时使用 Bootstrap 样本扰动和节点级特征"
        "随机化，是这两条思路的组合。Dietterich [15]、Zhou [16] 与 Biau 和 Scornet "
        "[4] 的理论和综述指出，集成性能依赖单个模型的精度与模型间的多样性；仅增加"
        "树的数量并不能消除相关性过高所导致的收益饱和。",
    )
    add_heading(doc, "2.3 视觉表征与本文定位", 2)
    add_para(
        doc,
        "Eigenfaces [10]、SIFT [11] 和 HOG [12] 等工作展示了不同历史阶段对视觉表征"
        "的设计方式。本文刻意不加入方向梯度、卷积或数据增广，以免把表征收益与集成"
        "收益混在一起。输入仅做线性归一化，所有模型使用同一个 64 维原始像素向量。"
        "这种限制使消融解释更直接：单树、Bagging 和随机森林的差异主要来自训练样本"
        "扰动与特征子空间，而不是输入特征或网络结构。",
    )

    add_heading(doc, "3 方法", 1)
    add_heading(doc, "3.1 问题设定与数据预处理", 2)
    add_para(
        doc,
        "给定图像 x∈R^{8×8} 和类别 y∈{0,…,9}，首先按行展开为 64 维向量，并将原始"
        "整数灰度除以 16，使所有像素落入 [0,1]。训练采用 80%/20% 分层抽样，随机种子"
        f"固定为 {split['split_seed']}。最终训练集包含 {split['train_samples']} 个样本，"
        f"测试集包含 {split['test_samples']} 个样本。测试集只在模型拟合完成后用于评估，"
        "不参与超参数选择。",
    )
    add_heading(doc, "3.2 单棵 CART 决策树", 2)
    add_para(
        doc,
        "基线模型为 CART 分类树 [17]。在每个内部节点 t，选择使加权 Gini 不纯度下降"
        "最大的像素阈值。节点不纯度定义为",
    )
    add_equation(doc, "G(t) = 1 − Σₖ pₖ(t)²", "1")
    add_para(
        doc,
        "其中 pₖ(t) 是节点 t 中第 k 类样本的比例。一次分裂的质量写为",
    )
    add_equation(doc, "ΔG = G(t) − (nL/n)G(tL) − (nR/n)G(tR)", "2")
    add_para(
        doc,
        "其中 nL 与 nR 分别表示左右子节点样本数。基线采用 Gini 准则，并将"
        "min_samples_leaf 设为 2，以避免单个训练样本形成叶节点。单树结构紧凑且可解释，"
        "但训练样本的局部变化可能导致树结构显著改变，是典型的高方差模型。",
    )
    add_heading(doc, "3.3 Bootstrap 聚合", 2)
    add_para(
        doc,
        "Bagging [3] 从 n 个训练样本中有放回抽取 n 个样本，训练一棵树。某个样本在"
        "一次 bootstrap 抽样中未被抽到的概率为 (1−1/n)^n，当 n 增大时趋近于 e^{-1}。"
        "因此每个训练子集平均包含约 63.2% 的不同原始样本，其余样本形成天然扰动。"
        "对 B 棵树，类别概率通过平均得到",
    )
    add_equation(doc, "p̂ₖ(x) = (1/B) Σᵦ pᵦ,ₖ(x)", "3")
    add_equation(doc, "ŷ(x) = argmaxₖ p̂ₖ(x)", "4")
    add_para(
        doc,
        "本文的消融模型使用 200 棵树，Bootstrap 开启，但每次分裂考虑全部 64 个特征。"
        "该设置只检验样本扰动带来的集成分散，不包含特征随机化。",
    )
    add_heading(doc, "3.4 随机特征子空间随机森林", 2)
    add_para(
        doc,
        "本文方法在 Bagging 的基础上，在每个节点只从随机选取的 m 个像素特征中寻找"
        "最佳分裂。对于 64 维输入，使用 m=⌊√64⌋=8。基树不限制最大深度，但叶节点至少"
        "包含一个样本。为缓解类别规模的小幅差异，使用 balanced_subsample 类权重："
        "在每棵树的 bootstrap 子集内按类别频率反比赋权。最终预测仍由式 (3) 和式 (4)"
        "给出。",
    )
    add_figure(
        doc,
        "fig01_framework.png",
        "图 1  实验总体框架。图像先展开为像素向量，集成模型同时引入样本扰动与特征扰动，"
        "最终通过 200 棵树的类别概率平均完成分类。",
    )
    add_para(
        doc,
        "随机森林的袋外（out-of-bag, OOB）估计使用每棵树未抽到的样本评估该树，再对"
        "整片森林进行汇总。OOB 结果不接触测试集，因此可作为训练过程内部的泛化近似。"
        "设 Nₒₒᵦ(xᵢ) 为未在 bootstrap 子集中使用 xᵢ 的树集合，则",
    )
    add_equation(doc, "p̂ₖᵒᵒᵦ(xᵢ) = (1/|Nₒₒᵦ(xᵢ)|) Σᵦ∈Nₒₒᵦ(xᵢ) pᵦ,ₖ(xᵢ)", "5")
    add_heading(doc, "3.5 计算复杂度", 2)
    add_para(
        doc,
        "若单树平均深度为 log n、每个节点考察 m 个特征，则随机森林训练复杂度近似"
        "O(Bmn log n)。本实验 B=200、m=8、n=1,437、p=64，输入规模很小，完整训练可在"
        "普通 CPU 上完成。推理时需遍历 B 棵树的路径并平均概率，计算量高于单树，但模型"
        "大小和延迟仍处于 KB 至 MB 级。本文报告的运行时间只用于数量级比较，不作为"
        "硬件基准。",
    )

    add_heading(doc, "4 实验", 1)
    add_heading(doc, "4.1 数据集与划分", 2)
    add_para(
        doc,
        f"数据集共 {split['n_samples']} 幅图像，来自 scikit-learn 内置的 UCI 手写数字"
        "数据版本 [13, 14, 19]。每类样本数为 174 至 183，整体接近平衡。图 2 展示每个"
        "类别的前两幅样例，图 3 给出类别分布。固定分层划分后，训练集每类 139 至 146 个"
        "样本，测试集每类 35 至 37 个样本。",
    )
    add_figure(
        doc,
        "fig02_digit_samples.png",
        "图 2  数据集样例。每列对应一个数字类别，上下两行展示同一类别的不同书写。",
    )
    add_figure(
        doc,
        "fig03_class_distribution.png",
        "图 3  十类手写数字的样本分布。类别频率接近均匀，但仍有小幅差异。",
        width_inches=5.9,
    )
    add_table(
        doc,
        "表 1  数据集划分与实验规模",
        ["项目", "数值", "说明"],
        [
            ["原始样本", split["n_samples"], "8×8 灰度图"],
            ["特征维数", split["n_features"], "像素值归一化至 [0,1]"],
            ["训练样本", split["train_samples"], "80%，分层抽样"],
            ["测试样本", split["test_samples"], "20%，只用于最终评估"],
            ["模型随机种子", 5, "主结果重复次数"],
            ["决策树数量", 200, "Bagging 与随机森林"],
        ],
        [1.25, 0.9, 4.35],
    )

    add_heading(doc, "4.2 实现细节", 2)
    add_para(
        doc,
        "全部实验使用 Python 3.12、scikit-learn 1.5.2、NumPy 2.1.3 和 Matplotlib 3.9.2。"
        "单树使用 Gini 准则与 min_samples_leaf=2；Bagging 与随机森林均使用 200 棵不剪枝"
        "树和 Bootstrap 抽样。Bagging 的 max_features=1.0，随机森林的 max_features="
        "sqrt，后者对 64 个像素每次考虑 8 个特征。随机森林开启 OOB 评估并使用 "
        "balanced_subsample 类权重。主对照在完全相同的训练集和测试集上进行，5 个种子"
        "只改变模型内部随机性。",
    )
    add_table(
        doc,
        "表 2  三类模型的配置对照",
        ["配置", "单棵决策树", "全特征 Bagging", "随机森林（本文方法）"],
        [
            ["基学习器", "CART", "CART", "CART"],
            ["树数量", "1", "200", "200"],
            ["Bootstrap", "否", "是", "是"],
            ["每次节点特征数", "64", "64", "8（sqrt）"],
            ["叶节点最小样本", "2", "2", "1"],
            ["类权重", "无", "无", "balanced_subsample"],
            ["OOB 评估", "不适用", "本文不报告", "是"],
        ],
        [1.72, 1.32, 1.70, 1.76],
        font_size=8.8,
    )

    add_heading(doc, "4.3 主结果", 2)
    add_para(
        doc,
        "表 3 给出 5 个模型种子的均值与标准差。单树准确率为 "
        f"{pct(acc_tree)}±{pct(summary['baseline']['accuracy_std'])}，全特征 Bagging "
        f"提高到 {pct(acc_bag)}±{pct(summary['ablation_bagging']['accuracy_std'])}，"
        f"随机森林进一步达到 {pct(acc_rf)}±{pct(summary['proposed']['accuracy_std'])}。"
        f"由单树到 Bagging 的增益为 {signed_pp(gain_bag)}，由 Bagging 到随机森林的"
        f"增益为 {signed_pp(gain_feature)}，总增益为 {signed_pp(gain_total)}。"
        "宏平均 F1 与准确率呈现一致趋势，说明提升并非由某一个多数类单独贡献。",
    )
    add_table(
        doc,
        "表 3  5 个随机种子下的测试结果（均值 ± 标准差）",
        ["方法", "准确率", "宏 F1", "Log Loss", "Brier", "模型大小"],
        [
            [
                "单棵决策树",
                f"{pct(acc_tree)} ± {pct(summary['baseline']['accuracy_std'])}",
                f"{pct(summary['baseline']['macro_f1_mean'])} ± "
                f"{pct(summary['baseline']['macro_f1_std'])}",
                f"{summary['baseline']['log_loss_mean']:.3f}",
                f"{summary['baseline']['brier_mean']:.3f}",
                f"{summary['baseline']['model_bytes_mean']/1024:.1f} KB",
            ],
            [
                "全特征 Bagging",
                f"{pct(acc_bag)} ± {pct(summary['ablation_bagging']['accuracy_std'])}",
                f"{pct(summary['ablation_bagging']['macro_f1_mean'])} ± "
                f"{pct(summary['ablation_bagging']['macro_f1_std'])}",
                f"{summary['ablation_bagging']['log_loss_mean']:.3f}",
                f"{summary['ablation_bagging']['brier_mean']:.3f}",
                f"{summary['ablation_bagging']['model_bytes_mean']/1024:.0f} KB",
            ],
            [
                "随机森林",
                f"{pct(acc_rf)} ± {pct(summary['proposed']['accuracy_std'])}",
                f"{pct(summary['proposed']['macro_f1_mean'])} ± "
                f"{pct(summary['proposed']['macro_f1_std'])}",
                f"{summary['proposed']['log_loss_mean']:.3f}",
                f"{summary['proposed']['brier_mean']:.3f}",
                f"{summary['proposed']['model_bytes_mean']/1024:.0f} KB",
            ],
        ],
        [1.30, 1.45, 1.45, 0.90, 0.85, 1.05],
        font_size=8.4,
    )
    add_para(
        doc,
        "保存权重的种子为 20261113。在该权重上，单树、Bagging 和随机森林的测试准确率"
        f"分别为 {pct(checkpoint['baseline']['accuracy'])}、"
        f"{pct(checkpoint['ablation_bagging']['accuracy'])} 和 "
        f"{pct(checkpoint['proposed']['accuracy'])}。随机森林 5 个种子的 OOB 准确率为 "
        f"{pct(oob_mean)}±{pct(oob_std)}，与独立测试集结果接近，说明模型并非只对某一次"
        "随机划分偶然有效。不过 OOB 与测试集相差约 0.7 个百分点，这种差异应结合小测试集"
        "的抽样波动理解。",
    )
    add_figure(
        doc,
        "fig04_main_results.png",
        "图 4  不同方法的平均准确率与宏平均 F1。误差棒为 5 个模型种子的标准差。",
        width_inches=6.0,
    )
    add_para(
        doc,
        "为检验主结果稳定性，本文对保存权重在 360 个测试样本上进行 2,000 次分层"
        "重采样以外的普通 Bootstrap。单树准确率的 95% 区间为 "
        f"[{pct(ci['baseline']['low'])}, {pct(ci['baseline']['high'])}]，Bagging 为 "
        f"[{pct(ci['ablation_bagging']['low'])}, "
        f"{pct(ci['ablation_bagging']['high'])}]，随机森林为 "
        f"[{pct(ci['proposed']['low'])}, {pct(ci['proposed']['high'])}]。"
        "虽然前两者区间不重叠已提示显著差异，本文仍采用逐样本配对的 McNemar 精确检验。"
        "随机森林相对单树获得 "
        f"{mcnemar_baseline['proposed_only_correct']} 个仅前者正确样本，丢失 "
        f"{mcnemar_baseline['baseline_only_correct']} 个仅后者正确样本，双侧 p="
        f"{mcnemar_baseline['two_sided_exact_p']:.2e}。相对 Bagging，对应数量为 "
        f"{mcnemar_bagging['proposed_only_correct']} 与 "
        f"{mcnemar_bagging['baseline_only_correct']}，p="
        f"{mcnemar_bagging['two_sided_exact_p']:.4f}。",
    )

    add_heading(doc, "4.4 集成与特征随机化消融", 2)
    add_para(
        doc,
        "表 2 的三层对照直接区分了两类机制。若随机森林相对单树的提升全部来自任意"
        "集成，那么全特征 Bagging 应接近随机森林；实际后者仍高出 "
        f"{signed_pp(gain_feature)}。因此 H1 与 H2 都得到支持：Bootstrap 聚合带来主要"
        "增益，而随机像素子空间带来额外增益。图 5 的树数量曲线进一步说明了收敛过程。"
        "随机森林在 10 棵树时达到约 95.6%，25 棵树达到 97.2%，50 棵树后基本稳定在"
        "98.1%；全特征 Bagging 在 25 至 50 棵树附近饱和，但平台低约 2 个百分点。",
    )
    add_figure(
        doc,
        "fig09_tree_count.png",
        "图 5  集成规模与测试准确率。横轴为对数尺度。随机森林在约 50 棵树后收敛，"
        "且平台高于全特征 Bagging。",
        width_inches=5.9,
    )
    add_para(
        doc,
        "这一曲线不是严格单调的，因为固定测试集只有 360 个样本，1 个样本的变化对应"
        "0.278 个百分点。本文不选择曲线上最高的偶然点作为最终模型，而是按预先设定的"
        "200 棵树报告结果。该做法避免了利用测试集反复挑选树数量带来的乐观偏差。",
    )

    add_heading(doc, "4.5 噪声与遮挡鲁棒性", 2)
    add_para(
        doc,
        "为检验 H3，本文在测试图像上加入标准差为 σ 的零均值高斯噪声并裁剪到 [0,1]，"
        "以及在图像中心置零一个 b×b 方块。两种扰动都不重新训练模型。每个 σ 或遮挡"
        "尺寸均在 5 个模型种子上评估，并保持同一模型在干净与扰动测试间可比。",
    )
    add_table(
        doc,
        "表 4  典型扰动下的平均准确率（5 个模型种子）",
        ["方法", "干净数据", "噪声 σ=0.15", "噪声 σ=0.25", "遮挡 2×2", "遮挡 4×4"],
        [
            ["单棵决策树", "86.39%", "60.94%", "50.39%", "34.72%", "10.00%"],
            ["全特征 Bagging", "95.67%", "81.72%", "70.56%", "41.67%", "10.33%"],
            ["随机森林", "98.28%", "95.28%", "89.78%", "85.72%", "56.44%"],
        ],
        [1.55, 0.95, 1.12, 1.12, 0.88, 0.88],
        font_size=8.7,
    )
    add_para(
        doc,
        "噪声增加时三个模型都退化，但斜率差异明显。σ=0.25 已相当于原始灰度范围的"
        "四分之一，随机森林仍保持 89.78%，而单树和 Bagging 分别降至 50.39% 与"
        "70.56%。中心 4×4 遮挡覆盖 64 个像素中的 16 个，单树和 Bagging 几乎失去"
        "判别能力，随机森林仍为 56.44%。这与特征随机化的预期一致：不同树依赖不同的"
        "像素子集，局部特征受损时，部分树仍能提供有效证据；全特征 Bagging 的树更容易"
        "共享同一组关键像素，因而相关性更高。",
    )
    add_figure(
        doc,
        "fig05_robustness.png",
        "图 6  高斯像素噪声与中心遮挡下的测试准确率。曲线为 5 个种子的均值，误差棒为"
        "标准差。",
        width_inches=6.2,
    )
    add_para(
        doc,
        "遮挡 2×2 时 Bagging 下降到 41.67%，而 3×3 时回升到 43.83%，这说明在小样本"
        "测试条件下遮挡位置与类别先验之间存在交互，不能把曲线解释为严格单调的容量"
        "度量。本文保留这一非单调现象，并在结论中把它列为实验局限。中心遮挡更靠近"
        "多数数字的笔画区域，也会使不同尺寸之间存在分布差异。",
    )

    add_heading(doc, "4.6 学习曲线", 2)
    add_para(
        doc,
        "为比较小样本条件下的数据利用率，本文从训练集中按类别分层抽取 10% 至 100% 的"
        "样本，使用 3 个种子重复训练。图 7 显示，随机森林在仅 140 个训练样本时已达到"
        "约 92.5%，单树约为 66.4%；在 580 个样本时随机森林约为 97.9%，已接近其完整"
        "训练集表现。单树在后期趋于平稳，且完整训练集平均甚至略低于 80% 点的 86.85%，"
        "这再次体现了单树对具体样本的敏感性和方差。",
    )
    add_figure(
        doc,
        "fig06_learning_curve.png",
        "图 7  训练样本数与测试准确率。随机森林在较小训练预算下即表现出更高的数据效率。",
        width_inches=6.0,
    )
    add_para(
        doc,
        "学习曲线不能直接证明随机森林拥有更低的真实误差上界，但可以说明在本数据的"
        "样本规模和像素表征下，多棵随机化树的平均比单一高方差树更稳定。随机森林在"
        "80% 与 100% 训练子集之间的波动也表明，其优势并非随样本增加无限扩大，"
        "而是接近此特征空间可达到的平台。",
    )

    add_heading(doc, "4.7 分类错误与可解释性", 2)
    add_para(
        doc,
        "图 8 为保存权重在测试集上的混淆矩阵。随机森林仅产生 7 个错误：真实 0 被预测"
        "为 4、3 被预测为 7、7 被预测为 9，以及 8 被预测为 1、3、7 各一次，另有 8 的"
        "两个样本误判为 1。类别 8 的 F1 最低，为 93.94%，但相较多棵对照已有明显改善。"
        "错误集中在笔画局部相近、分辨率又只有 8×8 的类别对，例如 3/7 与 7/9。",
    )
    add_figure(
        doc,
        "fig07_confusion_matrix.png",
        "图 8  保存的随机森林权重在 360 个测试样本上的混淆矩阵。颜色越深代表样本越多。",
        width_inches=4.8,
    )
    add_para(
        doc,
        "图 9 左侧为像素重要性热图。较高重要性主要分布在数字笔画可能出现的中心与"
        "上部区域，但没有任何单个像素占据决定性地位，说明随机特征子空间使模型利用了"
        "多组局部证据。右侧逐类 F1 显示，随机森林在 10 个类别上均不低于或接近"
        "Bagging，较好的类别 2、5、6 达到 100%，最低的类别 8 仍有 93.94%。",
    )
    add_figure(
        doc,
        "fig08_importance_per_class.png",
        "图 9  随机森林的像素重要性热图与单树/随机森林逐类 F1 对照。",
        width_inches=6.2,
    )

    add_heading(doc, "4.8 校准、效率与讨论", 2)
    add_para(
        doc,
        "主表还揭示了一个重要反例：随机森林准确率最高，但 Log Loss 为 "
        f"{summary['proposed']['log_loss_mean']:.3f}，高于 Bagging 的 "
        f"{summary['ablation_bagging']['log_loss_mean']:.3f}；10 箱期望校准误差 "
        f"（ECE）也从 {pct(summary['ablation_bagging']['ece_mean'])} 上升到 "
        f"{pct(summary['proposed']['ece_mean'])}。这说明随机森林虽然更常给出正确类别，"
        "却倾向于输出过度自信的概率。可能原因是随机特征使不同树的概率分布更尖锐，"
        "且较小的叶节点在叶内只见过极少样本。对于只关心 top-1 类别的场景，这一代价"
        "可以接受；对于风险控制、拒识或概率决策，需要额外校准。该结果说明本文并没有"
        "把“高准确率”误写成“所有指标全面更优”。",
    )
    add_para(
        doc,
        "从效率上看，单树保存权重约 6.9 KB，Bagging 约 936 KB，随机森林约 2.61 MB。"
        "随机森林权重更大，但在本实验的 360 个测试样本上，单次批量推理的中位延迟仍"
        "低于 0.1 ms/样本，具体数值受 CPU、内存和首次调用开销影响，只适合作为数量级"
        "说明。对于嵌入式或低存储场景，可通过并行训练、树剪枝或模型蒸馏进一步压缩。",
    )
    add_para(
        doc,
        "本文实验仍存在若干局限。第一，8×8 分辨率低于现代手写数字识别系统常用的"
        "28×28 或更高分辨率，像素级结论不能直接迁移到高分辨率图像。第二，测试集只有"
        "360 张，置信区间较宽，个别遮挡点出现非单调变化。第三，所有模型都使用原始像素，"
        "没有引入卷积、HOG 或仿射增广；因此本文不支持“随机森林优于深度网络”的结论。"
        "第四，超参数固定为历史常用设置，没有进行大规模搜索，这避免了测试集污染，"
        "但也意味着每个模型的潜力未必完全释放。第五，模型种子重复共享同一测试集，"
        "标准差反映的是模型随机性，而不是数据划分随机性。",
    )
    add_para(
        doc,
        "更完整的后续研究应至少采用嵌套交叉验证，报告 HOG/卷积特征、数据增广和概率"
        "校准后的结果；还可以引入旋转、平移、笔画断裂等更有结构意义的扰动，检查"
        "随机森林的鲁棒性是否仍成立。本文的贡献在于把这些边界条件明确写出，并提供"
        "可查证的原始表格、权重和运行日志。",
    )

    add_heading(doc, "5 结论", 1)
    add_para(
        doc,
        "本文在一个固定、可重复的手写数字视觉识别实验中复现了随机森林的历史核心"
        "思想。单棵 CART 树、全特征 Bagging 与随机森林的平均准确率分别为 "
        f"{pct(acc_tree)}、{pct(acc_bag)} 和 {pct(acc_rf)}。三层对照支持两个机制："
        "Bootstrap 聚合显著降低单树方差，随机像素子空间则在集成基础上进一步增加"
        f"多样性，贡献了 {100.0 * gain_feature:.2f} 个百分点的额外准确率。"
        "树数量曲线、学习曲线、"
        "逐类 F1 与特征重要性热图均与这一解释一致。",
    )
    add_para(
        doc,
        "扰动实验给出了更直接的证据。在高斯噪声和中心遮挡下，随机森林的退化幅度"
        "显著小于单树和全特征 Bagging，说明随机特征子空间不仅提高了干净测试集准确率，"
        "也使模型对局部像素损坏更不敏感。这一结果与引言中的 H1、H2、H3 相符。"
        "与此同时，随机森林的 Log Loss 和 ECE 较差，表明集成准确率提升和概率校准"
        "质量可以分离。",
    )
    add_para(
        doc,
        "因此，本文得到的结论是条件化的：在 64 维原始像素、约 1,400 个训练样本的"
        "历史设置下，随机森林相对其直接前身表现出稳定且可解释的优势；该结论不能被"
        "扩展为对现代深度视觉模型的普遍优越性。所有报告值均来自本机的真实训练与"
        "评估，代码、权重、原始 CSV、日志和图件共同构成可审计的复现材料。",
    )

    add_heading(doc, "参考文献", 1)
    references = [
        "[1] Breiman L. Random Forests[J]. Machine Learning, 2001, 45(1): 5-32.",
        "[2] Ho T K. The Random Subspace Method for Constructing Decision Forests[J]. "
        "IEEE Transactions on Pattern Analysis and Machine Intelligence, 1998, "
        "20(8): 832-844.",
        "[3] Breiman L. Bagging Predictors[J]. Machine Learning, 1996, 24(2): 123-140.",
        "[4] Biau G, Scornet E. A Random Forest Guided Tour[J]. TEST, 2016, 25: 197-227.",
        "[5] Amit Y, Geman D. Shape Quantization and Recognition with Randomized Trees[J]. "
        "Neural Computation, 1997, 9(7): 1545-1588.",
        "[6] LeCun Y, Bottou L, Bengio Y, et al. Gradient-based Learning Applied to "
        "Document Recognition[J]. Proceedings of the IEEE, 1998, 86(11): 2278-2324.",
        "[7] Krizhevsky A, Sutskever I, Hinton G E. ImageNet Classification with Deep "
        "Convolutional Neural Networks[C]//Advances in Neural Information Processing "
        "Systems. 2012: 1097-1105.",
        "[8] He K, Zhang X, Ren S, et al. Deep Residual Learning for Image Recognition[C]"
        "//IEEE Conference on Computer Vision and Pattern Recognition. 2016: 770-778.",
        "[9] Deng J, Dong W, Socher R, et al. ImageNet: A Large-Scale Hierarchical Image "
        "Database[C]//IEEE Conference on Computer Vision and Pattern Recognition. "
        "2009: 248-255.",
        "[10] Turk M, Pentland A. Eigenfaces for Recognition[J]. Journal of Cognitive "
        "Neuroscience, 1991, 3(1): 71-86.",
        "[11] Lowe D G. Distinctive Image Features from Scale-Invariant Keypoints[J]. "
        "International Journal of Computer Vision, 2004, 60(2): 91-110.",
        "[12] Dalal N, Triggs B. Histograms of Oriented Gradients for Human Detection[C]"
        "//IEEE Conference on Computer Vision and Pattern Recognition. 2005: 886-893.",
        "[13] Alpaydin E, Kaynak C. Optical Recognition of Handwritten Digits[DS]. "
        "UCI Machine Learning Repository, 1998.",
        "[14] Pedregosa F, Varoquaux G, Gramfort A, et al. Scikit-learn: Machine Learning "
        "in Python[J]. Journal of Machine Learning Research, 2011, 12: 2825-2830.",
        "[15] Dietterich T G. Ensemble Methods in Machine Learning[C]//Multiple Classifier "
        "Systems. 2000: 1-15.",
        "[16] Zhou Z H. Ensemble Methods: Foundations and Algorithms[M]. Boca Raton: "
        "Chapman and Hall/CRC, 2012.",
        "[17] Breiman L, Friedman J, Stone C J, et al. Classification and Regression "
        "Trees[M]. Belmont: Wadsworth, 1984.",
        "[18] Fernandez-Delgado M, Cernadas E, Barro S, et al. Do We Need Hundreds of "
        "Classifiers to Solve Real World Classification Problems?[J]. Journal of Machine "
        "Learning Research, 2014, 15: 3133-3181.",
        "[19] Dua D, Graff C. UCI Machine Learning Repository[DB/OL]. "
        "University of California, Irvine, 2019.",
    ]
    for reference in references:
        add_para(
            doc,
            reference,
            size=9.5,
            line=1.15,
            after=4,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )

    doc.add_page_break()
    add_heading(doc, "附录 A 代码、权重与复现说明", 1)
    add_para(
        doc,
        "本文使用独立的手写数字随机森林复现项目，不复用其他论文实验的代码或结果。"
        "仓库包含数据加载、三层模型训练、全部扰动分析、论文图表生成、权重独立回载验证"
        "和 DOCX 论文构建脚本。主实验命令如下：",
    )
    code = doc.add_paragraph()
    code.paragraph_format.left_indent = Inches(0.25)
    code.paragraph_format.space_before = Pt(4)
    code.paragraph_format.space_after = Pt(8)
    code.alignment = WD_ALIGN_PARAGRAPH.LEFT
    code_run = code.add_run(
        "python src/run_experiment.py\npython src/evaluate_checkpoints.py\n"
        "python src/build_paper.py"
    )
    set_run_font(
        code_run,
        latin="Consolas",
        cjk=BODY_FONT_CJK,
        size=9.2,
        color=NAVY,
    )
    link_paragraph = doc.add_paragraph()
    link_paragraph.paragraph_format.space_before = Pt(2)
    link_paragraph.paragraph_format.space_after = Pt(8)
    prefix = link_paragraph.add_run("GitHub 仓库：")
    set_run_font(prefix, size=10.5, bold=True, color=NAVY)
    add_hyperlink(
        link_paragraph,
        "打开 forest-ocr-reproduction GitHub 仓库",
        GITHUB_URL,
    )

    files = [
        ("baseline_tree.joblib", CHECKPOINT_DIR / "baseline_tree.joblib"),
        (
            "bagging_all_features.joblib",
            CHECKPOINT_DIR / "bagging_all_features.joblib",
        ),
        (
            "proposed_random_forest.joblib",
            CHECKPOINT_DIR / "proposed_random_forest.joblib",
        ),
        ("metrics.json", RESULT_DIR / "metrics.json"),
        ("predictions.npz", RESULT_DIR / "predictions.npz"),
    ]
    rows = []
    for label, path in files:
        rows.append(
            [
                label,
                f"{path.stat().st_size / 1024:.1f} KB",
                sha256(path)[:16] + "…",
            ]
        )
    add_table(
        doc,
        "表 A1  主要复现产物与 SHA-256 前缀",
        ["文件", "大小", "SHA-256 前缀"],
        rows,
        [3.0, 1.2, 2.3],
        font_size=8.5,
    )
    add_para(
        doc,
        "权重输入约定为按 C 顺序展开的 8×8 灰度图，像素值除以 16。"
        "加载 joblib 后，artifact['model'] 为可直接调用的分类器，"
        "artifact['metrics'] 保存训练时的评估元数据。保存权重的随机种子为 "
        f"{data['configuration']['checkpoint_seed']}，固定数据划分为 "
        f"{data['configuration']['eval_seeds']} 中的第一项。论文表中的均值来自全部 "
        f"{len(data['configuration']['eval_seeds'])} 个种子，而不是只选择权重种子。",
    )
    add_para(
        doc,
        "由于环境日期为 2026 年，本研究的“历史论文复现”指算法与论文结论层面的复现，"
        "不是对原作者代码、原始大规模数据集或当年硬件环境的逐位复刻。本文没有用生成"
        "数据冒充真实样本，也没有修改结果 CSV 或手工挑选测试用例。",
    )

    PAPER_DIR.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    print(f"Wrote {output_path}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--output",
        type=Path,
        default=PAPER_DIR / "Random_Forest_Digit_Reproduction_Paper.docx",
    )
    args = parser.parse_args()
    build_paper(args.output)


if __name__ == "__main__":
    main()
