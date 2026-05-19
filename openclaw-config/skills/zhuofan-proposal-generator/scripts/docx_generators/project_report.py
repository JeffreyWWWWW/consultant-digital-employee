import os
import re
from datetime import date

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


FONT_FANGSONG = "仿宋_GB2312"
FONT_HEITI = "黑体"
FONT_XIAOBIAOSONG = "方正小标宋_GBK"


def _set_run_font(run, font_name: str, size_pt: int, bold: bool = False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)


def _set_paragraph_format(p, first_line_indent: bool = True, space_before: int = 0, space_after: int = 0):
    fmt = p.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    fmt.space_before = Pt(space_before)
    fmt.space_after = Pt(space_after)
    if first_line_indent:
        fmt.first_line_indent = Cm(0.74)


def _setup_page(doc):
    section = doc.sections[0]
    section.top_margin = Cm(3.7)
    section.bottom_margin = Cm(3.5)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.6)


def _add_run(p, text: str, font: str = FONT_FANGSONG, size: int = 16, bold: bool = False):
    run = p.add_run(text or "")
    _set_run_font(run, font, size, bold)
    return run


def _heading(doc, text: str, level: int):
    p = doc.add_paragraph()
    _set_paragraph_format(p, first_line_indent=False, space_before=8, space_after=4)
    if level == 1:
        _add_run(p, text, FONT_HEITI, 16, True)
    else:
        _add_run(p, text, FONT_FANGSONG, 16, True)
    return p


def _add_paragraphs(doc, text):
    if not text:
        return
    if isinstance(text, list):
        for item in text:
            _add_paragraphs(doc, item)
        return
    for para in str(text).split("\n"):
        para = para.strip()
        if not para:
            continue
        p = doc.add_paragraph()
        _set_paragraph_format(p)
        _add_run(p, para)


def _clean_title(project_name: str) -> str:
    name = (project_name or "").strip(" 　《》")
    name = re.sub(r"(解决方案|方案初稿|初稿)$", "", name).strip(" 　-_")
    return name


def _project_report_title(project_name: str, customer_name: str = "") -> str:
    name = _clean_title(project_name)
    customer = (customer_name or "").strip(" 　《》")
    if "汇报" in name:
        if customer and customer not in name:
            return f"{customer}{name}"
        return name
    if name.endswith("建设项目"):
        report_name = f"{name}汇报"
    else:
        report_name = f"{name}建设项目汇报"
    if customer and customer not in report_name:
        return f"{customer}{report_name}"
    return report_name


def _safe_filename(text: str) -> str:
    text = (text or "").replace("/", "_").replace("\\", "_").replace(" ", "")
    return re.sub(r'[<>:"|?*]', "_", text)


def _default_output_path(region: str, project_name: str, customer_name: str = "") -> str:
    today = date.today().strftime("%Y%m%d")
    report_name = _safe_filename(_project_report_title(project_name, customer_name))
    return f"{report_name}_{today}.docx"


def _normalize_output_path(output_path: str, region: str, project_name: str, customer_name: str = "") -> str:
    filename = _default_output_path(region, project_name, customer_name)
    if not output_path:
        return filename
    output_dir = os.path.dirname(output_path) or "."
    return os.path.join(output_dir, filename)


def _project_contents(sections: dict):
    return sections.get("project_contents") or sections.get("modules") or []


def _format_source_item(item) -> str:
    if not isinstance(item, dict):
        return str(item)
    name = item.get("name") or item.get("title") or item.get("source_name") or item.get("policy_name") or "未命名来源"
    agency = item.get("agency") or item.get("publisher") or item.get("source") or item.get("source_org")
    doc_no = item.get("doc_no") or item.get("document_no")
    date_text = item.get("date") or item.get("publish_date") or item.get("published_at")
    url = item.get("url") or item.get("link") or item.get("source_url")
    status = item.get("status") or item.get("verification_status") or item.get("核验状态")
    used_for = item.get("used_for") or item.get("support") or item.get("purpose")
    parts = [str(name)]
    meta = [value for value in (agency, doc_no, date_text) if value]
    if meta:
        parts.append("（" + "，".join(map(str, meta)) + "）")
    if url:
        parts.append(f"：{url}")
    elif status:
        parts.append(f"：{status}")
    else:
        parts.append("：待核验原文")
    if used_for:
        parts.append(f"；支撑内容：{used_for}")
    return "".join(parts)


def _add_bullets(doc, items):
    if not items:
        return
    if isinstance(items, str):
        _add_paragraphs(doc, items)
        return
    for item in items:
        text = _format_source_item(item)
        p = doc.add_paragraph()
        _set_paragraph_format(p)
        _add_run(p, f"- {text}")


def _append_review_appendix(doc, sections: dict):
    sources = (
        sections.get("policy_sources")
        or sections.get("source_links")
        or sections.get("sources")
        or sections.get("source_note")
    )
    review_notes = sections.get("review_notes") or []

    _heading(doc, "附录A：政策与资料来源链接", 1)
    if sources:
        _add_bullets(doc, sources)
    else:
        _add_paragraphs(doc, "【待补充政策与资料来源链接】")

    _heading(doc, "附录B：需人工审核事项", 1)
    if review_notes:
        _add_bullets(doc, review_notes)
    else:
        _add_paragraphs(doc, "【待补充需人工审核事项】")


def build_docx(data: dict, output_path: str = None) -> str:
    project_name = data["project_name"]
    customer_name = data.get("customer_name", "")
    region = data.get("region", "")
    sections = data.get("sections", {})

    doc = Document()
    _setup_page(doc)

    style = doc.styles["Normal"]
    style.font.name = FONT_FANGSONG
    style.font.size = Pt(16)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_FANGSONG)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_paragraph_format(title_p, first_line_indent=False, space_after=18)
    _add_run(title_p, _project_report_title(project_name, customer_name), FONT_XIAOBIAOSONG, 22, True)

    _heading(doc, "项目建设的依据", 1)
    _add_paragraphs(doc, sections.get("project_basis") or sections.get("policy_background") or "【待补充项目建设依据】")

    _heading(doc, "建设目标", 1)
    _add_paragraphs(doc, sections.get("project_goals") or sections.get("overall_goal") or "【待补充建设目标】")

    _heading(doc, "建设内容", 1)
    contents = _project_contents(sections)
    if isinstance(contents, list) and contents:
        for item in contents:
            if isinstance(item, dict):
                name = item.get("name") or item.get("title")
                content = item.get("content") or item.get("description") or item.get("summary")
                if name:
                    _heading(doc, name, 2)
                _add_paragraphs(doc, content)
            else:
                _add_paragraphs(doc, item)
    else:
        _add_paragraphs(doc, contents or "【待补充建设内容】")

    _append_review_appendix(doc, sections)

    output_path = _normalize_output_path(output_path, region, project_name, customer_name)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    return os.path.abspath(output_path)
