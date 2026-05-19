import os
import re
from datetime import date

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


FONT_FANGSONG = "仿宋_GB2312"
FONT_HEITI = "黑体"
FONT_XIAOBIAOSONG = "方正小标宋_GBK"
REVIEW_MARKERS = (
    "待核验",
    "待确认",
    "待补充",
    "待测算",
    "需人工审核",
    "需人工复核",
    "推断",
    "建议",
    "未核验",
    "模型生成",
    "模型建议",
)
REVIEW_HIGHLIGHT_NOTE = "文中黄色标注内容为需人工核对事项，请结合原始材料、政策原文和客户确认结果复核。"
CN_NUMERALS = ("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")


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


def _add_hyperlink(p, text: str, url: str):
    part = p.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    r_style = OxmlElement("w:rStyle")
    r_style.set(qn("w:val"), "Hyperlink")
    r_pr.append(r_style)
    run.append(r_pr)
    text_element = OxmlElement("w:t")
    text_element.text = text or url
    run.append(text_element)
    hyperlink.append(run)
    p._p.append(hyperlink)


def _heading(doc, text: str, level: int):
    p = doc.add_paragraph()
    _set_paragraph_format(p, first_line_indent=False, space_before=8, space_after=4)
    if level == 1:
        _add_run(p, text, FONT_HEITI, 16, True)
    else:
        _add_run(p, text, FONT_FANGSONG, 16, True)
    return p


def _cn_number(index: int) -> str:
    if 1 <= index <= 10:
        return CN_NUMERALS[index - 1]
    if 11 <= index <= 19:
        return "十" + CN_NUMERALS[index - 11]
    if index == 20:
        return "二十"
    return str(index)


def _numbered_subheading(index: int, text: str) -> str:
    text = (text or "").strip()
    if re.match(r"^（[一二三四五六七八九十百千万]+）", text):
        return text
    if re.match(r"^[一二三四五六七八九十百千万]+、", text):
        return text
    return f"（{_cn_number(index)}）{text}"


def _page_break(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def _needs_review_highlight(text: str) -> bool:
    return any(marker in (text or "") for marker in REVIEW_MARKERS)


def _add_text_run(p, text: str, highlight_review: bool = True):
    if not highlight_review or not _needs_review_highlight(text):
        return _add_run(p, text)

    pattern = "(" + "|".join(re.escape(marker) for marker in REVIEW_MARKERS) + ")"
    last = 0
    highlighted_run = None
    for match in re.finditer(pattern, text):
        if match.start() > last:
            _add_run(p, text[last:match.start()])
        highlighted_run = _add_run(p, match.group(0))
        highlighted_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        last = match.end()
    if last < len(text):
        _add_run(p, text[last:])
    return highlighted_run


def _add_paragraphs(doc, text, highlight_review: bool = True):
    if not text:
        return
    if isinstance(text, list):
        for item in text:
            _add_paragraphs(doc, item, highlight_review=highlight_review)
        return
    for para in str(text).split("\n"):
        para = para.strip()
        if not para:
            continue
        p = doc.add_paragraph()
        _set_paragraph_format(p)
        _add_text_run(p, para, highlight_review=highlight_review)


def _highlight_terms_in_document(doc, terms):
    terms = [str(term).strip() for term in (terms or []) if str(term).strip()]
    if not terms:
        return
    for paragraph in doc.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        if not full_text or not any(term in full_text for term in terms):
            continue
        original_runs = paragraph.runs
        if not original_runs:
            continue
        base_font = original_runs[0].font.name or FONT_FANGSONG
        base_size = int(original_runs[0].font.size.pt) if original_runs[0].font.size else 16
        base_bold = bool(original_runs[0].font.bold)
        for run in list(original_runs):
            run._element.getparent().remove(run._element)

        pattern = "(" + "|".join(re.escape(term) for term in sorted(terms, key=len, reverse=True)) + ")"
        last = 0
        for match in re.finditer(pattern, full_text):
            if match.start() > last:
                _add_run(paragraph, full_text[last:match.start()], base_font, base_size, base_bold)
            highlighted_run = _add_run(paragraph, match.group(0), base_font, base_size, base_bold)
            highlighted_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            last = match.end()
        if last < len(full_text):
            _add_run(paragraph, full_text[last:], base_font, base_size, base_bold)


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
    elif name.endswith("建设"):
        report_name = f"{name}项目汇报"
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


def _source_fields(item):
    if not isinstance(item, dict):
        return {"text": str(item), "url": ""}
    source_names = item.get("source_names") or []
    source_urls = item.get("source_urls") or []
    if isinstance(source_names, str):
        source_names = [source_names]
    if isinstance(source_urls, str):
        source_urls = [source_urls]
    name = (
        item.get("name")
        or item.get("title")
        or item.get("source_name")
        or item.get("policy_name")
        or (source_names[0] if source_names else "")
        or "未命名来源"
    )
    agency = item.get("agency") or item.get("publisher") or item.get("source") or item.get("source_org")
    doc_no = item.get("doc_no") or item.get("document_no")
    date_text = item.get("date") or item.get("publish_date") or item.get("published_at")
    url = item.get("url") or item.get("link") or item.get("source_url") or (source_urls[0] if source_urls else "")
    status = item.get("status") or item.get("verification_status") or item.get("核验状态")
    used_for = item.get("used_for") or item.get("support") or item.get("purpose")
    meta = [value for value in (agency, doc_no, date_text) if value]
    prefix = f"- {name}"
    if meta:
        prefix += "（" + "，".join(map(str, meta)) + "）"
    suffix = ""
    if used_for:
        suffix += f"；支撑内容：{used_for}"
    if not url:
        suffix = f"：{status or '待核验原文'}" + suffix
    return {"prefix": prefix, "url": url or "", "suffix": suffix}


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


def _add_source_bullets(doc, items):
    if not items:
        return
    if isinstance(items, str):
        _add_paragraphs(doc, items)
        return
    for item in items:
        fields = _source_fields(item)
        p = doc.add_paragraph()
        _set_paragraph_format(p)
        if "text" in fields:
            _add_run(p, fields["text"])
            continue
        _add_run(p, fields["prefix"])
        if fields["url"]:
            _add_run(p, "：")
            _add_hyperlink(p, fields["url"], fields["url"])
        _add_run(p, fields["suffix"])


def _append_review_appendix(doc, sections: dict):
    sources = (
        sections.get("policy_sources")
        or sections.get("source_links")
        or sections.get("sources")
        or sections.get("source_note")
    )
    review_notes = sections.get("review_notes") or []

    _page_break(doc)
    _heading(doc, "附录A：政策与资料来源链接", 1)
    if sources:
        _add_source_bullets(doc, sources)
    else:
        _add_paragraphs(doc, "【待补充政策与资料来源链接】")

    _page_break(doc)
    _heading(doc, "附录B：需人工审核事项", 1)
    _add_paragraphs(doc, REVIEW_HIGHLIGHT_NOTE, highlight_review=False)
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

    _heading(doc, "一、项目建设的依据", 1)
    _add_paragraphs(doc, sections.get("project_basis") or sections.get("policy_background") or "【待补充项目建设依据】")

    _heading(doc, "二、建设目标", 1)
    _add_paragraphs(doc, sections.get("project_goals") or sections.get("overall_goal") or "【待补充建设目标】")

    _heading(doc, "三、建设内容", 1)
    contents = _project_contents(sections)
    if isinstance(contents, list) and contents:
        for idx, item in enumerate(contents, start=1):
            if isinstance(item, dict):
                name = item.get("name") or item.get("title")
                content = item.get("content") or item.get("description") or item.get("summary")
                if name:
                    _heading(doc, _numbered_subheading(idx, name), 2)
                _add_paragraphs(doc, content)
            else:
                _add_paragraphs(doc, item)
    else:
        _add_paragraphs(doc, contents or "【待补充建设内容】")

    _highlight_terms_in_document(doc, sections.get("review_highlights") or sections.get("manual_review_terms"))

    _append_review_appendix(doc, sections)

    output_path = _normalize_output_path(output_path, region, project_name, customer_name)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    return os.path.abspath(output_path)
