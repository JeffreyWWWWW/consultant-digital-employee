"""
ZX-01 方案初稿 → Word 文档生成器（Skill 内嵌版）

用法：
  python generate_docx.py --json proposal.json
  python generate_docx.py --json proposal.json --output /path/to/output.docx

输入 JSON 结构见 SCHEMA 常量。
依赖：pip install python-docx
"""

import json
import sys
import os
import argparse
from datetime import date

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
except ImportError:
    print("ERROR: python-docx 未安装，请执行 pip install python-docx", file=sys.stderr)
    sys.exit(1)


SCHEMA = """
输入 JSON 结构：
{
  "project_name": "项目名称",
  "customer_name": "客户单位名称",
  "customer_type": "数据局/政务中心/大数据中心/数据集团/企业",
  "region": "区域",
  "sections": {
    "policy_background": "政策背景正文（可多段，用\\n分隔）",
    "current_status": "现状概述正文",
    "pain_points": ["痛点1", "痛点2", "痛点3"],
    "overall_goal": "总体目标正文",
    "sub_goals": ["分项目标1", "分项目标2"],
    "modules": [
      {
        "name": "模块名称",
        "content": "模块详细内容",
        "product": "匹配的卓繁产品（无则写待确认）"
      }
    ],
    "architecture": "技术架构描述（空字符串则用默认五层架构）",
    "tech_selection": "技术选型说明",
    "phases": [
      {"name": "一期", "duration": "6个月", "content": "建设内容", "deliverables": "交付物"}
    ],
    "budget_items": [
      {"name": "模块名", "amount": "待定", "note": ""}
    ],
    "highlights": ["亮点1", "亮点2"],
    "benefits": ["效益1", "效益2"],
    "cases": [
      {"name": "案例名", "source": "来源", "summary": "摘要"}
    ],
    "review_notes": ["需人工审核项1", "需人工审核项2"],
    "ref_proposals": "参考的历史方案",
    "ref_products": "匹配的产品模块",
    "ref_policies_count": 0,
    "ref_cases_count": 0,
    "source_ratio": "知识库 30% / 政策库 20% / 模型生成 50%"
  }
}
"""


# ── 工具函数 ──────────────────────────────────────────────

def _set_cell_shading(cell, color_hex: str):
    tc_pr = cell._element.get_or_add_tcPr()
    shading = tc_pr.makeelement(
        qn("w:shd"), {qn("w:fill"): color_hex, qn("w:val"): "clear"}
    )
    tc_pr.append(shading)


def _heading(doc, text, level):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x1A, 0x3C, 0x6E)
    return h


def _add_paragraphs(doc, text: str):
    """将含 \\n 的长文本拆成多段落写入"""
    for line in text.split("\n"):
        line = line.strip()
        if line:
            doc.add_paragraph(line)


def _make_header_row(table, headers, bg="1A3C6E"):
    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = h
        _set_cell_shading(cell, bg)
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                r.bold = True


# ── 主函数 ────────────────────────────────────────────────

def build_proposal_docx(data: dict, output_path: str = None) -> str:
    project_name = data["project_name"]
    customer_name = data["customer_name"]
    customer_type = data["customer_type"]
    region = data["region"]
    s = data.get("sections", {})

    doc = Document()

    # 全局字体
    style = doc.styles["Normal"]
    style.font.name = "仿宋"
    style.font.size = Pt(14)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")
    style.paragraph_format.line_spacing = 1.5

    # ───── 封面 ─────
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.space_after = Pt(6)
    run = title_p.add_run(f"《{project_name}》解决方案")
    run.bold = True
    run.font.size = Pt(22)
    run.font.color.rgb = RGBColor(0x1A, 0x3C, 0x6E)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.space_after = Pt(24)
    run = sub_p.add_run("（初稿）")
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # 元信息
    meta_table = doc.add_table(rows=4, cols=4)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_rows = [
        ("编制单位", "卓繁信息集团股份有限公司", "目标客户", customer_name),
        ("客户类型", customer_type, "区域", region),
        ("编制日期", date.today().strftime("%Y年%m月%d日"), "版本", "V0.1"),
        ("状态", "AI辅助生成初稿 - 待人工审核", "", ""),
    ]
    for i, (k1, v1, k2, v2) in enumerate(meta_rows):
        cells = meta_table.rows[i].cells
        cells[0].text = k1
        cells[1].text = v1
        if k2:
            cells[2].text = k2
            cells[3].text = v2
        for j in [0, 2]:
            if cells[j].text:
                _set_cell_shading(cells[j], "E8EDF5")
                for p in cells[j].paragraphs:
                    for r in p.runs:
                        r.bold = True

    doc.add_paragraph("")

    # ───── 一、现状分析 ─────
    _heading(doc, "一、现状分析", 1)

    _heading(doc, "1.1 政策背景", 2)
    _add_paragraphs(doc, s.get("policy_background", "【待补充】"))

    _heading(doc, "1.2 现状概述", 2)
    _add_paragraphs(doc, s.get("current_status", "【待调研补充】"))

    _heading(doc, "1.3 痛点问题", 2)
    ordinals = ["一是", "二是", "三是", "四是", "五是", "六是", "七是"]
    for i, pt in enumerate(s.get("pain_points", [])):
        prefix = ordinals[i] if i < len(ordinals) else f"第{i+1}，"
        doc.add_paragraph(f"{prefix}，{pt}", style="List Bullet")

    # ───── 二、建设目标 ─────
    _heading(doc, "二、建设目标", 1)

    _heading(doc, "2.1 总体目标", 2)
    _add_paragraphs(doc, s.get("overall_goal", "【待补充】"))

    _heading(doc, "2.2 分项目标", 2)
    for g in s.get("sub_goals", []):
        doc.add_paragraph(g, style="List Bullet")

    # ───── 三、建设内容 ─────
    _heading(doc, "三、建设内容", 1)
    for i, mod in enumerate(s.get("modules", []), 1):
        _heading(doc, f"3.{i} {mod['name']}", 2)
        _add_paragraphs(doc, mod.get("content", ""))
        prod_p = doc.add_paragraph()
        run = prod_p.add_run(f"【产品匹配】{mod.get('product', '待确认')}")
        run.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
        run.bold = True

    # ───── 四、技术架构 ─────
    _heading(doc, "四、技术架构", 1)

    _heading(doc, "4.1 总体架构", 2)
    arch = s.get("architecture", "")
    if arch:
        _add_paragraphs(doc, arch)
    else:
        for layer in [
            "基础设施层：政务云、网络、计算存储等基础资源",
            "数据资源层：数据采集、治理、共享交换、数据中台",
            "应用支撑层：统一身份认证、工作流引擎、消息中心、开发框架",
            "业务应用层：各业务子系统与应用模块",
            "展示层：领导驾驶舱、门户网站、移动端",
        ]:
            doc.add_paragraph(layer, style="List Bullet")
        doc.add_paragraph("")
        doc.add_paragraph("两大支撑体系：标准规范体系、安全保障体系。")

    _heading(doc, "4.2 技术选型", 2)
    _add_paragraphs(doc, s.get("tech_selection", "【待补充技术选型说明】"))

    # ───── 五、实施计划 ─────
    _heading(doc, "五、实施计划", 1)

    _heading(doc, "5.1 分期实施路径", 2)
    phases = s.get("phases", [])
    if phases:
        t = doc.add_table(rows=len(phases) + 1, cols=4)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        _make_header_row(t, ["阶段", "周期", "建设内容", "交付物"])
        for i, ph in enumerate(phases, 1):
            t.rows[i].cells[0].text = ph.get("name", "")
            t.rows[i].cells[1].text = ph.get("duration", "")
            t.rows[i].cells[2].text = ph.get("content", "")
            t.rows[i].cells[3].text = ph.get("deliverables", "")

    _heading(doc, "5.2 项目组织", 2)
    for role in ["项目领导小组", "项目管理办公室（PMO）", "技术实施团队", "业务配合团队"]:
        doc.add_paragraph(role, style="List Bullet")

    # ───── 六、预算框架 ─────
    _heading(doc, "六、预算框架", 1)
    items = s.get("budget_items", [])
    if items:
        t = doc.add_table(rows=len(items) + 2, cols=4)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        _make_header_row(t, ["序号", "建设内容", "预算（万元）", "备注"])
        for i, item in enumerate(items, 1):
            t.rows[i].cells[0].text = str(i)
            t.rows[i].cells[1].text = item.get("name", "")
            t.rows[i].cells[2].text = item.get("amount", "待定")
            t.rows[i].cells[3].text = item.get("note", "")
        last = t.rows[-1]
        last.cells[1].text = "合计"
        last.cells[2].text = "待定"
        for c in last.cells:
            for p in c.paragraphs:
                for r in p.runs:
                    r.bold = True

    warn = doc.add_paragraph()
    run = warn.add_run("注：预算为估算框架，具体金额由商务部门确认。")
    run.font.color.rgb = RGBColor(0xCC, 0x66, 0x00)
    run.font.size = Pt(10)

    # ───── 七、亮点与预期效益 ─────
    _heading(doc, "七、亮点与预期效益", 1)

    _heading(doc, "7.1 方案亮点", 2)
    for h in s.get("highlights", []):
        doc.add_paragraph(h, style="List Bullet")

    _heading(doc, "7.2 预期效益", 2)
    for b in s.get("benefits", []):
        doc.add_paragraph(b, style="List Bullet")

    _heading(doc, "7.3 成功案例参考", 2)
    cases = s.get("cases", [])
    if cases:
        for case in cases:
            p = doc.add_paragraph()
            run = p.add_run(case.get("name", ""))
            run.bold = True
            if case.get("summary"):
                doc.add_paragraph(case["summary"])
            src_p = doc.add_paragraph()
            run = src_p.add_run(f"[来源] {case.get('source', '知识库')}")
            run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            run.font.size = Pt(9)
    else:
        doc.add_paragraph("【待补充成功案例】")

    # ───── 附录A：人工审核提示 ─────
    doc.add_page_break()
    _heading(doc, "附录A：人工审核提示", 1)
    default_reviews = [
        "政策引用是否准确（请核对原文及文号）",
        "产品模块匹配是否与公司最新产品清单一致",
        "预算金额需由商务部门填写",
        "方案创新点建议资深顾问补充差异化设计",
        "客户承诺与量化指标需确认可达性",
        "竞对策略如需补充请人工添加",
    ]
    for item in s.get("review_notes", default_reviews):
        doc.add_paragraph(f"[ ] {item}", style="List Bullet")

    # ───── 附录B：生成说明 ─────
    _heading(doc, "附录B：生成说明", 1)
    info_t = doc.add_table(rows=7, cols=2)
    info_t.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_data = [
        ("客户类型", customer_type),
        ("区域", region),
        ("参考历史方案", s.get("ref_proposals", "无（知识库未接入）")),
        ("匹配产品模块", s.get("ref_products", "无")),
        ("引用政策数量", f"{s.get('ref_policies_count', 0)} 条"),
        ("引用案例数量", f"{s.get('ref_cases_count', 0)} 个"),
        ("内容来源占比", s.get("source_ratio", "模型生成 100%（知识库未接入）")),
    ]
    for i, (k, v) in enumerate(info_data):
        info_t.rows[i].cells[0].text = k
        info_t.rows[i].cells[1].text = v
        _set_cell_shading(info_t.rows[i].cells[0], "E8EDF5")
        for p in info_t.rows[i].cells[0].paragraphs:
            for r in p.runs:
                r.bold = True

    disclaimer = doc.add_paragraph()
    disclaimer.space_before = Pt(12)
    run = disclaimer.add_run(
        "本方案初稿由「卓智」数字员工辅助生成，所有内容需经人工审核确认后方可使用。"
    )
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

    # ───── 保存 ─────
    if not output_path:
        today = date.today().strftime("%Y%m%d")
        safe = lambda t: t.replace("/", "_").replace("\\", "_").replace(" ", "")
        output_path = f"ZX01_{safe(region)}_{safe(project_name[:10])}_解决方案_初稿_{today}.docx"

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    return os.path.abspath(output_path)


# ── CLI 入口 ──────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="ZX-01 方案初稿 Word 生成器")
    parser.add_argument("--json", required=True, help="输入 JSON 文件路径")
    parser.add_argument("--output", default=None, help="输出 .docx 路径（可选）")
    parser.add_argument("--schema", action="store_true", help="打印输入 JSON 结构说明")
    args = parser.parse_args()

    if args.schema:
        print(SCHEMA)
        return

    with open(args.json, "r", encoding="utf-8") as f:
        data = json.load(f)

    path = build_proposal_docx(data, args.output)
    print(path)


if __name__ == "__main__":
    main()
