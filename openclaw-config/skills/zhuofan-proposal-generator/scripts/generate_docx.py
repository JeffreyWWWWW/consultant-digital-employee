"""
ZX-01 建设方案 Word 生成入口。

当前阶段只支持 document_type=construction_plan。

用法：
  python scripts/generate_docx.py --json proposal.json
  python scripts/generate_docx.py --json proposal.json --output output.docx
  python scripts/generate_docx.py --schema
"""

import argparse
import json
import sys
import time


SCHEMA = """
输入 JSON 结构：
{
  "document_type": "construction_plan",
  "project_name": "项目名称",
  "customer_name": "客户单位名称",
  "region": "区域",
  "sections": {
    "construction_background": "建设背景，可多段，用 \\n 分隔",
    "project_goals": "建设目标，可多段，用 \\n 分隔",
    "demand_analysis": "需求分析，可多段，用 \\n 分隔",
    "project_contents": [
      {
        "name": "建设事项名称",
        "content": "建设动作、服务对象、业务价值"
      }
    ],
    "construction_effects": "建设成效，可多段，用 \\n 分隔",
    "investment_estimate": "投资估算，可多段，用 \\n 分隔",
    "sources": [
      {
        "name": "来源名称",
        "agency": "发布单位或来源机构",
        "doc_no": "文号（如有）",
        "date": "发布时间（如有）",
        "url": "原文链接；无法核验时留空并填写 status",
        "source_type": "官网原文/政府公开平台/机构原文/主体公开页面/政府转载/媒体转载/线索页/未确认",
        "status": "已核验原文/待核验原文/待联网核验/未找到原文",
        "verification_status": "核验状态，同 status",
        "used_for": "支撑的正文内容",
        "query": "按 references/search/source-web-search.md 实际执行的检索词",
        "search_provider": "Tavily/百度/Tavily+百度/用户提供",
        "source_names": ["source-web-search.md 的来源名称数组"],
        "source_urls": ["source-web-search.md 的原文链接数组"]
      }
    ],
    "review_notes": [
      {
        "content": "需人工审核项",
        "target": "正文中对应的具体片段，必填，用于把批注挂到正文"
      }
    ],
    "review_highlights": ["正文中需要添加批注并人工核对的具体文字片段"],
    "source_note": "来源说明"
  }
}
"""


SUPPORTED_TYPES = {"construction_plan"}


def _load_json(path: str) -> dict:
    with open(path, "r", encoding="utf-8-sig") as f:
        return json.load(f)


def _validate_construction_plan(data: dict) -> None:
    doc_type = (data.get("document_type") or "").strip()
    if not doc_type:
        raise ValueError("缺少必填字段：document_type")
    if doc_type not in SUPPORTED_TYPES:
        raise ValueError(
            f"当前阶段只支持 document_type=construction_plan，不支持 {doc_type!r}。"
        )

    missing = [key for key in ("project_name", "customer_name", "region") if not data.get(key)]
    if missing:
        raise ValueError("缺少必填字段：" + ", ".join(missing))

    data["document_type"] = "construction_plan"
    data.setdefault("sections", {})


def generate_docx(json_path: str, output_path: str = None) -> str:
    data = _load_json(json_path)
    _validate_construction_plan(data)

    from docx_generators.construction_plan import build_docx

    return build_docx(data, output_path)


def main() -> None:
    start_time = time.perf_counter()
    parser = argparse.ArgumentParser(description="ZX-01 建设方案 Word 生成器")
    parser.add_argument("--json", default=None, help="输入 JSON 文件路径")
    parser.add_argument("--output", default=None, help="输出 .docx 路径（可选）")
    parser.add_argument("--schema", action="store_true", help="打印输入 JSON 结构说明")
    args = parser.parse_args()

    if args.schema:
        print(SCHEMA.strip())
        return

    if not args.json:
        parser.error("the following arguments are required: --json")

    try:
        path = generate_docx(args.json, args.output)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    elapsed = time.perf_counter() - start_time
    print(path)


if __name__ == "__main__":
    main()
