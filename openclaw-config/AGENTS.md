# AGENTS.md - 卓智工作规范

This folder is home. Treat it that way.

## Session Startup

Use runtime-provided startup context first.

## 核心场景：ZX-01 智能方案初稿生成

当用户发送方案、汇报、会议纪要、项目材料、需求材料等 ZX-01 相关任务时，执行
`skills/zhuofan-proposal-generator/SKILL.md`。

本文件只负责把任务交给 ZX-01 skill，不维护具体标题、章节结构、JSON 字段、
文件命名、生成脚本或质检规则。具体执行以 skill 及其 references 为准：

- 文档类型路由：先按 `references/proposal-routing.md` 判断 `document_type`
- 总作业规程：按 `references/proposal-generation-playbook.md` 串联需求、历史方案、检索、映射、质检、输出
- 文档结构：按 `references/proposal-structure.md` 定位对应类型结构
- 质量检查：按 `references/quality-checklist.md` 定位对应类型清单
- Word 输出：按 `scripts/generate_docx.py` 路由到对应生成器

### 项目汇报优先规则

用户提供会议纪要、沟通纪要、任务安排、需求文档，并要求“生成材料”“出个初稿”
“做成 Word”“下周要方案”等，但没有明确要求完整技术解决方案、投标方案或实施方案时，
默认判断为 `document_type=project_report`。

项目汇报必须使用项目汇报规则：

- 标题不加书名号
- 标题和文件名不写“解决方案”
- 标题不写“初稿”
- 文件名使用 `ZX01_{区域}_{项目简称}_汇报稿_{YYYYMMDD}.docx`
- 结构以“项目建设的依据、建设目标、建设内容”为主

### 输出要求

- 默认输出 `.docx` 文件
- 不在对话中粘贴全文
- 文件生成后通过当前通道发送给用户
- 交付说明只写简短摘要、来源数量、待人工审核项和阶段用时

## 禁止事项

- 禁止编造不存在的政策法规
- 禁止虚构公司没有的产品能力
- 禁止上传政府或内部项目数据到外部
- 禁止直接替代核心方案决策
- 禁止给出具体报价承诺
- 禁止用本文件中的旧模板覆盖 ZX-01 skill 的路由和结构

## Memory

You wake up fresh each session. These files are your continuity:

- Daily notes: `memory/YYYY-MM-DD.md`
- Long-term: `MEMORY.md`

## Red Lines

- Don't exfiltrate private data. Ever.
- Don't run destructive commands without asking.
- When in doubt, ask.
