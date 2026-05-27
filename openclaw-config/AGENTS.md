# AGENTS.md - 卓智工作规范

This folder is home. Treat it that way.

## Session Startup

Use runtime-provided startup context first.

## 核心场景：智能方案初稿生成

当用户发送方案、汇报、会议纪要、项目材料、需求材料等智能方案初稿生成相关任务时，执行
`skills/zhuofan-proposal-generator/SKILL.md`。

本文件只负责仓库级入口和工作约定，不维护具体标题、章节结构、JSON 字段、
文件命名、生成脚本、交付物路由或质检规则。具体执行以 skill 及其 references 为准。

权威入口：

- Skill 主入口：`skills/zhuofan-proposal-generator/SKILL.md`
- 文档类型路由：先按 `references/routing/proposal-routing.md` 判断 `document_type`
- 总作业规程：按 `references/workflow/proposal-generation-playbook.md` 串联需求、历史方案、检索、映射、质检、输出
- 文档结构：当前项目汇报阶段按 `references/structures/project-report.md`
- 质量检查：按 `references/routing/quality-routing.md` 定位对应质检清单
- 输出方式：按 skill 当前说明和对应脚本执行

## 核心场景：PPT Master 演示文稿生成

当用户要求创建、制作、生成、修改或导出演示文稿 / presentation / slides / deck，
或明确提到 `ppt-master` 时，优先执行
`skills/ppt-master/SKILL.md`。

这是本仓库的演示文稿权威入口。若它与通用演示文稿、Presentation 或文档生成类
skill 的触发范围重叠，除非用户明确指定其他工具，否则按 PPT Master skill 执行。

AGENTS.md 只负责把任务路由到 PPT Master，不复制其流水线、模板选择、SVG 生成、
质检或导出规则。具体执行以 `skills/ppt-master/SKILL.md` 及其 references /
workflows 为准。

AGENTS.md 不得覆盖、复制或简化 skill/reference 中的具体规则。发现冲突时，以
对应 `SKILL.md` 及其 references 为准，并清理本文件中的重复规则。

## 通用工作规范

- 优先按当前运行环境和用户最新指令执行。
- 所有面向用户交付的正式文件统一放到仓库根目录 `output/`，并按类型分开：文字材料放 `output/materials/`，演示文稿放 `output/presentations/`。
- Word、PDF、Excel、压缩包等文字材料成品完成后，直接交付文件给用户；最后回复必须给出 `output/materials/` 下的文件名或路径。
- 演示文稿、slides、deck 等成品完成后，直接交付文件给用户；最后回复必须给出 `output/presentations/` 下的文件名或路径。
- 发送 `MEDIA:` 文件前必须确认该绝对路径真实存在；如果演示文稿成品生成在 `output/` 根目录，先移动或复制到 `output/presentations/`，再使用实际存在的 `output/presentations/` 绝对路径发送，禁止凭记忆拼接路径。
- 临时文件和中间产物统一放到仓库根目录 `tmp/`，包括临时 JSON、渲染预览、源码 SVG、过程日志等；不要散落在各流程目录。
- `output/` 只放最终交付文件，不放临时文件和过程文件。
- 不在 AGENTS.md 中新增具体业务路由、标题、文件名、章节或质检细则。
- 需要修改智能方案初稿生成具体行为时，修改对应 skill 或 reference 文件。
- 需要修改 PPT Master 具体行为时，不得直接修改本仓库
  `openclaw-config/skills/ppt-master` submodule。必须先到
  `D:\repository\ppt-master` 修改并提交，再回本仓库执行
  `.\scripts\update_ppt_master_skill.ps1` 更新 submodule 指针。
- 不要在 AGENTS.md 里维护 PPT Master 的重复流水线规则；具体执行以
  `skills/ppt-master/SKILL.md` 及其 references / workflows 为准。

## 禁止事项

- 禁止编造不存在的政策法规
- 禁止虚构公司没有的产品能力
- 禁止上传政府或内部项目数据到外部
- 禁止直接替代核心方案决策
- 禁止给出具体报价承诺
- 禁止用本文件覆盖智能方案初稿生成 skill 的路由和结构
- 禁止用本文件覆盖 PPT Master skill 的流水线、模板、SVG 生成、质检和导出规则

## Memory

You wake up fresh each session. These files are your continuity:

- Daily notes: `memory/YYYY-MM-DD.md`
- Long-term: `MEMORY.md`

## Red Lines

- Don't exfiltrate private data. Ever.
- Don't run destructive commands without asking.
- When in doubt, ask.
