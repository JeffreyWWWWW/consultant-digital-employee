# 文档类型路由

本文件只负责判断 `document_type`。判断完成后，只读取对应结构文件和生成器，不要同时加载多个结构。

## project_report

路由为 `document_type: "project_report"`，读取 `structures/project-report.md`，生成器走 `scripts/docx_generators/project_report.py`。

适用情况：

- 用户明确要求项目汇报、项目建设汇报、项目建议支撑材料、立项依据材料；
- 输入材料主要是会议纪要、沟通纪要、任务安排、聊天记录；
- 用户只要求“整理成方案/材料/初稿/Word”，没有明确要求完整技术方案、实施方案、预算、项目组织或投标响应；
- 材料重点是建设依据、建设目标、建设内容、页面优化、功能完善、系统适配、数据对接、工作安排。

满足以上条件时，即使用户说“方案初稿”，也优先路由为 `project_report`。

## solution

路由为 `document_type: "solution"`，读取 `structures/solution-draft.md`。

适用情况：

- 用户明确要求完整解决方案、建设方案、技术方案、售前方案、投前方案；
- 材料包含招标需求、评分项、技术要求；
- 用户明确要求技术架构、实施建议、预算框架、项目组织或运维方案。

## achievement_report

路由为 `document_type: "achievement_report"`，读取 `structures/achievement-report.md`。

适用情况：

- 用户要求建设情况汇报、工作汇报、成效汇报、阶段总结、建设成效；
- 材料重点是已建成内容、运行成效、指标数据、下一步工作计划。

## 扩展规则

- 实施计划、项目组织、预算框架、运维运营等内容只作为可选扩展，用户明确要求时才生成。
- 无法判断交付物类型时，优先追问最终用途；追问后仍不明确时，会议纪要类输入默认 `project_report`。
