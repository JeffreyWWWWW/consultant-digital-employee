# 质量检查路由

本文件只负责根据 `document_type` 选择质量检查清单。判断完成后，只读取对应 checklist。

当前阶段已实现：

- `project_report`：读取 `../quality/project-report-checklist.md`

暂未实现：

- `solution`
- `bid_support`
- `achievement_report`
- `other_material`

如果 `document_type` 不是 `project_report`，当前阶段不生成 Word，先追问或说明暂不展开该交付物。

输出前必须形成 `sections.review_notes`，并根据检查结果修正 JSON 后再生成 Word。
