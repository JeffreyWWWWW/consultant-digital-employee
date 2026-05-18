# 质量检查路由

本文件只负责根据 `document_type` 选择质量检查清单。判断完成后，只读取对应 checklist。

- `project_report`：读取 `quality/project-report-checklist.md`
- `solution`：读取 `quality/solution-draft-checklist.md`
- `achievement_report`：读取 `quality/achievement-report-checklist.md`

输出前必须形成 `sections.review_notes`，并根据检查结果修正 JSON 后再生成 Word。
