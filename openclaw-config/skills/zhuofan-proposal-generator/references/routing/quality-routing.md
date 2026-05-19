# 质量检查路由

本文件只负责根据 `document_type` 选择质量检查清单。判断完成后，只读取对应 checklist。

当前可用清单：

- `project_report`：读取 `../quality/project-report-checklist.md`

如果没有对应的质量检查清单，停止当前输出流程，先追问或说明当前交付物暂不支持。

质量检查结果必须写入 `sections.review_notes`，并根据检查结果修正 JSON。
