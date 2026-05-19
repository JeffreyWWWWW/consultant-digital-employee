# scripts 目录说明

当前阶段只生成项目需求阶段的项目汇报 Word。

## 文件职责

- `generate_docx.py`：Word 输出入口。负责读取 JSON、校验 `document_type`、分发到对应生成器，并输出生成耗时。
- `docx_generators/project_report.py`：项目汇报 Word 生成器。负责标题、章节、版式和文件名规则。

## 设计说明

`generate_docx.py` 是统一入口，不直接维护具体文档版式。这样做是为了把“怎么调用”和“怎么生成某类文档”分开：

- 调用方只需要记住一个入口命令；
- 入口负责校验当前支持的 `document_type`；
- 具体版式、标题和章节规则放在对应生成器中；
- 后续扩展其他交付物时，可新增生成器，不影响入口命令。

当前已启用的生成器：

- `project_report` -> `docx_generators/project_report.py`

## 当前命令

```bash
python scripts/generate_docx.py --json proposal.json
python scripts/generate_docx.py --json proposal.json --output output.docx
python scripts/generate_docx.py --schema
```
