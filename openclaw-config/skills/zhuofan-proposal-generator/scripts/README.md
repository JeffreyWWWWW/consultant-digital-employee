# scripts 目录说明

当前阶段只生成项目需求阶段的建设方案 Word。

## 文件职责

- `generate_docx.py`：Word 输出入口。负责读取 JSON、校验 `document_type`，并分发到对应生成器。
- `docx_generators/construction_plan.py`：建设方案 Word 生成器。负责标题、章节、版式和文件名规则。

## 设计说明

`generate_docx.py` 是统一入口，不直接维护具体文档版式。这样做是为了把“怎么调用”和“怎么生成某类文档”分开：

- 调用方只需要记住一个入口命令；
- 入口负责校验当前支持的 `document_type`；
- 具体版式、标题和章节规则放在对应生成器中；
- 后续扩展其他交付物时，可新增生成器，不影响入口命令。

当前已启用的生成器：

- `construction_plan` -> `docx_generators/construction_plan.py`

## 当前命令

```bash
python scripts/generate_docx.py --json proposal.json
python scripts/generate_docx.py --json proposal.json --output output/materials
python scripts/generate_docx.py --schema
```

不传 `--output` 时，最终 Word 成品默认输出到仓库根目录 `output/materials/`。演示文稿成品统一放到 `output/presentations/`。
临时 JSON、过程日志和其他中间产物统一放到仓库根目录 `tmp/`。

## 字体说明

建设方案 Word 采用固定字体规范：标题 `方正小标宋_GBK` 二号，正文 `仿宋_GB2312` 三号。
