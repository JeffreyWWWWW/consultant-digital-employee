# 数据资产目录

本目录用于沉淀 ZX-01 后续要接入的结构化知识资产。先保留目录骨架，避免产品库、案例库、政策素材继续散落在文档或临时输出里。

## 建议子目录

```text
data/
├── products/   # 卓繁产品与能力模块清单
├── cases/      # 成功案例库，含脱敏说明和可引用状态
├── policies/   # 政策法规、标准规范、权威资料
├── templates/  # 方案模板、公文模板、汇报材料模板
├── chroma/     # 本地向量库运行数据，不提交
└── tmp/        # 临时解析文件，不提交
```

## 推荐文件

- `products/product-capabilities.json`
- `cases/case-library.json`
- `policies/policy-sources.json`
- `templates/template-index.json`

结构化数据进入生成流程前，应先通过质检，确保产品、政策、案例均可追溯。
