---
name: proposal-generator
description: 智能方案初稿生成技能。当用户提出方案需求（如"帮我写一个XX方案"、"做一个XX解决方案"、"XX项目方案初稿"、"写个数字政府方案"、"帮我出个方案"等），或提供招标文件/需求文档要求生成方案时，触发此技能。
metadata:
  {
    "openclaw": { "version": "1.0.0" },
  }
---

# 智能方案初稿生成

> 核心能力：将用户的需求（关键词/一句话/招标文件）转化为结构化的解决方案初稿，并输出为 .docx 文件。

> 禁止在任何输出中使用 emoji 符号。

---

## 触发条件

以下任一情况触发本技能：
- 用户明确要求"写方案"、"出方案"、"生成方案"、"方案初稿"
- 用户提供项目名称 + 客户信息，要求生成解决方案
- 用户发送招标文件/需求文档，要求据此撰写方案
- 用户提到"解决方案"、"建设方案"、"实施方案"等关键词并期望输出文档

---

## 第一步：需求解析与确认

收到需求后，**必须先确认以下关键信息**，不完整则追问：

| 信息项 | 说明 | 是否必填 |
|--------|------|----------|
| 项目名称 | 如"城市大脑建设"、"数据共享平台" | 必填 |
| 客户名称 | 如"杭州市数据局" | 必填 |
| 客户类型 | 数据局/政务服务中心/大数据中心/数据集团/企业 | 必填 |
| 区域 | 如"杭州市"、"浙江省" | 必填 |
| 核心需求 | 用户想解决什么问题、建设什么系统 | 必填 |
| 预算范围 | 如有则记录 | 选填 |
| 特殊要求 | 信创要求、时间节点、技术偏好等 | 选填 |

**追问示例**：
- "我需要确认几个关键信息来生成方案：目标客户是哪个单位？属于什么类型（数据局/政务中心/大数据中心）？项目所在区域？"
- "项目名称和核心需求我已了解，请确认客户单位全称和所在区域。"

---

## 第二步：方案内容生成

确认信息后，按以下结构生成方案内容。每一章节都需要有实质内容，不允许留空或只写标题。

### 方案结构（七章）

**一、现状分析**
- 1.1 政策背景：引用相关国家/省/市政策文件（必须有真实政策依据，不确定的标注「待确认」）
- 1.2 现状概述：描述客户当前信息化建设状况
- 1.3 痛点问题：列出 3-5 个核心痛点，用"一是、二是、三是"句式

**二、建设目标**
- 2.1 总体目标：一段话概括，包含可量化的指标
- 2.2 分项目标：3-5 个具体子目标

**三、建设内容**
- 按功能模块拆分，每个模块包含：
  - 模块名称
  - 建设内容详述（不少于100字）
  - 产品匹配：从卓繁信息产品清单中匹配，不确定的标注「待匹配」

**四、技术架构**
- 4.1 总体架构：五层架构（基础设施层、数据资源层、平台支撑层、业务应用层、用户交互层）+ 两大支撑体系
- 4.2 技术选型：明确技术路线（是否信创），核心组件选型

**五、实施计划**
- 分期实施，每期包含：阶段名称、周期、核心交付物

**六、预算框架**
- 按建设内容分项列出预算条目
- 金额标注"待定"，由人工确认

**七、亮点与预期效益**
- 7.1 方案亮点：3-5 个差异化亮点
- 7.2 预期效益：可量化的效益指标
- 7.3 成功案例参考

---

## 第三步：生成 .docx 文件

方案内容确认后，调用 Python 脚本生成 Word 文档：

```bash
python "D:/repository/consultant-digital-employee/src/output/docx_builder.py"
```

但实际使用时，需要通过 shell tool 执行以下 Python 代码来调用 `build_proposal()` 函数：

```python
import sys
sys.path.insert(0, "D:/repository/consultant-digital-employee/src")
from output.docx_builder import build_proposal

output_path = build_proposal(
    project_name="项目名称",
    customer_name="客户名称",
    customer_type="客户类型",
    region="区域",
    sections={
        "policy_background": "政策背景文本",
        "current_status": "现状概述文本",
        "pain_points": ["痛点1", "痛点2", "痛点3"],
        "overall_goal": "总体目标文本",
        "sub_goals": ["子目标1", "子目标2"],
        "modules": [
            {"name": "模块名", "content": "详细内容", "product": "匹配产品"}
        ],
        "architecture": "",
        "tech_selection": "技术选型说明",
        "phases": [
            {"name": "一期", "duration": "6个月", "deliverables": "交付物"}
        ],
        "budget_items": [
            {"name": "项目名", "amount": "待定", "note": ""}
        ],
        "highlights": ["亮点1", "亮点2"],
        "benefits": ["效益1", "效益2"],
        "cases": [{"name": "案例名", "source": "来源", "summary": "摘要"}],
        "ref_proposals": "参考的历史方案",
        "ref_products": "匹配的产品模块",
        "ref_policies_count": 0,
        "ref_cases_count": 0,
    },
    output_path="D:/repository/consultant-digital-employee/output/生成的文件名.docx"
)
print(f"文件已生成：{output_path}")
```

### 文件命名规范

`ZX01_{区域}_{项目简称}_解决方案_初稿_{日期}.docx`

示例：`ZX01_杭州市_城市大脑_解决方案_初稿_20260512.docx`

### 输出目录

所有生成的文件存放在：`D:/repository/consultant-digital-employee/output/`

---

## 第四步：交付与提示

文件生成后：

1. **告知用户文件路径**：明确告诉用户 .docx 文件的完整路径
2. **列出审核要点**：
   - 政策引用是否准确（核对原文）
   - 产品模块匹配是否与公司最新产品清单一致
   - 预算估算是否合理
   - 方案创新点是否有差异化
   - 客户特殊需求是否完整覆盖
3. **标注来源**：说明哪些内容来自知识库检索，哪些是模型生成
4. **提醒人工确认**：方案为AI初稿，需人工审核修改后方可使用

---

## 客户类型差异化策略

| 客户类型 | 侧重点 | 语言风格 |
|----------|--------|----------|
| 数据局 | 数据治理、数据共享、数据安全 | 高站位，引用国家数据战略 |
| 政务服务中心 | 一网通办、便民服务、减证便民 | 突出群众获得感、服务效率 |
| 大数据中心 | 算力、数据中台、技术架构 | 技术深度，架构专业性 |
| 数据集团 | 数据资产、数据运营、商业模式 | 市场化视角，投入产出比 |
| 企业 | 降本增效、业务痛点、ROI | 务实直接，商业价值导向 |

---

## 禁止事项

- 禁止编造不存在的政策文件
- 禁止虚构卓繁不存在的产品
- 禁止使用"综上所述"、"全方位保障"、"一站式解决"等AI套话
- 禁止在任何输出中使用 emoji 符号
- 禁止给出确定的预算金额（统一标注"待定"）
- 禁止跳过需求确认直接生成方案

---

## sections 字段说明

调用 `build_proposal()` 时，`sections` 参数的完整字段定义：

| 字段 | 类型 | 说明 |
|------|------|------|
| policy_background | string | 政策背景文本 |
| current_status | string | 现状概述文本 |
| pain_points | string[] | 痛点问题列表（3-5条） |
| overall_goal | string | 总体目标文本 |
| sub_goals | string[] | 分项目标列表 |
| modules | object[] | 建设内容模块，每项含 name/content/product |
| architecture | string | 技术架构描述（留空则使用默认五层架构） |
| tech_selection | string | 技术选型说明 |
| phases | object[] | 实施阶段，每项含 name/duration/deliverables |
| budget_items | object[] | 预算条目，每项含 name/amount/note |
| highlights | string[] | 方案亮点（3-5条） |
| benefits | string[] | 预期效益（可量化） |
| cases | object[] | 成功案例，每项含 name/source/summary |
| ref_proposals | string | 参考的历史方案 |
| ref_products | string | 匹配的产品模块 |
| ref_policies_count | int | 引用政策数量 |
| ref_cases_count | int | 引用案例数量 |
