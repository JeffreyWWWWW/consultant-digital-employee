# CLAUDE.md - 项目上下文（每次新对话先读这个文件）

## 项目概述

卓繁信息集团「咨询顾问」岗位数字员工建设项目。通过 AI 辅助咨询顾问高效输出专业方案。

## 技术栈

| 组件 | 技术 | 状态 |
|------|------|------|
| AI 网关 | [OpenClaw](https://github.com/openclaw/openclaw) v2026.5.7 | ✅ 已部署 |
| 消息通道 | 企业微信（Bot Mode） | ✅ 已连接 |
| 主 LLM | Moonshot Kimi K2.6（262K 上下文） | ✅ 已配置 |
| 向量数据库 | ChromaDB（计划） | ❌ 待搭建 |
| 代码仓库 | https://github.com/JeffreyWWWWW/consultant-digital-employee | ✅ |

## 关键配置路径

- **OpenClaw 主配置：** `~/.openclaw/openclaw.json`
- **Agent 人设文件：** `~/.openclaw/workspace/`（SOUL.md, AGENTS.md, IDENTITY.md, USER.md）
- **Agent 模型配置：** `~/.openclaw/agents/main/agent/models.json`
- **项目实施计划：** 仓库根目录 `PLAN_ZX01.md`

## 当前聚焦：ZX-01 智能方案初稿生成（P1）

> 用户通过企微发送需求（关键词/一句话/招标文件/图片）→ AI 自动生成结构化方案初稿

### 五个场景总览

| 编号 | 场景 | 优先级 | 状态 |
|------|------|--------|------|
| **ZX-01** | **智能方案初稿生成** | **P1** | **🔨 开发中** |
| ZX-02 | 投标文件辅助生成 | P1 | 待启动 |
| ZX-03 | 政策与案例智能检索 | P2 | 待启动 |
| ZX-04 | 资料智能检索与知识库助手 | P2 | 待启动 |
| ZX-05 | PPT智能生成与优化 | P2 | 待启动 |

### ZX-01 进度

- [x] 需求文档分析（doc/ 目录下三个文件）
- [x] 实施计划编写（PLAN_ZX01.md）
- [x] OpenClaw Agent 人设配置（SOUL.md / AGENTS.md / IDENTITY.md / USER.md）
- [ ] 知识库搭建（等待业务方提供测试素材 4-8 项）
- [ ] RAG 检索引擎
- [ ] Prompt 工程优化
- [ ] 方案生成 Pipeline
- [ ] Word 模板输出
- [ ] OpenClaw Skill 集成
- [ ] 去 AI 化后处理
- [ ] 测试评估

### OpenClaw Agent 人设（已配置）

- **名字：** 卓智（卓繁 + 智慧）
- **定位：** 15年经验的数字政府首席咨询师
- **风格：** 方案写作用政府公文风格；日常对话像靠谱的资深同事
- **核心原则：** 准确第一、数据安全、辅助而非替代、来源可追溯
- **文件位置：** `~/.openclaw/workspace/` 下的 SOUL.md、AGENTS.md、IDENTITY.md、USER.md

## 需求文档速查

| 文件 | 内容 |
|------|------|
| `doc/咨询顾问岗位需求说明书_数字员工.docx` | 完整需求说明书：5个场景详细定义、人机分工边界、验收指标、实施路线 |
| `doc/数字政府汇报材料指令.docx` | 资深顾问的 Prompt 模板集：通用优化、专项场景（汇报/规划/方案）、去AI化、投标文件优化 |
| `doc/ZX-01_测试素材清单.xlsx` | ZX-01 测试方案：9类素材定义（1-3触发输入 + 4-8知识库 + 9人工成品对照） |

## 下一步待做

1. **等素材：** 业务方提供 ZX-01 测试素材（历史方案、产品清单、政策文件、成功案例、标准模板）
2. **搭知识库：** 文档摄入 pipeline → ChromaDB
3. **写 Prompt：** 基于指令模板文档，设计结构化生成 Prompt
4. **跑通 MVP：** 输入需求 → RAG 检索 → LLM 生成 → Word 输出

## 注意事项

- 政府项目数据禁止上传公开大模型
- OpenClaw 人设文件在 `~/.openclaw/workspace/` 而不是仓库里
- 仓库 `.gitignore` 已忽略 `.claude/` 和 `.openclaw/`
