# 咨询顾问数字员工

卓繁信息集团「咨询顾问-卓智」数字员工项目。当前聚焦 `ZX-01 智能方案初稿生成`：通过 OpenClaw + 企业微信 + Kimi，根据一句话需求、招标/采购材料、截图等输入，生成可人工复核的数字政府方案初稿，并输出 Word 文档。

## 当前状态

- 已完成 OpenClaw Agent 人设与工作规范配置。
- 已完成 `zhuofan-proposal-generator` Skill 的方案生成流程设计。
- 已完成 `generate_docx.py`，可将结构化方案 JSON 输出为 `.docx`。
- 已沉淀政策检索、来源追溯、方案结构、质检清单和测试用例。
- 待建设真实产品能力库、案例库、知识库摄入与自动质检器。

## 目录结构

```text
.
├── README.md
├── CLAUDE.md                         # 项目上下文，新的 AI 会话优先阅读
├── PLAN_ZX01.md                      # ZX-01 实施计划
├── assets/                           # 数字员工头像等视觉资产
├── data/                             # 待接入的结构化知识资产
│   ├── products/                     # 产品能力清单
│   ├── cases/                        # 成功案例库
│   ├── policies/                     # 政策法规素材
│   ├── templates/                    # 标准模板与方案样式
│   └── README.md
├── doc/                              # 需求文档、周报、架构说明和 TODO
├── openclaw-config/                  # OpenClaw Agent 配置与 Skills
│   ├── AGENTS.md
│   ├── IDENTITY.md
│   ├── SOUL.md
│   ├── USER.md
│   └── skills/
├── output/                           # 运行时生成文件，不提交内容
└── tests/                            # 测试用例与测试素材
```

## 核心 Skill

- `openclaw-config/skills/zhuofan-proposal-generator/`：ZX-01 智能方案初稿生成。
- `openclaw-config/skills/ppt-master/`：PPT Master OpenClaw Skill，以 submodule 方式跟踪 `JeffreyWWWWW/ppt-master` 的 `openclaw-skill` 分支。

## 更新 PPT Master Skill

```powershell
.\scripts\update_ppt_master_skill.ps1
```

脚本会更新 `openclaw-config\skills\ppt-master` submodule，并在主仓库记录新的 submodule 版本；如果没有变化，不会创建提交。

## PPT Master 维护约定

`D:\repository\ppt-master` 是 PPT Master 的唯一维护源，负责跟官方 `hugohe3/ppt-master` 同步、保留完整仓库结构，并维护本项目需要的定制内容。`openclaw-config\skills\ppt-master` 只是本项目消费的 submodule，不直接手改。

推荐流程：

```powershell
cd D:\repository\ppt-master
git switch openclaw-custom

# 修改 skills\ppt-master 下的文件后提交
git add skills\ppt-master
git commit -m "Modify ppt-master skill"
git push origin openclaw-custom

# 导出 OpenClaw 使用的 skill-only 分支
git branch -D openclaw-skill
git subtree split --prefix=skills/ppt-master -b openclaw-skill
git push -f origin openclaw-skill

# 回本项目更新 submodule 指针
cd D:\repository\consultant-digital-employee
.\scripts\update_ppt_master_skill.ps1
```

## Word 输出脚本

```powershell
uv run python openclaw-config\skills\zhuofan-proposal-generator\scripts\generate_docx.py --json proposal.json --output output\proposal.docx
```

当前本地环境需要确保已安装 `python-docx`。

## 下一步建议

1. 在 `data/products/` 建立产品能力清单。
2. 在 `data/cases/` 建立成功案例库。
3. 基于 `references/quality-checklist.md` 实现可执行质检器。
4. 接入 ChromaDB/RAG 文档摄入流程。
