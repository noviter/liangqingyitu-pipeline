# 两清一图自动化流水线

面向交通运输数据治理场景的“两清一图”自动化工具集。项目把已经验证的人工梳理方法工程化为三个可审核阶段：职责拆解与事项清单、数字化支撑与业务流程、业务泳道图。

所有正式成果均为业务人员可直接打开和修改的 Excel、Draw.io 或图片文件。JSON 只用于程序内部处理，不能替代人工审核成果。

## 核心原则

- 职责原文、人工核验后的 Excel 是各阶段的权威输入。
- 每个阶段生成“人工核验前”成果后停止，等待业务人员确认。
- 不因资料缺失直接判断“无数字化”，不补造主体、权限、审批或退回路径。
- 人工样例只用于校验格式，不作为新处室的答案来源。
- 客户原始资料、运行缓存、密钥和人工样例不提交到 Git 仓库。

## 仓库结构

```text
skills/
  liangqingyitu-orchestrator/          # 全流程编排规则
  liangqingyitu-matter-decomposition/  # 阶段 1：事项拆解
  liangqingyitu-digital-process/       # 阶段 2：数字化与流程
  liangqingyitu-swimlane/              # 阶段 3：泳道图
tools/
  prepare_pipeline.py                  # 原始资料预处理与索引
  run_stage1_matter_decomposition.mjs  # 生成阶段 1 审核表
  retrieve_stage2_context.mjs          # 检索阶段 2 证据上下文
  run_stage2_prompt_analysis.mjs       # 形成阶段 2 分析结果
  run_stage2_digital_process.mjs       # 生成阶段 2 审核表
  run_stage3_swimlane.mjs              # 生成 Draw.io、预览图和核验表
  run_full_delivery_package.mjs        # 串联正式交付流程
03_规范文件/
  泳道图绘制规范.md
plan.md                                # 产品计划与业务口径
```

## 环境要求

- Node.js 22 或兼容版本
- Python 3.11 或兼容版本
- 可选：OpenAI API 密钥，用于阶段 2 模型分析

安装依赖：

```powershell
npm install
python -m pip install -r requirements.txt
```

如需调用模型，在本机环境变量中设置 `OPENAI_API_KEY`。不要把 `.env` 或密钥提交到仓库。

## 本地资料目录

以下目录由使用者在本地创建，已被 `.gitignore` 排除：

```text
01_职责原文/       # 职责 Excel
02_系统资料/       # 系统说明、制度、操作手册等
04_人工梳理样例/   # 仅供版式和结果对照
工作区/            # 预处理生成的索引与文本
outputs/           # 阶段输出
runs/              # 完整运行包
tmp/               # 临时文件
```

准备索引：

```powershell
python tools/prepare_pipeline.py
```

随后按 `skills/liangqingyitu-orchestrator/SKILL.md` 的阶段顺序执行。阶段 1 和阶段 2 的“核验后”文件必须由业务人员另存确认，自动流程不得用旧初稿覆盖。

## 阶段 2 模型配置

`tools/run_stage2_prompt_analysis.mjs` 默认读取：

- `OPENAI_API_KEY`
- `OPENAI_MODEL`（默认 `gpt-4.1-mini`）
- `OPENAI_BASE_URL`（默认 `https://api.openai.com/v1`）

也可以通过命令行参数传入模型名、API 地址或密钥环境变量名。密钥只从环境变量读取。

## 数据安全

本仓库仅保存自动化代码、业务规则和脱敏后的规范说明。交通系统设计文档、客户职责表、人工样例及预处理文本可能包含敏感业务信息或系统字段，必须留在受控的本地环境中。需要团队共享业务资料时，应使用组织批准的文档系统或私有存储，而不是提交到代码仓库。

## 当前状态

第一版已具备三阶段脚本、四项业务 Skill、阶段 2 证据检索与模型分析、Draw.io/Excel 交付能力。详细范围、字段口径和验收规则见 `plan.md`。
