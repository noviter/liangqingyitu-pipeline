---
name: liangqingyitu-orchestrator
description: 编排“两清一图”处室梳理全流程，适用于已有职责原文、系统资料和人工样例，需要生成待人工核验的事项台账、数字化流程和泳道图时。
---

# 两清一图总编排

先执行资料预处理，确保只读取每份职责 Excel 的第一个工作表：

```powershell
python tools/prepare_pipeline.py
```

以 `工作区/资料基线报告.md`、`职责来源清单.json`、`系统资料索引.json` 为唯一输入索引。原始资料不可改写；人工样例是对照基线，不是自动覆盖新结果的来源。

按以下顺序调用工作 skill：

1. `liangqingyitu-matter-decomposition`：产出事项台账和待核验事项。
2. `liangqingyitu-digital-process`：为每个事项形成数字化结论、证据和流程模型。
3. `liangqingyitu-swimlane`：只为证据充分、数字化结论为“有”的事项生成 Draw.io 页面。

所有阶段都必须使用稳定标识关联，不能按 Excel 行号关联：`source_id`、`responsibility_id`、`matter_id`、`evidence_id`、`flow_id`。人工合并或拆分事项时，保留被合并/拆分的来源 ID 和变更原因。

默认交付为“待人工核验初稿”。每条结论标注 `直接证据`、`业务推导`、`人工补充` 或 `待确认`；不得将资料缺失推定为现实中没有数字化。
