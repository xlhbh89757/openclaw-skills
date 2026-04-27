---
name: resume-risk-assessor
description: Use when assessing authenticity risk across one or more resumes from local files or provided text, especially for batch screening, shortlist triage, or generating a structured resume risk report before manual review.
---

# Resume Risk Assessor

## Overview

批量评估简历的真实性风险，输出结构化 Excel 报告，适合在初筛阶段快速标记高风险候选人。

## When to Use

- 用户要批量分析简历风险、筛选高风险简历、生成简历风险报告。
- 用户需要评估简历真假
- 输入已经在本地文件中，或可以直接提供为纯文本 / JSON。
- 目标是做初筛和人工复核前的风险提示，不是直接做录用决策。

## When Not to Use

- 需要联网核验学历、公司真实性、证书或背调信息。
- 需要对单份图片简历做 OCR。当前脚本说明提到 OCR，但现有主脚本并未实现图片输入流程。
- 需要法律、合规或最终雇佣结论。

## Inputs

- `--folder` 支持批量处理：`.pdf`、`.docx`、`.doc`、`.xlsx`、`.xls`，默认递归扫描子目录
- `--no-recursive` 可限制 `--folder` 只扫描当前目录
- `--file` 当前只明确处理：`.pdf`、`.docx`、`.doc`
- 也支持从标准输入传入 JSON 数组，每项包含 `text`，可选 `name`

## Run

脚本路径：

```text
C:\Users\EDY\.openclaw\workspace\skills\resume-risk-assessor\scripts\analyze_resume_v3.py
```

常用命令：

```bash
python C:\Users\EDY\.openclaw\workspace\skills\resume-risk-assessor\scripts\analyze_resume_v3.py --folder "C:\Resumes"
python C:\Users\EDY\.openclaw\workspace\skills\resume-risk-assessor\scripts\analyze_resume_v3.py --folder "C:\Resumes" --no-recursive
python C:\Users\EDY\.openclaw\workspace\skills\resume-risk-assessor\scripts\analyze_resume_v3.py --folder "C:\Resumes" --debug-dir "C:\Resumes\extraction_debug"
python C:\Users\EDY\.openclaw\workspace\skills\resume-risk-assessor\scripts\analyze_resume_v3.py --file "C:\Resumes\candidate.pdf"
echo '[{"name":"张三","text":"简历正文"}]' | python C:\Users\EDY\.openclaw\workspace\skills\resume-risk-assessor\scripts\analyze_resume_v3.py
```

## Output

- 默认输出文件名：`resume_risk_report.xlsx`
- 输出位置：评估文件夹下；单文件评估时输出到该文件所在目录。`--output` 如果传相对路径，也会相对评估目录解析。
- 标准输出会返回 JSON，包含 `success`、`output`、`count`、风险分布和候选人摘要
- Excel「风险总览」包含文件名、文本长度和解析质量，帮助区分真实风险和 PDF 抽取质量问题
- Excel「详细分析」包含「触发原文」，用于快速复核每个风险点的原文依据
- `--debug-dir` / `--save-extracted-text` 会为每份简历保存抽取原文 `.txt` 和章节分区 `.sections.json`，相对路径同样写到评估目录下，用于定位“解析复核”样本到底是抽取问题、分区问题还是规则问题。

## Review Guidance

- 优先关注 `overall` 为 `高` 或 `中` 的候选人
- 重点复核时间线重叠、工作经历过于模糊和表述模糊
- `analyze_resume_v3.py` 会先做章节标题标准化，兼容 `工作经历：`、`工作经验`、`项目经验`、`技术专长` 等常见标题写法，降低因标题格式导致的分区失败。
- 项目 / 系统 / 平台类时间段会从工作时间线重叠判断中剔除，避免把项目周期误判为候选人的多段雇佣经历。
- **技能夸大维度**：v4 脚本已改为上下文感知判定（结合工作年限、项目证据），大幅降低误报。但仍建议对标记为「技能夸大」的候选人做以下 AI 复核：
  1. 提取被标记候选人的简历原文（通过 `extract_section_lines` 或直接读 PDF）
  2. 结合经验年限、项目详实程度、用词分寸综合判断：
     - **无夸大**：经验足以支撑、项目证据充分、用词克制（如「熟悉」「了解」）
     - **轻微夸大**：个别「精通」过满但整体可信，面试追问即可
     - **明显夸大**：年限短却大量「精通」，技术栈覆盖面过广不可信
  3. 根据复核结果调整风险等级，降级误判的候选人
- 把结果当作人工复核线索，不要把脚本输出当成事实证明
- `analyze_resume_v3.py` 会优先按教育 / 工作 / 技能 / 项目分区判断，能降低项目经历对时间线和模糊表述的误报。
- 风险结果只适合初筛和复核排序，不替代背调、学历核验或最终录用决策。

## Failure Modes

- 如果缺少依赖，常见报错会来自 `pdfplumber`、`python-docx`、`openpyxl`
- 如果 `--file` 传入 Excel，当前主脚本不会像 `--folder` 那样处理它
- 复杂 PDF 可能抽取文本质量较差，报告会受原始文本质量影响
- 如果 PDF 带有水印、扫描页或缺少明显分区标题，文本抽取质量会直接影响风险判断，必要时应回看原件。

## Changelog

- **v7.1 (2026-04-27)**：统一报告和诊断目录的相对路径解析，默认写入被评估简历所在目录，避免输出散落到命令执行目录。
- **v7 (2026-04-27)**：新增抽取诊断输出，支持 `--debug-dir` 和 `--save-extracted-text`，批量保存每份简历的抽取文本和结构化章节结果，便于复核 PDF 抽取失败导致的中风险误报。
- **v6 (2026-04-24)**：优化章节标题识别，支持冒号、空格和常见别名；工作时间线判断会过滤项目 / 系统 / 平台类时间段，减少项目经历被误判为工作经历重叠的问题。在 `E:\简历` 130 份样本上复测，中风险从 54 降至 48。
- **v5 (2026-04-23)**：报告新增触发原文证据列、文本长度和解析质量字段；`--folder` 默认递归扫描并支持 `--no-recursive`；风险对象保留 evidence，便于复核和后续系统集成。
- **v4 (2026-04-23)**：技能夸大判定从简单关键词计数改为上下文感知模式——结合工作年限减免、项目证据支撑度，避免对经验丰富的候选人误判。新增 AI 复核步骤建议。
