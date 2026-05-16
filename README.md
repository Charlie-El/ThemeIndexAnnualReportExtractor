# ThemeIndexAnnualReportExtractor 使用说明

[English Guide](README_EN.md)

ThemeIndexAnnualReportExtractor 是一个面向美股年报的主题指数数据提取工具。它会从公司年报 HTML 中整理主营业务、经营分部、收入构成、收入确认和主题相关证据，再可选调用兼容 OpenAI SDK 的大模型接口，将证据转成结构化 Excel / JSON 结果。

本工具适合已经准备好年报 HTML 的用户。你可以先让程序提取可追溯的原文证据，再用大模型把证据整理成标准字段，最后按公司逐个查看结果和原文依据。

## 主要功能

- 批量读取公司文件夹中的 `.html` / `.htm` 年报。
- 支持 10-K、20-F、40-F 等常见美股年报 HTML。
- 自动识别 Business、MD&A、Financial Statements、Notes、Revenue、Segment Information 等关键章节。
- 提取主营业务、产品/服务、经营分部、收入拆分、地区/客户/合同类型收入、收入确认等候选证据。
- 识别 AI、半导体、机器人、商业航天、地缘风险、能源、矿产资源、生物医疗等主题线索。
- 为每家公司生成规则提取结果，便于人工复核。
- 为每家公司生成 `大模型证据包.json`，把关键上下文、相关表格和质量提示整理成可喂给模型的结构化输入。
- 可选调用豆包、方舟或其他 OpenAI 兼容接口，输出标准化结果。
- 每家公司单独保存 JSON / Excel / 原始响应，方便逐家公司追溯。
- 支持小批量测试、指定公司运行、跳过已有结果和失败明细记录。

## 适用场景

| 场景 | 说明 |
| --- | --- |
| 主题指数研究 | 从年报中整理公司业务、收入构成和主题相关线索。 |
| 行业/主题标签维护 | 结合年报原文判断公司是否与 AI、半导体、能源、矿产等主题相关。 |
| 大模型辅助抽取 | 先用规则程序准备可追溯证据，再交给大模型结构化输出。 |
| 批量年报复核 | 每家公司保留证据包和模型原始响应，便于人工抽查。 |

本工具不是投资建议工具，也不会自动判断证券买卖价值。抽取结果应作为研究辅助材料，关键字段建议人工复核。

## 项目结构

```text
ThemeIndexAnnualReportExtractor/
  scripts/
    annual_report_blocks.py        # 年报章节、上下文和表格定位工具
    theme_index_extractor.py       # 规则提取主流程
    build_llm_section_inputs.py    # 生成每家公司大模型证据包 JSON
    run_theme_index_llm_batch.py   # 调用 OpenAI 兼容模型并输出结果
  templates/
    keyword_overrides.json         # 可审核关键词扩展模板
    keyword_feedback_template.json # 模型反馈关键词建议模板
    README.md                      # 模板说明
  data/
    downloads/                     # 年报 HTML 放在这里，每家公司一个文件夹
  outputs/                         # 总表、报告和批量结果输出目录
  requirements.txt
  README.md
  README_EN.md
```

## 安装环境

建议使用 Python 3.10 或更高版本。

```powershell
cd ThemeIndexAnnualReportExtractor
python -m venv .venv
.\.venv\Scripts\activate
python -m pip install -r requirements.txt
```

依赖包括：

| 依赖 | 用途 |
| --- | --- |
| `beautifulsoup4` | 解析 HTML 年报。 |
| `lxml` | 提高 HTML 解析稳定性。 |
| `openpyxl` | 读写 Excel。 |
| `openai` | 调用 OpenAI 兼容模型接口。 |

## 准备输入数据

把年报 HTML 放到 `data/downloads` 下，每家公司一个文件夹：

```text
data/downloads/
  Example Company A/
    2024_Form_10-K.html
  Example Company B/
    2024_Form_20-F.html
```

注意事项：

- 每个公司文件夹至少需要一个 `.html` 或 `.htm` 文件。
- 如果同一公司有多个 HTML，脚本会优先选择更像年报的文件。
- 建议每家公司保留最新、最完整的年报 HTML。
- 公司文件夹名称会作为默认公司名称使用。

## 第一步：规则提取

运行：

```powershell
python scripts\theme_index_extractor.py
```

默认读取：

```text
data/downloads
```

默认输出：

```text
outputs/主题指数提取总表.xlsx
outputs/主题指数提取运行报告.xlsx
data/downloads/<公司名称>/主题指数提取结果.xlsx
```

常用参数：

```powershell
python scripts\theme_index_extractor.py --limit 10
python scripts\theme_index_extractor.py --companies "Example Company A" "Example Company B"
python scripts\theme_index_extractor.py --downloads-dir "path\to\downloads"
```

这一步不调用大模型，只做本地规则提取。

## 第二步：生成大模型证据包

运行：

```powershell
python scripts\build_llm_section_inputs.py
```

默认输出：

```text
data/downloads/<公司名称>/大模型证据包.json
outputs/大模型证据包生成报告.csv
```

证据包包含：

| 字段 | 说明 |
| --- | --- |
| `company` | 公司文件夹信息。 |
| `report` | 选中的年报文件、文件类型和日期线索。 |
| `source_quality_notes` | 文件质量提示，例如 40-F 只有 Exhibit 索引、10-K/A 可能缺正文等。 |
| `section_index` | Business、MD&A、Notes 等章节定位索引。 |
| `contexts` | 与主营业务、分部、收入构成、收入确认、主题线索相关的原文上下文。 |
| `relevant_tables` | 与收入、分部、产品、地区、客户、毛利等相关的表格。 |
| `text_stats` | 字符数、上下文数量、表格数量等统计信息。 |

这一步也不调用大模型。它的作用是把年报中最可能对应目标字段的原文证据整理好，供后续模型和人工复核使用。

## 第三步：调用大模型

本项目使用 OpenAI SDK 调用兼容接口。以火山方舟/豆包为例，先设置环境变量：

```powershell
$env:ARK_API_KEY="your-api-key"
```

先小批量测试：

```powershell
python scripts\run_theme_index_llm_batch.py --limit 3
```

指定公司测试：

```powershell
python scripts\run_theme_index_llm_batch.py --companies "Example Company A" "Example Company B" --rerun
```

如果你的模型接入点不是默认值，可以指定：

```powershell
python scripts\run_theme_index_llm_batch.py --model-name "your-model-or-endpoint-id" --limit 3
```

如果 API Key 环境变量不是 `ARK_API_KEY`：

```powershell
python scripts\run_theme_index_llm_batch.py --api-key-env YOUR_API_KEY_ENV --limit 3
```

如需传入公司映射表，表头需要包含：

```text
CUSIP
SECU_CODE
SECUNAME
```

运行示例：

```powershell
python scripts\run_theme_index_llm_batch.py --mapping-file "path\to\company_mapping.xlsx" --limit 3
```

模型阶段默认输出：

```text
outputs/主题指数大模型正式输出/主题指数大模型总表.xlsx
outputs/主题指数大模型正式输出/主题指数大模型运行报告.xlsx
data/downloads/<公司名称>/主题指数大模型提取结果.xlsx
data/downloads/<公司名称>/主题指数大模型提取结果.json
data/downloads/<公司名称>/主题指数大模型原始响应.txt
```

## 推荐完整流程

首次使用建议按这个顺序：

```powershell
python scripts\theme_index_extractor.py --limit 3
python scripts\build_llm_section_inputs.py --limit 3
python scripts\run_theme_index_llm_batch.py --limit 3
```

确认结果口径后再跑全量：

```powershell
python scripts\theme_index_extractor.py
python scripts\build_llm_section_inputs.py
python scripts\run_theme_index_llm_batch.py
```

如果公司数量很多，建议分批运行模型，例如：

```powershell
python scripts\run_theme_index_llm_batch.py --limit 50
```

模型调用可能产生费用，也可能受接口限速影响。批量运行前请先确认模型额度、上下文长度和输出稳定性。

## 输出字段说明

模型阶段输出五类内容：

| 输出 | 说明 |
| --- | --- |
| `table1_company_business` | 公司主营业务、报告期、总收入、披露维度和数据质量。 |
| `table2_segment_detail` | 业务分部、产品、行业/应用、地区、客户类型等收入构成明细。 |
| `table3_theme_dictionary` | 当前公司材料中实际命中的主题关键词。 |
| `table4_theme_mapping` | 公司与主题之间的相关性、证据、置信度和质量标记。 |
| `field_evidence_detail` | 面向人工复核的字段证据，包括原文、中文解释、来源和置信度。 |

`field_evidence_detail` 不只给总结，还保留可追溯的原文句子、段落或表格行。

## 关键词扩展

默认规则已经包含常见年报措辞，例如：

- `Segment Information`
- `Reportable Segments`
- `Revenue by Category`
- `Disaggregated Revenue`
- `Revenue Recognition`
- `Products and Services`
- `Geographic Information`
- `Major Customers`

如果某类公司经常漏提，可以在 `templates/keyword_overrides.json` 中增加经过人工审核的关键词。建议原则：

- 只添加短而稳定、可复用的表达。
- 避免添加公司专属名称、董事姓名或一次性描述。
- 避免把风险因素中的宽泛词直接当主题关键词。
- 模型可以给关键词建议，但建议由人工确认后再写入模板。

`templates/keyword_feedback_template.json` 提供了模型反馈关键词的格式。

## 常见问题

### 1. 为什么要先生成证据包再调用模型？

直接把完整年报交给模型很容易超出上下文，且结果不容易追溯。证据包会先用规则程序提取关键章节、上下文和表格，让模型基于更清晰的材料进行结构化抽取。

### 2. 证据包会不会漏内容？

证据包会围绕主营业务、收入构成、分部、收入确认和主题线索提取大量上下文和相关表格。SEC 年报格式差异很大，建议抽样复核。遇到稳定漏提的表达时，可以扩展关键词模板。

### 3. 40-F 为什么经常输出为空？

部分 40-F 文件只是 Exhibit 索引或引用页，正文业务和财务附注可能在附件里。如果 HTML 不包含附件正文，工具会标记材料不足，避免编造收入和分部。

### 4. 可以不用大模型吗？

可以。规则提取阶段和证据包生成阶段都不调用模型。大模型阶段只是把证据转成更完整的结构化结果。

## 致谢

这个项目来自真实批量年报处理流程中的工程化整理：先让程序承担稳定、重复、可追踪的证据筛选，再把复杂语义判断交给大模型和人工复核。希望它能把繁琐的年报阅读压缩成更清楚、更可检查的数据工作流。
