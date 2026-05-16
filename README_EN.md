# ThemeIndexAnnualReportExtractor

[中文说明](README.md)

ThemeIndexAnnualReportExtractor is a Python toolkit for extracting theme-index evidence from U.S. annual report HTML files. It prepares business descriptions, segment disclosures, revenue breakdowns, revenue recognition evidence, and theme-related clues, then can optionally call an OpenAI-compatible LLM endpoint to normalize the evidence into JSON and Excel outputs.

This repository contains reusable code, templates, and documentation only. It does not include real annual reports, company lists, API keys, model responses, logs, or generated workbooks.

## Features

- Batch-process `.html` and `.htm` annual reports stored by company folder.
- Supports common 10-K, 20-F, and 40-F style annual report HTML files.
- Locates Business, MD&A, Financial Statements, Notes, Revenue, and Segment Information sections.
- Extracts candidate evidence for business overview, operating segments, product/service revenue, geographic/customer revenue, gross margin, and revenue recognition.
- Detects theme clues for AI, semiconductors, robotics, space, geopolitical risk, energy, minerals, and biomedicine.
- Generates a per-company JSON evidence package for LLM input and human review.
- Optionally calls an OpenAI-compatible endpoint, such as Ark/Doubao, to produce structured JSON and Excel results.
- Keeps company-level outputs separate for easier audit and reruns.
- Provides strict `.gitignore` rules to avoid publishing private data.

## Project Layout

```text
ThemeIndexAnnualReportExtractor/
  scripts/
    annual_report_blocks.py        # Annual report section and table helpers
    theme_index_extractor.py       # Rule-based extraction pipeline
    build_llm_section_inputs.py    # Builds per-company JSON evidence packages
    run_theme_index_llm_batch.py   # Runs OpenAI-compatible LLM extraction
  templates/
    keyword_overrides.json
    keyword_feedback_template.json
    README.md
  data/
    downloads/
      .gitkeep
  outputs/
    .gitkeep
  docs/
    PRIVACY.md
  requirements.txt
  README.md
  README_EN.md
```

## Installation

Python 3.10 or newer is recommended.

```powershell
cd ThemeIndexAnnualReportExtractor
python -m venv .venv
.\.venv\Scripts\activate
python -m pip install -r requirements.txt
```

## Prepare Input Files

Place annual report HTML files under `data/downloads`, one folder per company:

```text
data/downloads/
  Example Company A/
    2024_Form_10-K.html
  Example Company B/
    2024_Form_20-F.html
```

Each company folder should contain at least one `.html` or `.htm` file. If a folder contains multiple HTML files, the scripts try to select the most likely annual report, but it is better to keep the latest complete annual report only.

## Step 1: Rule-Based Extraction

```powershell
python scripts\theme_index_extractor.py
```

Default outputs:

```text
outputs/主题指数提取总表.xlsx
outputs/主题指数提取运行报告.xlsx
data/downloads/<company>/主题指数提取结果.xlsx
```

Useful options:

```powershell
python scripts\theme_index_extractor.py --limit 10
python scripts\theme_index_extractor.py --companies "Example Company A" "Example Company B"
python scripts\theme_index_extractor.py --downloads-dir "path\to\downloads"
```

This step does not call any model.

## Step 2: Build LLM Evidence Packages

```powershell
python scripts\build_llm_section_inputs.py
```

Default outputs:

```text
data/downloads/<company>/大模型证据包.json
outputs/大模型证据包生成报告.csv
```

The evidence package contains company metadata, report metadata, quality notes, section indexes, relevant context blocks, relevant tables, and extraction statistics. This step also does not call any model.

## Step 3: Run LLM Extraction

Set your API key first:

```powershell
$env:ARK_API_KEY="your-api-key"
```

Run a small test:

```powershell
python scripts\run_theme_index_llm_batch.py --limit 3
```

Run specific companies:

```powershell
python scripts\run_theme_index_llm_batch.py --companies "Example Company A" "Example Company B" --rerun
```

Use a custom model or endpoint:

```powershell
python scripts\run_theme_index_llm_batch.py --model-name "your-model-or-endpoint-id" --limit 3
```

Use a custom API key environment variable:

```powershell
python scripts\run_theme_index_llm_batch.py --api-key-env YOUR_API_KEY_ENV --limit 3
```

Optional company mapping workbooks should contain these columns:

```text
CUSIP
SECU_CODE
SECUNAME
```

Example:

```powershell
python scripts\run_theme_index_llm_batch.py --mapping-file "path\to\company_mapping.xlsx" --limit 3
```

## Outputs

The LLM stage writes:

```text
outputs/主题指数大模型正式输出/主题指数大模型总表.xlsx
outputs/主题指数大模型正式输出/主题指数大模型运行报告.xlsx
data/downloads/<company>/主题指数大模型提取结果.xlsx
data/downloads/<company>/主题指数大模型提取结果.json
data/downloads/<company>/主题指数大模型原始响应.txt
```

The model result contains:

- `table1_company_business`
- `table2_segment_detail`
- `table3_theme_dictionary`
- `table4_theme_mapping`
- `field_evidence_detail`

`field_evidence_detail` is designed for review. It keeps original evidence text along with the extracted value, Chinese summary, source id, and confidence.

## Keyword Extensions

Use `templates/keyword_overrides.json` for reviewed, reusable keyword additions. Keep terms short and stable. Avoid company-specific names, person names, broad risk words, and private notes.

Use `templates/keyword_feedback_template.json` when asking a model to suggest missing keywords. The model should only provide suggestions; code changes should remain human-reviewed.

## Privacy

Before publishing, make sure you do not commit:

- Annual report HTML/PDF files.
- Internal screening workbooks or company mapping files.
- API keys or `.env` files.
- Model responses, logs, JSONL progress files, or generated Excel outputs.
- Files under `data/downloads/` or `outputs/`, except `.gitkeep`.

See [docs/PRIVACY.md](docs/PRIVACY.md) for the release checklist.

## Notes

- This project is a research workflow helper, not investment advice.
- SEC filings vary widely. 40-F, 10-K/A, 20-F/A, and NT filings may not contain complete business or financial statement content.
- LLM output should be sampled and reviewed before production use.
- Run a small batch first, then scale up.

## Suggested Repository Description

```text
Extract structured theme-index evidence from annual report HTML files, with optional LLM-assisted field normalization.
```
