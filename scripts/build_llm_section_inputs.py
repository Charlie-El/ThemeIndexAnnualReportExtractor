import argparse
import csv
import json
import re
import warnings
from datetime import datetime
from pathlib import Path

from bs4 import XMLParsedAsHTMLWarning

import annual_report_blocks as report_blocks


ROOT_DIR = Path.cwd()
DATA_DIR = ROOT_DIR / "data"
OUTPUTS_DIR = ROOT_DIR / "outputs"
DOWNLOADS_DIR = DATA_DIR / "downloads"
EVIDENCE_NAME = "大模型证据包.json"
REPORT_NAME = "大模型证据包生成报告.csv"
HTML_SUFFIXES = {".htm", ".html"}
MAX_MODEL_JSON_CHARS = 360000
MAX_CONTEXT_CHARS = 52000

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)


CONTEXT_GROUPS = [
    {
        "name": "main_business_overview",
        "title": "主营业务/公司概况原文",
        "sources": {"business"},
        "priority": 3,
        "before": 1800,
        "after": 30000,
        "terms": [
            r"company overview",
            r"business overview",
            r"\bour business(?:es)?\b",
            r"description of business",
            r"products and services",
            r"product offerings",
            r"principal products",
            r"principal services",
            r"product portfolio",
            r"market ic solutions",
            r"applications",
            r"end markets?",
        ],
    },
    {
        "name": "segment_information",
        "title": "经营分部/业务分部原文",
        "sources": {"notes", "mda", "business"},
        "priority": 1,
        "before": 1500,
        "after": 32000,
        "terms": [
            r"segment information",
            r"reportable segments?",
            r"operating segments?",
            r"business segments?",
            r"chief operating decision maker",
            r"\bCODM\b",
            r"single reportable segment",
            r"one reportable segment",
        ],
    },
    {
        "name": "revenue_breakdown",
        "title": "收入构成/收入拆分原文",
        "sources": {"notes", "mda", "business"},
        "priority": 1,
        "before": 1500,
        "after": 36000,
        "terms": [
            r"revenue by category",
            r"revenue by segment",
            r"revenues by segment",
            r"revenue by",
            r"revenues by",
            r"sales by",
            r"net sales by",
            r"disaggregated revenues?",
            r"disaggregation of revenues?",
            r"revenue from contracts with customers",
            r"product revenue",
            r"service revenue",
            r"contract type",
            r"customer category",
            r"geographic(?:al)? revenue",
            r"geographic(?:al)? information",
            r"major customers?",
            r"customer concentration",
        ],
    },
    {
        "name": "revenue_recognition",
        "title": "收入确认原文",
        "sources": {"notes", "mda"},
        "priority": 4,
        "before": 1200,
        "after": 24000,
        "terms": [
            r"revenue recognition",
            r"recognition of revenue",
            r"revenue is recognized",
            r"revenues are recognized",
            r"net sales are recognized",
            r"recognized when",
            r"transfer of control",
            r"performance obligations?",
            r"deferred revenue",
            r"contract liabilities?",
        ],
    },
    {
        "name": "theme_relevance",
        "title": "主题相关业务线索原文",
        "sources": {"business", "mda", "notes"},
        "priority": 5,
        "before": 1400,
        "after": 14000,
        "terms": [
            r"artificial intelligence",
            r"\bAI\b",
            r"accelerated computing",
            r"data center",
            r"semiconductor",
            r"integrated circuits?",
            r"sensor",
            r"robotics?",
            r"autonomous",
            r"satellite",
            r"aerospace",
            r"defense",
            r"cybersecurity",
            r"export controls?",
            r"oil",
            r"natural gas",
            r"\bLNG\b",
            r"renewable",
            r"gold",
            r"copper",
            r"lithium",
            r"mineral",
            r"clinical trials?",
            r"biotechnology",
            r"pharmaceutical",
        ],
    },
]


PDF_REQUIREMENT_FIELDS = {
    "table1_company_business": "公司主营业务原文、中文概括、总收入、分部披露状态、主要披露维度、收入拆分状态。",
    "table2_segment_detail": "经营分部、产品/服务、行业/应用、地区、客户类型等收入构成明细；需要金额、占比、毛利/毛利率、原文证据和中文总结。",
    "table3_theme_dictionary": "从公司年报中实际命中的主题关键词字典。",
    "table4_theme_mapping": "公司与主题的关联关系、证据原文、置信度和数据质量。",
}


def html_files(folder: Path) -> list[Path]:
    return sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in HTML_SUFFIXES])


def parse_report_meta(path: Path) -> tuple[str, str]:
    form_match = re.search(r"Form\s+([A-Za-z0-9-]+)", path.name, re.I)
    dates = re.findall(r"20\d{2}-\d{2}-\d{2}", path.name)
    return (form_match.group(1).upper() if form_match else "", dates[-1] if dates else "")


def select_report_file(files: list[Path]) -> Path:
    def score(path: Path) -> tuple[int, str, int]:
        name = path.name.lower()
        form_score = 0
        if "form 10-k" in name or "form 20-f" in name or "form 40-f" in name:
            form_score += 30
        if "form nt" in name or "12b-25" in name:
            form_score -= 40
        dates = re.findall(r"20\d{2}-\d{2}-\d{2}", name)
        return (form_score, dates[-1] if dates else "", path.stat().st_size)

    return sorted(files, key=score, reverse=True)[0]


def source_quality_notes(report_path: Path, form_type: str, text: str) -> list[str]:
    notes = []
    lower_name = report_path.name.lower()
    lower_text = text[:25000].lower()
    if form_type.startswith("NT") or "notification of late filing" in lower_name or "form 12b-25" in lower_text:
        notes.append("当前文件疑似 NT 延迟提交通知，不是完整年报；如无完整业务和财务披露，应输出无法提取。")
    if "amend" in lower_name or "amendment no." in lower_text or "explanatory note" in lower_text:
        notes.append("当前文件疑似年报修订版；请确认是否包含业务、财务报表和收入附注，若只修订 Part III/治理信息则不要编造原年报内容。")
    if "no financial statements" in lower_text or "no financial statements have been included" in lower_text:
        notes.append("文件明确说明未包含财务报表；收入构成和收入确认通常无法从该文件完整提取。")
    if form_type == "40-F" and "principal documents" in lower_text and "incorporated by reference" in lower_text:
        notes.append("当前 40-F 似乎主要是 Principal Documents / Exhibit 引用页；如果未提供 Exhibit 99.1/99.2/99.3 正文，不要从 forward-looking statements 或风险提示中推断主营业务、分部或收入数据。")
    if len(text) < 12000:
        notes.append("当前 HTML 可读正文较短，可能不是完整年报或正文提取不完整。")
    return notes


def paragraph_bounds(text: str, start: int, end: int, lower: int, upper: int) -> tuple[int, int]:
    left = text.rfind("\n\n", lower, start)
    if left >= 0 and start - left < 3000:
        start = left + 2
    else:
        start = max(lower, start)
    right = text.find("\n\n", end, upper)
    if right >= 0 and right - end < 3000:
        end = right
    else:
        end = min(upper, end)
    return start, end


def section_for_absolute(position: int, sections: list[report_blocks.SectionBlock]) -> report_blocks.SectionBlock | None:
    for section in sections:
        if section.start <= position <= section.end:
            return section
    return None


def add_interval(intervals: list[dict], *, start: int, end: int, label: str, title: str, source: str, priority: int, matched: str) -> None:
    if end <= start:
        return
    intervals.append(
        {
            "start": start,
            "end": end,
            "labels": {label},
            "titles": {title},
            "source_section": source,
            "priority": priority,
            "matched_terms": {matched} if matched else set(),
        }
    )


def trim_interval(text: str, start: int, end: int, max_chars: int = MAX_CONTEXT_CHARS) -> tuple[int, int]:
    if end - start <= max_chars:
        return start, end
    end = start + max_chars
    paragraph_end = text.rfind("\n\n", start + int(max_chars * 0.72), end)
    if paragraph_end > start:
        end = paragraph_end
    return start, end


def merge_intervals(intervals: list[dict], distance: int = 2500, max_merged_chars: int = 52000) -> list[dict]:
    merged = []
    for item in sorted(intervals, key=lambda row: (row["source_section"], row["start"])):
        if not merged:
            merged.append(item)
            continue
        last = merged[-1]
        same_source = last["source_section"] == item["source_section"]
        fallback_source = last["source_section"].startswith("fallback_") or item["source_section"].startswith("fallback_")
        local_distance = 400 if fallback_source else distance
        local_max = 18000 if fallback_source else max_merged_chars
        merged_len = max(last["end"], item["end"]) - last["start"]
        if same_source and item["start"] <= last["end"] + local_distance and merged_len <= local_max:
            last["end"] = max(last["end"], item["end"])
            last["priority"] = min(last["priority"], item["priority"])
            last["labels"].update(item["labels"])
            last["titles"].update(item["titles"])
            last["matched_terms"].update(item["matched_terms"])
        else:
            merged.append(item)
    return merged


def build_context_blocks(text: str, sections: list[report_blocks.SectionBlock], priority_blocks: list[report_blocks.EvidenceBlock]) -> list[dict]:
    intervals = []
    fallback_counts: dict[str, int] = {}
    for block in priority_blocks:
        add_interval(
            intervals,
            start=block.start,
            end=block.end,
            label=block.label,
            title=block.label,
            source=block.source_section,
            priority=block.priority,
            matched=block.matched_term,
        )

    for group in CONTEXT_GROUPS:
        for pattern in group["terms"]:
            for match in re.finditer(pattern, text, re.I | re.S):
                section = section_for_absolute(match.start(), sections)
                if section and section.source_section not in group["sources"]:
                    continue
                if not section:
                    current = fallback_counts.get(group["name"], 0)
                    if current >= 3:
                        continue
                    fallback_counts[group["name"]] = current + 1
                lower = section.start if section else 0
                upper = section.end if section else len(text)
                before = group["before"] if section else min(group["before"], 1200)
                after = group["after"] if section else min(group["after"], 9000)
                raw_start = max(lower, match.start() - before)
                raw_end = min(upper, match.end() + after)
                start, end = paragraph_bounds(text, raw_start, raw_end, lower, upper)
                add_interval(
                    intervals,
                    start=start,
                    end=end,
                    label=group["name"],
                    title=group["title"],
                    source=section.source_section if section else f"fallback_{group['name']}",
                    priority=group["priority"],
                    matched=match.group(0),
                )

    contexts = []
    for index, item in enumerate(merge_intervals(intervals), start=1):
        start, end = trim_interval(text, item["start"], item["end"])
        content = report_blocks.clean_text(text[start:end])
        if len(content) < 180:
            continue
        contexts.append(
            {
                "context_id": f"C{index:03d}",
                "source_section": item["source_section"],
                "priority": item["priority"],
                "context_types": sorted(item["labels"]),
                "title": " / ".join(sorted(item["titles"])),
                "start": start,
                "end": end,
                "matched_terms": sorted(term for term in item["matched_terms"] if term),
                "text": content,
            }
        )
    contexts.sort(key=lambda row: (row["priority"], {"notes": 0, "business": 1, "mda": 2}.get(row["source_section"], 9), row["start"]))
    final_contexts = []
    seen_ranges: list[tuple[int, int]] = []
    for row in contexts:
        if row["source_section"].startswith("fallback_"):
            covered = False
            for start, end in seen_ranges:
                overlap = max(0, min(end, row["end"]) - max(start, row["start"]))
                if overlap >= 0.75 * max(1, row["end"] - row["start"]):
                    covered = True
                    break
            if covered:
                continue
        final_contexts.append(row)
        if not row["source_section"].startswith("fallback_"):
            seen_ranges.append((row["start"], row["end"]))
    contexts = final_contexts
    for index, row in enumerate(contexts, start=1):
        row["context_id"] = f"C{index:03d}"
    return contexts


def table_to_dict(table: report_blocks.TableBlock) -> dict:
    keywords = [item.strip() for item in table.keywords.split(",") if item.strip()]
    return {
        "table_id": f"T{table.table_id:04d}",
        "original_table_id": table.table_id,
        "relevance_score": table.score,
        "latest_year": table.latest_year or None,
        "keywords": keywords,
        "source_section": table.source_section,
        "nearby_heading": table.heading,
        "text": table.text,
    }


def select_relevant_tables(tables: list[report_blocks.TableBlock], max_tables: int = 36) -> list[report_blocks.TableBlock]:
    selected: list[report_blocks.TableBlock] = []
    seen_text = set()
    for table in sorted(tables, key=lambda row: (row.score, row.latest_year, len(row.text)), reverse=True):
        compact = re.sub(r"\s+", " ", table.text).strip().lower()
        key = compact[:900]
        if key in seen_text:
            continue
        seen_text.add(key)
        selected.append(table)
        if len(selected) >= max_tables:
            break
    selected.sort(key=lambda row: row.table_id)
    return selected


def package_size(package: dict) -> int:
    return len(json.dumps(package, ensure_ascii=False))


def fit_package_budget(package: dict, max_chars: int = MAX_MODEL_JSON_CHARS) -> None:
    package["text_stats"]["evidence_json_characters_estimate"] = package_size(package)
    while package["text_stats"]["evidence_json_characters_estimate"] > max_chars:
        if package["relevant_tables"] and len(package["relevant_tables"]) > 24:
            package["relevant_tables"] = package["relevant_tables"][:-4]
        else:
            contexts = package["contexts"]
            fallback_indices = [i for i, row in enumerate(contexts) if str(row.get("source_section", "")).startswith("fallback_")]
            if fallback_indices:
                contexts.pop(fallback_indices[-1])
            elif len(contexts) > 6:
                contexts.pop()
            else:
                longest = max(contexts, key=lambda row: len(row.get("text", "")), default=None)
                if not longest or len(longest.get("text", "")) <= 24000:
                    break
                longest["text"] = longest["text"][:24000]
                longest["note"] = "该上下文因模型输入预算被截到 24000 字符；原 HTML 文件仍保留在公司目录。"
                longest["end"] = longest["start"] + len(longest["text"])
        package["text_stats"]["contexts"] = len(package["contexts"])
        package["text_stats"]["context_characters"] = sum(len(row["text"]) for row in package["contexts"])
        package["text_stats"]["relevant_tables"] = len(package["relevant_tables"])
        package["text_stats"]["relevant_table_characters"] = sum(len(table["text"]) for table in package["relevant_tables"])
        package["text_stats"]["evidence_json_characters_estimate"] = package_size(package)


def build_evidence_package(folder: Path) -> tuple[dict, Path]:
    files = html_files(folder)
    if not files:
        raise FileNotFoundError("公司目录下没有 html/htm 年报文件")

    report_path = select_report_file(files)
    form_type, report_date = parse_report_meta(report_path)
    full_text, soup = report_blocks.read_html_document(report_path)
    sections = report_blocks.extract_sections(full_text, max_chars=None)
    priority_blocks = report_blocks.extract_priority_blocks(full_text, sections, max_chars=None)
    relevant_tables = select_relevant_tables(report_blocks.extract_table_blocks(soup, full_text, sections, limit=None, max_chars=None))
    contexts = build_context_blocks(full_text, sections, priority_blocks)
    quality_notes = source_quality_notes(report_path, form_type, full_text)

    package = {
        "schema_version": "3.0",
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "model_called": False,
        "purpose": "本文件是本地生成的主题指数抽取证据包，不调用大模型；后续脚本会把该 JSON 连同提示词交给豆包长上下文模型抽取。",
        "company": {
            "folder_name": folder.name,
            "folder_path": str(folder),
        },
        "report": {
            "selected_file": report_path.name,
            "selected_file_path": str(report_path),
            "available_html_files": [path.name for path in files],
            "form_type_from_filename": form_type,
            "report_date_from_filename": report_date,
        },
        "pdf_requirement_fields": PDF_REQUIREMENT_FIELDS,
        "source_quality_notes": quality_notes,
        "extraction_policy": {
            "html_preserved": True,
            "non_html_outputs_cleaned_before_generation": "由清理步骤保证；本脚本不会删除文件。",
            "context_strategy": "围绕 PDF 需求字段抽取完整上下文段落和完整相关表格；宁可多给上下文，不只截一句。",
            "content_truncation": "selected contexts are capped by model input budget; source HTML remains complete in the company folder",
            "full_text_included": False,
            "full_text_reason": "完整 HTML 原文已保留在公司目录；为保证 128k 模型稳定输出，JSON 输入保留关键字段上下文和完整相关表格。",
        },
        "section_index": [
            {
                "name": section.name,
                "source_section": section.source_section,
                "priority": section.priority,
                "start": section.start,
                "end": section.end,
                "matched_heading": section.matched_heading,
            }
            for section in sections
        ],
        "contexts": contexts,
        "relevant_tables": [table_to_dict(table) for table in relevant_tables],
        "text_stats": {
            "full_text_characters": len(full_text),
            "detected_sections": len(sections),
            "contexts": len(contexts),
            "context_characters": sum(len(row["text"]) for row in contexts),
            "relevant_tables": len(relevant_tables),
            "relevant_table_characters": sum(len(table.text) for table in relevant_tables),
            "evidence_json_characters_estimate": 0,
        },
    }
    fit_package_budget(package)
    return package, report_path


def process_company(folder: Path) -> dict:
    package, report_path = build_evidence_package(folder)
    output_path = folder / EVIDENCE_NAME
    output_path.write_text(json.dumps(package, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        "company": folder.name,
        "status": "完成",
        "report_file": report_path.name,
        "contexts": package["text_stats"]["contexts"],
        "context_characters": package["text_stats"]["context_characters"],
        "relevant_tables": package["text_stats"]["relevant_tables"],
        "table_characters": package["text_stats"]["relevant_table_characters"],
        "json_characters": package["text_stats"]["evidence_json_characters_estimate"],
        "quality_notes": "；".join(package["source_quality_notes"]),
        "output": str(output_path),
        "note": "",
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Build JSON evidence packages from annual report HTML files. This script does not call any LLM.")
    parser.add_argument("--downloads-dir", default=str(DOWNLOADS_DIR))
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--companies", nargs="*", default=[])
    args = parser.parse_args()

    downloads_dir = Path(args.downloads_dir)
    folders = sorted([p for p in downloads_dir.iterdir() if p.is_dir()], key=lambda p: p.name.lower())
    if args.companies:
        wanted = {name.lower() for name in args.companies}
        folders = [folder for folder in folders if folder.name.lower() in wanted]
    if args.limit > 0:
        folders = folders[: args.limit]

    rows = []
    for index, folder in enumerate(folders, start=1):
        print(f"[{index}/{len(folders)}] {folder.name}", flush=True)
        try:
            rows.append(process_company(folder))
        except Exception as exc:
            rows.append(
                {
                    "company": folder.name,
                    "status": "失败",
                    "report_file": "",
                    "contexts": 0,
                    "context_characters": 0,
                    "relevant_tables": 0,
                    "table_characters": 0,
                    "json_characters": 0,
                    "quality_notes": "",
                    "output": "",
                    "note": f"{type(exc).__name__}: {exc}",
                }
            )

    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    report_path = OUTPUTS_DIR / REPORT_NAME
    fieldnames = [
        "company",
        "status",
        "report_file",
        "contexts",
        "context_characters",
        "relevant_tables",
        "table_characters",
        "json_characters",
        "quality_notes",
        "output",
        "note",
    ]
    with report_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(report_path)
    print(f"processed={len(rows)}")


if __name__ == "__main__":
    main()
