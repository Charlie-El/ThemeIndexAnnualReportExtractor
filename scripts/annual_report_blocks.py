import re
import warnings
from dataclasses import dataclass

from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning


warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)


MAX_SECTION_CHARS = 70000
MAX_BLOCK_CHARS = 26000
MAX_TABLE_CHARS = 16000


def clean_text(value: str) -> str:
    value = re.sub(r"\u200b", "", value or "")
    value = value.replace("\xa0", " ")
    value = re.sub(r"[ \t\r\f\v]+", " ", value)
    value = re.sub(r" *\n *", "\n", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def one_line(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def loose_word(word: str) -> str:
    return r"\s*".join(re.escape(ch) for ch in word)


BUSINESS = loose_word("business")
FINANCIAL = loose_word("financial")
STATEMENTS = loose_word("statements")
MANAGEMENT = loose_word("management")
DISCUSSION = loose_word("discussion")
ANALYSIS = loose_word("analysis")
RISK = loose_word("risk")
FACTORS = loose_word("factors")


@dataclass
class SectionBlock:
    name: str
    source_section: str
    priority: int
    start: int
    end: int
    matched_heading: str
    text: str


@dataclass
class EvidenceBlock:
    label: str
    source_section: str
    priority: int
    start: int
    end: int
    matched_term: str
    text: str


@dataclass
class TableBlock:
    table_id: int
    score: int
    latest_year: int
    keywords: str
    source_section: str
    heading: str
    text: str


SECTION_DEFS = [
    {
        "name": "Item 8 / Notes to Consolidated Financial Statements",
        "source": "notes",
        "priority": 1,
        "starts": [
            rf"\bitem\s*8\.?\s*{FINANCIAL}\s+{STATEMENTS}",
            rf"\bitem\s*18\.?\s*{FINANCIAL}\s+{STATEMENTS}",
            r"\bnotes\s+to\s+(?:the\s+)?consolidated\s+financial\s+statements\b",
            r"\bnotes\s+to\s+(?:the\s+)?financial\s+statements\b",
        ],
        "ends": [
            r"\bitem\s*9\.?\s*changes\s+in\s+and\s+disagreements\b",
            r"\bitem\s*19\.?\s*exhibits\b",
            r"\bsignatures\b",
        ],
    },
    {
        "name": "Item 1 Business / Item 4 Information on the Company",
        "source": "business",
        "priority": 3,
        "starts": [
            rf"\bitem\s*1\.?\s*{BUSINESS}",
            r"\bitem\s*4\.?\s*information\s+on\s+the\s+company\b",
            r"\bcompany\s+overview\b",
            r"\bbusiness\s+overview\b",
            r"\boverview\s+of\s+the\s+business\b",
            r"\bannual\s+information\s+form\b",
        ],
        "ends": [
            rf"\bitem\s*1a\.?\s*{RISK}\s+{FACTORS}",
            r"\bitem\s*4a\.?\s*unresolved\s+staff\s+comments\b",
            r"\bitem\s*5\.?\s*operating\s+and\s+financial\s+review\b",
            r"\bmanagement.s\s+discussion\s+and\s+analysis\b",
        ],
    },
    {
        "name": "Item 7 MD&A / Item 5 Operating and Financial Review",
        "source": "mda",
        "priority": 3,
        "starts": [
            rf"\bitem\s*7\.?\s*{MANAGEMENT}.{{0,20}}{DISCUSSION}.{{0,20}}{ANALYSIS}",
            r"\bitem\s*5\.?\s*operating\s+and\s+financial\s+review\b",
            r"\bmanagement.s\s+discussion\s+and\s+analysis\b",
            r"\bmanagement.s\s+discussion\s+and\s+analysis\s+for\s+the\s+year\b",
            r"\bresults\s+of\s+operations\b",
        ],
        "ends": [
            r"\bitem\s*7a\.?\s*quantitative\s+and\s+qualitative\b",
            rf"\bitem\s*8\.?\s*{FINANCIAL}\s+{STATEMENTS}",
            r"\bitem\s*6\.?\s*directors\b",
            r"\bannual\s+audited\s+consolidated\s+financial\s+statements\b",
        ],
    },
]


TARGET_BLOCK_DEFS = [
    {
        "label": "P1 Notes - SEGMENTS / Segment Information",
        "priority": 1,
        "sources": {"notes"},
        "terms": [
            r"segment information",
            r"reportable segments?",
            r"operating segments?",
            r"business segments?",
            r"chief operating decision maker",
            r"\bCODM\b",
        ],
        "before": 900,
        "after": 21000,
    },
    {
        "label": "P1 Notes - Revenue by segment / product / service",
        "priority": 1,
        "sources": {"notes"},
        "terms": [
            r"revenue by category",
            r"revenue by segment",
            r"revenues by segment",
            r"disaggregated revenues?",
            r"disaggregation of revenues?",
            r"revenue from contracts with customers",
        ],
        "before": 800,
        "after": 24000,
    },
    {
        "label": "P2 Notes - REVENUES / Revenue composition",
        "priority": 2,
        "sources": {"notes"},
        "terms": [
            r"\brevenues?\b",
            r"\bnet sales\b",
            r"\bsales by\b",
            r"\brevenue by\b",
            r"\bmajor customers?\b",
            r"\bcustomer concentration\b",
        ],
        "before": 700,
        "after": 18000,
    },
    {
        "label": "P3 Business - overview / products / services",
        "priority": 3,
        "sources": {"business"},
        "terms": [
            r"company overview",
            r"business overview",
            r"\bour mission\b",
            r"\bour business(?:es)?\b",
            r"products and services",
            r"product offerings",
            r"principal products",
            r"principal services",
        ],
        "before": 300,
        "after": 22000,
    },
    {
        "label": "P3 Business - product matrix / applications",
        "priority": 3,
        "sources": {"business"},
        "terms": [
            r"market ic solutions",
            r"product portfolio",
            r"products\s+applications",
            r"applications",
            r"solutions",
            r"end markets?",
        ],
        "before": 900,
        "after": 16000,
    },
    {
        "label": "P3 MD&A - results / revenue composition",
        "priority": 3,
        "sources": {"mda"},
        "terms": [
            r"results of operations",
            r"components of our results",
            r"\bnet sales\b",
            r"\brevenue by\b",
            r"\brevenues?\b",
            r"product revenue",
            r"service revenue",
            r"geographic",
        ],
        "before": 700,
        "after": 22000,
    },
    {
        "label": "P4 Notes - Revenue Recognition",
        "priority": 4,
        "sources": {"notes", "mda"},
        "terms": [
            r"revenue recognition",
            r"recognition of revenue",
            r"revenue is recognized",
            r"revenues are recognized",
            r"recognized when",
            r"transfer of control",
            r"performance obligations?",
            r"deferred revenue",
            r"contract liabilities?",
        ],
        "before": 700,
        "after": 16000,
    },
]


TABLE_KEYWORDS = [
    "revenue",
    "revenues",
    "net sales",
    "net revenue",
    "sales",
    "segment",
    "segments",
    "product",
    "products",
    "service",
    "services",
    "contract type",
    "customer category",
    "customer",
    "customers",
    "geographic",
    "geographical",
    "domestic",
    "international",
    "region",
    "regions",
    "gross profit",
    "gross margin",
    "operating income",
    "automotive",
    "industrial",
    "data center",
    "applications",
    "ffp",
    "cpff",
    "t&m",
    "u.s. government",
    "non-u.s. government",
    "magnetic sensors",
    "power integrated circuits",
]


def remove_noise(soup: BeautifulSoup) -> None:
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    for tag in soup.find_all(True):
        if getattr(tag, "attrs", None) is None:
            continue
        name = (tag.name or "").lower()
        style = (tag.get("style") or "").lower().replace(" ", "")
        if name.startswith("ix:") or name in {"ix:header", "ix:hidden", "ix:resources"}:
            tag.decompose()
        elif "display:none" in style or "visibility:hidden" in style:
            tag.decompose()


def read_html_document(path) -> tuple[str, BeautifulSoup]:
    raw = path.read_text(encoding="utf-8", errors="ignore")
    soup = BeautifulSoup(raw, "lxml")
    remove_noise(soup)
    return clean_text(soup.get_text("\n")), soup


def looks_like_toc(block: str) -> bool:
    compact = one_line(block[:1800]).lower()
    if "notes to the consolidated financial statements" in compact[:220] and compact.count("table of contents") >= 4:
        return False
    item_count = len(re.findall(r"\bitem\s+\d+[a-z]?\.", compact, re.I))
    dotted_count = compact.count("....")
    if item_count >= 4 or dotted_count >= 2:
        return True
    if ("table of contents" in compact[:800] or "index to form" in compact[:800]) and item_count >= 2:
        return True
    if re.search(r"\bitem\s+7\b", compact) and re.search(r"\bitem\s+8\b", compact) and re.search(r"\bitem\s+9\b", compact):
        return True
    if re.search(r"\bf-\s*4\b", compact) and re.search(r"\bf-\s*8\b", compact) and "notes to consolidated financial statements" in compact:
        return True
    return False


def looks_like_statement_index(block: str) -> bool:
    compact = one_line(block[:2200]).lower()
    return (
        "consolidated balance sheets" in compact
        and "consolidated statements" in compact
        and "notes to consolidated financial statements" in compact
    )


def looks_like_financial_statement_page(block: str) -> bool:
    compact = one_line(block[:1400]).lower()
    if "notes to the consolidated financial statements" in compact and "table of contents" in compact:
        return False
    return (
        "see accompanying notes to consolidated financial statements" in compact
        and (
            "consolidated statements of" in compact
            or "consolidated balance sheets" in compact
            or "consolidated statements of incom" in compact
        )
    )


def nearest_end(text: str, start: int, end_patterns: list[str]) -> int | None:
    nearest = None
    search_from = min(len(text), start + 450)
    for pattern in end_patterns:
        for match in re.finditer(pattern, text[search_from:], re.I | re.S):
            candidate = search_from + match.start()
            nearest = candidate if nearest is None else min(nearest, candidate)
            break
    return nearest


def cap_end(text_length: int, start: int, end: int | None, max_chars: int | None) -> int:
    if end is None:
        return text_length if max_chars is None else min(text_length, start + max_chars)
    return end if max_chars is None else min(end, start + max_chars)


def cap_text(value: str, max_chars: int | None) -> str:
    return value if max_chars is None else value[:max_chars]


def find_section(text: str, definition: dict, max_chars: int | None = MAX_SECTION_CHARS) -> SectionBlock | None:
    candidates = []
    for pattern in definition["starts"]:
        for match in re.finditer(pattern, text, re.I | re.S):
            start = match.start()
            if start < 1200:
                continue
            preview = text[start : start + 2400]
            if looks_like_toc(preview) or (definition["source"] == "notes" and looks_like_statement_index(preview)):
                continue
            if definition["source"] == "notes" and looks_like_financial_statement_page(text[max(0, start - 900) : start + 1600]):
                continue
            end = cap_end(len(text), start, nearest_end(text, start, definition["ends"]), max_chars)
            if end - start < 900:
                continue
            heading = one_line(match.group(0))
            formal_score = 100 if re.search(r"\bitem\s*\d|notes\s+to|annual\s+information\s+form", heading, re.I) else 30
            if definition["source"] == "notes":
                following = one_line(text[start : start + 1200]).lower()
                local = one_line(text[max(0, start - 120) : start + 120]).lower()
                if "table of contents" in following and len(re.findall(r"table of contents", following)) >= 4:
                    formal_score += 260
                if "amounts in thousands" in following or "significant accounting policies" in following:
                    formal_score += 80
                if re.search(r"\bnotes\s+to\s+(?:the\s+)?consolidated\s+financial\s+statements\b", heading, re.I):
                    formal_score += 30
                if "see accompanying notes" in local:
                    formal_score -= 100
                if "financial statement schedules" in following or "exhibits" in following[:600]:
                    formal_score -= 90
                if re.search(r"\bitem\s*7a\b|\bitem\s*9\b|\bquantitative\s+and\s+qualitative\b", following[:1200]):
                    formal_score -= 120
                formal_score += min(start // 20000, 20)
            if "overview" in heading.lower() or "our business" in heading.lower():
                formal_score -= 10
            candidates.append((formal_score, start, end, heading))
    if not candidates:
        if definition["source"] == "notes":
            upper_match = re.search(r"\bNOTES\s+TO\s+THE\s+CONSOLIDATED\s+FINANCIAL\s+STATEMENTS\b", text)
            if upper_match:
                start = upper_match.start()
                end = cap_end(len(text), start, nearest_end(text, start, definition["ends"]), max_chars)
                return SectionBlock(
                    name=definition["name"],
                    source_section=definition["source"],
                    priority=definition["priority"],
                    start=start,
                    end=end,
                    matched_heading=one_line(upper_match.group(0)),
                    text=cap_text(clean_text(text[start:end]), max_chars),
                )
        return None
    score, start, end, heading = sorted(candidates, key=lambda item: (-item[0], item[1]))[0]
    return SectionBlock(
        name=definition["name"],
        source_section=definition["source"],
        priority=definition["priority"],
        start=start,
        end=end,
        matched_heading=heading,
        text=cap_text(clean_text(text[start:end]), max_chars),
    )


def extract_sections(text: str, max_chars: int | None = MAX_SECTION_CHARS) -> list[SectionBlock]:
    sections: list[SectionBlock] = []
    for definition in SECTION_DEFS:
        section = find_section(text, definition, max_chars=max_chars)
        if section:
            sections.append(section)
    sections.sort(key=lambda item: (item.priority, item.start))
    return sections


def paragraph_start(text: str, position: int, lower_bound: int) -> int:
    prev_blank = text.rfind("\n\n", lower_bound, position)
    if prev_blank >= 0 and position - prev_blank < 2500:
        return prev_blank + 2
    return max(lower_bound, position - 900)


def section_for_position(position: int, sections: list[SectionBlock]) -> str:
    for section in sections:
        if section.start <= position <= section.end:
            return section.source_section
    return ""


def previous_heading(text: str, position: int, limit: int = 1800) -> str:
    start = max(0, position - limit)
    lines = [one_line(line) for line in text[start:position].splitlines() if one_line(line)]
    for line in reversed(lines[-12:]):
        if 3 <= len(line) <= 120 and not re.fullmatch(r"[\d\s$%(),.-]+", line):
            return line[:120]
    return ""


def extract_priority_blocks(
    text: str,
    sections: list[SectionBlock],
    max_chars: int | None = MAX_BLOCK_CHARS,
) -> list[EvidenceBlock]:
    blocks: list[EvidenceBlock] = []
    seen = set()
    for section in sections:
        if section.source_section == "business":
            default_end = min(section.end, section.start + 22000)
            blocks.append(
                EvidenceBlock(
                    label="P3 Business - section beginning",
                    source_section=section.source_section,
                    priority=3,
                    start=section.start,
                    end=default_end,
                    matched_term=section.matched_heading,
                    text=cap_text(clean_text(text[section.start:default_end]), max_chars),
                )
            )
        for definition in TARGET_BLOCK_DEFS:
            if section.source_section not in definition["sources"]:
                continue
            matches = []
            section_text = text[section.start : section.end]
            for pattern in definition["terms"]:
                for match in re.finditer(pattern, section_text, re.I | re.S):
                    absolute = section.start + match.start()
                    if any(abs(absolute - old) < 5000 for old in matches):
                        continue
                    matches.append(absolute)
                    break
                if len(matches) >= 3:
                    break
            for absolute in matches[:3]:
                start = paragraph_start(text, absolute - definition["before"], section.start)
                end = min(section.end, absolute + definition["after"])
                key = (definition["label"], start // 3000)
                if key in seen:
                    continue
                seen.add(key)
                content = clean_text(text[start:end])
                if len(content) < 300 or looks_like_toc(content):
                    continue
                blocks.append(
                    EvidenceBlock(
                        label=definition["label"],
                        source_section=section.source_section,
                        priority=definition["priority"],
                        start=start,
                        end=end,
                        matched_term=previous_heading(text, absolute) or one_line(text[absolute : absolute + 80]),
                        text=cap_text(content, max_chars),
                    )
                )
    blocks.sort(key=lambda item: (item.priority, {"notes": 0, "business": 1, "mda": 2}.get(item.source_section, 9), item.start))
    return blocks


def table_to_text(table) -> str:
    rows = []
    for tr in table.find_all("tr"):
        cells = []
        for cell in tr.find_all(["th", "td"]):
            cell_text = one_line(cell.get_text(" "))
            if cell_text:
                cells.append(cell_text)
        if cells:
            rows.append(" | ".join(cells))
    return "\n".join(rows)


def latest_year(text: str) -> int:
    years = [int(item) for item in re.findall(r"\b20\d{2}\b", text or "")]
    return max(years) if years else 0


def score_table_text(table_text: str) -> tuple[int, list[str]]:
    lower = table_text.lower()
    matched = [keyword for keyword in TABLE_KEYWORDS if keyword in lower]
    if not matched:
        return 0, []
    score = 0
    for keyword in matched:
        if keyword in {"revenue", "revenues", "net sales", "net revenue"}:
            score += 16
        elif keyword in {"segment", "segments"}:
            score += 16
        elif keyword in {"product", "products", "service", "services"}:
            score += 10
        elif keyword in {"geographic", "geographical", "domestic", "international", "region", "regions", "customer", "customers"}:
            score += 9
        elif keyword in {"gross profit", "gross margin", "operating income"}:
            score += 5
        else:
            score += 4
    if re.search(r"\b20\d{2}\b", table_text):
        score += 6
    if re.search(r"\d", table_text):
        score += 5
    if "total" in lower:
        score += 4
    if ("revenue" in lower or "net sales" in lower or "sales" in lower) and any(
        term in lower
        for term in [
            "segment",
            "product",
            "service",
            "contract type",
            "customer",
            "geographic",
            "domestic",
            "international",
            "automotive",
            "industrial",
        ]
    ):
        score += 30
    if "products" in lower and "applications" in lower:
        score += 28
    if "market ic solutions" in lower:
        score += 35
    if re.search(r"\bitem\s+\d+[a-z]?\.", lower) and ("part i" in lower or "part ii" in lower):
        score -= 60
    if any(term in lower for term in ["assets:", "liabilities", "stockholders", "cash and cash equivalents"]):
        score -= 20
    if lower.startswith("●") and score < 70:
        score -= 25
    return score, matched


def find_table_position(text: str, table_text: str) -> int:
    lines = [one_line(line) for line in table_text.splitlines() if len(one_line(line)) >= 20]
    best = -1
    for line in lines[:5]:
        start = 0
        needle = line[:80]
        while True:
            idx = text.find(needle, start)
            if idx < 0:
                break
            best = max(best, idx)
            start = idx + max(1, len(needle))
    if best >= 0:
        return best
    return -1


def extract_table_blocks(
    soup: BeautifulSoup,
    text: str,
    sections: list[SectionBlock],
    limit: int | None = 18,
    max_chars: int | None = MAX_TABLE_CHARS,
) -> list[TableBlock]:
    candidates: list[TableBlock] = []
    for table_id, table in enumerate(soup.find_all("table"), start=1):
        table_text = clean_text(table_to_text(table))
        if len(table_text) < 40:
            continue
        score, keywords = score_table_text(table_text)
        year = latest_year(table_text)
        if year:
            score += max(0, year - 2020)
        if score <= 0:
            continue
        position = find_table_position(text, table_text)
        source = section_for_position(position, sections) if position >= 0 else ""
        candidates.append(
            TableBlock(
                table_id=table_id,
                score=score,
                latest_year=year,
                keywords=", ".join(keywords),
                source_section=source,
                heading=previous_heading(text, position) if position >= 0 else "",
                text=cap_text(table_text, max_chars),
            )
        )
    candidates.sort(key=lambda item: (item.score, item.latest_year, len(item.text)), reverse=True)
    return candidates if limit is None else candidates[:limit]


def format_priority_evidence(blocks: list[EvidenceBlock], tables: list[TableBlock], sections: list[SectionBlock]) -> str:
    parts = [
        "===== PDF需求优先证据开始 =====",
        "说明：以下内容按《主题指数数据》PDF 的披露优先级组织。文本尽量保留年报原文和原始换行；表格使用每行单独一行、单元格以 | 分隔。",
        "优先级口径：P1=Item 8/Notes 的 Segment Information 或收入构成；P2=Notes 的 REVENUES；P3=Item 1 Business 或 Item 7 MD&A；P4=Revenue Recognition。",
    ]
    if not blocks:
        parts.append("未识别到稳定的 PDF 需求优先证据块。")
    for index, block in enumerate(blocks, start=1):
        parts.append(
            "\n"
            f"----- 证据块 {index} | {block.label} | source={block.source_section} | priority={block.priority} -----\n"
            f"定位标题/关键词：{block.matched_term}\n"
            f"{block.text}"
        )
    parts.append("===== PDF需求优先证据结束 =====")

    parts.append("\n===== 高分表格候选开始 =====")
    if not tables:
        parts.append("未识别到高分表格。")
    for index, table in enumerate(tables, start=1):
        parts.append(
            "\n"
            f"----- 表格 {index} | 原始表格序号={table.table_id} | score={table.score} | latest_year={table.latest_year or ''} | source={table.source_section or 'unknown'} -----\n"
            f"附近标题：{table.heading}\n"
            f"命中关键词：{table.keywords}\n"
            f"{table.text}"
        )
    parts.append("===== 高分表格候选结束 =====")

    parts.append("\n===== 年报章节索引开始 =====")
    for section in sections:
        parts.append(
            f"{section.name} | source={section.source_section} | start={section.start} | end={section.end} | matched={section.matched_heading}"
        )
    parts.append("===== 年报章节索引结束 =====")
    return "\n".join(parts)
