# parsers/horizon.py
from __future__ import annotations
import re
from io import BytesIO
from typing import Dict, Any, List
import fitz  # PyMuPDF
import pandas as pd

# ========== PDF Parsing ==========
def extract_text_from_pdf(file_like: BytesIO) -> str:
    with fitz.open(stream=file_like.read(), filetype="pdf") as doc:
        return "\n".join(page.get_text() for page in doc)

# ========== Utility ==========
def normalize_text(text: str) -> str:
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = re.sub(r"\xa0", " ", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()

# ========== Topic Extraction ==========
def extract_topic_blocks(text: str) -> List[Dict[str, Any]]:
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    fixed_lines = []
    i = 0
    while i < len(lines):
        if re.match(r"^HORIZON-[A-Z0-9\-]+:?$", lines[i]) and i + 1 < len(lines):
            fixed_lines.append(f"{lines[i]} {lines[i + 1]}")
            i += 2
        else:
            fixed_lines.append(lines[i])
            i += 1

    topic_pattern = r"^(HORIZON-[A-Za-z0-9\-]+):\s*(.*)$"
    candidate_topics = []
    for i, line in enumerate(fixed_lines):
        match = re.match(topic_pattern, line)
        if match:
            lookahead_text = "\n".join(fixed_lines[i+1:i+20]).lower()
            if any(key in lookahead_text for key in ["call:", "type of action"]):
                candidate_topics.append({
                    "code": match.group(1),
                    "title": match.group(2).strip(),
                    "start_line": i
                })

    topic_blocks = []
    for idx, topic in enumerate(candidate_topics):
        start = topic["start_line"]
        end = candidate_topics[idx + 1]["start_line"] if idx + 1 < len(candidate_topics) else len(fixed_lines)
        for j in range(start + 1, end):
            if fixed_lines[j].lower().startswith("this destination"):
                end = j
                break
        topic_blocks.append({
            "code": topic["code"],
            "title": topic["title"],
            "full_text": "\n".join(fixed_lines[start:end]).strip()
        })

    return topic_blocks

# ========== Field Extraction ==========
def extract_data_fields(topic: Dict[str, Any]) -> Dict[str, Any]:
    text = normalize_text(topic["full_text"])

    def extract_budget(txt: str):
        m = re.search(r"around\s+eur\s+([\d.,]+)", txt.lower())
        if m:
            return int(float(m.group(1).replace(",", "")) * 1_000_000)
        m = re.search(r"between\s+eur\s+[\d.,]+\s+and\s+([\d.,]+)", txt.lower())
        if m:
            return int(float(m.group(1).replace(",", "")) * 1_000_000)
        return None

    def extract_total_budget(txt: str):
        m = re.search(r"indicative budget.*?eur\s?([\d.,]+)", txt.lower())
        return int(float(m.group(1).replace(",", "")) * 1_000_000) if m else None

    def get_section(keyword: str, stop_keywords: List[str]):
        lines = text.splitlines()
        collecting = False
        section = []
        for line in lines:
            l = line.lower()
            if not collecting and keyword in l:
                collecting = True
                section.append(line.split(":", 1)[-1].strip())
            elif collecting and any(l.startswith(k) for k in stop_keywords):
                break
            elif collecting:
                section.append(line)
        return "\n".join(section).strip() if section else None

    def extract_type_of_action(txt: str):
        lines = txt.splitlines()
        for i, line in enumerate(lines):
            if "type of action" in line.lower():
                for j in range(i + 1, len(lines)):
                    if lines[j].strip():
                        return lines[j].strip()
        return None

    def extract_topic_title(txt: str):
        lines = txt.strip().splitlines()
        title_lines = []
        found = False
        for line in lines:
            if not found:
                m = re.match(r"^(HORIZON-[A-Za-z0-9-]+):\s*(.*)", line)
                if m:
                    found = True
                    title_lines.append(m.group(2).strip())
            else:
                if re.match(r"^\s*Call[:\-]", line, re.IGNORECASE):
                    break
                elif line.strip():
                    title_lines.append(line.strip())
        return " ".join(title_lines) if title_lines else None

    def extract_call_name_topic(txt: str):
        txt = normalize_text(txt)
        m = re.search(r"(?i)^\s*Call:\s*(.+)$", txt, re.MULTILINE)
        if m:
            return m.group(1).strip()
        return None

    return {
        "title": extract_topic_title(text),
        "budget_per_project": extract_budget(text),
        "indicative_total_budget": extract_total_budget(text),
        "type_of_action": extract_type_of_action(text),
        "expected_outcome": get_section("expected outcome:", ["scope:", "objective:", "expected impact:", "eligibility:", "budget"]),
        "scope": get_section("scope:", ["objective:", "expected outcome:", "expected impact:", "budget"]),
        "call": extract_call_name_topic(text),
        "trl": (m := re.search(r"TRL\s*(\d+)[^\d]*(\d+)?", text, re.IGNORECASE)) and (
            f"{m.group(1)}-{m.group(2)}" if m.group(2) else m.group(1)
        )
    }

# ========== Metadata (Opening / Deadlines / Destination) ==========
def extract_metadata_blocks(text: str) -> Dict[str, Dict[str, Any]]:
    """
    Keeps single-deadline behaviour but writes into 'deadline1'.
    Also detects two-stage forms like:
      'Deadline(s): 23 Sep 2025 (First Stage), 14 Apr 2026 (Second Stage)'
    and fills deadline1 / deadline2 accordingly.
    """
    lines = normalize_text(text).splitlines()

    metadata_map: Dict[str, Dict[str, Any]] = {}
    current = {
        "opening_date": None,
        "deadline1": None,
        "deadline2": None,
        "destination": None
    }

    topic_pattern = re.compile(r"^(HORIZON-[A-Z0-9\-]+):")

    # Header variants (lenient startswith)
    def is_opening(line: str) -> bool:
        return line.lower().startswith(("opening:", "opening date:", "opens:"))

    def is_deadline(line: str) -> bool:
        return line.lower().startswith(("deadline", "deadlines", "deadline(s):", "cut-off", "cut off"))

    def is_destination(line: str) -> bool:
        return line.lower().startswith("destination")

    # Date patterns
    months = r"Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t|tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?"
    date_wordy = re.compile(rf"\b(\d{{1,2}})\s+({months})\s+(\d{{4}})\b", re.IGNORECASE)   # 23 Sep 2025 / 23 September 2025
    date_iso   = re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b")                                 # 2025-09-23
    date_slash = re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b")                             # 23/09/2025

    def find_dates(s: str) -> List[str]:
        out: List[str] = []
        out += [f"{d} {m} {y}" for d, m, y in date_wordy.findall(s)]
        out += [f"{y}-{m}-{d}" for y, m, d in date_iso.findall(s)]
        for d, m, y in date_slash.findall(s):  # assume dd/mm/yyyy
            out.append(f"{int(y):04d}-{int(m):02d}-{int(d):02d}")
        # de-dup preserve order
        seen = set()
        dedup = []
        for v in out:
            if v not in seen:
                seen.add(v)
                dedup.append(v)
        return dedup

    collecting = False
    for line in lines:
        if is_opening(line):
            dates = find_dates(line)
            current["opening_date"] = dates[0] if dates else None
            collecting = True
            continue

        if is_deadline(line):
            dates = find_dates(line)
            # Keep legacy "one deadline" behaviour as Deadline1
            current["deadline1"] = dates[0] if dates else None
            # If a second date is present (two-stage), capture as Deadline2
            current["deadline2"] = dates[1] if len(dates) > 1 else None
            collecting = True
            continue

        if is_destination(line):
            current["destination"] = line.split(":", 1)[-1].strip()
            collecting = True
            continue

        if collecting:
            m = topic_pattern.match(line)
            if m:
                code = m.group(1)
                metadata_map[code] = {
                    "opening_date": current.get("opening_date"),
                    "deadline1": current.get("deadline1"),
                    "deadline2": current.get("deadline2"),
                    "destination": current.get("destination"),
                }

    return metadata_map

# ========== Public API ==========
def parse_pdf(file_like, *, source_filename: str = "", version_label: str = "Unknown", parsed_on_utc: str = "") -> pd.DataFrame:
    """
    Returns a DataFrame equivalent to your app's output,
    but with 'Deadline1' and 'Deadline2' columns.
    """
    pdf_bytes = file_like.read()
    raw_text = extract_text_from_pdf(BytesIO(pdf_bytes))

    topic_blocks = extract_topic_blocks(raw_text)
    metadata_by_code = extract_metadata_blocks(raw_text)

    enriched = [
        {
            **topic,
            **extract_data_fields(topic),
            **metadata_by_code.get(topic["code"], {})
        }
        for topic in topic_blocks
    ]

    df = pd.DataFrame([{
        "Code": t["code"],
        "Title": t["title"],
        "Opening Date": t.get("opening_date"),
        "Deadline1": t.get("deadline1"),           # ← single or first deadline
        "Deadline2": t.get("deadline2"),           # ← second deadline when present
        "Destination": t.get("destination"),
        "Budget Per Project": t.get("budget_per_project"),
        "Total Budget": t.get("indicative_total_budget"),
        "Number of Projects": int(t["indicative_total_budget"] / t["budget_per_project"])
            if t.get("budget_per_project") and t.get("indicative_total_budget") else None,
        "Type of Action": t.get("type_of_action"),
        "TRL": t.get("trl"),
        "Call Name": t.get("call"),
        "Expected Outcome": t.get("expected_outcome"),
        "Scope": t.get("scope"),
        "Description": t.get("full_text"),
        "Source Filename": source_filename,
        "Version Label": version_label,
        "Parsed On (UTC)": parsed_on_utc,
    } for t in enriched])

    return df
