# parsers/horizon.py
from __future__ import annotations
import re
from io import BytesIO
from typing import Dict, Any, List, Optional
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
        m = re.match(topic_pattern, line)
        if m:
            lookahead_text = "\n".join(fixed_lines[i+1:i+20]).lower()
            if any(key in lookahead_text for key in ["call:", "type of action"]):
                candidate_topics.append({
                    "code": m.group(1),
                    "title": m.group(2).strip(),
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

    def extract_budget(txt: str) -> Optional[int]:
        m = re.search(r"around\s+eur\s+([\d.,]+)", txt.lower())
        if m:
            return int(float(m.group(1).replace(",", "")) * 1_000_000)
        m = re.search(r"between\s+eur\s+[\d.,]+\s+and\s+([\d.,]+)", txt.lower())
        if m:
            return int(float(m.group(1).replace(",", "")) * 1_000_000)
        return None

    def extract_total_budget(txt: str) -> Optional[int]:
        m = re.search(r"indicative budget.*?eur\s?([\d.,]+)", txt.lower())
        return int(float(m.group(1).replace(",", "")) * 1_000_000) if m else None

    def get_section(keyword: str, stop_keywords: List[str]) -> Optional[str]:
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

    def extract_type_of_action(txt: str) -> Optional[str]:
        lines = txt.splitlines()
        for i, line in enumerate(lines):
            if "type of action" in line.lower():
                for j in range(i + 1, len(lines)):
                    if lines[j].strip():
                        return lines[j].strip()
        return None

    def extract_topic_title(txt: str) -> Optional[str]:
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

    def extract_call_name_topic(txt: str) -> Optional[str]:
        txt = normalize_text(txt)
        m = re.search(r"(?i)^\s*Call:\s*(.+)$", txt, re.MULTILINE)
        return m.group(1).strip() if m else None

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

# ========== Metadata (Opening + multi-deadline) ==========
_DATE_LABEL_PAIR = re.compile(
    r"(\d{1,2}\s+\w+\s+\d{4})"          # date, e.g., 23 Sep 2025
    r"(?:\s*\(([^)]+)\))?"              # optional label in parentheses, e.g., (First Stage)
)

def _parse_deadlines(line: str) -> Dict[str, Optional[str]]:
    """
    Handles:
      'Deadline: 15 Oct 2025'
      'Deadlines: 23 Sep 2025 (First Stage), 14 Apr 2026 (Second Stage)'
      'Deadline(s): ...'
    Returns dict with deadline1, deadline1_label, deadline2, deadline2_label (strings as found).
    """
    results = _DATE_LABEL_PAIR.findall(line)
    d1 = l1 = d2 = l2 = None
    if results:
        if len(results) >= 1:
            d1, l1 = results[0][0], (results[0][1] or None)
        if len(results) >= 2:
            d2, l2 = results[1][0], (results[1][1] or None)
    return {
        "deadline1": d1,
        "deadline1_label": l1,
        "deadline2": d2,
        "deadline2_label": l2,
    }

def extract_metadata_blocks(text: str) -> Dict[str, Dict[str, Any]]:
    lines = normalize_text(text).splitlines()

    metadata_map: Dict[str, Dict[str, Any]] = {}
    current = {
        "opening_date": None,
        "deadline1": None,
        "deadline1_label": None,
        "deadline2": None,
        "deadline2_label": None,
        "destination": None
    }

    topic_pattern = re.compile(r"^(HORIZON-[A-Z0-9\-]+):")
    collecting = False
    for line in lines:
        lower = line.lower()

        if lower.startswith("opening:"):
            m = re.search(r"(\d{1,2} \w+ \d{4})", line)
            current["opening_date"] = m.group(1) if m else None
            # reset deadlines when we see a new 'Opening:'
            current["deadline1"] = current["deadline1_label"] = None
            current["deadline2"] = current["deadline2_label"] = None
            collecting = True

        elif collecting and lower.startswith("deadline"):
            # covers 'deadline', 'deadlines', 'deadline(s)'
            dl = _parse_deadlines(line)
            # fill first, then second if present
            current.update(dl)

        elif collecting and lower.startswith("destination"):
            current["destination"] = line.split(":", 1)[-1].strip()

        elif collecting:
            match = topic_pattern.match(line)
            if match:
                code = match.group(1)
                metadata_map[code] = current.copy()

    return metadata_map

# ========== Public API ==========
def parse_pdf(file_like, *, source_filename: str = "", version_label: str = "Unknown", parsed_on_utc: str = "") -> pd.DataFrame:
    """
    Returns a DataFrame like your current output, but with:
      - 'Deadline 1' and 'Deadline 2' columns (labels optional)
      - Legacy 'Deadline' kept (equal to 'Deadline 1') for compatibility
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

    rows = []
    for t in enriched:
        deadline1 = t.get("deadline1")
        deadline2 = t.get("deadline2")
        rows.append({
            "Code": t["code"],
            "Title": t["title"],
            "Opening Date": t.get("opening_date"),
            # NEW multi-stage fields
            "Deadline 1": deadline1,
            "Deadline 1 Label": t.get("deadline1_label"),
            "Deadline 2": deadline2,
            "Deadline 2 Label": t.get("deadline2_label"),
            # Legacy single deadline (kept = Deadline 1)
            "Deadline": deadline1,
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
        })

    df = pd.DataFrame(rows)

    return df

