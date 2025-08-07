import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Horizon Topic Extractor", layout="wide")
st.title("📄 Horizon Europe Topic Extractor Dashboard")

# ========== File Upload ==========
uploaded_file = st.file_uploader("Upload a Horizon PDF", type=["pdf"])

# ========== PDF Text Extraction ==========
def extract_text_from_pdf(file):
    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        return "\n".join(page.get_text() for page in doc)

# ========== Utility ==========
def normalize_text(text):
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = re.sub(r"\xa0", " ", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()

# ========== Topic Extraction ==========
def extract_topic_blocks(text):
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
def extract_data_fields(topic):
    text = normalize_text(topic["full_text"])

    def extract_budget(text):
        match = re.search(r"around\s+eur\s+([\d.,]+)", text.lower())
        if match:
            return int(float(match.group(1).replace(",", "")) * 1_000_000)
        match = re.search(r"between\s+eur\s+[\d.,]+\s+and\s+([\d.,]+)", text.lower())
        if match:
            return int(float(match.group(1).replace(",", "")) * 1_000_000)
        return None

    def extract_total_budget(text):
        match = re.search(r"indicative budget.*?eur\s?([\d.,]+)", text.lower())
        return int(float(match.group(1).replace(",", "")) * 1_000_000) if match else None

    def get_section(keyword, stop_keywords):
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

    def extract_type_of_action(text):
        lines = text.splitlines()
        for i, line in enumerate(lines):
            if "type of action" in line.lower():
                for j in range(i + 1, len(lines)):
                    if lines[j].strip():
                        return lines[j].strip()
        return None

    def extract_call_name_topic(text):
        text = normalize_text(text)
        match = re.search(r"(?i)^\s*Call:\s*(.+)$", text, re.MULTILINE)
        if match:
            return match.group(1).strip()
        return None

    def extract_trl(text):
        match = re.search(r"TRL\s*(\d+)[^\d]*(\d+)?", text, re.IGNORECASE)
        if match:
            return f"{match.group(1)}-{match.group(2)}" if match.group(2) else match.group(1)
        return None

    def clean_text_section(text):
        if not text:
            return None
        text = re.sub(r"\n{2,}", "\n", text)
        text = re.sub(r" {2,}", " ", text)
        text = text.strip().replace("\n", "  \n")  # for markdown line breaks
        return text

    return {
        "budget_per_project": extract_budget(text),
        "indicative_total_budget": extract_total_budget(text),
        "type_of_action": extract_type_of_action(text),
        "expected_outcome": clean_text_section(get_section("expected outcome:", ["scope:", "objective:", "expected impact:", "eligibility:", "budget"])),
        "scope": clean_text_section(get_section("scope:", ["objective:", "expected outcome:", "expected impact:", "budget"])),
        "call": extract_call_name_topic(text),
        "trl": extract_trl(text),
        "full_text": clean_text_section(topic["full_text"])
    }

def extract_metadata_blocks(text):
    lines = normalize_text(text).splitlines()
    metadata_map = {}
    current_metadata = {"opening_date": None, "deadline": None, "destination": None}
    topic_pattern = re.compile(r"^(HORIZON-[A-Z0-9\-]+):")
    collecting = False

    for i, line in enumerate(lines):
        lower = line.lower()

        if lower.startswith("opening:"):
            match = re.search(r"(\d{1,2} \w+ \d{4})", line)
            current_metadata["opening_date"] = match.group(1) if match else None
            collecting = True

        elif collecting and lower.startswith("deadline"):
            match = re.search(r"(\d{1,2} \w+ \d{4})", line)
            current_metadata["deadline"] = match.group(1) if match else None

        elif collecting and lower.startswith("destination"):
            current_metadata["destination"] = line.split(":", 1)[-1].strip()

        elif collecting:
            match = topic_pattern.match(line)
            if match:
                code = match.group(1)
                metadata_map[code] = current_metadata.copy()

    return metadata_map

# ========== Main ==========
if uploaded_file:
    raw_text = extract_text_from_pdf(uploaded_file)
    topic_blocks = extract_topic_blocks(raw_text)
    metadata_by_code = extract_metadata_blocks(raw_text)

    enriched = []
    for topic in topic_blocks:
        data_fields = extract_data_fields(topic)
        enriched.append({
            "code": topic["code"],
            "title": topic["title"],
            **data_fields,
            **metadata_by_code.get**

