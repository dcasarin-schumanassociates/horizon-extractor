import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Horizon Topic Extractor", layout="centered")
st.title("📄 Horizon Topic Extractor")
st.write("Upload a Horizon Europe PDF file and get an Excel sheet with parsed topics.")

# ========== File Upload ==========
uploaded_file = st.file_uploader("Upload a Horizon PDF", type=["pdf"])

# ========== PDF Parsing ==========
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

    def extract_topic_title(text):
        lines = text.strip().splitlines()
        title_lines = []
        found = False
        for line in lines:
            if not found:
                match = re.match(r"^(HORIZON-[A-Za-z0-9-]+):\s*(.*)", line)
                if match:
                    found = True
                    title_lines.append(match.group(2).strip())
            else:
                if re.match(r"^\s*Call[:\-]", line, re.IGNORECASE):
                    break
                elif line.strip():
                    title_lines.append(line.strip())
        return " ".join(title_lines) if title_lines else None

    def extract_call_name_topic(text):
        text = normalize_text(text)
        match = re.search(r"(?i)^\s*Call:\s*(.+)$", text, re.MULTILINE)
        if match:
            return match.group(1).strip()
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

def extract_metadata_blocks(text):
    lines = normalize_text(text).splitlines()

    metadata_map = {}
    current_metadata = {
        "opening_date": None,
        "deadline": None,
        "destination": None
    }

    topic_pattern = re.compile(r"^(HORIZON-[A-Z0-9\-]+):")

    collecting = False
    for i, line in enumerate(lines):
        lower = line.lower()

        if lower.startswith("opening:"):
            current_metadata["opening_date"] = re.search(r"(\d{1,2} \w+ \d{4})", line)
            current_metadata["opening_date"] = (
                current_metadata["opening_date"].group(1)
                if current_metadata["opening_date"]
                else None
            )
            current_metadata["deadline"] = None
            collecting = True

        elif collecting and lower.startswith("deadline"):
            current_metadata["deadline"] = re.search(r"(\d{1,2} \w+ \d{4})", line)
            current_metadata["deadline"] = (
                current_metadata["deadline"].group(1)
                if current_metadata["deadline"]
                else None
            )

        elif collecting and lower.startswith("destination"):
            current_metadata["destination"] = line.split(":", 1)[-1].strip()

        elif collecting:
            match = topic_pattern.match(line)
            if match:
                code = match.group(1)
                metadata_map[code] = current_metadata.copy()

    return metadata_map

# ========== Main Streamlit App ==========
if uploaded_file:
    raw_text = extract_text_from_pdf(uploaded_file)

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
        "Deadline": t.get("deadline"),
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
        "Description": t.get("full_text")
    } for t in enriched])

    st.subheader("📊 Preview of Extracted Topics")
    st.dataframe(df.drop(columns=["Description"]).head(10), use_container_width=True)

    # ========== 🔍 New Word Search ==========
    st.subheader("🔍 Search Topics by Keyword")
    keyword = st.text_input("Enter keyword to filter topics:")

    if keyword:
        keyword = keyword.lower()
        filtered_df = df[df.apply(lambda row: row.astype(str).str.lower().str.contains(keyword).any(), axis=1)]
        filtered_df = filtered_df.drop_duplicates()

        st.markdown(f"**Results containing keyword: `{keyword}`**")
        st.dataframe(filtered_df.drop(columns=["Description"]), use_container_width=True)
        st.write(f"🔎 Found {len(filtered_df)} matching topics.")

    # ========== Download ==========
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    st.success(f"✅ Extracted {len(df)} topics!")
    st.download_button(
        label="⬇️ Download Excel File",
        data=output,
        file_name="horizon_topics.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

