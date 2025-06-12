# %%
import fitz  # PyMuPDF
import re
import pandas as pd
from tkinter import Tk, filedialog
import os

# ========== File Dialogs ==========
def select_pdf_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    root.destroy()
    return file_path

def save_excel_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
    root.destroy()
    return file_path

# ========== PDF Parsing ==========
def extract_text_from_pdf(path):
    with fitz.open(path) as doc:
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

# ========== Function to find nearest call info ================
def find_nearest_call_info(text):
    text = normalize_text(text)
    call_info_list = []

    # Same date pattern as before
    date_pattern = re.compile(
        r"Opening\s*[:]*\s*(\d{1,2} \w+ \d{4})\s*"
        r"Deadline\(s\)\s*[:]*\s*([^\n]+)", re.IGNORECASE
    )

    for match in date_pattern.finditer(text):
        opening_date = match.group(1).strip()
        deadlines_raw = match.group(2).strip()

        # Extract deadline(s)
        deadlines = re.findall(r"\d{1,2} \w+ \d{4}", deadlines_raw)
        deadline_1 = deadlines[0] if len(deadlines) > 0 else None
        deadline_2 = deadlines[1] if len(deadlines) > 1 else None

        # Text before the current match (search backward)
        preceding_text = text[:match.start()]

        # Stop the extraction of call name when 'HORIZON' or similar terms are encountered
        call_match = list(re.finditer(r"Call\s*[-:]?\s*(.*?)\s*(?=HORIZON|Deadline|Opening|$)", preceding_text))
        call_title = call_match[-1].group(1).strip() if call_match else "Unknown"

        # Get last HORIZON ID before this
        id_match = list(re.finditer(r"(HORIZON-[A-Z]+-\d{4}-[A-Za-z0-9-]+)", preceding_text))
        call_id = id_match[-1].group(1).strip() if id_match else "Unknown"

        call_info_list.append((call_title, call_id, opening_date, deadline_1, deadline_2))

    return pd.DataFrame(call_info_list, columns=[
        "Call Name", "Call ID", "Opening Date", "Deadline 1", "Deadline 2"
    ])


# ========== Pipeline ==========
def run_pipeline():
    pdf_path = select_pdf_file()
    if not pdf_path:
        print("❌ No file selected.")
        return

    raw_text = extract_text_from_pdf(pdf_path)
    topic_blocks = extract_topic_blocks(raw_text)
    print(f"✅ Found {len(topic_blocks)} topic blocks")

    enriched = []
    for topic in topic_blocks:
        enriched.append({**topic, **extract_data_fields(topic)})

    df = pd.DataFrame([{
        "Code": t["code"],
        "Title": t["title"],
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

    calls_with_dates_df = find_nearest_call_info(raw_text)

    # Normalize Call Names before merge
    df["Call Name"] = df["Call Name"].str.strip().str.lower()
    calls_with_dates_df["Call Name"] = calls_with_dates_df["Call Name"].str.strip().str.lower()

    # 🔍 Show sample call names from both dataframes
    print("\n🔍 Sample topic-level Call Names:")
    print(df["Call Name"].dropna().unique()[:5])

    print("\n🔍 Sample date-level Call Names:")
    print(calls_with_dates_df["Call Name"].dropna().unique()[:5])

    # 🧪 Show full unmatched call names (those in df but not in calls_with_dates_df)
    unmatched = df[~df["Call Name"].isin(calls_with_dates_df["Call Name"])]
    print(f"\n❌ {len(unmatched)} Call Names from topic-level data had no date match.")
    print(unmatched["Call Name"].dropna().unique()[:5])

    # 🧪 Optional: also see unmatched in reverse
    unmatched_dates = calls_with_dates_df[~calls_with_dates_df["Call Name"].isin(df["Call Name"])]
    print(f"\n📅 {len(unmatched_dates)} date Call Names had no match in topic-level data.")
    print(unmatched_dates["Call Name"].dropna().unique()[:5])

    # 📅 See how many rows were matched
    df_final = df.merge(calls_with_dates_df, on="Call Name", how="left")
    print(f"\n✅ Successfully merged {df_final['Opening Date'].notnull().sum()} rows with dates.")

    df_final = df.merge(calls_with_dates_df, on="Call Name", how="left")

    output_path = save_excel_file()
    if output_path:
        df_final.to_excel(output_path, index=False)
        print(f"✅ Saved Excel to: {output_path}")
    else:
        print("⚠️ Save cancelled.")

# ========== Run ==========
if __name__ == "__main__":
    run_pipeline()



