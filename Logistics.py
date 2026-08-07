from bs4 import BeautifulSoup
import pandas as pd
import os
import re

# --- 1. FOLDER SETUP ---
main_folder = "Logistics"
os.makedirs(main_folder, exist_ok=True)
nested_folder = os.path.join(main_folder, "Logistics")
os.makedirs(nested_folder, exist_ok=True)

# --- 2. DATA LOADING ---
excel_path = r"C:\\Users\\leesil\\Python Projects\\Git Branch\\Master\\OJTI\\export.xlsx"
data = []

if not os.path.exists(excel_path):
    print(f"❌ Warning: Excel file not found at {excel_path}. Skipping data processing.")
else:
    try:
        df = pd.read_excel(excel_path, header=None)
        data = df.values.tolist()
        print(f"✅ Loaded {len(data)} rows from Excel.")
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        data = []

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
keywords = ["DPLP", "Subcon", "Shipping"]

SKILL_ROLE = {
    1: "Operator/Technician",
    2: "Operator",
    3: "Operator/Technician",
    4: "Line Technician",
    5: "PM Technician",
    6: "Technician",
}


def level_role_label(i: int) -> str:
    return f"Level {i} ({SKILL_ROLE.get(i, '')})"

LEVEL_HEADER = "<tr class='header-row'><th>Level/Role</th><th>Documents</th></tr>"

# Initialize dictionary
keyword_lists = {
    name: {f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)}
    for name in keywords
}

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def make_entry(title: str, doc_id: str) -> str:
    title = str(title).strip()
    doc_id = str(doc_id).strip()
    url = f"https://plmpublishing.icp.infineon.com/api/download-pdf/{doc_id}"
    return f"{title} - <a href='{url}' target='_blank'>{doc_id}</a>"

def detect_skill_level(title: str):
    """
    More robust skill level detection:
    Accepts:
      (Skill Level 2)
      (Skill level 2)
      (Skill Level2)
      (Skill Level 2)(TWI)
      (Skill Level 2) (TWI)
    """
    s = str(title)
    m = re.search(r"\(\s*Skill\s*[Ll]evel\s*([1-6])\s*\)", s)
    return int(m.group(1)) if m else None

def contains_keyword(title: str, keyword: str) -> bool:
    """
    Case-insensitive keyword check.
    Uses word-boundary-ish matching for safety with keywords like 'Subcon'/'Shipping'.
    """
    t = str(title)
    # If you want pure substring matching, replace with: return keyword.lower() in t.lower()
    return bool(re.search(rf"\b{re.escape(keyword)}\b", t, flags=re.IGNORECASE))

# ---------------------------------------------------------------------------
# 3. DATA PROCESSING
# ---------------------------------------------------------------------------
print("🔄 Processing data...")
total_processed = 0

for row in data:
    padded_row = list(row) + [""] * max(0, 2 - len(row))
    title = str(padded_row[0]).strip() if pd.notna(padded_row[0]) else ""
    doc_id = str(padded_row[1]).strip() if pd.notna(padded_row[1]) else ""

    if not title:
        continue

    lvl = detect_skill_level(title)
    if lvl is None:
        continue

    entry = make_entry(title, doc_id)

    for keyword in keywords:
        if contains_keyword(title, keyword):
            bucket = keyword_lists[keyword][f"Skill Level {lvl}"]
            if entry not in bucket["list"]:  # de-dup
                bucket["list"].append(entry)
                bucket["count"] += 1
                total_processed += 1

print(f"✅ Total entries processed: {total_processed}")

# ---------------------------------------------------------------------------
# SHARED STYLE
# ---------------------------------------------------------------------------
SHARED_STYLE = """
body { font-family: Arial, sans-serif; margin: 20px; background-color: #f0f0f0; }
.container { width: 80%; margin: 0 auto; background-color: #fff; padding: 20px; border: 1px solid #ddd; border-radius: 10px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); }
.header { background-color: #f0f0f0; padding: 10px; border-bottom: 1px solid #ddd; display: flex; justify-content: flex-start; align-items: center; }
.home-btn { display: inline-block; padding: 8px 16px; background-color: #08665c; color: white; text-decoration: none; border-radius: 5px; font-size: 14px; transition: background-color 0.3s; }
.home-btn:hover { background-color: #054a40; }
.content { padding: 20px; display: flex; }
.sidebar { width: 20%; background-color: #08665c; color: #fff; padding: 20px; font-size: 24px; text-align: center; border-radius: 10px 0 0 10px; display: flex; justify-content: center; align-items: center; }
.main-content { width: 80%; padding: 20px; }
.post-it-container { display: flex; flex-wrap: wrap; justify-content: center; }
.post-it { width: 150px; height: 150px; background-color: #ADD8E6; padding: 10px; border: 1px solid #ccc; border-radius: 10px; margin: 10px; display: flex; justify-content: center; align-items: center; cursor: pointer; transition: background-color 0.2s ease-in-out; }
.post-it:hover { background-color: #87CEEB; }
.post-it a { text-decoration: none; color: #000; width: 100%; height: 100%; display: flex; justify-content: center; align-items: center; transition: color 0.2s ease-in-out; font-weight: normal; }
.post-it:hover a { color: #fff; }
"""

# ---------------------------------------------------------------------------
# 4. MAIN DASHBOARD (Logistics/Logistics.html)
# ---------------------------------------------------------------------------
html = "<html><head><style>"
html += SHARED_STYLE
html += "</style></head><body>"
html += "<div class='container'>"
html += "<div class='header'><a href='../index.html' class='home-btn'>Return to Home</a></div>"
html += "<div class='content'>"
html += "<div class='sidebar'>Logistics</div>"
html += "<div class='main-content'><div class='post-it-container'>"

for keyword in keywords:
    html += f"<div class='post-it'><a href='Logistics/{keyword}.html'>{keyword}</a></div>"

html += "</div></div></div></div></body></html>"

with open(os.path.join(main_folder, "Logistics.html"), "w", encoding="utf-8") as file:
    print(BeautifulSoup(html, "html.parser").prettify(), file=file)

# ---------------------------------------------------------------------------
# 5. INDIVIDUAL KEYWORD PAGES
# ---------------------------------------------------------------------------
for keyword in keywords:
    keyword_html = "<html><head><style>"
    keyword_html += SHARED_STYLE
    keyword_html += """
    .back-btn { display: inline-block; padding: 10px 20px; background-color: #08665c; color: white; text-decoration: none; border-radius: 5px; font-size: 14px; margin-bottom: 20px; transition: background-color 0.3s; }
    .back-btn:hover { background-color: #054a40; }
    h1 { font-size: 24px; color: #333; margin-bottom: 20px; }
    .data-table { border-collapse: collapse; width: 100%; font-size: 16px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); background-color: #fff; }
    .header-row { background-color: #f0f0f0; color: #333; font-weight: bold; }
    .header-row th { padding: 12px; text-align: left; }
    .skill-level { text-align: left; font-size: 16px; padding: 12px; border-bottom: 1px solid #ddd; white-space: nowrap; vertical-align: top; }
    .data-cell { text-align: left; font-size: 16px; padding: 12px; border-bottom: 1px solid #ddd; line-height: 1.5; }
    .data-cell:hover { background-color: #f0f0f0; }
    """
    keyword_html += "</style></head><body>"
    keyword_html += "<a href='../Logistics.html' class='back-btn'>← Back to Dashboard</a>"
    keyword_html += f"<h1>{keyword}</h1>"
    keyword_html += "<table class='data-table'>"
    keyword_html += LEVEL_HEADER

    has_data = False
    for i in range(1, 7):
        bucket = keyword_lists[keyword][f"Skill Level {i}"]["list"]
        if bucket:
            has_data = True
            # Sort A-Z within each level
            bucket_sorted = sorted(bucket, key=lambda x: x.lower())
            keyword_html += (
                f"<tr>"
                f"<td class='skill-level'>{level_role_label(i)}</td>"
                f"<td class='data-cell'>{'<br/><br/>'.join(bucket_sorted)}</td>"
                f"</tr>"
            )

    if not has_data:
        keyword_html += "<tr><td class='skill-level' colspan='2'>No documents found for this keyword.</td></tr>"

    keyword_html += "</table></body></html>"

    with open(os.path.join(nested_folder, f"{keyword}.html"), "w", encoding="utf-8") as file:
        print(BeautifulSoup(keyword_html, "html.parser").prettify(), file=file)

print("\n✅ Structure Updated to Nested Folders:")
print("📁 Dashboard : Logistics/Logistics.html")
print("📁 Pages     : Logistics/Logistics/[Keyword].html")
print("📋 Column header : Level/Role")
print("📋 Cell format   : Level 1 (Operator) / Level 3 (Technician)")