from bs4 import BeautifulSoup
import pandas as pd
import os
import re

# --- 1. FOLDER SETUP ---
main_folder = "MSP"
os.makedirs(main_folder, exist_ok=True)
nested_folder = os.path.join(main_folder, "MSP")
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
equipment_names = ["KLA", "SRM", "ISMECA", "TTM", "ETM", "Peel Force Tester"]

# Skill Level -> Role mapping
SKILL_ROLE = {
    1: "Operator",
    2: "Operator",
    3: "Technician",
    4: "Technician",
    5: "Technician",
    6: "Technician",
}

def level_role_label(i: int) -> str:
    return f"Level {i} ({SKILL_ROLE.get(i, '')})"

LEVEL_HEADER = "<tr class='header-row'><th>Level/Role</th><th>Documents</th></tr>"

# Initialize dictionary
equipment_skill_lists = {
    equipment: {f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)}
    for equipment in equipment_names
}

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def detect_skill_level(title: str):
    """Robust detection: (Skill Level 2), (Skill level 2), (Skill Level2), etc."""
    m = re.search(r"\(\s*Skill\s*[Ll]evel\s*([1-6])\s*\)", str(title))
    return int(m.group(1)) if m else None

def make_entry(value, link):
    value = str(value).strip()
    link = str(link).strip()
    return (
        f"{value} - "
        f"<a href='https://plmpublishing.icp.infineon.com/api/download-pdf/{link}' target='_blank'>{link}</a>"
    )

def matches_equipment(title: str, equipment: str) -> bool:
    """
    Whole-word matching for most equipment.
    Special case: 'Peel Force Tester' should also match 'peel'.
    """
    t = str(title)

    if equipment.lower() == "peel force tester":
        # Match either phrase "peel force tester" OR the word "peel"
        return bool(re.search(r"\bpeel\b", t, flags=re.IGNORECASE) or
                    re.search(r"\bpeel\s+force\s+tester\b", t, flags=re.IGNORECASE))

    # Default whole-word match for other equipment (case-insensitive)
    return bool(re.search(rf"\b{re.escape(equipment)}\b", t, flags=re.IGNORECASE))

# ---------------------------------------------------------------------------
# DATA PROCESSING
# ---------------------------------------------------------------------------
print("🔄 Processing data with Whole Word Logic...")
total_processed = 0

for row in data:
    if len(row) < 2:
        continue

    value = str(row[0]).strip() if pd.notna(row[0]) else ""
    link  = str(row[1]).strip() if pd.notna(row[1]) else ""
    if not value:
        continue

    lvl = detect_skill_level(value)
    if lvl is None:
        continue

    for equipment in equipment_names:
        if matches_equipment(value, equipment):
            entry = make_entry(value, link)
            bucket = equipment_skill_lists[equipment][f"Skill Level {lvl}"]
            if entry not in bucket["list"]:  # de-dup
                bucket["list"].append(entry)
                bucket["count"] += 1
                total_processed += 1

print(f"✅ Total entries processed: {total_processed}")

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
dashboard_css = """
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

detail_css = """
body { font-family: Arial, sans-serif; margin: 20px; background-color: #f0f0f0; }
.back-btn { display: inline-block; padding: 10px 20px; background-color: #08665c; color: white; text-decoration: none; border-radius: 5px; font-size: 14px; margin-bottom: 20px; transition: background-color 0.3s; }
.back-btn:hover { background-color: #054a40; }
.data-table { border-collapse: collapse; width: 100%; font-size: 18px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); background-color: #fff; }
.header-row { background-color: #f0f0f0; color: #333; font-weight: bold; }
.header-row th { padding: 12px; text-align: left; }
.skill-level { text-align: left; font-size: 16px; padding: 12px; border-bottom: 1px solid #ddd; white-space: nowrap; vertical-align: top; }
.data-cell { text-align: left; font-size: 16px; padding: 12px; border-bottom: 1px solid #ddd; line-height: 1.5; }
.data-cell:hover { background-color: #f0f0f0; }
"""

# --- 3. MAIN DASHBOARD (MSP/MSP.html) ---
html = f"<html><head><style>{dashboard_css}</style></head><body>"
html += "<div class='container'>"
html += "<div class='header'><a href='../index.html' class='home-btn'>Return to Home</a></div>"
html += "<div class='content'>"
html += "<div class='sidebar'>MSP</div>"
html += "<div class='main-content'><div class='post-it-container'>"

for equipment in equipment_names:
    safe_filename = equipment.replace(" ", "_") + ".html"
    html += f"<div class='post-it'><a href='MSP/{safe_filename}'>{equipment}</a></div>"

html += "</div></div></div></div></body></html>"

with open(os.path.join(main_folder, "MSP.html"), "w", encoding="utf-8") as file:
    print(BeautifulSoup(html, "html.parser").prettify(), file=file)

# --- 4. INDIVIDUAL EQUIPMENT PAGES ---
for equipment in equipment_names:
    safe_filename = equipment.replace(" ", "_") + ".html"

    equipment_html = f"<html><head><style>{detail_css}</style></head><body>"
    equipment_html += "<a href='../MSP.html' class='back-btn'>← Back to Dashboard</a>"
    equipment_html += f"<h1>{equipment}</h1>"
    equipment_html += "<table class='data-table'>"
    equipment_html += LEVEL_HEADER

    has_data = False
    for i in range(1, 7):
        bucket = equipment_skill_lists[equipment][f"Skill Level {i}"]["list"]
        if bucket:
            has_data = True
            bucket_sorted = sorted(bucket, key=lambda x: x.lower())
            content = "<br/><br/>".join(bucket_sorted)
            equipment_html += (
                f"<tr>"
                f"<td class='skill-level'>{level_role_label(i)}</td>"
                f"<td class='data-cell'>{content}</td>"
                f"</tr>"
            )

    if not has_data:
        equipment_html += (
            "<tr><td class='skill-level' colspan='2'>"
            "No documents found for this equipment."
            "</td></tr>"
        )

    equipment_html += "</table></body></html>"

    with open(os.path.join(nested_folder, safe_filename), "w", encoding="utf-8") as file:
        print(BeautifulSoup(equipment_html, "html.parser").prettify(), file=file)

print("\n✅ Structure Updated to Nested Folders:")
print("📁 Dashboard : MSP/MSP.html")
print("📁 Pages     : MSP/MSP/[Equipment].html")
print("📋 Column header : Level/Role")
print("📋 Cell format   : Level 1 (Operator) / Level 3 (Technician)")
print("📋 Peel Force Tester matching: 'peel' OR 'peel force tester'")