from bs4 import BeautifulSoup
import pandas as pd
import os
import re

# --- 1. FOLDER SETUP ---
main_folder = "Packing"
os.makedirs(main_folder, exist_ok=True)
nested_folder = os.path.join(main_folder, "Packing")
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

# --- 3. DEFINITIONS & EXCLUSION LISTS ---
packing_equipment_names = ["DPL", "Despatch"]
ibis_names         = ["MGT", "Mirae", "Gen2", "Gen 3", "Gen 32"]
keywords           = ["DPLP", "Subcon", "Shipping"]
equipment_names    = ["KLA", "SRM", "ISMECA", "TTM", "ETM",  "Peel Force Tester"]
automation_names   = ["OHT", "AMR", "E-rack", "Strapping", "ASRS", "AMHS", "MMR", "Point to Point"]
tester_names       = ["V93K", "LTX", "MMCI", "Advantest", "Ultraflex", "Rasco", "EXAScale", "J750"]
handler_names      = ["North Star", "Delta Matrix", "Delta Castle", "OSAI", "Multitest", "JHT"]

all_exclusion_keywords = (
    packing_equipment_names +
    ibis_names +
    keywords +
    equipment_names +
    automation_names +
    tester_names +
    handler_names
)

# Initialize dictionaries
packing_equipment_skill_lists = {
    equipment: {f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)}
    for equipment in packing_equipment_names
}

# CHANGED: other_packing_skill_lists -> general_skill_lists
general_skill_lists = {
    f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)
}

# Track links already matched to specific equipment
matched_links = set()

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

print("🔄 Processing data...")
total_processed   = 0
debug_other_added = 0

for row in data:
    if len(row) < 2:
        continue

    value = str(row[0]) if pd.notna(row[0]) else ""
    link  = str(row[1]) if pd.notna(row[1]) else ""
    col6  = str(row[5]) if (len(row) > 5 and pd.notna(row[5])) else ""

    if not value.strip() or not link.strip():
        continue

    value_lower = value.lower()
    matched_specific = False

    # 1. CHECK FOR SPECIFIC EQUIPMENT (DPL) - whole word match
    for equipment in packing_equipment_names:
        pattern = r'\b' + re.escape(equipment.lower()) + r'\b'
        if re.search(pattern, value_lower):
            matched_specific = True
            matched_links.add(link)
            for i in range(1, 7):
                if f"(Skill Level {i})" in value:
                    entry = (
                        f"{value} - "
                        f"<a href='https://plmpublishing.icp.infineon.com/api/download-pdf/{link}'"
                        f" target='_blank'>{link}</a>"
                    )
                    packing_equipment_skill_lists[equipment][f"Skill Level {i}"]["list"].append(entry)
                    packing_equipment_skill_lists[equipment][f"Skill Level {i}"]["count"] += 1
                    total_processed += 1
            break

    if matched_specific:
        continue

    # 2. GENERAL (was Other Packing)
    #    col6 must contain "packing" AND not already matched to specific equipment
    if "packing" in col6.lower() and link not in matched_links:
        for i in range(1, 7):
            if f"(Skill Level {i})" in value:
                entry = (
                    f"{value} - "
                    f"<a href='https://plmpublishing.icp.infineon.com/api/download-pdf/{link}'"
                    f" target='_blank'>{link}</a>"
                )
                general_skill_lists[f"Skill Level {i}"]["list"].append(entry)
                general_skill_lists[f"Skill Level {i}"]["count"] += 1
                total_processed += 1
                debug_other_added += 1

# CHANGED: total_other_packing -> total_general
total_general = sum(general_skill_lists[f"Skill Level {i}"]["count"] for i in range(1, 7))

print(f"✅ Total entries processed : {total_processed}")
print(f"📌 General total           : {total_general}")
print(f"➕ Added to General         : {debug_other_added}")

# --- 4. CSS ---
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
.back-btn { display: inline-block; padding: 10px 20px; background-color: #08665c; color: white; text-decoration: none; border-radius: 5px; font-size: 14px; margin-bottom: 20px; }
.back-btn:hover { background-color: #054a40; }
.data-table { border-collapse: collapse; width: 100%; font-size: 18px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); background-color: #fff; }
.header-row { background-color: #f0f0f0; color: #333; font-weight: bold; }
.header-row th { padding: 12px; text-align: left; }
.skill-level { text-align: left; font-size: 16px; padding: 12px; border-bottom: 1px solid #ddd; }
.data-cell { text-align: left; font-size: 18px; padding: 12px; border-bottom: 1px solid #ddd; line-height: 1.5; }
.data-cell.empty { background-color: #cccccc; }
.data-cell:hover { background-color: #f0f0f0; }
"""

# --- 5. MAIN DASHBOARD (Packing/Packing.html) ---
# CHANGED: "Other Packing" tile -> "General" tile
html  = f"<html><head><style>{dashboard_css}</style></head><body>"
html += "<div class='container'>"
html += "<div class='header'><a href='../index.html' class='home-btn'>Return to Home</a></div>"
html += "<div class='content'><div class='sidebar'>Packing</div>"
html += "<div class='main-content'><div class='post-it-container'>"

for equipment in packing_equipment_names:
    html += f"<div class='post-it'><a href='Packing/{equipment}.html'>{equipment}</a></div>"

# CHANGED: Other Packing -> General, links to General.html
html += "<div class='post-it'><a href='Packing/General.html'>General</a></div>"

html += "</div></div></div></div></body></html>"

with open(os.path.join(main_folder, "Packing.html"), "w", encoding="utf-8") as file:
    print(BeautifulSoup(html, 'html.parser').prettify(), file=file)
print(f"✅ Written: Packing/Packing.html")

# --- 6. INDIVIDUAL PACKING EQUIPMENT PAGES ---
for equipment in packing_equipment_names:
    eq_html  = f"<html><head><style>{detail_css}</style></head><body>"
    eq_html += f"<a href='../Packing.html' class='back-btn'>← Back to Dashboard</a>"
    eq_html += f"<h1>{equipment}</h1>"
    eq_html += f"<table class='data-table'>{LEVEL_HEADER}"

    has_data = False
    for i in range(1, 7):
        if packing_equipment_skill_lists[equipment][f"Skill Level {i}"]["count"] > 0:
            has_data = True
            content = '<br/><br/>'.join(packing_equipment_skill_lists[equipment][f"Skill Level {i}"]["list"])
            eq_html += (
                f"<tr>"
                f"<td class='skill-level'>{level_role_label(i)}</td>"
                f"<td class='data-cell'>{content}</td>"
                f"</tr>"
            )

    if not has_data:
        eq_html += "<tr><td class='skill-level' colspan='2'>No documents found for this equipment.</td></tr>"

    eq_html += "</table></body></html>"

    with open(os.path.join(nested_folder, f"{equipment}.html"), "w", encoding="utf-8") as file:
        print(BeautifulSoup(eq_html, 'html.parser').prettify(), file=file)

print("✅ Written: Equipment detail pages")

# --- 7. GENERAL PAGE (was Other Packing) ---
general_html  = f"<html><head><style>{detail_css}</style></head><body>"
general_html += "<a href='../Packing.html' class='back-btn'>← Back to Dashboard</a>"
general_html += "<h1>General</h1>"
general_html += f"<table class='data-table'>{LEVEL_HEADER}"

has_data = False
for i in range(1, 7):
    if general_skill_lists[f"Skill Level {i}"]["count"] > 0:
        has_data = True
        content = '<br/><br/>'.join(general_skill_lists[f"Skill Level {i}"]["list"])
        general_html += (
            f"<tr>"
            f"<td class='skill-level'>{level_role_label(i)}</td>"
            f"<td class='data-cell'>{content}</td>"
            f"</tr>"
        )

if not has_data:
    general_html += "<tr><td class='skill-level' colspan='2'>No documents found.</td></tr>"

general_html += "</table></body></html>"

# CHANGED: saved as General.html
with open(os.path.join(nested_folder, "General.html"), "w", encoding="utf-8") as file:
    print(BeautifulSoup(general_html, 'html.parser').prettify(), file=file)

print("✅ Written: Packing/Packing/General.html")
print("\n✅ Done.")
print("\n📋 Logic Summary:")
print("   DPL     : whole word match in col1, NO column filter")
print("   General : col6 contains 'packing' + NOT already matched to DPL")
print("📋 Column header : Level/Role")
print("📋 Cell format   : Level 1 (Operator) / Level 3 (Technician)")