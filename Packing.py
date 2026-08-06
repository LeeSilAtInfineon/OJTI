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
packing_equipment_names = ["DPL", "MVH"]

ibis_names         = ["MGT", "Mirae", "Gen2", "Gen 3", "Gen 32"]
keywords           = ["DPLP", "Subcon", "Shipping"]
equipment_names = ["KLA", "SRM", "ISMECA", "TTM", "ETM", "Peel", "KLR", "McDry", "Despatch"]
automation_names   = ["OHT", "AMR", "E-rack", "Strapping", "ASRS", "AMHS", "MMR", "Point to Point"]
tester_names       = ["V93K", "LTX", "MMCI", "Advantest", "Ultraflex", "Rasco", "EXAScale", "J750"]
handler_names      = ["North Star", "Delta Matrix", "Delta Castle", "OSAI", "Multitest", "JHT"]

# Exclude these from Packing/General (in addition to the specific-equipment match)
# NOTE: We exclude packing_equipment_names themselves so DPL/Despatch docs don't double-appear in General.
GENERAL_PACKING_EXCLUDE_TERMS = (
    packing_equipment_names +
    ibis_names +
    keywords +
    equipment_names +
    automation_names +
    tester_names +
    handler_names
)

def contains_any_excluded_term(title: str) -> bool:
    """
    Case-insensitive substring match against exclusion terms.
    If True, do NOT include in Packing/General.
    """
    t = str(title).lower()
    return any(term.lower() in t for term in GENERAL_PACKING_EXCLUDE_TERMS)

# Initialize dictionaries
packing_equipment_skill_lists = {
    equipment: {f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)}
    for equipment in packing_equipment_names
}

general_skill_lists = {f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)}

matched_links = set()

# Skill Level -> Role mapping
SKILL_ROLE = {
    1: "Operator",
    2: "Operator",
    3: "Operator/Technician",
    4: "Line Technician",
    5: "PM Technician",
    6: "Technician",
}


def level_role_label(i: int) -> str:
    return f"Level {i} ({SKILL_ROLE.get(i, '')})"

LEVEL_HEADER = "<tr class='header-row'><th>Level/Role</th><th>Documents</th></tr>"

def detect_skill_level(title: str):
    """Robust detection: handles (Skill Level 2), (Skill level 2), (Skill Level2), etc."""
    m = re.search(r"\(\s*Skill\s*[Ll]evel\s*([1-6])\s*\)", str(title))
    return int(m.group(1)) if m else None

def make_entry(title: str, link: str) -> str:
    title = str(title).strip()
    link = str(link).strip()
    return (
        f"{title} - "
        f"<a href='https://plmpublishing.icp.infineon.com/api/download-pdf/{link}' target='_blank'>{link}</a>"
    )

print("🔄 Processing data...")

total_processed = 0
debug_general_added = 0
debug_general_skipped_exclusion = 0

for row in data:
    padded_row = list(row) + [""] * max(0, 6 - len(row))
    value = str(padded_row[0]).strip() if pd.notna(padded_row[0]) else ""
    link  = str(padded_row[1]).strip() if pd.notna(padded_row[1]) else ""
    col6  = str(padded_row[5]).strip() if pd.notna(padded_row[5]) else ""

    if not value or not link:
        continue

    lvl = detect_skill_level(value)
    if lvl is None:
        continue

    value_lower = value.lower()
    entry = make_entry(value, link)

    # 1) CHECK FOR SPECIFIC EQUIPMENT (DPL, Despatch) - whole word match
    matched_specific = False
    for equipment in packing_equipment_names:
        pattern = r"\b" + re.escape(equipment.lower()) + r"\b"
        if re.search(pattern, value_lower):
            matched_specific = True
            matched_links.add(link)

            bucket = packing_equipment_skill_lists[equipment][f"Skill Level {lvl}"]
            if entry not in bucket["list"]:  # de-dup
                bucket["list"].append(entry)
                bucket["count"] += 1
                total_processed += 1
            break

    if matched_specific:
        continue

    # 2) GENERAL: col6 must contain "packing" AND not already matched
    #    PLUS: exclude if contains any terms from the other lists (ibis/tester/handler/etc.)
    if "packing" in col6.lower() and link not in matched_links:
        if contains_any_excluded_term(value):
            debug_general_skipped_exclusion += 1
            continue

        bucket = general_skill_lists[f"Skill Level {lvl}"]
        if entry not in bucket["list"]:  # de-dup
            bucket["list"].append(entry)
            bucket["count"] += 1
            total_processed += 1
            debug_general_added += 1

total_general = sum(general_skill_lists[f"Skill Level {i}"]["count"] for i in range(1, 7))

print(f"✅ Total entries processed : {total_processed}")
print(f"📌 General total           : {total_general}")
print(f"➕ Added to General         : {debug_general_added}")
print(f"🚫 Skipped from General (exclusion match): {debug_general_skipped_exclusion}")

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
.skill-level { text-align: left; font-size: 16px; padding: 12px; border-bottom: 1px solid #ddd; white-space: nowrap; vertical-align: top; }
.data-cell { text-align: left; font-size: 18px; padding: 12px; border-bottom: 1px solid #ddd; line-height: 1.5; }
.data-cell.empty { background-color: #cccccc; }
.data-cell:hover { background-color: #f0f0f0; }
"""

# --- 5. MAIN DASHBOARD (Packing/Packing.html) ---
html  = f"<html><head><style>{dashboard_css}</style></head><body>"
html += "<div class='container'>"
html += "<div class='header'><a href='../index.html' class='home-btn'>Return to Home</a></div>"
html += "<div class='content'><div class='sidebar'>Packing</div>"
html += "<div class='main-content'><div class='post-it-container'>"

for equipment in packing_equipment_names:
    html += f"<div class='post-it'><a href='Packing/{equipment}.html'>{equipment}</a></div>"

html += "<div class='post-it'><a href='Packing/General.html'>General</a></div>"
html += "</div></div></div></div></body></html>"

with open(os.path.join(main_folder, "Packing.html"), "w", encoding="utf-8") as file:
    print(BeautifulSoup(html, "html.parser").prettify(), file=file)

print("✅ Written: Packing/Packing.html")

# --- 6. INDIVIDUAL PACKING EQUIPMENT PAGES ---
for equipment in packing_equipment_names:
    eq_html  = f"<html><head><style>{detail_css}</style></head><body>"
    eq_html += f"<a href='../Packing.html' class='back-btn'>← Back to Dashboard</a>"
    eq_html += f"<h1>{equipment}</h1>"
    eq_html += f"<table class='data-table'>{LEVEL_HEADER}"

    has_data = False
    for i in range(1, 7):
        bucket = packing_equipment_skill_lists[equipment][f"Skill Level {i}"]["list"]
        if bucket:
            has_data = True
            bucket_sorted = sorted(bucket, key=lambda x: x.lower())
            content = "<br/><br/>".join(bucket_sorted)
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
        print(BeautifulSoup(eq_html, "html.parser").prettify(), file=file)

print("✅ Written: Equipment detail pages")

# --- 7. GENERAL PAGE ---
general_html  = f"<html><head><style>{detail_css}</style></head><body>"
general_html += "<a href='../Packing.html' class='back-btn'>← Back to Dashboard</a>"
general_html += "<h1>General</h1>"
general_html += f"<table class='data-table'>{LEVEL_HEADER}"

has_data = False
for i in range(1, 7):
    bucket = general_skill_lists[f"Skill Level {i}"]["list"]
    if bucket:
        has_data = True
        bucket_sorted = sorted(bucket, key=lambda x: x.lower())
        content = "<br/><br/>".join(bucket_sorted)
        general_html += (
            f"<tr>"
            f"<td class='skill-level'>{level_role_label(i)}</td>"
            f"<td class='data-cell'>{content}</td>"
            f"</tr>"
        )

if not has_data:
    general_html += "<tr><td class='skill-level' colspan='2'>No documents found.</td></tr>"

general_html += "</table></body></html>"

with open(os.path.join(nested_folder, "General.html"), "w", encoding="utf-8") as file:
    print(BeautifulSoup(general_html, "html.parser").prettify(), file=file)

print("✅ Written: Packing/Packing/General.html")

print("\n✅ Done.")
print("\n📋 Logic Summary:")
print("   DPL/Despatch : whole word match in col1, NO column filter")
print("   General      : col6 contains 'packing' + NOT matched to DPL/Despatch + NOT matching exclusion lists")
print("📋 Column header : Level/Role")
print("📋 Cell format   : Level 1 (Operator) / Level 3 (Technician)")