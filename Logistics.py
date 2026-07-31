from bs4 import BeautifulSoup
import pandas as pd
import os
import re

# --- 1. FOLDER SETUP ---
main_folder = "Logistics"
if not os.path.exists(main_folder):
    os.makedirs(main_folder)

nested_folder = os.path.join(main_folder, "Logistics")
if not os.path.exists(nested_folder):
    os.makedirs(nested_folder)

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

# List of keywords
keywords = ["DPLP", "Subcon", "Shipping"]

# Initialize dictionary
keyword_lists = {
    name: {f"Skill Level {i}": {"list": [], "count": 0} for i in range(1, 7)}
    for name in keywords
}

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

# Single source of truth for the table header
LEVEL_HEADER = "<tr class='header-row'><th>Level/Role</th><th>Documents</th></tr>"

print("🔄 Processing data...")
total_processed = 0

for row in data:
    if len(row) < 2:
        continue

    value = str(row[0]) if pd.notna(row[0]) else ""
    link  = str(row[1]) if pd.notna(row[1]) else ""

    if value is not None and value.strip():
        for i in range(1, 7):
            if f"(Skill Level {i})" in str(value):
                for keyword in keywords:
                    if keyword.upper() in str(value).upper():
                        entry = f"{value} - <a href='https://plmpublishing.icp.infineon.com/api/download-pdf/{link}' target='_blank'>{link}</a>"
                        keyword_lists[keyword][f"Skill Level {i}"]["list"].append(entry)
                        keyword_lists[keyword][f"Skill Level {i}"]["count"] += 1
                        total_processed += 1

print(f"✅ Total entries processed: {total_processed}")

# --- 3. MAIN DASHBOARD (Logistics/Logistics.html) ---
html = "<html><head>"
html += "<style>"
html += """
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
html += "</style></head><body>"
html += "<div class='container'>"
html += "<div class='header'>"
html += "<a href='../index.html' class='home-btn'>Return to Home</a>"
html += "</div>"
html += "<div class='content'>"
html += "<div class='sidebar'>Logistics</div>"
html += "<div class='main-content'>"
html += "<div class='post-it-container'>"

for keyword in keywords:
    html += f"<div class='post-it'><a href='Logistics/{keyword}.html'>{keyword}</a></div>"

html += "</div></div></div></div></body></html>"

soup = BeautifulSoup(html, 'html.parser')
with open(os.path.join(main_folder, "Logistics.html"), "w", encoding="utf-8") as file:
    print(soup.prettify(), file=file)

# --- 4. INDIVIDUAL KEYWORD PAGES ---
for keyword in keywords:
    keyword_html  = "<html><head><style>"
    keyword_html += """
    body { font-family: Arial, sans-serif; margin: 20px; background-color: #f0f0f0; }
    .back-btn { display: inline-block; padding: 10px 20px; background-color: #08665c; color: white; text-decoration: none; border-radius: 5px; font-size: 14px; margin-bottom: 20px; transition: background-color 0.3s; }
    .back-btn:hover { background-color: #054a40; }
    .data-table { border-collapse: collapse; width: 100%; font-size: 18px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); background-color: #fff; }
    .header-row { background-color: #f0f0f0; color: #333; font-weight: bold; }
    .header-row th { padding: 12px; text-align: left; }
    .skill-level { text-align: left; font-size: 16px; padding: 12px; border-bottom: 1px solid #ddd; }
    .data-cell { text-align: left; font-size: 16px; padding: 12px; border-bottom: 1px solid #ddd; line-height: 1.5; }
    .data-cell.empty { background-color: #cccccc; }
    .data-cell:hover { background-color: #f0f0f0; }
    """
    keyword_html += "</style></head><body>"
    keyword_html += "<a href='../Logistics.html' class='back-btn'>← Back to Dashboard</a>"
    keyword_html += f"<h1>{keyword}</h1>"
    keyword_html += "<table class='data-table'>"

    # CHANGED: "Level/Role" header via shared constant
    keyword_html += LEVEL_HEADER

    has_data = False
    for i in range(1, 7):
        if keyword_lists[keyword][f"Skill Level {i}"]["count"] > 0:
            has_data = True
            label = level_role_label(i)
            keyword_html += (
                f"<tr>"
                f"<td class='skill-level'>{label}</td>"
                f"<td class='data-cell'>{'<br/><br/>'.join(keyword_lists[keyword][f'Skill Level {i}']['list'])}</td>"
                f"</tr>"
            )

    if not has_data:
        keyword_html += "<tr><td class='skill-level' colspan='2'>No documents found for this keyword.</td></tr>"

    keyword_html += "</table></body></html>"

    with open(os.path.join(nested_folder, f"{keyword}.html"), "w", encoding="utf-8") as file:
        print(BeautifulSoup(keyword_html, 'html.parser').prettify(), file=file)

print("\n✅ Structure Updated to Nested Folders:")
print("📁 Dashboard : Logistics/Logistics.html")
print("📁 Pages     : Logistics/Logistics/[Keyword].html")
print("📋 Column header : Level/Role")
print("📋 Cell format   : Level 1 (Operator) / Level 3 (Technician)")