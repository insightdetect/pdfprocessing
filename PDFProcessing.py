import pandas as pd
import pdfplumber
from weasyprint import HTML
from datetime import datetime
import os

# Paths
input_pdf_path = r"PDF Location"
output_pdf_path = r"ProcessedPDFlocation"

# Step 1: Extract table from PDF
all_rows = []
with pdfplumber.open(input_pdf_path) as pdf:
    for page in pdf.pages[1:]:  # Skip cover page
        table = page.extract_table()
        if table:
            for row in table:
                if any(cell and str(cell).strip() for cell in row):
                    all_rows.append(row)

# Step 2: Normalize
trimmed_rows = [row[:8] + [''] * (8 - len(row)) if len(row) < 8 else row for row in all_rows]
df = pd.DataFrame(trimmed_rows, columns=[
    'S. No', 'Associated Process', 'Problem', 'Observation',
    'Time Loss (min)', 'Root Cause', 'Action Taken', 'Suggested Action'
])

# === 3. Clean + Fill missing S. No ===
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
df['S. No'] = df['S. No'].replace('', None).ffill()

# === 4. Group by S. No ===
#grouped_df = df.groupby('S. No', dropna=False).agg(
#    lambda x: '<br>'.join(filter(None, map(str, x)))
#).reset_index()
# === 4. Group by S. No and merge clean text ===
grouped_df = df.groupby('S. No', dropna=False).agg(
    lambda x: '<br>'.join(
        str(item).replace('\\n', ' ').replace('\n', ' ').strip()
        for item in x if item and str(item).strip()
    )
).reset_index()

# === 5. Clean column names + Sort by S. No ===
grouped_df.columns = [col.strip().replace('\n', ' ') for col in grouped_df.columns]
# Drop rows where S. No is not numeric
grouped_df = grouped_df[grouped_df['S. No'].astype(str).str.strip().str.isnumeric()]

# Now convert to int and sort
grouped_df['S. No'] = grouped_df['S. No'].astype(int)
grouped_df = grouped_df.sort_values('S. No')

#grouped_df = grouped_df.sort_values('S. No')


# Step 6: Convert to HTML
html_table = grouped_df.to_html(index=False, escape=False, border=1)

# Inject colgroup with column width hints
colgroup = """
<colgroup>
    <col style="width:5%;">
    <col style="width:13%;">
    <col style="width:17%;">
    <col style="width:17%;">
    <col style="width:10%;">
    <col style="width:13%;">
    <col style="width:12%;">
    <col style="width:13%;">
</colgroup>
"""

# Inject colgroup right after <table> tag
#html_table = html_table.replace("<table border=\"1\" class=\"dataframe\">", f"<table border=\"1\" class=\"dataframe\">{colgroup}")
html_table = html_table.replace('<table border="1" class="dataframe">', f'<table border="1" class="dataframe">{colgroup}')

# === 7. Build full HTML with styling, page numbers, footer, and logo ===
today_str = datetime.today().strftime('%Y-%m-%d')

full_html = f"""
<html>
<head>
<meta charset="UTF-8">
<style>
    @page {{
        size: A4;
        margin: 20mm;
        @bottom-center {{
            content: "Generated on: {today_str} | Page " counter(page) " of " counter(pages);
            font-size: 9pt;
            color: black;
        }}
    }}
    body {{
        font-family: Arial, sans-serif;
        font-size: 9.5pt;
        margin: 0;        
    }}
    table {{
        width: 100%;
        table-layout: fixed;
        border-collapse: collapse;
        border: none;
    }}
    th, td {{
        display: table-cell;
        border: 1px solid #ccc;
        padding: 6px;
        text-align: left;
        vertical-align: top;        
        
        white-space: normal;
        word-break: break-word;
        overflow-wrap: break-word;
        hyphens: auto;

        min-height: 50px;      /* ✅ Ensure height even if content is short */
        box-sizing: border-box;
    }}
    th {{
        background-color: #f2f2f2;
        font-weight: bold;
        word-break: keep-all;  /* ✅ Prevent header splitting mid-word */
    }}
    h2 {{
        text-align: center;
        margin: 10px 0;
    }}
    .header {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 10px;
    }}
</style>
</head>
<body>
    <div class="header">
        <h2>Processed Denmaru Assist Report</h2>
        <div style="width: 50px;"></div>
    </div>
    {html_table}
</body>
</html>
"""

# Step 5: Generate PDF
print("Final DataFrame before PDF:")
print(df.head(10))  # Or grouped_df.head(10)
print(df.shape)     # See row count
HTML(string=full_html).write_pdf(output_pdf_path)
print(f"[SUCCESS] Processed PDF saved at:\n{output_pdf_path}")

