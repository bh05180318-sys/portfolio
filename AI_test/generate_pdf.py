import os
import subprocess
import sys
import markdown

sys.stdout.reconfigure(encoding='utf-8')

md_path = "correlation_report.md"
html_path = "correlation_report.html"
pdf_path = "correlation_report.pdf"

if not os.path.exists(md_path):
    print(f"Error: {md_path} not found.")
    sys.exit(1)

print("Reading Markdown...")
with open(md_path, "r", encoding="utf-8") as f:
    md_text = f.read()

# Convert markdown to HTML (with table extension support)
print("Converting to HTML...")
html_body = markdown.markdown(md_text, extensions=['tables'])

# Premium CSS for PDF printing
html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
    @page {{
        size: A4;
        margin: 20mm;
    }}
    body {{
        font-family: 'Malgun Gothic', '맑은 고딕', Arial, sans-serif;
        color: #2c3e50;
        line-height: 1.6;
        font-size: 14px;
        margin: 0;
        padding: 0;
    }}
    h1 {{
        font-size: 22px;
        color: #1b4f72;
        border-bottom: 2px solid #1f618d;
        padding-bottom: 8px;
        margin-top: 30px;
        margin-bottom: 15px;
        page-break-before: always;
    }}
    /* Do not page break before the very first H1 */
    h1:first-of-type {{
        page-break-before: avoid;
    }}
    h2 {{
        font-size: 16px;
        color: #2e4053;
        margin-top: 20px;
        margin-bottom: 10px;
    }}
    p {{
        margin-top: 0;
        margin-bottom: 10px;
        text-align: justify;
    }}
    ul, ol {{
        margin-top: 0;
        margin-bottom: 12px;
        padding-left: 20px;
    }}
    li {{
        margin-bottom: 6px;
    }}
    blockquote {{
        margin: 15px 0;
        padding: 10px 15px;
        background-color: #f8f9f9;
        border-left: 4px solid #1f618d;
        font-style: italic;
    }}
    table {{
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
        font-size: 13px;
    }}
    th, td {{
        border: 1px solid #d5dbdb;
        padding: 10px;
        text-align: center;
    }}
    th {{
        background-color: #1b4f72;
        color: #ffffff;
        font-weight: bold;
    }}
    tr:nth-child(even) {{
        background-color: #f2f4f4;
    }}
    img {{
        display: block;
        max-width: 100%;
        height: auto;
        margin: 20px auto;
        max-height: 380px;
        border-radius: 4px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }}
    em {{
        display: block;
        text-align: center;
        color: #7f8c8d;
        font-size: 11px;
        font-style: italic;
        margin-top: -15px;
        margin-bottom: 25px;
    }}
    /* Title layout */
    .doc-title {{
        text-align: center;
        font-size: 26px;
        font-weight: bold;
        color: #1b4f72;
        margin-top: 40px;
        margin-bottom: 10px;
        line-height: 1.3;
    }}
    .doc-subtitle {{
        text-align: center;
        font-size: 12px;
        color: #7f8c8d;
        margin-bottom: 50px;
    }}
    hr {{
        border: 0;
        height: 1px;
        background: #d5dbdb;
        margin: 30px 0;
    }}
</style>
</head>
<body>
    <!-- Replace the H1 title in body with styled title -->
    <div class="doc-title">삶의 만족도와 자살률 간의 상관관계 분석 보고서<br><span style="font-size: 20px;">(2020 ~ 2024)</span></div>
    <div class="doc-subtitle">작성일: 2026. 06. 11 &nbsp;|&nbsp; 분석기관: AI 코딩 어시스턴트</div>
    <hr>
    
    {html_body.replace('<h1>삶의 만족도와 자살률 간의 상관관계 분석 보고서 (2020 ~ 2024)</h1>', '')}
</body>
</html>
"""

# Save HTML
with open(html_path, "w", encoding="utf-8") as f:
    f.write(html_content)
print(f"HTML saved to {html_path}")

# Find Edge Executable
edge_paths = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    os.path.expandvars(r"%LocalAppData%\Microsoft\Edge\Application\msedge.exe")
]

edge_bin = None
for path in edge_paths:
    if os.path.exists(path):
        edge_bin = path
        break

if not edge_bin:
    # Try using PATH
    import shutil
    edge_bin = shutil.which("msedge")

if not edge_bin:
    print("Error: Microsoft Edge not found. Cannot convert to PDF.")
    sys.exit(1)

print(f"Edge located at: {edge_bin}")
print("Generating PDF via headless Edge...")

# Run Edge to print HTML to PDF
cmd = [
    edge_bin,
    "--headless",
    "--disable-gpu",
    f"--print-to-pdf={pdf_path}",
    html_path
]

try:
    subprocess.run(cmd, check=True)
    print(f"Successfully generated PDF: {pdf_path}")
    
    # Optionally clean up the temporary HTML file
    if os.path.exists(html_path):
        os.remove(html_path)
        print("Cleaned up temporary HTML file.")
        
except Exception as e:
    print(f"Failed to generate PDF: {e}")
    sys.exit(1)
