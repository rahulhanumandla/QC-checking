import fitz
import re
import os
import tempfile
import shutil
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors

# -------------------------
# PATHS
# -------------------------
BASE_PATH = r"C:\Users\Hanumandla Rahul\Desktop\PROXY"
PROXY_PDF = os.path.join(BASE_PATH, "proxy.pdf")
NOTICE_PDF = os.path.join(BASE_PATH, "notice.pdf")
OUTPUT_REPORT = os.path.join(BASE_PATH, "mismatch_report.pdf")

# -------------------------
# STEP 1: READ PDF TEXT
# -------------------------
def extract_text(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc:
        full_text += page.get_text()
    return full_text

# -------------------------
# STEP 2: EXTRACT PROPOSALS (FIXED: No more NOTE/FOR leakage)
# -------------------------
def extract_proposals(text):
    # Now captures NOTE, FOR, AGAINST, etc. as SEPARATE blocks
    pattern = r"(\b\d{1,2}[a-zA-Z]?\)|\b\d{1,2}[a-zA-Z]?\.|\b0\d\)|\b\d{1,2}[A-Za-z]?\.|NOTE\b|\bFOR\b|\bAGAINST\b|\bABSTAIN\b|\bWITHHOLD\b|\b1 YEAR\b|\b2 YEARS\b|\b3 YEARS\b)"
    matches = list(re.finditer(pattern, text, re.IGNORECASE))
    proposals = {}
    
    for i in range(len(matches)):
        label = matches[i].group().strip()
        start = matches[i].end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        content = text[start:end].strip()
        
        # Skip voting options & NOTE — they are not real proposals
        if re.match(r"(?i)^(FOR|AGAINST|ABSTAIN|WITHHOLD|NOTE|1 YEAR|2 YEARS|3 YEARS)", label):
            continue
            
        proposals[label] = content
    
    return proposals

# -------------------------
# STEP 3: COMPARE
# -------------------------
def compare_proposals(proxy_data, notice_data):
    results = []
    all_labels = set(proxy_data.keys()) | set(notice_data.keys())
    for label in all_labels:
        p = proxy_data.get(label, "")
        n = notice_data.get(label, "")
        if p == n:
            results.append((label, "MATCH", ""))
        else:
            results.append((label, "MISMATCH", f"Proxy:\n{p}\n\nNotice:\n{n}"))
    return results

# -------------------------
# BEAUTIFUL COLORFUL REPORT
# -------------------------
def create_report(results, output_path):
    fd, temp_path = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    
    doc = SimpleDocTemplate(temp_path, pagesize=letter, topMargin=50, leftMargin=40, rightMargin=40)
    elements = []

    elements.append(Paragraph("<font size=18><b>Proxy vs Notice QC Report</b></font>", getSampleStyleSheet()['Title']))
    elements.append(Spacer(1, 15))
    elements.append(Paragraph(f"<i>{os.path.basename(PROXY_PDF)} vs {os.path.basename(NOTICE_PDF)}</i>", getSampleStyleSheet()['Normal']))
    elements.append(Spacer(1, 30))

    mismatches = sum(1 for r in results if r[1] == "MISMATCH")
    summary = f"<font color='red' size=14><b>{mismatches} MISMATCH(ES) FOUND</b></font>" if mismatches else "<font color='green' size=14><b>ALL PERFECT - NO MISMATCH</b></font>"
    elements.append(Paragraph(summary, getSampleStyleSheet()['Normal']))
    elements.append(Spacer(1, 30))

    data = [["#", "Label", "Status", "Difference"]]
    for i, (label, status, info) in enumerate(results, 1):
        if status == "MATCH":
            row = [str(i), label, Paragraph("<font color='green'><b>MATCH</b></font>"), "Identical"]
        else:
            row = [str(i), label, Paragraph("<font color='red'><b>MISMATCH</b></font>"), Paragraph(info.replace("\n", "<br/>"))]
        data.append(row)

    table = Table(data, colWidths=[40, 100, 90, 300])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2E86C1")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 12),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F8F9F9")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTSIZE', (0,1), (-1,-1), 9),
        ('LEFTPADDING', (3,1), (3,-1), 15),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 30))
    elements.append(Paragraph("<i>Report generated automatically</i>", getSampleStyleSheet()['Normal']))

    doc.build(elements)
    shutil.move(temp_path, output_path)
    print(f"Report saved: {output_path}")

# -------------------------
# MAIN
# -------------------------
def run_qc():
    print("Starting Proxy vs Notice QC...\n")
    proxy_text = extract_text(PROXY_PDF)
    notice_text = extract_text(NOTICE_PDF)
    
    print("Extracting proposals...")
    proxy_props = extract_proposals(proxy_text)
    notice_props = extract_proposals(notice_text)
    
    print(f"Proxy items : {len(proxy_props)}")
    print(f"Notice items: {len(notice_props)}\n")
    
    results = compare_proposals(proxy_props, notice_props)
    create_report(results, OUTPUT_REPORT)
    
    mismatches = sum(1 for r in results if r[1] == "MISMATCH")
    print(f"QC COMPLETE → {mismatches} mismatch(es) found!")
    print(f"Report: {OUTPUT_REPORT}")

if __name__ == "__main__":
    run_qc()
