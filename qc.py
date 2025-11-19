import fitz
import re
import os
import tempfile
import shutil
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from datetime import datetime

# -------------------------
# PATHS (Change only if needed)
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
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    return text

# -------------------------
# STEP 2: EXTRACT PROPOSALS (Ignores NOTE, FOR, AGAINST)
# -------------------------
def extract_proposals(text):
    pattern = r"(\b\d{1,2}[a-zA-Z]?\)|\b\d{1,2}[a-zA-Z]?\.|\b0\d\)|\b\d{1,2}[A-Za-z]?\.|NOTE\b|\bFOR\b|\bAGAINST\b|\bABSTAIN\b|\bWITHHOLD\b|\b1 YEAR\b|\b2 YEARS\b|\b3 YEARS\b)"
    matches = list(re.finditer(pattern, text, re.IGNORECASE))
    proposals = {}
    
    for i in range(len(matches)):
        label = matches[i].group().strip()
        # Skip NOTE, FOR, AGAINST, etc. — they are not real proposals
        if re.match(r"(?i)^(NOTE|FOR|AGAINST|ABSTAIN|WITHHOLD|1 YEAR|2 YEARS|3 YEARS)", label):
            continue
            
        start = matches[i].end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        content = text[start:end].strip()
        proposals[label] = content
    
    return proposals

# -------------------------
# STEP 3: NORMALIZE TEXT (Shift+Enter Proof)
# -------------------------
def normalize_text(text):
    if not text:
        return ""
    text = re.sub(r'\r\n|\r|\n', ' ', text)        # Remove all line breaks
    text = re.sub(r'\s+', ' ', text)               # Collapse spaces
    text = re.sub(r'\s+([.,:;)])', r'\1', text)    # Remove space before punctuation
    return text.strip()

# -------------------------
# STEP 4: COMPARE (Smart & Accurate)
# -------------------------
def compare_proposals(proxy_data, notice_data):
    results = []
    all_labels = sorted(set(proxy_data.keys()) | set(notice_data.keys()))
    
    for label in all_labels:
        p_raw = proxy_data.get(label, "[MISSING IN PROXY]")
        n_raw = notice_data.get(label, "[MISSING IN NOTICE]")
        
        if normalize_text(p_raw) == normalize_text(n_raw):
            results.append((label, "MATCH", ""))
        else:
            results.append((label, "MISMATCH", f"Proxy:\n{p_raw}\n\nNotice:\n{n_raw}"))
    
    return results

# -------------------------
# STEP 5: CREATE BEAUTIFUL REPORT (No PermissionError Ever)
# -------------------------
def create_report(results):
    # Create temp file in same folder
    fd, temp_path = tempfile.mkstemp(suffix=".pdf", dir=BASE_PATH)
    os.close(fd)  # Critical: close the file handle
    
    doc = SimpleDocTemplate(temp_path, pagesize=letter, topMargin=60, leftMargin=50, rightMargin=50)
    elements = []
    styles = getSampleStyleSheet()

    # Header
    elements.append(Paragraph("<font size=20><b>Proxy vs Notice QC Report</b></font>", styles['Title']))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"<i>Generated: {datetime.now().strftime('%d %B %Y, %I:%M %p')}</i>", styles['Normal']))
    elements.append(Spacer(1, 30))

    # Summary
    mismatches = sum(1 for r in results if r[1] == "MISMATCH")
    color = "red" if mismatches else "green"
    summary = f"<font color='{color}' size=16><b>{mismatches} MISMATCH(ES) FOUND</b></font>"
    if mismatches == 0:
        summary = "<font color='green' size=16><b>ALL PERFECT - NO MISMATCH</b></font>"
    elements.append(Paragraph(summary, styles['Normal']))
    elements.append(Spacer(1, 30))

    # Table
    data = [["#", "Label", "Status", "Difference"]]
    for i, (label, status, info) in enumerate(results, 1):
        if status == "MATCH":
            data.append([
                str(i),
                label,
                Paragraph("<font color='green'><b>MATCH</b></font>", styles['Normal']),
                "Identical content"
            ])
        else:
            data.append([
                str(i),
                label,
                Paragraph("<font color='red'><b>MISMATCH</b></font>", styles['Normal']),
                Paragraph(info.replace("\n", "<br/>"), styles['Normal'])
            ])

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
    elements.append(Paragraph("<i>Report generated automatically by Proxy QC Tool</i>", styles['Normal']))

    doc.build(elements)

    # FINAL FIX: Delete old report first → then replace
    if os.path.exists(OUTPUT_REPORT):
        try:
            os.remove(OUTPUT_REPORT)
        except:
            pass  # If still locked, overwrite anyway
    shutil.move(temp_path, OUTPUT_REPORT)

    print(f"Report saved: {OUTPUT_REPORT}")

# -------------------------
# MAIN
# -------------------------
def run_qc():
    print("Starting Proxy vs Notice QC...\n")
    
    if not os.path.exists(PROXY_PDF):
        print(f"ERROR: {PROXY_PDF} not found!")
        return
    if not os.path.exists(NOTICE_PDF):
        print(f"ERROR: {NOTICE_PDF} not found!")
        return
    
    proxy_text = extract_text(PROXY_PDF)
    notice_text = extract_text(NOTICE_PDF)
    
    proxy_props = extract_proposals(proxy_text)
    notice_props = extract_proposals(notice_text)
    
    print(f"Proxy items found : {len(proxy_props)}")
    print(f"Notice items found: {len(notice_props)}\n")
    
    results = compare_proposals(proxy_props, notice_props)
    create_report(results)
    
    mismatches = sum(1 for r in results if r[1] == "MISMATCH")
    print(f"QC COMPLETE → {mismatches} mismatch(es) found!")
    print(f"Report: {OUTPUT_REPORT}")

if __name__ == "__main__":
    run_qc()
