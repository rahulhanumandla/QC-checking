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
# PATHS
# -------------------------
BASE_PATH = r"C:\Users\Hanumandla Rahul\Desktop\PROXY"
PROXY_PDF = os.path.join(BASE_PATH, "proxy.pdf")
NOTICE_PDF = os.path.join(BASE_PATH, "notice.pdf")
OUTPUT_REPORT = os.path.join(BASE_PATH, "mismatch_report.pdf")

# -------------------------
# STEP 1: Extract clean text + remove junk
# -------------------------
def clean_text(text, is_proxy=False):
    lines = text.splitlines()
    cleaned = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # Remove Board recommendations
        if re.search(r"board.*recommend|recommends you vote", line, re.I):
            continue
        # Remove signature block
        if is_proxy and re.search(r"please sign|joint owners|corporation|authorized officer", line, re.I):
            continue
        # Remove voting boxes
        if re.search(r"☐|For|Against|Abstain|Year", line):
            continue
        cleaned.append(line)
    
    return "\n".join(cleaned)

def extract_text(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text("text")
    doc.close()
    return clean_text(text, "proxy" in pdf_path.lower())

# -------------------------
# STEP 2: Extract proposals by indentation + numbering (HUMAN LOGIC)
# -------------------------
def extract_proposals_by_indent(text):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    proposals = []
    current = ""
    
    for line in lines:
        # Detect proposal start: number/letter at start (1., 1) , 2a., 2b., 01), etc.)
        if re.match(r"^(\d{1,3}[a-zA-Z]?\.?\s*[\.)]?)|(\d{1,3}[a-zA-Z]?\s*$)", line):
            if current:
                proposals.append(current.strip())
            current = line
        else:
            # This line belongs to previous proposal (indented)
            if current:
                current += " " + line
    
    if current:
        proposals.append(current.strip())
    
    return proposals

# -------------------------
# STEP 3: Normalize for comparison
# -------------------------
def normalize(text):
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\s+([.,;])', r'\1', text)
    return text.strip().lower()

# -------------------------
# STEP 4: Compare one-by-one (SIMPLE & ACCURATE)
# -------------------------
def compare_proposals(proxy_props, notice_props):
    results = []
    max_len = max(len(proxy_props), len(notice_props))
    
    for i in range(max_len):
        p = proxy_props[i] if i < len(proxy_props) else "[MISSING IN PROXY]"
        n = notice_props[i] if i < len(notice_props) else "[MISSING IN NOTICE]"
        
        label = f"Item {i+1}"
        if i < len(proxy_props) and re.match(r"^\d", proxy_props[i]):
            label = re.split(r'\s+', proxy_props[i], 1)[0]
        elif i < len(notice_props) and re.match(r"^\d", notice_props[i]):
            label = re.split(r'\s+', notice_props[i], 1)[0]
        
        if normalize(p) == normalize(n):
            results.append((label, "MATCH", ""))
        else:
            results.append((label, "MISMATCH", f"Proxy:\n{p}\n\nNotice:\n{n}"))
    
    return results

# -------------------------
# STEP 5: Beautiful Report
# -------------------------
def create_report(results):
    fd, temp_path = tempfile.mkstemp(suffix=".pdf", dir=BASE_PATH)
    os.close(fd)
    
    doc = SimpleDocTemplate(temp_path, pagesize=letter, topMargin=60, leftMargin=50)
    elements = []
    styles = getSampleStyleSheet()

    elements.append(Paragraph("<font size=20><b>Proxy vs Notice QC Report</b></font>", styles['Title']))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"<i>Generated: {datetime.now().strftime('%d %B %Y, %I:%M %p')}</i>", styles['Normal']))
    elements.append(Spacer(1, 30))

    mismatches = sum(1 for r in results if r[1] == "MISMATCH")
    color = "red" if mismatches else "green"
    summary = f"<font color='{color}' size=16><b>{mismatches} MISMATCH(ES) FOUND</b></font>"
    if mismatches == 0:
        summary = "<font color='green' size=16><b>ALL PERFECT - NO MISMATCH</b></font>"
    elements.append(Paragraph(summary, styles['Normal']))
    elements.append(Spacer(1, 30))

    data = [["#", "Label", "Status", "Difference"]]
    for i, (label, status, info) in enumerate(results, 1):
        if status == "MATCH":
            data.append([str(i), label, Paragraph("<font color='green'><b>MATCH</b></font>", styles['Normal']), "Identical"])
        else:
            data.append([str(i), label, Paragraph("<font color='red'><b>MISMATCH</b></font>", styles['Normal']), 
                        Paragraph(info.replace("\n", "<br/>"), styles['Normal'])])

    table = Table(data, colWidths=[40, 100, 90, 300])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2E86C1")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F8F9F9")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTSIZE', (0,1), (-1,-1), 9),
        ('LEFTPADDING', (3,1), (3,-1), 15),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    elements.append(table)
    doc.build(elements)

    if os.path.exists(OUTPUT_REPORT):
        try: os.remove(OUTPUT_REPORT)
        except: pass
    shutil.move(temp_path, OUTPUT_REPORT)
    print(f"Report saved: {OUTPUT_REPORT}")

# -------------------------
# MAIN — SIMPLE & HUMAN
# -------------------------
def run_qc():
    print("Starting Proxy vs Notice QC (Human Logic)...\n")
    
    proxy_text = extract_text(PROXY_PDF)
    notice_text = extract_text(NOTICE_PDF)
    
    proxy_props = extract_proposals_by_indent(proxy_text)
    notice_props = extract_proposals_by_indent(notice_text)
    
    print(f"Proxy proposals : {len(proxy_props)}")
    print(f"Notice proposals: {len(notice_props)}\n")
    
    results = compare_proposals(proxy_props, notice_props)
    create_report(results)
    
    mismatches = sum(1 for r in results if r[1] == "MISMATCH")
    print(f"QC COMPLETE → {mismatches} mismatch(es) found!")
    print(f"Report: {OUTPUT_REPORT}")

if __name__ == "__main__":
    run_qc()
