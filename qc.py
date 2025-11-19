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
# STEP 1: READ PDF TEXT + REMOVE SIGNATURE BLOCK
# -------------------------
def extract_text(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()

    # REMOVE SIGNATURE BLOCK COMPLETELY (Proxy only)
    if "proxy" in pdf_path.lower():
        # Common signature patterns — bulletproof
        text = re.sub(r"Please sign exactly as your name.*$", "", text, flags=re.I | re.DOTALL)
        text = re.sub(r"Joint owners should each sign.*$", "", text, flags=re.I | re.DOTALL)
        text = re.sub(r"If a corporation or partnership.*$", "", text, flags=re.I | re.DOTALL)
    
    return text

# -------------------------
# STEP 2: EXTRACT PROPOSALS (Now detects 2a., 2b., 10a. perfectly)
# -------------------------
def extract_proposals(text):
    # REMOVE BOARD RECOMMENDATION BLOCKS FIRST
    text = re.sub(r"The\s+Board\s+of\s+Directors\s+recommends\s+you\s+vote.*?(\d+\.|$)", " ", text, flags=re.I | re.DOTALL)
    text = re.sub(r"Board\s+Recommends.*?(?=\d+\.|$)", " ", text, flags=re.I | re.DOTALL)

    # PERFECT PATTERN — now catches 1., 01), 1a., 2a., 2b., 10a., 10b. etc.
    pattern = r"(\b\d{1,3}[a-zA-Z]?\.|\b\d{1,3}[a-zA-Z]?\)|\b\d{1,3}[a-zA-Z]?\s*[.)]\s*|\b\d{1,3}[a-zA-Z]?\b)"
    
    matches = list(re.finditer(pattern, text, re.IGNORECASE))
    proposals = {}
    
    for i in range(len(matches)):
        label_match = matches[i]
        label = label_match.group().strip().rstrip('.')
        
        # Skip junk
        if re.search(r"(?i)^(NOTE|FOR|AGAINST|ABSTAIN|WITHHOLD|YEAR)", label):
            continue
            
        start = label_match.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        content = text[start:end].strip()
        
        # Clean up content
        content = re.sub(r"^\s*[:\-–—]\s*", "", content)
        content = re.sub(r"\s+", " ", content)
        
        proposals[label] = content
    
    return proposals

# -------------------------
# STEP 3: NORMALIZE TEXT (Shift+Enter & spacing proof)
# -------------------------
def normalize_text(text):
    if not text:
        return ""
    text = re.sub(r'\r\n|\r|\n', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\s+([.,;:!?])', r'\1', text)
    return text.strip()

# -------------------------
# STEP 4: COMPARE
# -------------------------
def compare_proposals(proxy_data, notice_data):
    results = []
    all_labels = sorted(set(proxy_data.keys()) | set(notice_data.keys()), key=lambda x: str(x))
    
    for label in all_labels:
        p_raw = proxy_data.get(label, "[MISSING IN PROXY]")
        n_raw = notice_data.get(label, "[MISSING IN NOTICE]")
        
        if normalize_text(p_raw) == normalize_text(n_raw):
            results.append((label, "MATCH", ""))
        else:
            results.append((label, "MISMATCH", f"Proxy:\n{p_raw}\n\nNotice:\n{n_raw}"))
    
    return results

# -------------------------
# STEP 5: BEAUTIFUL REPORT
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
            data.append([str(i), label, Paragraph("<font color='green'><b>MATCH</b></font>", styles['Normal']), "Identical content"])
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
    elements.append(Spacer(1, 30))
    elements.append(Paragraph("<i>Report generated automatically by Proxy QC Master Tool</i>", styles['Normal']))

    doc.build(elements)

    if os.path.exists(OUTPUT_REPORT):
        try: os.remove(OUTPUT_REPORT)
        except: pass
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
    
    print(f"Proxy proposals found : {len(proxy_props)} → {list(proxy_props.keys())}")
    print(f"Notice proposals found: {len(notice_props)} → {list(notice_props.keys())}\n")
    
    results = compare_proposals(proxy_props, notice_props)
    create_report(results)
    
    mismatches = sum(1 for r in results if r[1] == "MISMATCH")
    print(f"QC COMPLETE → {mismatches} mismatch(es) found!")
    print(f"Report: {OUTPUT_REPORT}")

if __name__ == "__main__":
    run_qc()
