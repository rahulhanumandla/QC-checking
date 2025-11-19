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
# STEP 1: Extract raw text
# -------------------------
def extract_raw_text(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text("text")
    doc.close()
    return text

# -------------------------
# STEP 2: Clean text — remove junk but keep structure
# -------------------------
def clean_text_for_parsing(text, is_proxy=False):
    lines = text.splitlines()
    cleaned = []
    
    for line in lines:
        stripped = line.strip()
        
        # Remove signature block (proxy only)
        if is_proxy and re.search(r"please sign|joint owners|authorized officer|executor|administrator", stripped, re.I):
            continue
            
        # Remove voting boxes and "For Against Abstain"
        if re.search(r"☐|□|For\s+Against\s+Abstain|Year", stripped):
            continue
            
        # Remove "The Board of Directors recommends..." lines
        if re.search(r"board.*recommend|recommends you vote", stripped, re.I):
            continue
            
        # Remove "Board Recommends" column header and "For" bullets in Notice
        if re.search(r"Board\s+Recommends|For\s*$", stripped, re.I):
            continue
            
        if stripped:
            cleaned.append(line)  # Keep original spacing for indentation detection
    
    return "\n".join(cleaned)

# -------------------------
# STEP 3: Extract proposals using REAL indentation + numbering (HUMAN LOGIC)
# -------------------------
def extract_proposals(text):
    lines = text.splitlines()
    proposals = []
    current_proposal = []
    current_indent = 0

    # Detect proposal start pattern: 1.  1)  01.  2a.  2b.  10a. etc.
    start_pattern = re.compile(r'^(\s*)((\d{1,3}[a-zA-Z]?\.?)|(\d{1,3}[a-zA-Z]?\)))\s')

    for line in lines:
        match = start_pattern.match(line)
        if match:
            # Save previous proposal
            if current_proposal:
                proposals.append(" ".join(current_proposal).strip())
                current_proposal = []
            
            indent = len(match.group(1))
            label_and_text = line[match.end():].strip()
            current_proposal.append(label_and_text)
            current_indent = indent
        else:
            # Continuation line — belongs to current proposal
            if current_proposal:
                # Remove leading whitespace but keep content
                content = line.strip()
                if content:
                    current_proposal.append(content)
    
    # Don't forget the last one
    if current_proposal:
        proposals.append(" ".join(current_proposal).strip())
    
    return proposals

# -------------------------
# STEP 4: Normalize for comparison
# -------------------------
def normalize(text):
    if not text:
        return ""
    text = re.sub(r'\s+', ' ', text)           # Collapse all whitespace
    text = re.sub(r'\s+([.,;:!?)])', r'\1', text)
    return text.strip().lower()

# -------------------------
# STEP 5: Compare one-by-one
# -------------------------
def compare_proposals(proxy_list, notice_list):
    results = []
    max_len = max(len(proxy_list), len(notice_list))
    
    for i in range(max_len):
        p = proxy_list[i] if i < len(proxy_list) else "[MISSING IN PROXY]"
        n = notice_list[i] if i < len(notice_list) else "[MISSING IN NOTICE]"
        
        # Extract label from first few words
        label = f"Item {i+1}"
        if i < len(proxy_list):
            first_words = proxy_list[i].split()[:3]
            label = " ".join(first_words)
        elif i < len(notice_list):
            first_words = notice_list[i].split()[:3]
            label = " ".join(first_words)
        
        if normalize(p) == normalize(n):
            results.append((label, "MATCH", ""))
        else:
            results.append((label, "MISMATCH", f"Proxy:\n{p}\n\nNotice:\n{n}"))
    
    return results

# -------------------------
# STEP 6: Beautiful Report
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

    table = Table(data, colWidths=[40, 140, 90, 280])
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
    elements.append(Paragraph("<i>Report generated by Rahul's Final QC Master Tool</i>", styles['Normal']))

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
    print("Starting FINAL Proxy vs Notice QC (Human Accuracy)...\n")
    
    proxy_raw = extract_raw_text(PROXY_PDF)
    notice_raw = extract_raw_text(NOTICE_PDF)
    
    proxy_clean = clean_text_for_parsing(proxy_raw, is_proxy=True)
    notice_clean = clean_text_for_parsing(notice_raw, is_proxy=False)
    
    proxy_proposals = extract_proposals(proxy_clean)
    notice_proposals = extract_proposals(notice_clean)
    
    print(f"Proxy proposals found : {len(proxy_proposals)}")
    print(f"Notice proposals found: {len(notice_proposals)}\n")
    
    results = compare_proposals(proxy_proposals, notice_proposals)
    create_report(results)
    
    mismatches = sum(1 for r in results if r[1] == "MISMATCH")
    print(f"QC COMPLETE → {mismatches} mismatch(es) found!")
    print(f"Report: {OUTPUT_REPORT}")

if __name__ == "__main__":
    run_qc()
