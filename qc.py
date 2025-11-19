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

BASE_PATH = r"C:\Users\Hanumandla Rahul\Desktop\PROXY"
PROXY_PDF = os.path.join(BASE_PATH, "proxy.pdf")
NOTICE_PDF = os.path.join(BASE_PATH, "notice.pdf")
OUTPUT_REPORT = os.path.join(BASE_PATH, "mismatch_report.pdf")

def extract_text(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    return text

def extract_proposals(text):
    pattern = r"(\b\d{1,2}[a-zA-Z]?\)|\b\d{1,2}[a-zA-Z]?\.|\b0\d\)|\b\d{1,2}[A-Za-z]?\.|NOTE\b|\bFOR\b|\bAGAINST\b|\bABSTAIN\b|\bWITHHOLD\b|\b1 YEAR\b|\b2 YEARS\b|\b3 YEARS\b)"
    matches = list(re.finditer(pattern, text, re.IGNORECASE))
    proposals = {}
    
    for i in range(len(matches)):
        label = matches[i].group().strip()
        if re.match(r"(?i)^(NOTE|FOR|AGAINST|ABSTAIN|WITHHOLD|1 YEAR|2 YEARS|3 YEARS)", label):
            continue
        start = matches[i].end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        proposals[label] = text[start:end].strip()
    
    return proposals

def normalize_text(text):
    if not text:
        return ""
    text = re.sub(r'\r\n|\r|\n', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\s+([.,:;)])', r'\1', text)
    return text.strip()

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

def create_report(results):
    fd, temp_path = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    
    doc = SimpleDocTemplate(temp_path, pagesize=letter, topMargin=60, leftMargin=50)
    elements = []
    styles = getSampleStyleSheet()

    elements.append(Paragraph("<font size=20><b>Proxy vs Notice QC Report</b></font>", styles['Title']))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"<i>{datetime.now().strftime('%d %B %Y, %I:%M %p')}</i>", styles['Normal']))
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
            data.append([str(i), label, Paragraph("<font color='red'><b>MISMATCH</b></font>", styles['Normal']), Paragraph(info.replace("\n", "<br/>"), styles['Normal'])])

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
    shutil.move(temp_path, OUTPUT_REPORT)
    print(f"Report saved: {OUTPUT_REPORT}")

def run_qc():
    print("Starting Proxy vs Notice QC...\n")
    proxy_text = extract_text(PROXY_PDF)
    notice_text = extract_text(NOTICE_PDF)
    
    proxy_props = extract_proposals(proxy_text)
    notice_props = extract_proposals(notice_text)
    
    results = compare_proposals(proxy_props, notice_props)
    create_report(results)
    
    mismatches = sum(1 for r in results if r[1] == "MISMATCH")
    print(f"QC COMPLETE → {mismatches} mismatch(es) found!")

if __name__ == "__main__":
    run_qc()
