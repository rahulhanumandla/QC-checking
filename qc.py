import fitz
import re

proxy_path = "proxy.pdf"
notice_path = "notice.pdf"
output_pdf = "Agenda_Comparison_Report.pdf"

# -------------------------
# 1. Extract rectangle from annotations
# -------------------------
def get_rect_from_pdf(path):
    doc = fitz.open(path)
    for pno in range(len(doc)):
        page = doc[pno]
        annot = page.first_annot
        while annot:
            rect = annot.rect
            if rect:  
                return pno, rect
            annot = annot.next
    raise ValueError("No rectangle annotation found!")

# -------------------------
# 2. Extract text inside the rectangle only
# -------------------------
def get_text_in_rect(path, pno, rect):
    doc = fitz.open(path)
    page = doc[pno]
    txt = page.get_text("text", clip=rect)
    return txt

# -------------------------
# 3. Split into label → text dictionary
# -------------------------
def extract_proposals(text):
    label_re = re.compile(r'^\s*(\d{1,2}[a-z]?\.|0\d\)|\d{1,2}\))')
    proposals = {}
    current_label = None
    lines = text.splitlines()

    for line in lines:
        m = label_re.match(line)
        if m:
            label = m.group(1).strip()
            remainder = line[m.end():].strip()
            proposals[label] = remainder
            current_label = label
        else:
            if current_label:
                proposals[current_label] += " " + line.strip()

    # cleanup repeated spaces
    for k in proposals:
        proposals[k] = re.sub(r'\s+', ' ', proposals[k]).strip()

    return proposals

# -------------------------
# 4. Compare proxy vs notice
# -------------------------
def compare(proxy_dict, notice_dict):
    labels = sorted(set(proxy_dict.keys()) | set(notice_dict.keys()))
    results = []

    for lbl in labels:
        p = proxy_dict.get(lbl)
        n = notice_dict.get(lbl)

        if p is None:
            results.append((lbl, "MISSING_IN_PROXY"))
        elif n is None:
            results.append((lbl, "MISSING_IN_NOTICE"))
        elif p == n:
            results.append((lbl, "MATCH"))
        else:
            results.append((lbl, f"MISMATCH — Proxy: '{p}'  |  Notice: '{n}'"))

    return results

# -------------------------
# 5. Generate a PDF-only report
# -------------------------
def create_report(results, output_path):
    doc = fitz.open()
    page = doc.new_page()

    y = 40
    page.insert_text((40, 20), "AGENDA COMPARISON REPORT", fontsize=16)

    for lbl, status in results:
        if y > 760:
            page = doc.new_page()
            y = 40

        page.insert_text((40, y), f"{lbl}: {status}", fontsize=11)
        y += 20

    doc.save(output_path)
    doc.close()

# -----------------------------------------------------
# PIPELINE
# -----------------------------------------------------

# Extract rectangles
p_page, p_rect = get_rect_from_pdf(proxy_path)
n_page, n_rect = get_rect_from_pdf(notice_path)

# Extract agenda-only text
proxy_text = get_text_in_rect(proxy_path, p_page, p_rect)
notice_text = get_text_in_rect(notice_path, n_page, n_rect)

# Convert into dictionaries
proxy_dict = extract_proposals(proxy_text)
notice_dict = extract_proposals(notice_text)

# Compare
results = compare(proxy_dict, notice_dict)

# PDF Report
create_report(results, output_pdf)

print("DONE! Report saved as:", output_pdf)
