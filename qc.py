import fitz
import re

PROXY_FILE = "proxy.pdf"
NOTICE_FILE = "notice.pdf"
OUTPUT_PDF = "comparison_report.pdf"

LEFT_LIMIT = 540     # 7.5 inches × 72 dpi

# -----------------------------
# Extract left-side text (agenda area)
# -----------------------------
def extract_left_text(path):
    doc = fitz.open(path)
    label_re = re.compile(r'^\s*(\d{1,2}(?:[a-z])?\.|0\d\)|\d{1,2}\))', re.MULTILINE)

    for pno in range(len(doc)):
        page = doc[pno]
        clip = fitz.Rect(0, 0, LEFT_LIMIT, page.rect.height)
        txt = page.get_text("text", clip=clip)

        if label_re.search(txt):
            return txt

    # fallback if no agenda detected
    page = doc[0]
    clip = fitz.Rect(0, 0, LEFT_LIMIT, page.rect.height)
    return page.get_text("text", clip=clip)


# -----------------------------
# Convert text into proposal entries
# -----------------------------
def split_proposals(text):
    lines = text.splitlines()
    proposals = []
    cur_label = None
    buf = []

    label_pat = re.compile(r'^\s*(\d{1,2}(?:[a-z])?\.|0\d\)|\d{1,2}\))')

    for ln in lines:
        m = label_pat.match(ln)
        if m:
            if cur_label:
                proposals.append((cur_label, " ".join(buf).strip()))
            cur_label = m.group(1).strip()
            content = ln[m.end():].strip()
            buf = [content] if content else []
        else:
            if cur_label:
                buf.append(ln.strip())

    if cur_label:
        proposals.append((cur_label, " ".join(buf).strip()))

    return proposals


# -----------------------------
# Compare logic
# -----------------------------
def compare(proxy, notice):
    proxy_dict = {p[0]: p[1] for p in proxy}
    notice_dict = {p[0]: p[1] for p in notice}

    all_labels = sorted(set(proxy_dict.keys()) | set(notice_dict.keys()))

    results = []
    for lbl in all_labels:
        p = proxy_dict.get(lbl)
        n = notice_dict.get(lbl)

        if p is None:
            results.append((lbl, "MISSING_IN_PROXY", "Label missing in Proxy"))
        elif n is None:
            results.append((lbl, "MISSING_IN_NOTICE", "Label missing in Notice"))
        elif p == n:
            results.append((lbl, "MATCH", ""))
        else:
            results.append((lbl, "MISMATCH", f"Proxy: {p}\nNotice: {n}"))

    return results


# -----------------------------
# Create PDF report
# -----------------------------
def create_report(results):
    doc = fitz.open()
    page = doc.new_page()

    y = 50
    page.insert_text((30, 20), "Proxy vs Notice – Comparison Report", fontsize=14, fontname="helv")

    for lbl, status, reason in results:
        color = (0, 0.6, 0) if status == "MATCH" else \
                (1, 0, 0) if status == "MISMATCH" else \
                (1, 0.5, 0)

        page.insert_text((30, y), f"{lbl} → {status}", fontsize=11, color=color)

        if reason:
            lines = reason.split("\n")
            for ln in lines:
                y += 14
                page.insert_text((50, y), ln, fontsize=9)

        y += 20
        if y > 750:  # new page
            page = doc.new_page()
            y = 40

    doc.save(OUTPUT_PDF)
    doc.close()


# -----------------------------
# MAIN
# -----------------------------
proxy_text = extract_left_text(PROXY_FILE)
notice_text = extract_left_text(NOTICE_FILE)

proxy_entries = split_proposals(proxy_text)
notice_entries = split_proposals(notice_text)

results = compare(proxy_entries, notice_entries)

create_report(results)

print("Comparison completed →", OUTPUT_PDF)
