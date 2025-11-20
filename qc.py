#!/usr/bin/env python3
"""
proxy_notice_compare.py

Compare agenda proposals between a proxy card PDF and a notice PDF, focusing only on the agenda
left-column area (left 7.5 inches by default) or a user-marked rectangle annotation if present.
Produces a single PDF report listing matches/mismatches and exact reasons for mismatches.

Usage:
    python proxy_notice_compare.py
    python proxy_notice_compare.py --proxy /mnt/data/proxy.pdf --notice /mnt/data/notice.pdf --out comparison_report.pdf
"""

import fitz  # PyMuPDF
import re
import difflib
import argparse
from datetime import datetime

# ---------- Config ----------
LEFT_CLIP_WIDTH_PTS = 540  # 7.5 inches * 72 pts/in
DEFAULT_PROXY = r"C:\Users\Hanumandla Rahul\Desktop\PROXY\proxy.pdf"
DEFAULT_NOTICE = r"C:\Users\Hanumandla Rahul\Desktop\PROXY\notice.pdf"
DEFAULT_OUT = "comparison_report.pdf"
# ----------------------------

label_pattern = re.compile(r'^\s*(\d{1,2}(?:[a-z])?\.|0\d\)|\d{1,2}\))')

def find_rectangle_annotation(doc):
    """Return (page_no, rect) of first rectangle/square annotation found, else None."""
    for pno in range(len(doc)):
        page = doc[pno]
        annot = page.first_annot
        while annot:
            try:
                r = annot.rect
                # If there's any rectangle-like annot, return it.
                if r and (r.x1 - r.x0) > 10 and (r.y1 - r.y0) > 10:
                    return pno, fitz.Rect(r)
            except Exception:
                pass
            annot = annot.next
    return None

def pick_agenda_clip(path):
    """
    Open PDF and try to determine page & clip rect for agenda:
      1) Use rectangle annotation if present.
      2) Otherwise look for page(s) where label_pattern matches in left 7.5in (540 pts).
      Returns (page_no, clip_rect, extracted_text)
    """
    doc = fitz.open(path)
    found = find_rectangle_annotation(doc)
    if found:
        pno, rect = found
        page = doc[pno]
        # expand a little to be safe
        page_rect = page.rect
        rect = fitz.Rect(max(page_rect.x0, rect.x0 - 6),
                         max(page_rect.y0, rect.y0 - 6),
                         min(page_rect.x1, rect.x1 + 6),
                         min(page_rect.y1, rect.y1 + 6))
        txt = page.get_text("text", clip=rect)
        doc.close()
        return pno, rect, txt

    # fallback: search pages for label pattern in left width
    for pno in range(len(doc)):
        page = doc[pno]
        clip = fitz.Rect(0, 0, LEFT_CLIP_WIDTH_PTS, page.rect.height)
        txt = page.get_text("text", clip=clip)
        if label_pattern.search(txt):
            doc.close()
            return pno, clip, txt

    # last fallback: first page left area
    page = doc[0]
    clip = fitz.Rect(0, 0, LEFT_CLIP_WIDTH_PTS, page.rect.height)
    txt = page.get_text("text", clip=clip)
    doc.close()
    return 0, clip, txt

def collapse_linebreaks(text):
    """Replace sequences of newlines and multiple spaces with a single space, but keep punctuation."""
    # Replace newlines with spaces, collapse repeated spaces
    # Keep other chars as-is (case & punctuation preserved)
    return re.sub(r'\s+', ' ', text).strip()

def split_proposals(text):
    """
    Given the text of the agenda column, split into ordered list of (label, content).
    Collapses internal linebreaks so proposal content is a single-line string.
    """
    lines = text.splitlines()
    entries = []
    cur_label = None
    cur_lines = []
    for ln in lines:
        m = label_pattern.match(ln)
        if m:
            if cur_label is not None:
                content = " ".join([l.strip() for l in cur_lines if l.strip()])
                entries.append((cur_label, collapse_linebreaks(content)))
            cur_label = m.group(1).strip()
            remainder = ln[m.end():].strip()
            cur_lines = [remainder] if remainder else []
        else:
            if cur_label is None:
                # skip leading lines (e.g., headers) until first label
                continue
            cur_lines.append(ln)
    if cur_label is not None:
        content = " ".join([l.strip() for l in cur_lines if l.strip()])
        entries.append((cur_label, collapse_linebreaks(content)))
    return entries

def compare_texts(a, b):
    """Return (status, reason, ndiff_snippet). Status: MATCH / MISMATCH."""
    if a == b:
        return "MATCH", "Exact match", ""
    if a == "" and b != "":
        return "MISSING_IN_PROXY", "Label missing in proxy", ""
    if a != "" and b == "":
        return "MISSING_IN_NOTICE", "Label missing in notice", ""
    # classify reason heuristically:
    reason = []
    if a.lower() == b.lower():
        reason.append("case-only difference")
    # strip punctuation and compare
    stripped_punc = lambda s: re.sub(r'[^\w\s]', '', s)
    if stripped_punc(a) == stripped_punc(b):
        reason.append("punctuation-only difference")
    # compare after collapsing whitespace
    if re.sub(r'\s+', ' ', a).strip() == re.sub(r'\s+', ' ', b).strip():
        reason.append("whitespace/linebreak differences only (after collapse)")
    if not reason:
        reason_text = "content differs"
    else:
        reason_text = ", ".join(reason)
    # produce a short ndiff for visibility (only lines surrounding changes)
    seq = difflib.ndiff([a], [b])
    nd = "\n".join(list(seq))
    return "MISMATCH", reason_text, nd

def build_report(proxy_path, notice_path, out_pdf_path, include_note=True):
    # Extract clips
    p_pno, p_rect, p_text_raw = pick_agenda_clip(proxy_path)
    n_pno, n_rect, n_text_raw = pick_agenda_clip(notice_path)

    # For both, collapse interior newlines into single spaces BEFORE splitting (so label detection still works)
    p_text = collapse_linebreaks(p_text_raw)
    n_text = collapse_linebreaks(n_text_raw)

    # split into proposals
    p_list = split_proposals(p_text_raw)  # we pass raw lines to preserve label detection then collapse content
    n_list = split_proposals(n_text_raw)

    # If user wants Note point included: we assume 'NOTE' or 'Note:' starts a note block.
    # We'll keep it as another label 'NOTE' if present at the end of agenda.
    def extract_note_block(text):
        m = re.search(r'\bNOTE\b[:\s-]*(.*)$', text, re.IGNORECASE)
        if m:
            return "NOTE", collapse_linebreaks(m.group(1))
        return None

    if include_note:
        pn = extract_note_block(p_text_raw)
        nn = extract_note_block(n_text_raw)
        # Add or merge into lists if present and not already included
        if pn:
            if not any(lbl.upper().startswith("NOTE") for lbl, _ in p_list):
                p_list.append(pn)
        if nn:
            if not any(lbl.upper().startswith("NOTE") for lbl, _ in n_list):
                n_list.append(nn)

    # convert to dict preserving order (proxy order preferred then notice extras)
    def to_dict(lst):
        d = {}
        order = []
        for lbl, txt in lst:
            key = lbl.strip()
            if key not in d:
                d[key] = txt
                order.append(key)
        return d, order

    p_dict, p_order = to_dict(p_list)
    n_dict, n_order = to_dict(n_list)

    all_labels = []
    for lbl in p_order:
        if lbl not in all_labels:
            all_labels.append(lbl)
    for lbl in n_order:
        if lbl not in all_labels:
            all_labels.append(lbl)

    comparisons = []
    for lbl in all_labels:
        a = p_dict.get(lbl, "")
        b = n_dict.get(lbl, "")
        # already collapsed content, but keep exact punctuation/case
        status, reason, ndiff_snip = compare_texts(a, b)
        comparisons.append({
            "label": lbl,
            "proxy_text": a,
            "notice_text": b,
            "status": status,
            "reason": reason,
            "ndiff": ndiff_snip
        })

    # Build PDF report
    out_doc = fitz.open()
    # Title page
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    title_page = out_doc.new_page()
    title_text = f"Agenda Comparison Report\nGenerated: {now}\n\nProxy: {proxy_path}\nNotice: {notice_path}\n\nRules applied:\n - Left column up to 7.5 in (540 pts) or rectangle annotation if present\n - Soft-returns / line-breaks collapsed before comparing\n - Exact match required (case & punctuation). Any difference is a mismatch\n - Note block is included (if present)\n\nSummary:\n"
    title_page.insert_text((36, 36), title_text, fontsize=10)

    # Summary table text
    y = 170
    for c in comparisons:
        line = f"{c['label']}: {c['status']}"
        title_page.insert_text((40, y), line, fontsize=9)
        y += 12
        if y > 740:
            # start a new page
            title_page = out_doc.new_page()
            y = 36

    # Detailed pages for mismatches (and also include MATCH lines optionally)
    for c in comparisons:
        page = out_doc.new_page()
        header = f"Label: {c['label']}   Status: {c['status']}\nReason: {c['reason']}\n\n"
        page.insert_text((36,36), header, fontsize=10)
        page.insert_text((36, 80), "Proxy text:", fontsize=9)
        page.insert_text((36, 96), c['proxy_text'] or "[EMPTY]", fontsize=9)
        page.insert_text((36, 140), "Notice text:", fontsize=9)
        page.insert_text((36, 156), c['notice_text'] or "[EMPTY]", fontsize=9)
        if c['status'] != "MATCH":
            page.insert_text((36, 220), "Exact diff (ndiff):", fontsize=9)
            nd = c['ndiff'] or "[ndiff unavailable]"
            # nd may be long; insert in a box area with smaller font
            page.insert_text((36, 236), nd, fontsize=8)
    # Save
    out_doc.save(out_pdf_path)
    out_doc.close()
    return out_pdf_path, comparisons

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Compare proxy and notice agenda proposals and produce a PDF report.")
    parser.add_argument("--proxy", default=DEFAULT_PROXY, help="Path to proxy PDF (default /mnt/data/proxy.pdf)")
    parser.add_argument("--notice", default=DEFAULT_NOTICE, help="Path to notice PDF (default /mnt/data/notice.pdf)")
    parser.add_argument("--out", default=DEFAULT_OUT, help="Output PDF report path")
    parser.add_argument("--no-note", dest="include_note", action="store_false", help="Do not include NOTE block in comparison")
    args = parser.parse_args()

    outpath, comps = build_report(args.proxy, args.notice, args.out, include_note=args.include_note)
    print("Report written to:", outpath)
    # Print quick console summary
    for c in comps:
        if c['status'] != "MATCH":
            print(f"{c['label']}: {c['status']} -> {c['reason']}")
