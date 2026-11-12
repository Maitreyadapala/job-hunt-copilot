#!/usr/bin/env python3
"""Count how many of a posting's terms appear in a CV. This is a plain keyword count,
NOT an official ATS score (no real ATS publishes one).

Usage:
  python3 keyword_match.py cv.pdf terms.txt
terms.txt: one term per line. Use | for accepted alternatives: "A/B testing|split testing".
Lines starting with # are ignored. Matching is case-insensitive, on whole words/phrases.
Prints the percentage, the found list and the missing list.
"""
import sys, re, subprocess

def text_of(path):
    if path.lower().endswith(".pdf"):
        return subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True).stdout
    return open(path, encoding="utf-8", errors="ignore").read()

def norm(s):
    return re.sub(r"\s+", " ", s.lower().replace("‑", "-").replace("–", "-"))

def has(term, hay):
    pat = r"(?<![a-z0-9])" + re.escape(norm(term).strip()) + r"(?![a-z0-9])"
    return re.search(pat, hay) is not None

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__); sys.exit(1)
    hay = norm(text_of(sys.argv[1]))
    terms = [t.strip() for t in open(sys.argv[2], encoding="utf-8") if t.strip() and not t.startswith("#")]
    found, missing = [], []
    for t in terms:
        (found if any(has(a, hay) for a in t.split("|")) else missing).append(t)
    pct = round(100 * len(found) / max(1, len(terms)))
    print(f"Keyword match: {pct}% ({len(found)} of {len(terms)} terms). A keyword count, not an ATS score.")
    print("Found:  ", ", ".join(found) or "none")
    print("Missing:", ", ".join(missing) or "none")
