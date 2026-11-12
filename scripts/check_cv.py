#!/usr/bin/env python3
"""Final checks before a CV is sent. Usage: python3 check_cv.py cv.pdf

Checks: page count, first-person pronouns, em dashes, repeated tool names in the Skills section
(read from the Software line automatically; --tools adds more names), and that the text extracts in reading order.
Exit code 0 = all clear, 1 = something to fix.
"""
import sys, re, subprocess, argparse

ap = argparse.ArgumentParser()
ap.add_argument("pdf"); ap.add_argument("--tools", default="")
a = ap.parse_args()
bad = False

info = subprocess.run(["pdfinfo", a.pdf], capture_output=True, text=True).stdout
pages = int(re.search(r"Pages:\s+(\d+)", info).group(1))
print(f"Pages: {pages}" + ("" if pages == 1 else "  <-- must be 1"))
bad |= pages != 1

text = subprocess.run(["pdftotext", a.pdf, "-"], capture_output=True, text=True).stdout
if len(text.strip()) < 200:
    print("Text did not extract: the PDF may be an image. An ATS cannot read it."); bad = True

pron = re.findall(r"(?<![A-Za-z])(I|my|me|he|his|she|her)(?![A-Za-z'])", text)
if pron:
    print("First or third person pronouns found:", sorted(set(pron))); bad = True
else:
    print("Pronouns: none")

if "—" in text:
    print("Em dashes found: replace with commas, colons, periods or parentheses."); bad = True
else:
    print("Dashes: none")

tools = [x.strip() for x in a.tools.split(",") if x.strip()]
layout_text = subprocess.run(["pdftotext", "-layout", a.pdf, "-"], capture_output=True, text=True).stdout
m = re.search(r"Software:[ \t]*(.+)", layout_text)
if m:
    tools += [x.strip() for x in re.split(r",|;", m.group(1)) if 1 < len(x.strip()) < 30]
for t in dict.fromkeys(tools):
    n = len(re.findall(r"(?<![A-Za-z0-9])" + re.escape(t) + r"(?![A-Za-z0-9])", text, flags=re.I))
    if n > 2:
        print(f"Tool '{t}' appears {n} times: once in Software, and once in a bullet only if a result is attached."); bad = True
    elif n == 2:
        print(f"Warning: tool '{t}' appears twice. Keep the bullet mention only if a result is attached to it.")
print("ALL CLEAR" if not bad else "FIX THE ITEMS ABOVE")
sys.exit(1 if bad else 0)
