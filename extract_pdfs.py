# -*- coding: utf-8 -*-
import sys
import os
sys.stdout.reconfigure(encoding='utf-8')

from pdfminer.high_level import extract_text
import json
import glob

resume_dir = r'C:/Users/EDY/Documents/数据岗-1.20'
pdf_files = glob.glob(os.path.join(resume_dir, '*.pdf'))

print(f"Found {len(pdf_files)} PDF files")

resumes = []
for pf in sorted(pdf_files):
    fname = os.path.basename(pf)
    print(f"Extracting: {fname}")
    try:
        text = extract_text(pf)
        # Clean up the text
        text = text.replace('\n\n\n', '\n').replace('\r', '')
        if len(text.strip()) < 50:
            print(f"  WARNING: Very little text extracted ({len(text)} chars)")
        print(f"  Extracted {len(text)} chars")
        resumes.append({'name': fname.replace('.pdf',''), 'text': text})
    except Exception as e:
        print(f"  ERROR: {e}")

print(f"\nTotal resumes extracted: {len(resumes)}")

# Save to temp json
output_path = os.path.join(os.path.expanduser('~'), 'resume_risk_report', 'resumes_for_analysis.json')
os.makedirs(os.path.dirname(output_path), exist_ok=True)
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(resumes, f, ensure_ascii=False)

print(f"Saved to {output_path}")
print("\nFirst resume preview (first 500 chars):")
if resumes:
    print(resumes[0]['text'][:500])
