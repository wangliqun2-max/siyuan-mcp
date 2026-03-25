"""Smoke test: verify physical page tracking works for DOCX."""
import sys
sys.path.insert(0, 'backend')
from services.doc_parser import parse_document, get_section_text

result = parse_document('Section V Part- B2 - Standard AIS Rev1.docx')
chunks = result['chunks']
headings = result['headings']

# Find POWER TRANSFORMERS L1 heading
pt = [h for h in headings if 'POWER TRANSFORMER' in h['text'].upper() and h.get('level', 99) == 1]
if pt:
    h = pt[0]
    phys = h.get('physical_page', '?')
    print(f"POWER TRANSFORMERS: chunk={h['chunk_page']}, physical={phys}")
    text = get_section_text(chunks, phys, phys + 31)
    print(f"Text chars from physical range {phys}-{phys+31}: {len(text)}")
    print(f"First 300 chars:\n{text[:300]}")
else:
    print("Not found! All headings with TRANSFORMER:")
    for h in headings:
        if 'TRANSFORMER' in h['text'].upper():
            print(f"  chunk={h['chunk_page']} phys={h.get('physical_page','?')} L{h.get('level','?')} {h['text']}")
