import fitz
import json

doc = fitz.open('2026 Bid Schedule 3-10-2026.pdf')
for i in range(doc.page_count):
    page = doc[i]
    print(f'=== PAGE {i+1} ===')
    print(page.get_text())
    print()
doc.close()
