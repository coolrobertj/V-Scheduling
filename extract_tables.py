import fitz
import json

# Extract table structure and colors from bid schedule
doc = fitz.open('2026 Bid Schedule 3-10-2026.pdf')

for page_idx in range(doc.page_count):
    page = doc[page_idx]
    print(f"=== PAGE {page_idx+1} ===")
    
    # Try to get tables
    tabs = page.find_tables()
    if tabs and tabs.tables:
        for t_idx, tab in enumerate(tabs.tables):
            print(f"Table {t_idx}: {tab.row_count} rows x {tab.col_count} cols")
            data = tab.extract()
            for r_idx, row in enumerate(data):
                print(f"  Row {r_idx}: {row}")
    print()

doc.close()
