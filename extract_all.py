import fitz
import json

# Write all pages to a text file
doc = fitz.open('2026 Bid Schedule 3-10-2026.pdf')
with open('bid_schedule_text.txt', 'w', encoding='utf-8') as f:
    for i in range(doc.page_count):
        page = doc[i]
        f.write(f'=== PAGE {i+1} ===\n')
        f.write(page.get_text())
        f.write('\n\n')
doc.close()

# Also extract table data using get_text("dict") for structured info
doc = fitz.open('2026 Bid Schedule 3-10-2026.pdf')
with open('bid_schedule_colors.txt', 'w', encoding='utf-8') as f:
    for i in range(doc.page_count):
        page = doc[i]
        f.write(f'=== PAGE {i+1} ===\n')
        # Get drawings/rects for color info
        drawings = page.get_drawings()
        for d in drawings[:20]:
            f.write(f"Drawing: fill={d.get('fill')}, rect={d.get('rect')}\n")
        f.write('\n')
doc.close()

# Extract Master Daily Schedule structure
doc2 = fitz.open('Master Daily Schedule 2026.pdf')
with open('master_schedule_text.txt', 'w', encoding='utf-8') as f:
    page = doc2[0]
    f.write(page.get_text())
    f.write('\n\n--- DICT DATA ---\n')
    d = page.get_text("dict")
    for block in d["blocks"]:
        if block["type"] == 0:  # text block
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    if text:
                        color_int = span["color"]
                        r = (color_int >> 16) & 0xFF
                        g = (color_int >> 8) & 0xFF
                        b = color_int & 0xFF
                        f.write(f"Text: '{text}' | Font: {span['font']} | Size: {span['size']:.1f} | Color: ({r},{g},{b}) | Bbox: {span['bbox']}\n")
doc2.close()

# Extract colors/fills from Master Daily Schedule
doc2 = fitz.open('Master Daily Schedule 2026.pdf')
with open('master_schedule_colors.txt', 'w', encoding='utf-8') as f:
    page = doc2[0]
    drawings = page.get_drawings()
    f.write(f"Total drawings: {len(drawings)}\n")
    for idx, d in enumerate(drawings):
        fill = d.get('fill')
        rect = d.get('rect')
        if fill:
            f.write(f"Drawing {idx}: fill={fill}, rect={rect}\n")
doc2.close()

print("Done extracting all data!")
