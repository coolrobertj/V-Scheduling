import fitz
import json

doc = fitz.open('Master Daily Schedule 2026.pdf')
page = doc[0]

# Get detailed text with positions
d = page.get_text("dict")
print("=== ALL TEXT SPANS WITH POSITIONS ===")
for block in d["blocks"]:
    if block["type"] == 0:
        for line in block["lines"]:
            for span in line["spans"]:
                text = span["text"].strip()
                if text:
                    bbox = span["bbox"]
                    print(f"  x={bbox[0]:.1f}-{bbox[2]:.1f}  y={bbox[1]:.1f}-{bbox[3]:.1f}  '{text}'")

print("\n=== ALL DRAWINGS (RECTANGLES) ===")
drawings = page.get_drawings()
for idx, d in enumerate(drawings):
    fill = d.get('fill')
    rect = d.get('rect')
    stroke = d.get('color')
    width = d.get('width')
    print(f"  D{idx}: fill={fill} stroke={stroke} width={width} rect={rect}")

doc.close()
