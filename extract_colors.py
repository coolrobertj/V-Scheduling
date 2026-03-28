import fitz
import json

# Extract cell background colors from bid schedule
doc = fitz.open('2026 Bid Schedule 3-10-2026.pdf')

# For each page, get the table cells and their fill colors
for page_idx in range(min(2, doc.page_count)):  # Just first 2 pages for analysis
    page = doc[page_idx]
    print(f"=== PAGE {page_idx+1} ===")
    
    # Get all drawings (rectangles) with fills - these are cell backgrounds
    drawings = page.get_drawings()
    
    # Filter to only colored (non-white, non-black) rectangles
    colored_rects = []
    for d in drawings:
        fill = d.get('fill')
        if fill and fill != (1.0, 1.0, 1.0) and fill != (0.0, 0.0, 0.0):
            colored_rects.append({
                'fill': fill,
                'rect': d['rect'],
                'r': int(fill[0]*255),
                'g': int(fill[1]*255),
                'b': int(fill[2]*255)
            })
    
    print(f"Found {len(colored_rects)} colored rectangles")
    for cr in colored_rects:
        print(f"  Color: ({cr['r']},{cr['g']},{cr['b']}) Rect: {cr['rect']}")
    
    # Also get tables to map positions to cells
    tabs = page.find_tables()
    if tabs and tabs.tables:
        tab = tabs.tables[0]
        print(f"\nTable has {tab.row_count} rows x {tab.col_count} cols")
        # Print header row cell positions
        for c in range(tab.col_count):
            cell = tab.cells[0 * tab.col_count + c] if hasattr(tab, 'cells') else None
            
    print()

doc.close()
