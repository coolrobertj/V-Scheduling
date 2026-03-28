import fitz
import json

# Extract cell colors for every data cell in the bid schedule
doc = fitz.open('2026 Bid Schedule 3-10-2026.pdf')

# Color to lot mapping (from header legend)
# We need to identify which colored rectangle each cell falls within

all_shifts = []

for page_idx in range(doc.page_count):
    page = doc[page_idx]
    
    # Get colored rectangles (cell backgrounds)
    drawings = page.get_drawings()
    colored_rects = []
    for d in drawings:
        fill = d.get('fill')
        if fill:
            colored_rects.append({
                'fill': fill,
                'rect': d['rect'],
            })
    
    # Get table data
    tabs = page.find_tables()
    if not tabs or not tabs.tables:
        continue
    tab = tabs.tables[0]
    data = tab.extract()
    
    # Get cell positions from the table
    # tab.cells gives us cell coordinates
    cells = tab.cells  # list of (x0, y0, x1, y1) tuples
    
    # Day column indices: Mon=2,3  Tue=4,5  Wed=6,7  Thur=8,9  Fri=10,11  Sat=12,13  Sun=14,15
    day_cols = {
        'Mon': (2, 3),
        'Tue': (4, 5),
        'Wed': (6, 7),
        'Thur': (8, 9),
        'Fri': (10, 11),
        'Sat': (12, 13),
        'Sun': (14, 15),
    }
    
    for row_idx in range(2, len(data)):  # Skip header rows
        row = data[row_idx]
        shift_num = row[0].strip() if row[0] else ''
        name = row[1].strip().replace('\n', ' ') if row[1] else ''
        wkly_hrs = row[16].strip() if row[16] else ''
        
        if not shift_num:
            continue
        
        shift_data = {
            'shift': shift_num,
            'name': name,
            'wkly_hrs': wkly_hrs,
            'days': {}
        }
        
        for day_name, (time_col, hrs_col) in day_cols.items():
            time_val = row[time_col].strip() if row[time_col] else ''
            hrs_val = row[hrs_col].strip() if row[hrs_col] else ''
            
            if not time_val and not hrs_val:
                shift_data['days'][day_name] = {'time': '', 'hrs': '', 'lot': ''}
                continue
            
            if time_val == '0' or (not time_val and hrs_val == '0'):
                shift_data['days'][day_name] = {'time': 'OFF', 'hrs': '0', 'lot': 'OFF'}
                continue
            
            # Find the background color for this cell
            cell_idx = row_idx * tab.col_count + time_col
            if cell_idx < len(cells):
                cell_rect = cells[cell_idx]
                cx = (cell_rect[0] + cell_rect[2]) / 2
                cy = (cell_rect[1] + cell_rect[3]) / 2
                
                # Find which colored rect contains this cell center
                lot = 'EAST LOT'  # default (white)
                for cr in colored_rects:
                    r = cr['rect']
                    f = cr['fill']
                    if r.x0 <= cx <= r.x1 and r.y0 <= cy <= r.y1:
                        # Map color to lot
                        ri, gi, bi = int(f[0]*255), int(f[1]*255), int(f[2]*255)
                        if ri == 128 and gi == 128 and bi == 128:
                            lot = 'WEST LOT'
                        elif ri == 216 and gi == 216 and bi == 216:
                            lot = 'SOUTH LOT'
                        elif abs(ri-191) < 20 and abs(gi-230) < 20 and abs(bi-244) < 20:
                            lot = 'BP LOT'
                        elif abs(ri-218) < 20 and abs(gi-242) < 20 and abs(bi-207) < 20:
                            lot = 'CBC'
                        elif ri == 255 and gi == 255 and bi == 0:
                            lot = 'BREAKER'
                        elif abs(ri-181) < 20 and abs(gi-230) < 20 and abs(bi-161) < 20:
                            lot = 'SOUTH LOT'  # another green shade
                        elif ri == 255 and gi == 255 and bi == 255:
                            lot = 'EAST LOT'
                        # Don't break - last matching colored rect wins
                
                shift_data['days'][day_name] = {'time': time_val, 'hrs': hrs_val, 'lot': lot}
            else:
                shift_data['days'][day_name] = {'time': time_val, 'hrs': hrs_val, 'lot': 'EAST LOT'}
        
        all_shifts.append(shift_data)

doc.close()

# Save as JSON for verification
with open('bid_data.json', 'w') as f:
    json.dump(all_shifts, f, indent=2)

# Print summary
print(f"Total shifts extracted: {len(all_shifts)}")

# Print first few for verification
for s in all_shifts[:5]:
    print(f"\nShift {s['shift']}: {s['name']}")
    for day in ['Sun', 'Mon', 'Tue', 'Wed', 'Thur', 'Fri', 'Sat']:
        d = s['days'].get(day, {})
        if d.get('time'):
            print(f"  {day}: {d['time']} ({d['hrs']}hrs) -> {d['lot']}")
        else:
            print(f"  {day}: --")

# Also print lot distribution per day
print("\n\n=== LOT DISTRIBUTION ===")
for day in ['Sun', 'Mon', 'Tue', 'Wed', 'Thur', 'Fri', 'Sat']:
    lots = {}
    for s in all_shifts:
        d = s['days'].get(day, {})
        lot = d.get('lot', '')
        if lot and lot != 'OFF' and lot != '':
            lots[lot] = lots.get(lot, 0) + 1
    print(f"\n{day}:")
    for lot, count in sorted(lots.items()):
        print(f"  {lot}: {count}")
