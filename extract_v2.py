import fitz
import json
from collections import Counter

doc = fitz.open('2026 Bid Schedule 3-10-2026.pdf')

def classify_lot(r, g, b):
    """Classify a pixel color to a lot"""
    # Black/near-black = grid line, skip
    if r < 50 and g < 50 and b < 50:
        return None
    # Dark colors = grid line artifacts
    if r < 80 and g < 80:
        return None
    # White or near-white = EAST LOT
    if r > 235 and g > 235 and b > 235:
        return 'EAST LOT'
    # Gray ~(128,128,128) = WEST LOT
    if 100 < r < 170 and 100 < g < 170 and 100 < b < 170 and abs(r-g) < 20 and abs(g-b) < 20:
        return 'WEST LOT'
    # Light gray ~(216,216,216) = SOUTH LOT
    if 190 < r < 235 and 190 < g < 235 and 190 < b < 235 and abs(r-g) < 20 and abs(g-b) < 20:
        return 'SOUTH LOT'
    # Light blue ~(191,230,244) or (218,232,248) = BP LOT
    if g > 200 and b > 220 and r < g and r < b:
        return 'BP LOT'
    # Green shades = CBC  ~(218,242,207) or (181,230,161)
    if g > 200 and g > r and g > b and abs(r - b) < 100:
        return 'CBC'
    # Yellow (255,255,0) = BREAKER
    if r > 200 and g > 200 and b < 80:
        return 'BREAKER'
    # Yellowish
    if r > 150 and g > 150 and b < 50 and abs(r-g) < 80:
        return 'BREAKER'
    return None  # Unknown, keep sampling

def get_cell_lot(pix, cell_rect, scale=2):
    """Sample multiple points in a cell to reliably detect background color"""
    x0, y0, x1, y1 = cell_rect
    w = x1 - x0
    h = y1 - y0
    
    # Sample at multiple points (avoid edges where grid lines are)
    offsets = [
        (0.3, 0.3), (0.3, 0.5), (0.3, 0.7),
        (0.5, 0.3), (0.5, 0.5), (0.5, 0.7),
        (0.7, 0.3), (0.7, 0.5), (0.7, 0.7),
        (0.25, 0.25), (0.75, 0.75), (0.4, 0.6),
    ]
    
    lots = []
    for ox, oy in offsets:
        px = int((x0 + w * ox) * scale)
        py = int((y0 + h * oy) * scale)
        if 0 <= px < pix.width and 0 <= py < pix.height:
            pixel = pix.pixel(px, py)
            r, g, b = pixel[0], pixel[1], pixel[2]
            lot = classify_lot(r, g, b)
            if lot:
                lots.append(lot)
    
    if not lots:
        return 'EAST LOT'  # Default
    
    # Return most common
    counter = Counter(lots)
    return counter.most_common(1)[0][0]

all_shifts = []

for page_idx in range(doc.page_count):
    page = doc[page_idx]
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
    
    tabs = page.find_tables()
    if not tabs or not tabs.tables:
        continue
    tab = tabs.tables[0]
    data = tab.extract()
    cells = tab.cells
    
    day_cols = {
        'Mon': (2, 3),
        'Tue': (4, 5),
        'Wed': (6, 7),
        'Thur': (8, 9),
        'Fri': (10, 11),
        'Sat': (12, 13),
        'Sun': (14, 15),
    }
    
    for row_idx in range(2, len(data)):
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
                shift_data['days'][day_name] = {'time': '', 'hrs': '', 'lot': '', 'start': '', 'end': ''}
                continue
            
            if time_val == '0' or (not time_val and hrs_val == '0'):
                shift_data['days'][day_name] = {'time': 'OFF', 'hrs': '0', 'lot': 'OFF', 'start': '', 'end': ''}
                continue
            
            # Parse start/end times
            start_time = ''
            end_time = ''
            if time_val and '-' in time_val:
                parts = time_val.split('-')
                start_time = parts[0].strip().replace(':', '')
                end_time = parts[1].strip().replace(':', '')
            
            # Get lot from pixel color
            cell_idx = row_idx * tab.col_count + time_col
            lot = 'EAST LOT'
            if cell_idx < len(cells):
                cell_rect = cells[cell_idx]
                lot = get_cell_lot(pix, cell_rect)
            
            shift_data['days'][day_name] = {
                'time': time_val,
                'hrs': hrs_val,
                'lot': lot,
                'start': start_time,
                'end': end_time
            }
        
        all_shifts.append(shift_data)

doc.close()

with open('bid_data.json', 'w') as f:
    json.dump(all_shifts, f, indent=2)

print(f"Total shifts: {len(all_shifts)}")

print("\n=== LOT DISTRIBUTION PER DAY ===")
for day in ['Sun', 'Mon', 'Tue', 'Wed', 'Thur', 'Fri', 'Sat']:
    lots = {}
    for s in all_shifts:
        d = s['days'].get(day, {})
        lot = d.get('lot', '')
        if lot and lot != 'OFF' and lot != '':
            lots[lot] = lots.get(lot, 0) + 1
    print(f"\n{day}: (total working: {sum(lots.values())})")
    for lot, count in sorted(lots.items()):
        print(f"  {lot}: {count}")

print("\n=== SAMPLE SHIFTS ===")
for s in all_shifts[:8]:
    print(f"\nShift {s['shift']}: {s['name']}")
    for day in ['Sun', 'Mon', 'Tue', 'Wed', 'Thur', 'Fri', 'Sat']:
        d = s['days'].get(day, {})
        if d.get('time') and d['time'] != 'OFF':
            print(f"  {day}: {d['time']} ({d['hrs']}hrs) -> {d['lot']}")
        elif d.get('time') == 'OFF':
            print(f"  {day}: OFF")
        else:
            print(f"  {day}: --")
