import fitz
import json

doc = fitz.open('2026 Bid Schedule 3-10-2026.pdf')

# Color ranges to lot mapping (RGB from pixel sampling)
def classify_lot(r, g, b):
    """Classify a pixel color to a lot"""
    # White or near-white = EAST LOT
    if r > 240 and g > 240 and b > 240:
        return 'EAST LOT'
    # Gray (128,128,128) = WEST LOT
    if 100 < r < 160 and 100 < g < 160 and 100 < b < 160 and abs(r-g) < 15 and abs(g-b) < 15:
        return 'WEST LOT'
    # Light gray (216,216,216) = SOUTH LOT
    if 190 < r < 230 and 190 < g < 230 and 190 < b < 230 and abs(r-g) < 15 and abs(g-b) < 15:
        return 'SOUTH LOT'
    # Light blue (191,230,244) = BP LOT
    if 170 < r < 210 and 210 < g < 250 and 230 < b < 255:
        return 'BP LOT'
    # Light green (218,242,207) or (181,230,161) = CBC
    if r < g and b < g and 180 < g < 250 and 140 < r < 230 and abs(r-b) < 80:
        # Greenish   
        return 'CBC'
    # Yellow (255,255,0) = BREAKER
    if r > 230 and g > 230 and b < 50:
        return 'BREAKER'
    # Another green shade
    if 150 < r < 200 and 200 < g < 250 and 140 < b < 180:
        return 'CBC'
    return f'UNKNOWN({r},{g},{b})'

all_shifts = []

for page_idx in range(doc.page_count):
    page = doc[page_idx]
    
    # Render page at 2x resolution for better color sampling
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
            
            # Sample pixel color at cell center
            cell_idx = row_idx * tab.col_count + time_col
            lot = 'EAST LOT'
            if cell_idx < len(cells):
                cell_rect = cells[cell_idx]
                # Sample at center of cell, scale by 2 for the pixmap
                cx = int((cell_rect[0] + cell_rect[2]) / 2 * 2)
                cy = int((cell_rect[1] + cell_rect[3]) / 2 * 2)
                
                if 0 <= cx < pix.width and 0 <= cy < pix.height:
                    pixel = pix.pixel(cx, cy)
                    r, g, b = pixel[0], pixel[1], pixel[2]
                    lot = classify_lot(r, g, b)
            
            shift_data['days'][day_name] = {
                'time': time_val, 
                'hrs': hrs_val, 
                'lot': lot,
                'start': start_time,
                'end': end_time
            }
        
        all_shifts.append(shift_data)

doc.close()

# Save as JSON
with open('bid_data.json', 'w') as f:
    json.dump(all_shifts, f, indent=2)

print(f"Total shifts extracted: {len(all_shifts)}")

# Print lot distribution
print("\n=== LOT DISTRIBUTION PER DAY ===")
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

# Print a few specific shifts for spot-check
print("\n=== SAMPLE SHIFTS ===")
for s in all_shifts[:10]:
    print(f"\nShift {s['shift']}: {s['name']}")
    for day in ['Sun', 'Mon', 'Tue', 'Wed', 'Thur', 'Fri', 'Sat']:
        d = s['days'].get(day, {})
        if d.get('time') and d['time'] != 'OFF':
            print(f"  {day}: {d['time']} ({d['hrs']}hrs) -> {d['lot']}")
        elif d.get('time') == 'OFF':
            print(f"  {day}: OFF")
        else:
            print(f"  {day}: --")
