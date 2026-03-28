# V-Scheduling Preferences

## Workflow
- Close the workbook (not Excel) before regenerating to avoid permission errors
- Use COM automation: `[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')` to close just the workbook
- After generating, reopen the file in Excel
- Python: `C:/Users/rojorda/AppData/Local/Python/pythoncore-3.14-64/python.exe`
- Main script: `generate_excel_v3.py`
- Data source: `bid_data.json` (117 named employees extracted from PDF)
- GitHub repo: https://github.com/coolrobertj/V-Scheduling.git

## Shift Definitions
- **Day shift**: start times 00:00–10:00
- **Swing shift**: start times 11:00–18:00
- **Graveyard shift**: start times 19:00–23:59 (wraps past midnight)

## Excel Structure
- **22 tabs**: 1 Master (blank template) + 21 daily shift tabs (7 days × 3 shifts)
- **Days**: Sunday–Saturday
- **Shifts per day**: Day, Swing, Graveyard
- **3 blank rows** at top of each sheet (header starts at row 4)
- **Two-column layout**: Left side (cols A–J), Right side (cols K–T)
  - Left sections: EAST LOT, WEST LOT, BREAKER
  - Right sections: SOUTH LOT, CBC, BP LOT
- **2-column absences section** when right side is shorter

## Fonts
- Section titles: Arial 13 bold
- Data cells: Arial 10
- Small text (headers): Arial 8 bold
- Data bold: Arial 10 bold
- White bold (on black headers): Arial 10 bold white
- **No word wrap** on any cells

## Lot Fill Colors
- EAST LOT: White (`FFFFFF`)
- SOUTH LOT: Gray (`BFBFBF`)
- WEST LOT: Dark Gray (`808080`)
- CBC: Green (`C4D79B`)
- BP LOT: Blue (`8DB4E2`)
- BREAKER: Yellow (`FFFF00`)
- Fill colors apply to **section headers only**, NOT to data/blank rows

## Column Widths (proportional, not equal)
- A, K (BUS#): 10.36
- B, L (Relief): 5
- C, M (Name): 16
- D, N (Start): 8
- E, O (End): 8
- F, P (Hours): 5
- G–J, Q–T (Trade Day/Overtime): 5

## Time Format
- Military time stored as integers (e.g., 700 for 07:00, 1430 for 14:30)
- Number format: `00":"00` — typing `0700` displays as `07:00`
- Applied to Start and End columns on both data and blank rows

## Hours Column
- Formula-based: auto-calculates from Start and End times
- Handles overnight shifts (end < start adds 24 hours)
- Uses `INT()` to **round down** (floor), not up
- Number format: `0` (whole number)
- Applied to both data rows and blank rows (including border-extension rows)

## Headers
- **Trade Day** and **Overtime** headers rotated 90° (bottom to top) using `text_rotation=90`

## Print Settings
- **Portrait** orientation
- **Letter** size paper
- Fit to page: `fitToWidth=1, fitToHeight=1, fitToPage=True`
- Margins: 0.1" left/right, 0.25" top/bottom
- Print area set to used range

## Things That Were Tried and Reverted
- Arial 12 bold for everything (reverted to original sizes)
- Auto-fit column widths (reverted to fixed proportional widths)
- Equal column widths (reverted to proportional)
- Section fill colors on data/blank rows (reverted to headers only)
