#!/usr/bin/env python3
"""
Apply all URL, price, and description corrections to the audited spreadsheet.
This fixes every broken/mismatched/discontinued link found during verification.
"""

import openpyxl
import re

INPUT_FILE = 'Claim_700_Estimate_AUDITED.xlsx'
OUTPUT_FILE = 'Claim_700_Estimate_FINAL.xlsx'

wb = openpyxl.load_workbook(INPUT_FILE)

def h(url, price):
    return f'=HYPERLINK("{url}",{price})'

def get_parts(cell_value):
    if isinstance(cell_value, str) and 'HYPERLINK' in cell_value:
        m = re.match(r'=HYPERLINK\("([^"]+)",\s*([0-9.]+)\)', cell_value)
        if m:
            return m.group(1), float(m.group(2))
    return None, None

def fix(sheet, row, url=None, price=None, desc=None, notes=None):
    ws = wb[sheet]
    cell = ws[f'E{row}']
    old_url, old_price = get_parts(cell.value)
    u = url or old_url
    p = price if price is not None else old_price
    if u and p is not None:
        cell.value = h(u, p)
    if desc:
        ws[f'B{row}'] = desc
    if notes:
        existing = ws[f'G{row}'].value or ''
        # Replace any existing AUDIT notes
        if 'AUDIT:' in str(existing):
            parts = [x.strip() for x in str(existing).split(';') if 'AUDIT:' not in x]
            existing = '; '.join(parts)
        new_note = notes
        if existing:
            ws[f'G{row}'] = f"{existing}; {new_note}" if existing else new_note
        else:
            ws[f'G{row}'] = new_note

def clear_audit_note(sheet, row):
    """Remove AUDIT notes from column G, keep other notes."""
    ws = wb[sheet]
    existing = ws[f'G{row}'].value
    if existing and 'AUDIT:' in str(existing):
        parts = [x.strip() for x in str(existing).split(';') if 'AUDIT:' not in x]
        ws[f'G{row}'] = '; '.join(parts) if parts else None


print("=== APPLYING ALL URL/PRODUCT FIXES ===\n")

# ============================================================
# 1. DAP Kwik Seal Plus - WRONG SKU → correct SKU 100189580
# ============================================================
print("1. DAP Kwik Seal Plus → correct SKU 100189580")
kwik_seal_plus_url = "https://www.homedepot.com/p/DAP-Kwik-Seal-Plus-10-1-oz-White-Premium-Kitchen-and-Bath-Siliconized-Caulk-18510/100189580"
for row in [15, 25]:
    fix('Downstairs Bathroom', row, url=kwik_seal_plus_url, price=7.98,
        desc="DAP Kwik Seal Plus caulk - bath (10.1oz white)")
    clear_audit_note('Downstairs Bathroom', row)

# ============================================================
# 2. Loctite Power Grab Express → 9oz All-Purpose cartridge
# ============================================================
print("2. Loctite Power Grab Express → 9oz cartridge SKU 206432092")
loctite_url = "https://www.homedepot.com/p/Loctite-Power-Grab-All-Purpose-Instant-Grab-9-oz-Latex-Construction-Adhesive-White-Cartridge-2022554/206432092"
fix('Downstairs Bathroom', 16, url=loctite_url, price=5.48,
    desc="Panel adhesive - 9oz cartridge (per DreamLine specs)")
clear_audit_note('Downstairs Bathroom', 16)
fix('Family Room', 26, url=loctite_url, price=5.48,
    desc="Panel adhesive (heavy duty, 9oz cartridge)")
clear_audit_note('Family Room', 26)

# ============================================================
# 3. Roberts Moisture Barrier → new SKU 100578718
# ============================================================
print("3. Roberts Moisture Barrier → SKU 100578718")
roberts_url = "https://www.homedepot.com/p/ROBERTS-Moisture-Barricade-120-sq-ft-Roll-120-in-W-x-12-ft-L-x-6-mil-T-Underlayment-for-Vinyl-Laminate-SPC-Wood-70-115/100578718"
fix('Laundry Room', 10, url=roberts_url, price=0.19,
    notes="Updated SKU; ~$22.84/roll ÷ 120SF = $0.19/SF")
clear_audit_note('Laundry Room', 10)

# ============================================================
# 4. Masonite Lincoln Park 32" door slab → new SKU 302895972
# ============================================================
print("4. Masonite door slab 32\" → SKU 302895972")
door32_url = "https://www.homedepot.com/p/Masonite-32-in-x-80-in-1-Panel-Lincoln-Park-Primed-Solid-Core-Composite-Interior-Door-Slab-92453/302895972"
fix('Laundry Room', 19, url=door32_url,
    desc="Interior door slab - solid core (32\")")
clear_audit_note('Laundry Room', 19)

# ============================================================
# 5. Kwikset Juno passage knob → new SKU 205176758
# ============================================================
print("5. Kwikset Juno passage → SKU 205176758")
kwikset_passage_url = "https://www.homedepot.com/p/Kwikset-Juno-Satin-Nickel-Passage-Hall-Closet-Door-Knob-with-Microban-Antimicrobial-Technology-720J-15-CP/205176758"
fix('Laundry Room', 21, url=kwikset_passage_url, price=29.98,
    desc="Door hardware - passage knob w/Microban (satin nickel)")
clear_audit_note('Laundry Room', 21)

# ============================================================
# 6. Everbilt door stop → RIGID type SKU 314150759
# ============================================================
print("6. Everbilt door stop → rigid SKU 314150759")
doorstop_url = "https://www.homedepot.com/p/Everbilt-Satin-Nickel-Solid-Door-Stop-28447/314150759"
fix('Laundry Room', 23, url=doorstop_url, price=3.47,
    desc="Rigid door stop (satin nickel)")
clear_audit_note('Laundry Room', 23)

# ============================================================
# 7. Cement board screws → Pro-Twist SKU 301548302
# ============================================================
print("7. Cement board screws → Pro-Twist SKU 301548302")
screws_url = "https://www.homedepot.com/p/Pro-Twist-8-x-1-5-8-in-2-Phillips-Reduced-Flat-Head-Fiber-Cement-Board-Screws-1-lb-Box-NFCBS1581/301548302"
fix('Downstairs Bathroom', 20, url=screws_url,
    desc="Cement board screws 1-5/8\" (1 lb)")
clear_audit_note('Downstairs Bathroom', 20)

# ============================================================
# 8. Daltile Restore 3x6 tile → new SKU 302575146
# ============================================================
print("8. Daltile 3x6 tile → SKU 302575146")
tile_url = "https://www.homedepot.com/p/Daltile-Restore-Bright-White-3-in-x-6-in-Glossy-Ceramic-Subway-Wall-Tile-12-5-sq-ft-case-RE1536MODHD1P4/302575146"
# Price is $14.98/case for 12.5 SF = $1.20/SF. Original was $5.98/SF.
# Update URL and price to per-SF
fix('Downstairs Bathroom', 17, url=tile_url, price=1.20,
    desc="Ceramic wall tile - Bright White 3x6 (12.5 SF/case)",
    notes="$14.98/case ÷ 12.5 SF = $1.20/SF; updated from old SKU")

# ============================================================
# 9. Daltile bullnose → new SKU 302603039
# ============================================================
print("9. Daltile bullnose → SKU 302603039")
bullnose_url = "https://www.homedepot.com/p/Daltile-Restore-Bright-White-2-in-x-6-in-Glazed-Ceramic-Mudd-Bullnose-Trim-Tile-0-1-sq-ft-each-RE15A4200CC1P2/302603039"
fix('Downstairs Bathroom', 18, url=bullnose_url, price=1.19,
    desc="Bullnose trim tile - Bright White 2x6 (each)")

# ============================================================
# 10. Wooster roller covers 3-pack → SKU 203247625
# ============================================================
print("10. Wooster roller covers → SKU 203247625")
roller_url = "https://www.homedepot.com/p/Wooster-9-in-x-3-8-in-High-Density-Fabric-Wooster-Pro-White-Woven-Roller-Cover-Applicator-Tool-3-Pack-0HR4810090/203247625"
for sheet, row in [('Office', 33), ('Hallway', 34), ('Family Room', 38),
                   ('Laundry Room', 33), ('Downstairs Bathroom', 48)]:
    fix(sheet, row, url=roller_url, price=11.48)
    clear_audit_note(sheet, row)

# ============================================================
# 11. Wooster Sherlock roller frame + pole → SKU 321249479
# ============================================================
print("11. Wooster Sherlock combo → SKU 321249479")
sherlock_url = "https://www.homedepot.com/p/Wooster-4ft-8ft-Sherlock-Extension-Pole-and-9-in-Sherlock-Frame-0X10160000/321249479"
for sheet, row in [('Office', 34), ('Laundry Room', 34)]:
    fix(sheet, row, url=sherlock_url)
    clear_audit_note(sheet, row)

# ============================================================
# 12. HDX paint tray kit → SKU 205871766
# ============================================================
print("12. HDX paint tray → SKU 205871766")
tray_url = "https://www.homedepot.com/p/3-Piece-Plastic-Tray-Applicator-Kit-HD-MS-3-N/205871766"
for sheet, row in [('Office', 35), ('Laundry Room', 35)]:
    fix(sheet, row, url=tray_url, price=3.27)
    clear_audit_note(sheet, row)

# ============================================================
# 13. Wooster sash brush → SKU 206852097
# ============================================================
print("13. Wooster sash brush → SKU 206852097")
brush_url = "https://www.homedepot.com/p/Wooster-2-1-2-in-Pro-Nylon-Polyester-Angle-Sash-Brush-0H21410024/206852097"
for sheet, row in [('Office', 36), ('Laundry Room', 36)]:
    fix(sheet, row, url=brush_url)
    clear_audit_note(sheet, row)

# ============================================================
# 14. ScotchBlue painters tape → SKU 100550611
# ============================================================
print("14. ScotchBlue tape → SKU 100550611")
tape_url = "https://www.homedepot.com/p/3M-ScotchBlue-1-88-In-x-60-Yds-Original-Multi-Surface-Painter-s-Tape-2090-48CP/100550611"
for sheet, row in [('Office', 37), ('Laundry Room', 37)]:
    fix(sheet, row, url=tape_url, price=7.98)
    clear_audit_note(sheet, row)

# ============================================================
# 15. HDX drop cloth → 3-pack SKU 204711646 at $7.47
# ============================================================
print("15. HDX drop cloth → 3-pack SKU 204711646")
drop_url = "https://www.homedepot.com/p/HDX-9-ft-x-12-ft-Clear-Plastic-Drop-Cloths-3-Pack-DCHD-07-3/204711646"
for sheet, row in [('Office', 38), ('Laundry Room', 39), ('Hallway', 35),
                   ('Family Room', 39), ('Downstairs Bathroom', 49)]:
    fix(sheet, row, url=drop_url, price=7.47,
        desc="Plastic drop cloths 9x12 (3-pack)")
    clear_audit_note(sheet, row)

# ============================================================
# 16. Zamma T-molding → Millstead SKU 100661539
# ============================================================
print("16. T-molding → Millstead SKU 100661539")
tmold_url = "https://www.homedepot.com/p/Millstead-Unfinished-Oak-647-Thick-x-2-in-Wide-x-78-in-Length-T-Molding-LM5719/100661539"
fix('Laundry Room', 11, url=tmold_url,
    desc="Transition T-molding - unfinished oak (78\")",
    notes="Millstead replacement for discontinued Zamma")
clear_audit_note('Laundry Room', 11)

# ============================================================
# 17. Quarter round → shoe moulding SKU 206178777
# ============================================================
print("17. Quarter round → shoe moulding SKU 206178777")
shoe_url = "https://www.homedepot.com/p/Woodgrain-Millwork-126-1-2-in-x-3-4-in-x-96-in-Primed-Finger-Jointed-Shoe-Moulding-1-Piece-8-Total-Linear-Feet-10004360/206178777"
fix('Laundry Room', 12, url=shoe_url, price=5.37,
    desc="Shoe molding - primed 1/2\"x3/4\"x96\" (8ft)")
clear_audit_note('Laundry Room', 12)

# ============================================================
# 18. Alexandria door casing → individual casing piece SKU 205576626
# ============================================================
print("18. Door casing → Alexandria individual SKU 205576626")
casing_url = "https://www.homedepot.com/p/Alexandria-Moulding-5-8-in-x-2-3-4-in-x-84-in-Primed-MDF-Casing-Moulding-07705-96169C/205576626"
# Need 3 pieces per door (2 sides + 1 header) - update qty and pricing
for sheet, row in [('Laundry Room', 20), ('Downstairs Bathroom', 34)]:
    ws = wb[sheet]
    fix(sheet, row, url=casing_url, price=9.68,
        desc="Door casing - primed MDF colonial 84\" (per piece)")
    ws[f'C{row}'] = 3  # 3 pieces per door opening
    ws[f'D{row}'] = 'piece'
    clear_audit_note(sheet, row)

# ============================================================
# 19. Laundry Room row 38 - Paint supplies URL was wrong (roller frame)
# ============================================================
print("19. Laundry Room row 38 → paint supply kit URL")
# This row says "Paint supplies (rollers, brushes, tape, tray)" at $39.98
# but the URL points to the roller frame. Fix to a generic paint supply search
# or remove the hyperlink since this is a catch-all supply line.
# Best option: remove misleading hyperlink, keep as flat cost
ws = wb['Laundry Room']
ws['E38'] = 39.98  # Just the price, no misleading hyperlink
clear_audit_note('Laundry Room', 38)

# ============================================================
# 20. LifeProof carpet - category page is actually appropriate
# ============================================================
print("20. LifeProof carpet - keeping category page (product selected at store)")
# This is sold as installed carpet; category page is the correct reference
for sheet, row in [('Office', 8), ('Hallway', 9), ('Family Room', 9)]:
    ws = wb[sheet]
    existing = ws[f'G{row}'].value or ''
    if 'AUDIT:' in str(existing):
        parts = [x.strip() for x in str(existing).split(';') if 'AUDIT:' not in x]
        new = '; '.join(parts) if parts else ''
        ws[f'G{row}'] = f"{new}; Category page - specific style selected at store".strip('; ') if new else "Category page - specific style selected at store"

# ============================================================
# SAVE
# ============================================================
wb.save(OUTPUT_FILE)
print(f"\nSaved to: {OUTPUT_FILE}")
print("All broken/mismatched/discontinued URLs have been corrected.")
print("Open in Excel to see recalculated totals (F column formulas preserved).")
