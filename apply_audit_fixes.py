#!/usr/bin/env python3
"""
Apply audit fixes to Claim_700_Estimate_UPDATED.xlsx based on URL verification.
Only updates E column (HYPERLINK formulas with prices) and B column (descriptions).
Leaves F column formulas (=C*E) and subtotal SUM formulas intact for Excel to recalculate.
"""

import openpyxl
import re

INPUT_FILE = 'Claim_700_Estimate_UPDATED.xlsx'
OUTPUT_FILE = 'Claim_700_Estimate_AUDITED.xlsx'

wb = openpyxl.load_workbook(INPUT_FILE)

changes_log = []

def h(url, price):
    """Build HYPERLINK formula."""
    return f'=HYPERLINK("{url}",{price})'

def get_hyperlink_parts(cell_value):
    """Extract URL and price from HYPERLINK formula."""
    if isinstance(cell_value, str) and 'HYPERLINK' in cell_value:
        match = re.match(r'=HYPERLINK\("([^"]+)",\s*([0-9.]+)\)', cell_value)
        if match:
            return match.group(1), float(match.group(2))
    return None, None

def update_price(sheet_name, row, new_price, reason):
    """Update price in HYPERLINK formula, preserving the URL. Leave F column formula intact."""
    ws = wb[sheet_name]
    cell = ws[f'E{row}']
    url, old_price = get_hyperlink_parts(cell.value)
    if url and old_price is not None:
        cell.value = h(url, new_price)
        changes_log.append(f"  {sheet_name} Row {row}: E price ${old_price:.2f} → ${new_price:.2f} - {reason}")

def update_desc(sheet_name, row, new_desc, reason):
    """Update description in column B."""
    ws = wb[sheet_name]
    old_desc = ws[f'B{row}'].value
    ws[f'B{row}'] = new_desc
    changes_log.append(f"  {sheet_name} Row {row}: desc '{old_desc}' → '{new_desc}' - {reason}")

def add_notes(sheet_name, row, note):
    """Add/update note in column G."""
    ws = wb[sheet_name]
    existing = ws[f'G{row}'].value or ''
    if existing:
        ws[f'G{row}'] = f"{existing}; {note}"
    else:
        ws[f'G{row}'] = note


# ============================================================
# 1. CONFIRMED PRICE UPDATES (E column only)
# ============================================================
print("=== APPLYING CONFIRMED PRICE CHANGES ===")

# Mold Tough Drywall: $14.98 → $21.48
changes_log.append("Mold Tough Drywall $14.98 → $21.48:")
update_price('Laundry Room', 13, 21.48, "+43% increase")

# Zinsser Mold Killing Primer: $28.98 → $45.98
changes_log.append("\nZinsser Mold Killing Primer $28.98 → $45.98:")
for sheet, row in [('Laundry Room', 30), ('Downstairs Bathroom', 45)]:
    update_price(sheet, row, 45.98, "+59% increase")

# Lightweight Finishing Compound: $20.98 → $23.68
changes_log.append("\nLightweight Finishing Compound $20.98 → $23.68:")
for sheet, row in [('Office', 22), ('Hallway', 23), ('Family Room', 23),
                   ('Laundry Room', 17), ('Downstairs Bathroom', 43)]:
    update_price(sheet, row, 23.68, "+13% increase")

# All-Purpose Joint Compound: $19.48 → $20.58
changes_log.append("\nAll-Purpose Joint Compound $19.48 → $20.58:")
for sheet, row in [('Office', 21), ('Hallway', 22), ('Family Room', 22),
                   ('Laundry Room', 16), ('Downstairs Bathroom', 42)]:
    update_price(sheet, row, 20.58, "+6% increase")

# HardieBacker: $14.68 → $16.85
changes_log.append("\nHardieBacker Cement Board $14.68 → $16.85:")
update_price('Downstairs Bathroom', 19, 16.85, "+15% increase")

# DAP Wood Filler: $6.98 → $7.98
changes_log.append("\nDAP Wood Filler $6.98 → $7.98:")
for sheet, row in [('Office', 28), ('Hallway', 29), ('Family Room', 33),
                   ('Laundry Room', 28), ('Downstairs Bathroom', 41)]:
    update_price(sheet, row, 7.98, "+14% increase")

# Door Hinges: $9.98 → $10.98
changes_log.append("\nEverbilt Door Hinges $9.98 → $10.98:")
for sheet, row in [('Laundry Room', 22), ('Downstairs Bathroom', 36)]:
    update_price(sheet, row, 10.98, "+10% increase")

# UltraLight Drywall: $16.25 → $16.68
changes_log.append("\nUltraLight Drywall $16.25 → $16.68:")
for sheet, row in [('Office', 18), ('Hallway', 19), ('Family Room', 19)]:
    update_price(sheet, row, 16.68, "+3% increase")

# Hydraulic Cement: $13.98 → $14.52
changes_log.append("\nHydraulic Cement $13.98 → $14.52:")
for sheet, row in [('Office', 15), ('Hallway', 16), ('Family Room', 16)]:
    update_price(sheet, row, 14.52, "+4% increase")

# Drywall Screws: $8.47 → $6.97
changes_log.append("\nDrywall Screws $8.47 → $6.97:")
for sheet, row in [('Office', 19), ('Hallway', 20), ('Family Room', 20), ('Laundry Room', 14)]:
    update_price(sheet, row, 6.97, "-18% decrease")

# Baseboard MDF: $8.98 → $6.51
changes_log.append("\nBaseboard MDF 3-1/4\" $8.98 → $6.51:")
for sheet, row in [('Office', 24), ('Hallway', 25), ('Family Room', 29),
                   ('Laundry Room', 24), ('Downstairs Bathroom', 37)]:
    update_price(sheet, row, 6.51, "-27% decrease")

# VersaBond Thinset: $17.97 → $14.97
changes_log.append("\nVersaBond Thinset $17.97 → $14.97:")
update_price('Downstairs Bathroom', 21, 14.97, "-17% decrease")


# ============================================================
# 2. DESCRIPTION FIXES
# ============================================================
print("=== APPLYING DESCRIPTION FIXES ===")

changes_log.append("\n\n--- Description Fixes ---")
update_desc('Downstairs Bathroom', 31, "Plumbers putty (stain-free)",
            "link only covers putty, not teflon tape")
update_desc('Laundry Room', 23, "Spring door stop (satin nickel)",
            "URL points to spring door stop, not rigid")


# ============================================================
# 3. ADD AUDIT NOTES (column G)
# ============================================================
print("=== ADDING AUDIT NOTES ===")

changes_log.append("\n\n--- Audit Notes Added ---")

# LifeProof carpet - category page
for sheet, row in [('Office', 8), ('Hallway', 9), ('Family Room', 9)]:
    add_notes(sheet, row, "AUDIT: URL is category page, not specific product SKU")

# LVP Trail Oak vs Sterling Oak
for sheet, row in [('Laundry Room', 9), ('Downstairs Bathroom', 9)]:
    add_notes(sheet, row, "AUDIT: SKU 300699284 may now be Sterling Oak; verify product")

# DAP Kwik Seal mismatch
for row in [15, 25]:
    add_notes('Downstairs Bathroom', row, "AUDIT: URL may be Kwik Seal (18032) not Kwik Seal Plus (18510); verify SKU")

# Loctite Power Grab size
for sheet, row in [('Downstairs Bathroom', 16), ('Family Room', 26)]:
    add_notes(sheet, row, "AUDIT: Verify 9oz cartridge vs 6oz tube")

# Roberts moisture barrier
add_notes('Laundry Room', 10, "AUDIT: SKU 100564393 may be replaced by 100578718; price may be lower")

# Potentially discontinued items
for sheet, row, note in [
    ('Laundry Room', 11, "AUDIT: Zamma T-molding SKU 205721044 may be discontinued"),
    ('Laundry Room', 12, "AUDIT: Quarter round SKU 206082006 not found in search"),
    ('Laundry Room', 19, "AUDIT: Door slab SKU 204418170 may be replaced by 302895972"),
    ('Laundry Room', 20, "AUDIT: Door casing SKU 309671432 not found in search"),
    ('Laundry Room', 21, "AUDIT: Kwikset SKU 100689825 may be replaced by 205176758"),
    ('Downstairs Bathroom', 17, "AUDIT: Daltile tile SKU 307562453 not found in search"),
    ('Downstairs Bathroom', 18, "AUDIT: Daltile bullnose SKU 307562448 not found in search"),
    ('Downstairs Bathroom', 20, "AUDIT: Cement board screws SKU 205690610 not found in search"),
]:
    add_notes(sheet, row, note)

# Paint supplies that may be discontinued
for note, locations in [
    ("AUDIT: Roller covers SKU may be superseded",
     [('Office', 33), ('Hallway', 34), ('Family Room', 38), ('Laundry Room', 33), ('Downstairs Bathroom', 48)]),
    ("AUDIT: Roller frame SKU may be discontinued",
     [('Office', 34), ('Laundry Room', 34)]),
    ("AUDIT: Paint tray SKU may be discontinued",
     [('Office', 35), ('Laundry Room', 35)]),
    ("AUDIT: Sash brush SKU may be discontinued",
     [('Office', 36), ('Laundry Room', 36)]),
    ("AUDIT: ScotchBlue tape may be rebranded PROSharp",
     [('Office', 37), ('Laundry Room', 37)]),
    ("AUDIT: Drop cloth SKU not found in search",
     [('Office', 38), ('Laundry Room', 39), ('Hallway', 35), ('Family Room', 39), ('Downstairs Bathroom', 49)]),
]:
    for sheet, row in locations:
        add_notes(sheet, row, note)

# Laundry Room row 38 URL mismatch
add_notes('Laundry Room', 38, "AUDIT: URL is roller frame, not paint supply kit; verify")


# ============================================================
# 4. SAVE (F column formulas and subtotals left intact for Excel)
# ============================================================
wb.save(OUTPUT_FILE)

print("\n" + "\n".join(changes_log))
print(f"\nSaved audited spreadsheet to: {OUTPUT_FILE}")
print("NOTE: F column formulas (=C*E) and subtotal SUM formulas preserved.")
print("Open in Excel to see recalculated totals.")
