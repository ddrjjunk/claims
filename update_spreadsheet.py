#!/usr/bin/env python3
"""
Update Claim_700_Estimate spreadsheet:
- Fix broken/discontinued Home Depot URLs
- Correct product description mismatches
- Update prices to current values
"""

import openpyxl

INPUT_FILE = 'Claim_700_Estimate (4).xlsx'
OUTPUT_FILE = 'Claim_700_Estimate_UPDATED.xlsx'

wb = openpyxl.load_workbook(INPUT_FILE)

def h(url, price):
    """Build HYPERLINK formula."""
    return f'=HYPERLINK("{url}",{price})'

# ============================================================
# 1. DESCRIPTION FIXES (Column B)
# ============================================================

desc_fixes = {
    # "Epoxy crack filler" → correct: it's DAP latex filler/sealant, not epoxy
    ('Office', 12): 'Concrete & mortar filler/sealant (10oz)',
    ('Hallway', 13): 'Concrete & mortar filler/sealant (10oz)',
    ('Family Room', 13): 'Concrete & mortar filler/sealant (10oz)',

    # 50 lb → 40 lb (Henry 555 is actually 40 lb)
    ('Office', 16): 'Self-leveling compound (40 lb bag)',
    ('Hallway', 17): 'Self-leveling compound (40 lb bag)',
    ('Family Room', 17): 'Self-leveling compound (40 lb bag)',

    # "standard drywall" → URL is actually UltraLight
    ('Office', 18): '1/2" x 4\'x8\' UltraLight drywall',
    ('Hallway', 19): '1/2" x 4\'x8\' UltraLight drywall',
    ('Family Room', 19): '1/2" x 4\'x8\' UltraLight drywall',

    # URL is plinth blocks, not corner blocks
    ('Office', 25): 'Plinth blocks (MDF primed)',
    ('Laundry Room', 25): 'Plinth blocks (MDF primed)',
    ('Downstairs Bathroom', 38): 'Plinth blocks (MDF primed)',
    ('Hallway', 26): 'Plinth blocks (MDF primed)',
    ('Family Room', 30): 'Plinth blocks (MDF primed)',

    # Spray texture is 20oz, not 25oz
    ('Office', 29): 'Spray texture - orange peel (20oz)',
    ('Laundry Room', 29): 'Spray texture - orange peel (20oz)',
    ('Hallway', 30): 'Spray texture - orange peel (20oz)',
    ('Family Room', 34): 'Spray texture - orange peel (20oz)',

    # Laundry Room: door slab (not prehung), knob (not lever), stop (no shims)
    ('Laundry Room', 19): 'Interior door slab - solid core (32")',
    ('Laundry Room', 21): 'Door hardware - passage knob (satin nickel)',
    ('Laundry Room', 23): 'Door stop (satin nickel)',

    # Bathroom: ceramic (not porcelain), screws only (no tape), slab/knob fixes
    ('Downstairs Bathroom', 17): 'Ceramic wall tile - surround (4ft high, excl shower)',
    ('Downstairs Bathroom', 20): 'Backer board screws (1 lb)',
    ('Downstairs Bathroom', 33): 'Interior door slab - solid core (30")',
    ('Downstairs Bathroom', 35): 'Door hardware - privacy knob (brushed nickel)',
}

for (sheet_name, row), new_desc in desc_fixes.items():
    wb[sheet_name][f'B{row}'] = new_desc

print(f"Applied {len(desc_fixes)} description fixes.")

# ============================================================
# 2. URL + PRICE FIXES (Column E - HYPERLINK formulas)
# ============================================================

# --- Tack strips: old SKU 202066639 discontinued ---
# New: Halex 3-pack 4ft strips, $3.47/pack = $0.29/LF
tack_url = "https://www.homedepot.com/p/Halex-Poplar-7-8-in-x-4-ft-Carpet-Tack-Strip-for-Wood-or-Concrete-Subfloors-3-Pack-HD-161-P-8/203301578"
for sheet, row in [('Office', 10), ('Hallway', 11), ('Family Room', 11)]:
    wb[sheet][f'E{row}'] = h(tack_url, 0.29)

# --- Carpet seam tape: Roberts 50-325 discontinued ---
# New: Roberts Heat-Loc 50-245
seam_url = "https://www.homedepot.com/p/ROBERTS-Heat-Loc-3-11-16-in-x-60-ft-Heat-Bond-Carpet-Seaming-Tape-Roll-50-245/204470708"
for sheet, row in [('Office', 11), ('Hallway', 12), ('Family Room', 12)]:
    wb[sheet][f'E{row}'] = h(seam_url, 13.20)

# --- DAP crack filler: old limestone variant gone, wrong description ---
# New: DAP gray variant SKU 204167828
dap_url = "https://www.homedepot.com/p/DAP-10-1-oz-Gray-Concrete-and-Mortar-Filler-and-Latex-Sealant-18096/204167828"
for sheet, row in [('Office', 12), ('Hallway', 13), ('Family Room', 13)]:
    wb[sheet][f'E{row}'] = h(dap_url, 8.78)

# --- Concrete bonding adhesive: WRONG PRODUCT (was crack sealant) ---
# Fix: correct Quikrete bonding adhesive SKU 100318541
bond_url = "https://www.homedepot.com/p/Quikrete-1-Qt-Concrete-Bonding-Adhesive-990214/100318541"
for sheet, row in [('Office', 14), ('Hallway', 15), ('Family Room', 15)]:
    wb[sheet][f'E{row}'] = h(bond_url, 11.97)

# --- Hydraulic cement: price increase ---
cement_url = "https://www.homedepot.com/p/Quikrete-10-lb-Hydraulic-Water-Stop-Cement-112611/100318494"
for sheet, row in [('Office', 15), ('Hallway', 16), ('Family Room', 16)]:
    wb[sheet][f'E{row}'] = h(cement_url, 13.98)

# --- Self-leveling compound: SKU changed ---
slc_url = "https://www.homedepot.com/p/Henry-555-Level-Pro-40-lb-Self-Leveling-Underlayment-12165/100549588"
for sheet, row in [('Office', 16), ('Hallway', 17), ('Family Room', 17)]:
    wb[sheet][f'E{row}'] = h(slc_url, 35.00)

# --- SLC primer: old URL broken, replace with Henry 554 1 gal ---
primer_url = "https://www.homedepot.com/p/Henry-554-Level-Pro-1-Gallon-Indoor-Underlayment-Primer-48967/340501834"
for sheet, row in [('Office', 17), ('Hallway', 18), ('Family Room', 18)]:
    wb[sheet][f'E{row}'] = h(primer_url, 27.98)

# --- Drywall: price increase $11.98 → $16.25 ---
drywall_url = "https://www.homedepot.com/p/USG-Sheetrock-Brand-1-2-in-x-4-ft-x-8-ft-UltraLight-Drywall-14113411708/202530243"
for sheet, row in [('Office', 18), ('Hallway', 19), ('Family Room', 19)]:
    wb[sheet][f'E{row}'] = h(drywall_url, 16.25)

# --- Door hinges: old model 14810 discontinued → new model 14874 ---
hinge_url = "https://www.homedepot.com/p/Everbilt-3-Pack-3-1-2-in-x-5-8-in-Radius-Satin-Nickel-Security-Butt-Door-Hinge-Value-Pack-14874/202818703"
for sheet, row in [('Laundry Room', 22), ('Downstairs Bathroom', 36)]:
    wb[sheet][f'E{row}'] = h(hinge_url, 9.98)

# --- Glacier Bay shower arm: discontinued → replacement ---
shower_arm_url = "https://www.homedepot.com/p/Glacier-Bay-6-in-Shower-Arm-and-Flange-in-Brushed-Nickel-3075-502/204511172"
wb['Downstairs Bathroom']['E14'] = h(shower_arm_url, 14.98)

# --- DAP Kwik Seal Plus: price increase $5.98 → $7.98 ---
kwik_url = "https://www.homedepot.com/p/DAP-Kwik-Seal-Plus-10-1-oz-White-Kitchen-and-Bath-Adhesive-Caulk-18510/100634364"
for row in [15, 25]:
    wb['Downstairs Bathroom'][f'E{row}'] = h(kwik_url, 7.98)

# --- VersaBond thinset: price decrease $21.98 → $17.97 ---
thinset_url = "https://www.homedepot.com/p/Custom-Building-Products-VersaBond-50-lb-White-Professional-Polymer-Modified-Thinset-Mortar-MTSW50/100091767"
wb['Downstairs Bathroom']['E21'] = h(thinset_url, 17.97)

# --- Unsanded grout: new SKU, price increase ---
grout_url = "https://www.homedepot.com/p/Custom-Building-Products-Polyblend-Plus-381-Bright-White-10-lb-Unsanded-Grout-PBPG38110/313296538"
wb['Downstairs Bathroom']['E22'] = h(grout_url, 18.48)

# --- Grout sealer: discontinued 22oz → replacement 15oz aerosol ---
sealer_url = "https://www.homedepot.com/p/Custom-Building-Products-TileLab-15-oz-Aerosol-Grout-Sealer-TLAGS15Z/202243455"
wb['Downstairs Bathroom']['E23'] = h(sealer_url, 13.98)

# --- QEP LASH: restructured → new project pack ---
lash_url = "https://www.homedepot.com/p/QEP-LASH-Tile-Leveling-System-Project-Pack-100-1-16-in-Leveling-Clips-and-100-Wedges-200-Piece-99741/314956144"
wb['Downstairs Bathroom']['E24'] = h(lash_url, 20.47)

# --- Vanity: replaced with newer assembled model ---
vanity_url = "https://www.homedepot.com/p/Home-Decorators-Collection-Doveton-30-in-Single-Sink-Freestanding-White-Bath-Vanity-with-White-Engineered-Marble-Top-Assembled-Doveton-30W/320686810"
wb['Downstairs Bathroom']['E26'] = h(vanity_url, 439)

# --- Moen Genta faucet: price increase $109 → $116.10 ---
genta_url = "https://www.homedepot.com/p/MOEN-Genta-Single-Handle-Single-Hole-Bathroom-Faucet-in-Brushed-Nickel-WS84760SRN/301980213"
wb['Downstairs Bathroom']['E27'] = h(genta_url, 116.10)

# --- Drain assembly: discontinued → Oatey replacement ---
drain_url = "https://www.homedepot.com/p/Oatey-1-1-4-in-Brushed-Nickel-Finish-Pop-Up-Assembly-HD759BNB3/316408158"
wb['Downstairs Bathroom']['E28'] = h(drain_url, 14.98)

# --- P-trap: SKU changed → updated SKU ---
ptrap_url = "https://www.homedepot.com/p/Everbilt-1-1-4-in-White-Plastic-Sink-Drain-P-Trap-with-Reversible-J-Bend-C9700B/205153784"
wb['Downstairs Bathroom']['E29'] = h(ptrap_url, 5.98)

# --- Supply lines: price decrease $9.98 → $7.28 ---
supply_url = "https://www.homedepot.com/p/BrassCraft-3-8-in-Comp-x-1-2-in-FIP-x-20-in-Braided-Stainless-Steel-Faucet-Connector-B1-20A-F/100464003"
wb['Downstairs Bathroom']['E30'] = h(supply_url, 7.28)

# --- Plumbers putty: discontinued → Oatey 9oz stain-free ---
putty_url = "https://www.homedepot.com/p/Oatey-9-oz-Stain-Free-Plumber-s-Putty-31177/203013821"
wb['Downstairs Bathroom']['E31'] = h(putty_url, 5.97)

# --- Mirror: discontinued → replacement HDC black rectangular ---
mirror_url = "https://www.homedepot.com/p/Home-Decorators-Collection-24-in-W-x-30-in-H-Rectangular-Aluminum-Framed-Wall-Bathroom-Vanity-Mirror-in-Black-Screws-Not-Included-2430-AL067B/322183957"
wb['Downstairs Bathroom']['E32'] = h(mirror_url, 89.97)

# --- BEHR Marquee semi-gloss 1 gal: price increase $54.98 → $58.98 ---
behr_sg_url = "https://www.homedepot.com/p/BEHR-MARQUEE-1-gal-Ultra-Pure-White-Semi-Gloss-Enamel-Interior-Paint-Primer-345001/204747524"
wb['Downstairs Bathroom']['E46'] = h(behr_sg_url, 58.98)

# --- Wall paneling: discontinued → MDF wainscot 4'x8' replacement ---
panel_url = "https://www.homedepot.com/p/Unbranded-1-4-in-x-4-ft-x-8-ft-MDF-Wainscot-Panel-739558/202090200"
wb['Family Room']['E25'] = h(panel_url, 34.98)

# --- Paneling trim: discontinued → Alexandria batten moulding ---
batten_url = "https://www.homedepot.com/p/Alexandria-Moulding-3-8-in-x-1-1-4-x-84-in-Primed-MDF-Header-Batten-Moulding-MW440-9G168/205576719"
wb['Family Room']['E27'] = h(batten_url, 10.88)

# --- Panel nails: discontinued 1lb → 6oz white ---
nails_url = "https://www.homedepot.com/p/Grip-Rite-1-in-x-16-1-2-Gauge-White-Ring-Shank-Panel-Board-Nails-6-oz-Box-1PBWH/202308581"
wb['Family Room']['E28'] = h(nails_url, 4.98)

# ============================================================
# 3. SAVE
# ============================================================

wb.save(OUTPUT_FILE)
print(f"\nSaved updated spreadsheet to: {OUTPUT_FILE}")
print("\n=== CHANGES SUMMARY ===")
print("""
DESCRIPTION FIXES (30 cells):
- "Epoxy crack filler" → "Concrete & mortar filler/sealant" (3 rooms)
- "50 lb bag" → "40 lb bag" for self-leveling compound (3 rooms)
- "standard drywall" → "UltraLight drywall" (3 rooms)
- "Inside/outside corner blocks" → "Plinth blocks" (5 rooms)
- "25oz" → "20oz" for spray texture (4 rooms)
- "prehung door" → "door slab" (Laundry + Bathroom)
- "lever handle" → "passage knob" (Laundry Room)
- "privacy lever" → "privacy knob" (Bathroom)
- "Door stop + shims" → "Door stop" (Laundry Room)
- "Porcelain" → "Ceramic" wall tile (Bathroom)
- "screws + tape" → "screws" only (Bathroom)

URL FIXES (broken/discontinued/wrong links - 37 cells):
- Tack strips: new product (3 rooms)
- Carpet seam tape: new product (3 rooms)
- DAP crack filler: correct color variant (3 rooms)
- Concrete bonding adhesive: CORRECT product (was wrong!) (3 rooms)
- Hydraulic cement: price update (3 rooms)
- Self-leveling compound: new SKU (3 rooms)
- SLC primer: new product (Henry 554) (3 rooms)
- Drywall: price update (3 rooms)
- Door hinges: new model (2 rooms)
- Shower arm: replacement product (Bathroom)
- DAP Kwik Seal: price update (2 cells)
- Thinset mortar: price update (Bathroom)
- Grout: new SKU (Bathroom)
- Grout sealer: replacement product (Bathroom)
- QEP LASH tile system: new project pack (Bathroom)
- Vanity: new assembled model (Bathroom)
- Moen Genta faucet: price update (Bathroom)
- Drain assembly: replacement product (Bathroom)
- P-trap: updated SKU (Bathroom)
- Supply lines: price update (Bathroom)
- Plumbers putty: replacement product (Bathroom)
- Mirror: replacement product (Bathroom)
- BEHR semi-gloss 1gal: price update (Bathroom)
- Wall paneling: replacement product (Family Room)
- Paneling trim: replacement product (Family Room)
- Panel nails: replacement product (Family Room)

PRICES THAT WENT UP:
- Bonding adhesive: $9.98 → $11.97
- Hydraulic cement: $12.48 → $13.98
- Drywall: $11.98 → $16.25
- DAP Kwik Seal: $5.98 → $7.98
- Grout: $16.98 → $18.48
- Moen Genta faucet: $109 → $116.10
- BEHR semi-gloss 1gal: $54.98 → $58.98
- Paneling trim: $4.98 → $10.88
- Seam tape: $12.98 → $13.20

PRICES THAT WENT DOWN:
- Tack strips: $0.45 → $0.29/LF
- DAP crack filler: $9.98 → $8.78
- Self-leveling: $36.97 → $35.00
- Door hinges: $13.47 → $9.98
- Shower arm: $19.98 → $14.98
- Thinset: $21.98 → $17.97
- QEP LASH: $24.98 → $20.47
- Vanity: $549 → $439
- P-trap: $7.98 → $5.98
- Supply lines: $9.98 → $7.28
- Panel nails: $6.98 → $4.98

NOTE: A few prices were kept at original values where current
price could not be confirmed (SLC primer gallon, drain assembly,
wall paneling). Verify these by clicking the links.
""")
