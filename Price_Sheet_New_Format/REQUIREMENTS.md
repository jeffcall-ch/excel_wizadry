# Price Sheet Generator — Requirements & Business Rules

## Goal

Build a fully working **Excel-based solution** that transforms Bill of Material (BoM) input files into a formatted price sheet. The final deliverable is an `.xlsx` workbook that a supplier receives, fills in unit prices, and returns.

The output **does not have to match the current format exactly** — the goal is correctness and usability. A different layout is acceptable if it works better.

The **Excel file setup may be done programmatically** (e.g. Python + openpyxl), but once created, the workbook must work purely in Excel — no VBA macros, no external scripts at runtime. Power Query is acceptable if it remains inside the Excel file. The user has **Microsoft 365**.

---

## Current Workflow (what we're replacing)

Today, the price sheet is built manually through these steps:

1. Open each BoM input file, copy the "Total Piping Material" tab data into a pivot table
2. Run a VBA macro to calculate spare quantities
3. Manually assemble the painting section (CS TUBI feedback, paint cans)
4. Manually add blind disks, flange guards, welding nozzles, orifice plates
5. Compile everything into the price sheet format

The goal is to automate steps 1–4 completely. Orifice plates remain manual.

---

## Input Files

All inputs are `.xlsm` files with macro-generated BoM data. They come in revisions (rev0, rev1, etc.).

### Location

```
Price_Sheet_New_Format/
  rev0/input/    — revision 0 BoM files
  rev1/input/    — revision 1 BoM files
```

### Input file types

**3 main material BoMs** (each has "Total Piping Material" + "KKS Piping Material" sheets):
- Stainless Steel (SS)
- Carbon Steel (CS)
- MAPRESS (Geberit press-fit)

**3 erection material BoMs** (each has "Total Piping Material" sheet):
- SS Erection
- CS Erection
- MAPRESS Erection

**1 painting BoM** (has "PIPING PAINTING MATERIAL" sheet):
- Piping painting material

### Reference files

- Existing VBA spare logic: `Spare_calcs_for_price_sheet/spare_calcs_main_material.VBA`
- Existing VBA erection spare logic: `Spare_calcs_for_price_sheet/spare_calcs_erection_material.VBA`
- Example output (rev0): `rev0/output_current_format/` — the finished price sheet for reference
- Revision comparison script: `Price_Sheet_Excel_Compare/price_sheet_compare.py`

---

## Business Rules

### Spare Calculation — Main Material (SS, CS, MAPRESS)

Spare quantities are determined by a **priority chain**: SEAL > CRITICAL > TUBI > STANDARD. The first matching rule wins.

**SEAL components** — description contains ALL of: "seal", "ring", "fkm" AND NONE of: "elbow", "flange", "tee", "reducer"
- QTY ≤ 10 → spare = 5
- QTY > 10 → spare = MAX(5, CEILING(QTY × 10%))

**CRITICAL components** — description contains ANY of: "nipple", "union", "connection", "welding", "adaptor", "coupling", "threaded"
- QTY ≤ 10 → spare = MAX(3, CEILING(QTY × 50%))
- QTY > 10 → spare = MAX(3, CEILING(QTY × 10%))

**TUBI (pipes)** — Type column = "TUBI"
- DN < 50: spare base = CEILING(QTY_m × 15%)
- DN ≥ 50: spare base = CEILING(QTY_m × 10%)
- Total order length = CEILING((QTY_m + spare_base), 6) — rounded up to next multiple of 6 meters
- Spare = total order length − QTY_m

**STANDARD (everything else)**
- Pieces (QTY_pcs):
  - QTY ≤ 5 → spare = 0
  - QTY 6–20 → spare = 1
  - QTY > 20 and DN < 50 → spare = CEILING(QTY × 7.5%)
  - QTY > 20 and DN ≥ 50 → spare = CEILING(QTY × 5%)
- Meters (QTY_m, non-TUBI items with length):
  - DN < 50 → CEILING(QTY_m × 15%)
  - DN ≥ 50 → CEILING(QTY_m × 10%)

### Spare Calculation — Erection Material

Priority chain: NUT/WASHER > BOLT > GASKET > STANDARD.

**NUT/WASHER** — description contains "nut" or "washer"
- QTY ≤ 30 → base = 100%
- QTY 31–100 → base = 50%
- QTY 101–400 → base = 35%
- QTY > 400 → base = 20%
- Rounding the total (QTY + spare_base): ≤100 → round to next 10, 101–500 → round to next 50, >500 → round to next 100
- Spare = rounded total − QTY

**BOLT** — description contains "bolt"
- Same percentage tiers as nut/washer
- Always round total to next 10 (regardless of size)
- Spare = rounded total − QTY

**GASKET** — description contains "gasket"
- QTY ≤ 30 → spare = CEILING(QTY × 100%)
- QTY 31–500 → spare base = 50%, round total to next 10, spare = rounded − QTY
- QTY > 500 → spare base = 30%, round total to next 10, spare = rounded − QTY

**STANDARD (all others)**
- QTY < 400 → spare = CEILING(QTY × 35%)
- QTY ≥ 400 → spare = CEILING(QTY × 20%)

### CS TUBI ↔ Painting Feedback Loop

This is a critical interaction between the painting BoM and the Carbon Steel output:

1. The painting BoM lists every CS pipe by description (e.g. "Tube 33.7x2.6 EN10216-2") with length and surface
2. Group painting TUBI rows by pipe description + painting colour
3. For each group: apply spare % based on DN (same as TUBI spare rule), then round up to next multiple of 6 meters
4. Sum the rounded lengths per pipe description across all colours → this becomes the CS TUBI "For Order" quantity
5. The CS output sheet's TUBI spare is NOT calculated by the normal TUBI formula — instead, spare = painting for-order − CS base QTY

This ensures CS pipe order quantities match what painting needs after rounding.

### Painting — Surface and Colour

- Painting colour comes from the painting BoM column "Painting colour (acc. BS4800)" (column 16 in the actual file)
- Each painting row has an external surface value in m²
- The painting section of the output shows pipes grouped by description + colour, with surface calculated from for-order length × (surface per meter ratio)

### Paint Cans

Paint comes in cans for site touch-up and fitting painting:

- **Touch-up (pipes/TUBI)**: 5% of the TUBI pipe surface area
- **Fittings**: 100% of the fitting (non-TUBI) surface area
- **Primer**: needs to cover ALL surfaces (both touch-up and fittings). For "basic only" colour items, the primer surface is doubled (since there's no final coat, primer is all they get — so double coverage)
- **Conversion**: 1 litre covers 9 m²
- **Rounding**: round litre quantity up to next multiple of 20
- **Large quantity rule**: if the raw litre amount exceeds 15 litres, buy CEILING(litres, 20) + 20 (i.e. an extra 20L can)

### Blind Disks (Section 6A)

Blind disks are for hydrotesting. Quantity depends on total flange count across all material BoMs (SS + CS + MAPRESS), for flange types FLAN and FBLI:

- Total flanges > 50 → order 10 blind disks
- Total flanges 30–50 → order 4
- Total flanges ≤ 30 → order 2

Each DN size has a fixed S1 thickness (from a lookup table).
Material: S235JRG2/1.0038 or equal. PN16 between Flanges EN 1092-1 Form AA.

### Flange Guards (Section 6B)

Steel flange guards are needed only for specific systems:
- **Fuel oil** systems
- **Ammonia** (ammonia water) systems
- **Urea** systems

The system name comes from the KKS Piping Material sheet "System" column (column 1). Filter flanges (Type = "FLAN") where the system name matches any of the above keywords.

For each unique (Material, DN) combination of matching flanges, the for-order quantity = flange count from BoQ + 1.

### Welding Nozzles (Section 6C)

Currently quantity = 0. Placeholder for "Welding Nozzle for Temperature Measurement Weldolet for Thermowell OD40/ID25.5 L=50 EN 10222-1", material P235GH or equal.

### Orifice Plates

These are **manually calculated** — not automated. The solution should leave space/placeholder for them but not attempt to calculate.

---

## Output Requirements

### Price Sheet Sections

The price sheet should contain these sections (whether as separate sheets or sections within sheets is flexible):

1. **Mapress material** — sorted/grouped by material + type + DN + description, with BOM qty, spare, for-order qty, weight, unit price (blank for supplier), total price
2. **Stainless Steel material** — same structure
3. **Carbon Steel material** — same structure, but TUBI for-order comes from painting feedback
4. **Erection material** — combined from all 3 erection BoMs, with erection spare rules
5. **Painting** — CS pipe painting by description + colour, surface areas, plus paint cans section
6. **Additional parts** — blind disks, flange guards, welding nozzles, orifice plate placeholders
7. **Price summary** — grand totals from all sections above

### Supplier Interaction

- Unit price columns must be clearly marked and left blank for the supplier to fill in
- Total price = for-order quantity × unit price
- Currency field should be editable
- Grand totals should sum all total prices

### Revision Support

The input BoM files contain "QTY prev. rev." and "Difference" columns. Ideally the solution should support:
- Generating the first revision price sheet
- Updating with a new revision and highlighting what changed

---

## Constraints

- Must work in **Microsoft 365 Excel** (desktop, not web)
- No VBA at runtime
- No external dependencies at runtime
- The workbook setup CAN be done via Python/scripting — that's a one-time generation step
- Power Query is acceptable if embedded in the workbook
