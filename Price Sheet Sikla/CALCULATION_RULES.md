# Price Sheet Calculation Rules

Source file: `CA100-KVI-50296373_0.0 - BoM General Hangers and support material.xlsm`, sheet `Total Qty`

---

## Material Type Classification

Items are classified into 5 base groups (plus 5 C5 variants) by matching description keywords/prefixes in order:

| Group | Label | Triggers |
|-------|-------|----------|
| 04 | Beam Sections | prefixes `TP F`, `FLS`, `MS`; keywords `channel`, `rail`, `profile`, `beam section ms`, `beam section tp` |
| 02 | Brackets Consoles | prefixes `WD F`, `WBD`, `SA F`; keywords `bracket`, `console`, `base`, `pivot joint`, `connector`, `mounting angle`, `beam clip`, `beam connection`, `holder`, `eye`, `plate`, `end support`, `adapter`, `joining` |
| 01 | Primary Supports | prefixes `Stabil`, `Ratio`; keywords `clamp`, `hanger`, `pipe clamp`, `assembly set`, `fixed point`, `guided`, `sliding`, `guiding`, `assembly`, `axial`, `threaded rod` |
| 03 | Bolts Screws Nuts | prefixes `HZM`, `SK`, `Formlock`; keywords `bolt`, `screw`, `nut`, `washer`, `resin anchor` |
| 05 | Installation Material | keywords `glass fabric`, `end cap` |

**C5 remapping**: if `"C5"` is in the coating value, the group is remapped:
01→06, 02→07, 03→08, 04→09, 05→10.

---

## Aggregation Routing

Each raw row is routed to one of six paths before quantity/weight summation:

| Path | Condition | Cut length cleared? |
|------|-----------|---------------------|
| `rod_groups` | Group 01 + `"threaded rod"` in description + cut length present | yes (aggregated by article+coating) |
| `beam_groups` | Group 04 + `"beam section ms"/"beam section tp"` in description + cut length present | yes (aggregated by article+coating) |
| `tape_groups` | Group 05 + `"glass fabric"` in description + cut length present | yes (aggregated by article+coating) |
| `bracket_groups` | Group 02, not angled | yes (aggregated by article+coating+cleaned description) |
| `individual` | Group 02 angled brackets, or any item with cut length that didn't match above | no — kept individually |
| `groups` | Everything else (no cut length) | n/a — no cut length to begin with |

**Bracket description cleaning**: the suffix ` x NNN[suffix]` (e.g. ` x 250 HCP`) is stripped before aggregation. The coating suffix (e.g. `HCP`) is preserved in the `Coating` column.

**Final de-duplication pass**: after all paths are assembled, rows sharing the same `(Article_Number, Coating, Cut_Length_mm, Remarks)` key are merged by summing `Qty`, `Total_Weight_kg`, `Spare`, `Order_Qty`, `Order_Weight_kg`, and `Net_Cut_Weight_kg`.

---

## Column Semantics

| Column | Meaning |
|--------|---------|
| `Qty` | Net quantity needed (no spare). For cut items: minimum whole bars/rolls. |
| `Total_Weight_kg` | Net weight of `Qty` items (no spare). For cut items: net cut weight only. |
| `Spare` | Extra units ordered on top of `Qty`. `NULL` for standard items with no spare rule yet defined. |
| `Order_Qty` | Total units to order = `Qty + Spare`. |
| `Order_Weight_kg` | Weight of `Order_Qty` (scaled proportionally from `Total_Weight_kg`). |
| `Net_Cut_Length_m` | Sum of all piece lengths needed (cut items only). |
| `Net_Cut_Weight_kg` | Weight of net cut material only (= `Total_Weight_kg` for cut items). |
| `Total_Order_Length_m` | Total bar/roll length ordered (= `Order_Qty × bar/roll length`). |
| `Utilisation_pct` | `Net_Cut_Length_m / Total_Order_Length_m × 100`. |

---

## Spare Rules by Group

### 01 Primary Supports — Threaded Rods

Cut-length items with `"threaded rod"` in description. Uses **First Fit Decreasing (FFD) packing** algorithm.

- Bar length: parsed from description suffix `/NNNN` (e.g. `M10/2000` → 2000 mm physical)
- Effective bar length: `physical_mm` (no end-loss scrap)
- Kerf: `3 mm` per cut
- `Qty` = `ceil(net_cut_mm / physical_mm)` — naive minimum bars
- FFD packs pieces into effective bars; result = minimum bar count
- `Order_Qty` = `ceil(ffd_bar_count × physical_mm × 1.10 / physical_mm)` — 10% buffer, rounded up to whole bars
- `Spare` = `Order_Qty − Qty`
- `Order_Weight_kg` = `Total_Weight_kg × (Total_Order_Length_m / Net_Cut_Length_m)`

### 01 Primary Supports — All other items (clamps, hangers, assemblies, etc.)

**Spare: 5%** — `Spare = ceil(Qty × 0.05)`, `Order_Qty = Qty + Spare`

### 02 Brackets Consoles

**Spare: 5%** — `Spare = ceil(Qty × 0.05)`, `Order_Qty = Qty + Spare`

### 03 Bolts Screws Nuts

**Spare: 15%** — `Spare = ceil(Qty × 0.15)`, `Order_Qty = Qty + Spare`

### 04 Beam Sections — Cut items (beam section ms / beam section tp)

Items with `"beam section ms"` or `"beam section tp"` in description and a cut length. Uses **First Fit Decreasing (FFD) packing**.

- Physical bar: `6000 mm`; effective bar: `6000 mm` (no end-loss scrap)
- Kerf: `3 mm` per cut
- `Qty` = `ceil(net_cut_mm / 6000)` — naive minimum bars
- `Order_Qty` = `ceil(ffd_bar_count × 1.10)` — 10% buffer, rounded up to whole bars
- `Spare` = `Order_Qty − Qty`
- `Order_Weight_kg` = `Total_Weight_kg × (Total_Order_Length_m / Net_Cut_Length_m)`

### 04 Beam Sections — Channel rows

Items where `"channel"` appears in the description (no cut-length optimiser).

**Spare: 10%** — `Spare = ceil(Qty × 0.10)`, `Order_Qty = Qty + Spare`

### 05 Installation Material — Glass Fabric Tape

Cut-length items with `"glass fabric"` in description. Aggregated by article + coating.

- Roll length: `10 m`
- `Qty` = `ceil(net_cut_m / 10)` — naive minimum rolls
- `Order_Qty` = `ceil(net_cut_m × 1.30 / 10)` — 30% spare factor
- `Spare` = `Order_Qty − Qty`
- `Order_Weight_kg` = `Total_Weight_kg × (Total_Order_Length_m / Net_Cut_Length_m)`
- `Remarks` = `"10m roll"`

### 05 Installation Material — End Caps (and other non-tape items)

**Spare: 15%** — `Spare = ceil(Qty × 0.15)`, `Order_Qty = Qty + Spare`

---

## Weight Scaling Formula

For all cut-to-length items (beams, rods, tape), ordered weight is scaled proportionally:

```
Order_Weight_kg = Total_Weight_kg × (Total_Order_Length_m / Net_Cut_Length_m)
```

For percentage-spare items, ordered weight is scaled proportionally from net weight:

```
Order_Weight_kg = Total_Weight_kg × (Order_Qty / Qty)
```

---

## Constants (code)

| Constant | Value | Purpose |
|----------|-------|---------|
| `_KERF_MM` | `3` | Saw-cut kerf per cut |
| `_ORDER_BUFFER` | `1.10` | 10% workshop safety buffer for FFD items |
| `_BEAM_PHYSICAL_MM` | `6000` | Beam bar physical length (mm) |
| `_BEAM_BAR_MM` | `6000` | Beam effective cut length (mm) — no end-loss |
| `_TAPE_ROLL_M` | `10` | Tape roll length (m) |
| `_TAPE_SPARE_FACTOR` | `1.30` | 30% spare factor for tape |
| `_SPARE_PCT["01 Primary Supports"]` | `0.05` | 5% for non-rod primary support items |
| `_SPARE_PCT["02 Brackets Consoles"]` | `0.05` | 5% |
| `_SPARE_PCT["03 Bolts Screws Nuts"]` | `0.15` | 15% |
| `_SPARE_PCT["05 Installation Material"]` | `0.15` | 15% for end caps etc. |
| C5 variants (06/07/08/10) | same as base | identical spare rules apply |
| Channel rows in 04/09 | `0.10` | 10% spare (hardcoded, not in `_SPARE_PCT`) |

---

## Known Source Data Quirks

| Article | Description | Issue |
|---------|-------------|-------|
| 160803 | Corner Bracket EW 41 | `Total_Weight_kg` is NULL in source Excel — no weight data available |
| 190775 | Resin Anchor Rod VMZ-A 80 M12-25/125 | `Weight_kg` stored as `0.13` but actual unit weight is `~0.128 kg`; `Total_Weight_kg` was computed at higher precision before rounding — V05 WARN expected |
| 190784 | Resin Anchor Rod VMZ-A 80 M12-50/150 | Same unit-weight rounding as above — V05 WARN expected |

---

## Validation Checks

Run automatically at the end of every `test-run`, `populate`, and `export` command. Results appear in the **VALIDATION** sheet of the output xlsx.

Colour coding: 🟢 PASS · 🔴 FAIL · 🟡 WARN · 🔵 INFO

### Category 1 — Raw Data Integrity

| ID | Level | Check |
|----|-------|-------|
| V01 | FAIL | No NULL Article Numbers in raw |
| V02 | FAIL | No NULL/blank Descriptions in raw |
| V03 | FAIL | All raw Qty values > 0 |
| V04 | WARN | No NULL Total_Weight_kg in raw |
| V05 | WARN | Standard items: `Weight_kg × Qty = Total_Weight_kg` (tolerance ±0.02 kg). Cut-length rows excluded — `Weight_kg` is linear density there, not per-piece weight. Source rounding may trigger this on a small number of rows. |
| V06 | FAIL | `Cut_Length_mm > 0` whenever present |

### Category 2 — Classification Completeness

| ID | Level | Check |
|----|-------|-------|
| V07 | FAIL | All raw article numbers are present in the aggregated output |
| V08 | FAIL | No unexpected article numbers appear in the aggregated output |

### Category 3 — Aggregation Fidelity

| ID | Level | Check |
|----|-------|-------|
| V09 | FAIL | For no-cut-length items: `AGG.Qty = SUM(raw Qty)` per `Article + Coating` |
| V10 | FAIL | For FFD/tape items: `AGG.Net_Cut_Length_m × 1000 = SUM(raw Qty × Cut_Length_mm)` |
| V11 | FAIL | Grand total net weight: raw `SUM(Total_Weight_kg)` ≈ AGG `SUM(Total_Weight_kg)` (threshold < 1 kg) |

### Category 4 — Spare / Order Logic

| ID | Level | Check |
|----|-------|-------|
| V12 | FAIL | `Order_Qty >= Qty` for every row |
| V13 | FAIL | `Spare = Order_Qty − Qty` wherever `Spare` is not NULL |
| V14 | WARN | All rows in groups with a spare rule (01–03, 05–08, 10) have `Spare` set |
| V15 | FAIL | `Utilisation_pct <= 100%` for all FFD/tape rows |
| V16 | FAIL | `Total_Order_Length_m >= Net_Cut_Length_m` for FFD/tape rows |
| V17 | FAIL | `Order_Weight_kg >= Net_Cut_Weight_kg` for FFD/tape rows |

### Category 5 — Weight Sanity

| ID | Level | Check |
|----|-------|-------|
| V18 | FAIL | No negative `Total_Weight_kg` in aggregated output |
| V19 | FAIL | No negative `Order_Weight_kg` in aggregated output |
| V20 | FAIL | For percentage-spare items: `Order_Weight_kg / Order_Qty ≈ Total_Weight_kg / Qty` (tolerance ±0.05 kg/unit) |
| V21 | WARN | No NULL `Total_Weight_kg` in aggregated output |

### Category 6 — AGG Data Quality

| ID | Level | Check |
|----|-------|-------|
| V22 | FAIL | No duplicate `(Article_Number, Coating, Cut_Length_mm, Remarks)` keys |
| V23 | WARN | No zero-Qty rows in aggregated output |
| V24 | FAIL | Every row has a non-empty `Spare_Calculation_Rule` |
| V25 | WARN | 04/09 Beam Section rows that are neither FFD cut items nor channel items have a spare assigned |

### Category 7 — Summary (INFO only)

| ID | Content |
|----|----------|
| V26 | Raw row count vs AGG row count |
| V27 | Grand totals: Net Qty / Spare / Order Qty |
| V28 | Grand totals: Net Weight (kg) / Order Weight (kg) |
| V29 | Per-material-type breakdown table: rows, net qty, spare, order qty, net weight, order weight |
