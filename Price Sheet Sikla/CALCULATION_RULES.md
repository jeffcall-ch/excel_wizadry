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
- Effective bar length: `physical_mm − 100 mm` (thread engagement allowance)
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

- Physical bar: `6000 mm`; effective bar: `5900 mm` (`−100 mm` end-loss)
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
- `Order_Qty` = `ceil(net_cut_m × 1.50 / 10)` — 50% spare factor
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
| `_BEAM_BAR_MM` | `5900` | Beam effective cut length (mm) |
| `_TAPE_ROLL_M` | `10` | Tape roll length (m) |
| `_TAPE_SPARE_FACTOR` | `1.50` | 50% spare factor for tape |
| `_SPARE_PCT["01 Primary Supports"]` | `0.05` | 5% for non-rod primary support items |
| `_SPARE_PCT["02 Brackets Consoles"]` | `0.05` | 5% |
| `_SPARE_PCT["03 Bolts Screws Nuts"]` | `0.15` | 15% |
| `_SPARE_PCT["05 Installation Material"]` | `0.15` | 15% for end caps etc. |
| C5 variants (06/07/08/10) | same as base | identical spare rules apply |
