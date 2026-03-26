# Price Sheet Settings (Code-Read Only)

This file contains only settings that are read by code.
All sections below are parsed and used at runtime.

Section numbering follows existing code labels.

## 1. General

| Key | Value |
| --- | --- |
| shop_suffix | shop |
| erection_suffix | erection |
| paint_material_type_label | 05 Paint shop |
| aggregate_table_suffix | aggreated |

## 2. MaterialTypes

| MaterialType |
| --- |
| ss |
| cs |
| mapress |

## 3. RequiredFileCategories

| MaterialType | IsErection |
| --- | --- |
| ss | 0 |
| ss | 1 |
| cs | 0 |
| cs | 1 |
| mapress | 0 |
| mapress | 1 |

## 4. MaterialTypeLabels

| MaterialType | Label |
| --- | --- |
| ss | SS |
| cs | CS |
| mapress | Mapress |

## 5. ColumnMappings_Normal

| SourceKey | SQLiteColumn | Required |
| --- | --- | --- |
| system | System | 1 |
| pipe | KKS | 1 |
| type | Type | 1 |
| pipe_component | Description | 1 |
| dn | Size | 1 |
| material | Material | 1 |
| total_weight | Weight kg | 1 |
| qty_pcs | Qty pcs | 1 |
| qty_m | Qty m | 1 |

## 6. ColumnMappings_Erection

| SourceKey | SQLiteColumn | Required |
| --- | --- | --- |
| pipe_kks | KKS | 1 |
| type | Type | 1 |
| pipe_component | Description | 1 |
| dn | Size | 1 |
| material | Material | 1 |
| qty_pcs_per_m | Qty pcs | 1 |

## 7. ColumnMappings_Paint

| SourceKey | SQLiteColumn | Required |
| --- | --- | --- |
| name_of_system | System | 1 |
| name_of_pipe | KKS | 1 |
| type | Type | 1 |
| description | Description | 1 |
| dn_1 | Size | 1 |
| material | Material | 1 |
| quantity_pcs | Qty pcs | 1 |
| quantity_m | Qty m | 1 |
| external_surface_m2 | Surface m2 | 1 |
| corrosion_class | Corrosion category | 1 |
| painting_colour | Paint colour | 1 |
| insulated | Insulated | 1 |

## 8. Aliases

| Key | Alias |
| --- | --- |
| system | system |
| pipe | pipe |
| pipe_kks | pipe kks |
| pipe_kks | pipe |
| type | type |
| pipe_component | pipe component |
| pipe_component | pipecomponent |
| material | material |
| total_weight | total weight |
| total_weight | total weight kg |
| qty_pcs | qty pcs |
| qty_m | qty m |
| qty_pcs_per_m | qty pcs m |
| name_of_system | name of system |
| name_of_pipe | name of pipe |
| description | description |
| dn_1 | dn 1 |
| dn_1 | dn1 |
| quantity_pcs | quantity pcs |
| quantity_m | quantity m |
| external_surface_m2 | external surface m2 |
| external_surface_m2 | external surface |
| corrosion_class | corrosion class |
| insulated | insulated |

## 9. FileKeywords

| Kind | Keyword |
| --- | --- |
| paint | paint |
| paint | painting |
| orifice | orifice |
| orifice | orifices |

## 10. SqliteNumericColumns

| Column |
| --- |
| UID |
| Size |
| Qty pcs |
| Qty m |
| Weight kg |
| Surface m2 |
| Order qty |
| Weight w spare |
| Unit price |
| Total price |
| Delta qty to previous revision |

## 11. AggregateSumColumns

| Column |
| --- |
| Qty pcs |
| Qty m |
| Weight kg |
| Surface m2 |

## 12. AggregationMaterialTypeRules

| MatchType | Pattern | Output |
| --- | --- | --- |
| endswith | erection | 04 Erection |
| contains | paint shop | 05 Paint shop |

## 13. SpareGeneral

| Key | Value |
| --- | --- |
| aggregate_table_suffix | aggreated |
| spare_column | Spare |
| spare_info_column | Spare info |
| material_type_column | Material type |
| description_column | Description |
| type_label_column | Type label |
| type_column | Type |
| size_column | Size |
| qty_pcs_column | Qty pcs |
| qty_m_column | Qty m |
| weight_column | Weight kg |
| total_qty_column | Order qty |
| weight_w_spare_column | Weight w spare |
| paint_prefix | paint |
| erection_exact | 04 Erection |
| shop_suffix | shop |

## 14. SpareMainRules

Used for shop rows (and paint shop rows that follow main spare logic).

| Key | Value | Meaning in calculation |
| --- | --- | --- |
| seal_qty_threshold | 10 | If SEAL qty is <= this, use fixed low spare. |
| seal_min_spare_low | 5 | Fixed spare for low-qty SEAL rows. |
| seal_min_spare_high | 5 | Minimum spare for high-qty SEAL rows. |
| seal_percent | 0.10 | High-qty SEAL percent (10%). |
| critical_qty_threshold | 10 | Split point for CRITICAL low/high logic. |
| critical_min_spare | 3 | Minimum spare for CRITICAL rows. |
| critical_percent_low | 0.50 | CRITICAL percent when qty <= threshold (50%). |
| critical_percent_high | 0.10 | CRITICAL percent when qty > threshold (10%). |
| threshold_1 | 5 | Standard row lower threshold. |
| threshold_2 | 20 | Standard row middle threshold. |
| size_threshold | 50 | DN split between small and large percentages. |
| percent_pcs_small | 0.075 | Standard pcs spare for small DN (7.5%). |
| percent_pcs_large | 0.05 | Standard pcs spare for large DN (5%). |
| percent_m_small | 0.15 | Standard meter spare for small DN (15%). |
| percent_m_large | 0.10 | Standard meter spare for large DN (10%). |
| tubi_round_multiple | 6 | TUBI total order qty is rounded up to this multiple. |

## 15. SpareErectionRules

Used for erection rows only.

| Key | Value | Meaning in calculation |
| --- | --- | --- |
| nwb_threshold_1 | 30 | Nut/Washer first threshold. |
| nwb_threshold_2 | 100 | Nut/Washer second threshold. |
| nwb_threshold_3 | 400 | Nut/Washer third threshold. |
| nwb_percent_low | 1.0 | Nut/Washer spare percent for lowest range (100%). |
| nwb_percent_mid | 0.5 | Nut/Washer spare percent for mid range (50%). |
| nwb_percent_high | 0.35 | Nut/Washer spare percent for high range (35%). |
| nwb_percent_very_high | 0.2 | Nut/Washer spare percent for very high range (20%). |
| gasket_threshold_1 | 30 | Gasket first threshold. |
| gasket_threshold_2 | 500 | Gasket second threshold. |
| gasket_percent_low | 1.0 | Gasket spare percent low range (100%). |
| gasket_percent_mid | 0.5 | Gasket spare percent mid range (50%). |
| gasket_percent_high | 0.3 | Gasket spare percent high range (30%). |
| standard_threshold | 400 | Non-special erection threshold. |
| standard_percent_low | 0.35 | Non-special erection spare below threshold (35%). |
| standard_percent_high | 0.2 | Non-special erection spare at/above threshold (20%). |

## 16. SpareKeywords

How keyword matching works:
- Matching is case-insensitive.
- `seal_required` needs all listed words to be present.
- `seal_excluded` blocks SEAL logic if any listed word is present.
- All other groups trigger when any listed word is present.
- Classification order in code is: `seal` -> `critical` -> `tubi` -> standard.

| Name | Keywords | What it means |
| --- | --- | --- |
| seal_required | seal,ring,fkm | Row is treated as SEAL only when all words exist in description. |
| seal_excluded | elbow,flange,tee,reducer | Even if seal_required matches, these words cancel SEAL class. |
| critical | nipple,nipples,union,connection,connections,welding,adaptor,coupling,threaded | Any match makes row CRITICAL (if not already SEAL). |
| tubi | tubi | Any match makes row TUBI (if not already SEAL/CRITICAL). |
| nwb | nut,washer | Erection class for nut/washer logic. |
| bolt | bolt | Erection class for bolt logic. |
| gasket | gasket | Erection class for gasket logic. |

## 17. SpareRuleGuide

Documentation-only section: this explains the rule flow and formulas used by the code.

Rule priority (main/shop):
- First matching class wins: SEAL -> CRITICAL -> TUBI -> STANDARD.

Rule priority (erection):
- First matching class wins: NWB -> BOLT -> GASKET -> STANDARD.

| Rule ID | Applies To | Trigger | Spare Formula | Rounding / Notes |
| --- | --- | --- | --- | --- |
| MAIN#1 | Main/Shop | SEAL and qty <= seal_qty_threshold | spare = seal_min_spare_low | Fixed integer spare. |
| MAIN#2 | Main/Shop | SEAL and qty > seal_qty_threshold | spare = max(seal_min_spare_high, ceil(qty * seal_percent)) | Uses ceiling on percent result. |
| MAIN#3 | Main/Shop | CRITICAL and qty <= critical_qty_threshold | spare = max(critical_min_spare, ceil(qty * critical_percent_low)) | Uses ceiling on percent result. |
| MAIN#4 | Main/Shop | CRITICAL and qty > critical_qty_threshold | spare = max(critical_min_spare, ceil(qty * critical_percent_high)) | Uses ceiling on percent result. |
| MAIN#T1 | Main/Shop | TUBI match | base = ceil(qty * pct); total = round_up_to_multiple(qty + base, tubi_round_multiple); spare = total - qty | pct = percent_m_small when 0 < DN < size_threshold, else percent_m_large. |
| MAIN#5 | Main/Shop | STANDARD pcs and qty <= threshold_1 | spare = 0 | Pieces path only. |
| MAIN#6 | Main/Shop | STANDARD pcs and threshold_1 < qty <= threshold_2 | spare = 1 | Pieces path only. |
| MAIN#7a | Main/Shop | STANDARD pcs and qty > threshold_2 and DN < size_threshold | spare = ceil(qty * percent_pcs_small) | Pieces path only. |
| MAIN#7b | Main/Shop | STANDARD pcs and qty > threshold_2 and DN >= size_threshold | spare = ceil(qty * percent_pcs_large) | Pieces path only. |
| MAIN#8a | Main/Shop | STANDARD meter and DN < size_threshold | spare = ceil(qty_m * percent_m_small) | Meter path is used when Qty pcs is 0. |
| MAIN#8b | Main/Shop | STANDARD meter and DN >= size_threshold | spare = ceil(qty_m * percent_m_large) | Meter path is used when Qty pcs is 0. |
| ER#1 | Erection | NWB and qty <= nwb_threshold_1 | base = ceil(qty * nwb_percent_low); total = round_up_nut_washer(qty + base); spare = total - qty | NWB rounding steps: <=100 to 10, 101-500 to 50, >500 to 100. |
| ER#2 | Erection | NWB and nwb_threshold_1 < qty <= nwb_threshold_2 | base = ceil(qty * nwb_percent_mid); total = round_up_nut_washer(qty + base); spare = total - qty | Same NWB rounding steps. |
| ER#3 | Erection | NWB and nwb_threshold_2 < qty <= nwb_threshold_3 | base = ceil(qty * nwb_percent_high); total = round_up_nut_washer(qty + base); spare = total - qty | Same NWB rounding steps. |
| ER#4 | Erection | NWB and qty > nwb_threshold_3 | base = ceil(qty * nwb_percent_very_high); total = round_up_nut_washer(qty + base); spare = total - qty | Same NWB rounding steps. |
| ER#1B | Erection | BOLT and qty <= nwb_threshold_1 | base = ceil(qty * nwb_percent_low); total = round_up_to_next_10(qty + base); spare = total - qty | Bolt totals are rounded to next 10. |
| ER#2B | Erection | BOLT and nwb_threshold_1 < qty <= nwb_threshold_2 | base = ceil(qty * nwb_percent_mid); total = round_up_to_next_10(qty + base); spare = total - qty | Bolt totals are rounded to next 10. |
| ER#3B | Erection | BOLT and nwb_threshold_2 < qty <= nwb_threshold_3 | base = ceil(qty * nwb_percent_high); total = round_up_to_next_10(qty + base); spare = total - qty | Bolt totals are rounded to next 10. |
| ER#4B | Erection | BOLT and qty > nwb_threshold_3 | base = ceil(qty * nwb_percent_very_high); total = round_up_to_next_10(qty + base); spare = total - qty | Bolt totals are rounded to next 10. |
| ER#5 | Erection | GASKET and qty <= gasket_threshold_1 | spare = ceil(qty * gasket_percent_low) | Integer ceiling, no extra rounding stage. |
| ER#6 | Erection | GASKET and gasket_threshold_1 < qty <= gasket_threshold_2 | base = ceil(qty * gasket_percent_mid); total = round_up_to_next_10(qty + base); spare = total - qty | Total rounded to next 10. |
| ER#7 | Erection | GASKET and qty > gasket_threshold_2 | base = ceil(qty * gasket_percent_high); total = round_up_to_next_10(qty + base); spare = total - qty | Total rounded to next 10. |
| ER#8 | Erection | STANDARD and qty < standard_threshold | spare = ceil(qty * standard_percent_low) | Used when not NWB/BOLT/GASKET. |
| ER#9 | Erection | STANDARD and qty >= standard_threshold | spare = ceil(qty * standard_percent_high) | Used when not NWB/BOLT/GASKET. |

## 18. BlindDiskThickness

Used for both blind disks and orifice plate thickness lookup by DN.

| DN | S1 mm |
| --- | --- |
| 10 | 6 |
| 15 | 6 |
| 20 | 6 |
| 25 | 6 |
| 32 | 8 |
| 40 | 8 |
| 50 | 10 |
| 65 | 10 |
| 80 | 12 |
| 100 | 12 |
| 125 | 12 |
| 150 | 12 |
| 200 | 15 |
| 250 | 20 |

## 19. PaintErectionDefaults

| Column | Value |
| --- | --- |
| Material type | 06 Paint erection |
| Status | N/A |
| Material | Paint |
| Type | PAINT |
| Description | Touch up paint 20L can - {paint_colour} (fittings 100% + TUBI 5%; colored includes 2x basic layer). |
| Size | N/A |
| Paint colour | {paint_colour} |
| Corrosion category | N/A |
| Add. Info | N/A |
| Qty m | N/A |
| Weight kg | N/A |
| Spare | N/A |
| Spare info | Paint erection touch-up order. Surface={surface_m2} m2; liters=surface/5={liters}; 20L cans={cans}; rule: +1 extra can when liters in not-full can >= 15 (current={liters_in_not_full_can}). |

## 20. BlindDiskDefaults

| Column | Value |
| --- | --- |
| Material type | 07 Blind disks |
| Status | N/A |
| Material | S235JRG2/1.0038 or equal |
| Type | BLIND DISK |
| Size | {size} |
| Description | Blind disk between flanges EN 1092-1 Form AA (unfinished) PN16 for hydrostatic pressure test. See TSD General Piping Material; Additional Parts. |
| Qty m | N/A |
| Paint colour | N/A |
| Corrosion category | N/A |
| Spare info | Collected flange qty for DN {dn}: {flange_qty}; blind disk rule applied (>50 -> 10, >30 -> 4, else 2) => order qty {order_qty}. |
| Add. Info | Blind disk thickness S1={thickness} mm (DN {dn}). |
| Add. Info Missing Thickness | Blind disk thickness S1=N/A (DN {dn}). |

## 21. FlangeGuardSystems

Only FLAN rows whose System contains one of these keywords are used for calculated 09 Flange Guards rows.

| SystemKeyword |
| --- |
| fuel oil |
| ammonia |
| urea |

## 22. FlangeGuardDefaults

Defaults/templates for calculated 09 Flange Guards rows.

| Column | Value |
| --- | --- |
| Material type | 09 Flange Guards |
| Status | N/A |
| Material | Flange material {material} |
| Type | FLANGE GUARD |
| Description | Flange guard. See TSD General Piping Material; Additional Parts. |
| Qty m | N/A |
| Weight kg | N/A |
| Corrosion category | N/A |
| Paint colour | N/A |
| Add. Info | N/A |
| Spare | N/A |
| Spare info | Collected base rows where Type contains FLAN (paint rows excluded) for target systems. System={system}; Material={material}; Size={size}; Base flange qty={flange_qty}; Spare=ceil(5% of base), min 1 => {spare_qty}; Order qty={order_qty}. |

## 23. ColumnMappings_Orifice

| SourceKey | SQLiteColumn | Required |
| --- | --- | --- |
| kks | KKS | 1 |
| nominal_pipe_size | Size | 1 |

## 24. OrificeDefaults

| Key | Value |
| --- | --- |
| material_type_label | 08 Orifice Plates |
| type_value | ORIFICE PLATE |
| description | Orifice Plate between flanges EN-1092-1/11/B1/PN16 |
| material_value | 1.4571 or 1.4404 |
| add_info_template | Orifice plate thickness S1={thickness} mm (DN {dn}). |
| add_info_missing_template | Orifice plate thickness S1=N/A (DN {dn}). |
| qty_per_row | 1 |

## 25. MapressSealOrderSize

Order number to DN size mapping for generated MAPRESS FKM seal rows in aggregation.

| OrderNo | Size |
| --- | --- |
| 90883 | 15 |
| 90884 | 20 |
| 90885 | 25 |
| 90886 | 32 |
| 90887 | 40 |
| 90888 | 50 |
| 90891 | 65 |
| 90892 | 80 |
| 90893 | 100 |
