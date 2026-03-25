# Price Sheet Settings

This file controls import classification, column mapping, aggregation, and spare calculations.

How to edit safely:
- Keep section numbers and titles unchanged.
- Keep table headers unchanged.
- Add or update rows in tables.
- Booleans accept: 1/0, true/false, yes/no.

## 1. General

Core labels and suffixes used throughout import and aggregation.

| Key | Value | Description |
| --- | --- | --- |
| shop_suffix | shop | Suffix appended to base material labels for shop rows. |
| erection_suffix | erection | Suffix appended to base material labels for erection rows. |
| paint_material_type_label | Paint shop | Material type written for paint source rows. |
| aggregate_table_suffix | aggreated | Suffix appended to revision table name to create aggregated table. |

## 2. MaterialTypes

Allowed base material families used for required file detection.

| MaterialType | Description |
| --- | --- |
| ss | Stainless steel family token in source filename. |
| cs | Carbon steel family token in source filename. |
| mapress | Mapress family token in source filename. |

## 3. RequiredFileCategories

Defines mandatory source file categories that must exist before import starts.

| MaterialType | IsErection | Description |
| --- | --- | --- |
| ss | 0 | SS shop material file is required. |
| ss | 1 | SS erection material file is required. |
| cs | 0 | CS shop material file is required. |
| cs | 1 | CS erection material file is required. |
| mapress | 0 | Mapress shop material file is required. |
| mapress | 1 | Mapress erection material file is required. |

## 4. MaterialTypeLabels

Display label per base material type.

| MaterialType | Label | Description |
| --- | --- | --- |
| ss | SS | Prefix for generated material type value. |
| cs | CS | Prefix for generated material type value. |
| mapress | Mapress | Prefix for generated material type value. |

## 5. ColumnMappings_Normal

Column mapping for non-erection (shop) material files.

| SourceKey | SQLiteColumn | Required | Description | Example |
| --- | --- | --- | --- | --- |
| system | System | 1 | Source system column. | SYSTEM |
| pipe | KKS | 1 | Source pipe/KKS column. | 10-AB-123 |
| type | Type | 1 | Source type column. | TUBI |
| pipe_component | Description | 1 | Source description column. | Elbow 90 |
| dn | Size | 1 | Source DN column; exact DN header required. | DN 50 |
| material | Material | 1 | Source material column. | 1.4571 |
| total_weight | Weight kg | 1 | Source total weight column. | 120.5 |
| qty_pcs | Qty pcs | 1 | Source quantity in pieces. | 25 |
| qty_m | Qty m | 1 | Source quantity in meters. | 12.8 |

## 6. ColumnMappings_Erection

Column mapping for erection material files.

| SourceKey | SQLiteColumn | Required | Description | Example |
| --- | --- | --- | --- | --- |
| pipe_kks | KKS | 1 | Erection source KKS column. | 10-AB-123 |
| type | Type | 1 | Erection source type column. | BOLT |
| pipe_component | Description | 1 | Erection source description column. | Hex bolt |
| dn | Size | 1 | Erection source DN column; exact DN header required. | DN 50 |
| material | Material | 1 | Erection source material column. | A2 |
| qty_pcs_per_m | Qty pcs | 1 | Erection quantity source mapped to Qty pcs. | 40 |

## 7. ColumnMappings_Paint

Column mapping for paint files.

| SourceKey | SQLiteColumn | Required | Description | Example |
| --- | --- | --- | --- | --- |
| name_of_system | System | 1 | Paint system column. | SYS-A |
| name_of_pipe | KKS | 1 | Paint pipe/KKS column. | 10-AB-123 |
| type | Type | 1 | Paint type column. | PAINT |
| description | Description | 1 | Paint description column. | Surface coating |
| dn_1 | Size | 1 | Paint DN column (header expected as DN 1). | DN 50 |
| material | Material | 1 | Paint material column. | CS |
| quantity_pcs | Qty pcs | 1 | Paint quantity in pieces. | 10 |
| quantity_m | Qty m | 1 | Paint quantity in meters. | 30 |
| external_surface_m2 | Surface m2 | 1 | Paint surface quantity column. | 85.4 |
| corrosion_class | Corrosion category | 1 | Paint corrosion class column. | C3 |
| painting_colour | Paint colour | 1 | Paint colour column; smart match supports color/colour. | RAL 7035 |
| insulated | Insulated | 1 | Insulation state column. | Yes |

## 8. Aliases

Alternate header names used when matching required source columns.

| Key | Alias | Description |
| --- | --- | --- |
| system | system | Alias for system source key. |
| pipe | pipe | Alias for pipe source key. |
| pipe_kks | pipe kks | Alias for erection KKS source key. |
| pipe_kks | pipe | Alias for erection KKS source key. |
| type | type | Alias for type source key. |
| pipe_component | pipe component | Alias for component source key. |
| pipe_component | pipecomponent | Alias for component source key. |
| material | material | Alias for material source key. |
| total_weight | total weight | Alias for total_weight source key. |
| total_weight | total weight kg | Alias for total_weight source key. |
| qty_pcs | qty pcs | Alias for qty_pcs source key. |
| qty_m | qty m | Alias for qty_m source key. |
| qty_pcs_per_m | qty pcs m | Alias for qty_pcs_per_m source key. |
| name_of_system | name of system | Alias for paint system source key. |
| name_of_pipe | name of pipe | Alias for paint pipe source key. |
| description | description | Alias for description source key. |
| dn_1 | dn 1 | Alias for paint DN source key. |
| dn_1 | dn1 | Alias for paint DN source key. |
| quantity_pcs | quantity pcs | Alias for paint quantity_pcs source key. |
| quantity_m | quantity m | Alias for paint quantity_m source key. |
| external_surface_m2 | external surface m2 | Alias for paint surface source key. |
| external_surface_m2 | external surface | Alias for paint surface source key. |
| corrosion_class | corrosion class | Alias for paint corrosion class source key. |
| insulated | insulated | Alias for paint insulated source key. |

## 9. FileKeywords

Keywords used for file classification.

| Kind | Keyword | Description |
| --- | --- | --- |
| paint | paint | If token starts with this, file is treated as paint candidate. |
| paint | painting | If token starts with this, file is treated as paint candidate. |

## 10. SqliteNumericColumns

Columns created as NUMERIC in SQLite.

| Column | Description |
| --- | --- |
| UID | Unique row identifier. |
| Size | Size/DN value. |
| Qty pcs | Quantity in pieces. |
| Qty m | Quantity in meters. |
| Weight kg | Weight in kilograms. |
| Surface m2 | Surface area in square meters. |
| Order qty | Base quantity plus spare. |
| Weight w spare | Weight including spare quantity. |
| Unit price | Unit price value. |
| Total price | Total price value. |
| Delta qty to previous revision | Delta quantity against previous revision. |

## 11. AggregateSumColumns

Columns summed during aggregation grouping.

| Column | Description |
| --- | --- |
| Qty pcs | Summed in group. |
| Qty m | Summed in group. |
| Weight kg | Summed in group. |
| Surface m2 | Summed in group. |

## 12. AggregationMaterialTypeRules

Normalization rules applied before grouping/sorting aggregated material type.

| MatchType | Pattern | Output | Description |
| --- | --- | --- | --- |
| endswith | erection | Erection | Any material type ending with erection is normalized to Erection. |
| startswith | paint | Paint shop | Any material type starting with paint is normalized to Paint shop. |

## 13. SpareGeneral

General spare calculation column names and behavior switches.

| Key | Value | Description | Example |
| --- | --- | --- | --- |
| aggregate_table_suffix | aggreated | Suffix appended to revision table to get aggregated table. | rev0 + aggreated -> rev0aggreated |
| spare_column | Spare | Target column to store calculated spare quantity. | Spare |
| spare_info_column | Spare info | Dedicated column with applied spare rule explanation text. | Rule MAIN#3 applied |
| material_type_column | Material type | Column used to classify shop/erection/paint rows. | SS shop |
| description_column | Description | Description source for keyword-based classification. | Hexagon Nut DIN-EN-ISO-4032 M12 |
| type_label_column | Type label | Optional type label source used by TUBI detection. | TUBI |
| type_column | Type | Type source used by some rules and TUBI detection. | PIPE |
| size_column | Size | DN/size source used for threshold split logic. | 50 |
| qty_pcs_column | Qty pcs | Piece quantity source column. | 120 |
| qty_m_column | Qty m | Meter quantity source column when Qty pcs is 0. | 35.5 |
| weight_column | Weight kg | Weight source used to derive unit weight. | 120.5 |
| total_qty_column | Order qty | Output total quantity column: base qty + spare. | 137 |
| weight_w_spare_column | Weight w spare | Output total weight with spare: total qty * unit weight. | 138.22 |
| paint_prefix | paint | Material types starting with this are paint rows. | Paint shop |
| erection_exact | Erection | Material type value treated as erection logic. | Erection |
| shop_suffix | shop | Material type suffix that marks main/shop rows. | SS shop |

## 14. SpareMainRules

Main/shop spare percentages, thresholds, and TUBI rounding controls.

| Key | Value | Description | Example |
| --- | --- | --- | --- |
| seal_qty_threshold | 10 | Seal low/high split quantity threshold. | qty <= 10 |
| seal_min_spare_low | 5 | Fixed spare for low-qty seal rows. | qty=7 -> spare=5 |
| seal_min_spare_high | 5 | Minimum spare for high-qty seal rows. | max(5, ceil(qty*seal_percent)) |
| seal_percent | 0.10 | Seal percentage for qty above threshold. | qty=40 -> ceil(4) |
| critical_qty_threshold | 10 | Critical low/high split threshold. | qty <= 10 |
| critical_min_spare | 3 | Minimum critical spare regardless of percent output. | qty=2 -> spare>=3 |
| critical_percent_low | 0.50 | Critical percentage for low quantity. | qty=8 -> ceil(4) |
| critical_percent_high | 0.10 | Critical percentage for high quantity. | qty=80 -> ceil(8) |
| threshold_1 | 5 | Standard rule first threshold. | qty <= 5 |
| threshold_2 | 20 | Standard rule second threshold. | 5 < qty <= 20 |
| size_threshold | 50 | Size split between small and large percentages. | size < 50 |
| percent_pcs_small | 0.075 | Standard pcs percentage when size is small. | qty=100 -> 8 |
| percent_pcs_large | 0.05 | Standard pcs percentage when size is large. | qty=100 -> 5 |
| percent_m_small | 0.15 | Standard meter percentage when size is small. | qty_m=30 -> 5 |
| percent_m_large | 0.10 | Standard meter percentage when size is large. | qty_m=30 -> 3 |
| tubi_round_multiple | 6 | Round TUBI total quantity to this multiple. | total 39 -> 42 |

## 15. SpareErectionRules

Erection spare thresholds and percentages for NWB, bolts, gaskets, and standard rows.

| Key | Value | Description | Example |
| --- | --- | --- | --- |
| nwb_threshold_1 | 30 | NWB first threshold. | qty <= 30 |
| nwb_threshold_2 | 100 | NWB second threshold. | 30 < qty <= 100 |
| nwb_threshold_3 | 400 | NWB third threshold. | 100 < qty <= 400 |
| nwb_percent_low | 1.0 | NWB percent for lowest range. | 100% base spare |
| nwb_percent_mid | 0.5 | NWB percent for mid range. | 50% base spare |
| nwb_percent_high | 0.35 | NWB percent for high range. | 35% base spare |
| nwb_percent_very_high | 0.2 | NWB percent for very high range. | 20% base spare |
| gasket_threshold_1 | 30 | Gasket first threshold. | qty <= 30 |
| gasket_threshold_2 | 500 | Gasket second threshold. | 30 < qty <= 500 |
| gasket_percent_low | 1.0 | Gasket percent for low range. | 100% |
| gasket_percent_mid | 0.5 | Gasket percent for mid range. | 50% |
| gasket_percent_high | 0.3 | Gasket percent for high range. | 30% |
| standard_threshold | 400 | Standard erection threshold. | qty < 400 |
| standard_percent_low | 0.35 | Standard erection percent below threshold. | 35% |
| standard_percent_high | 0.2 | Standard erection percent at or above threshold. | 20% |

## 16. SpareKeywords

Keyword groups used by spare rule classification.

| Name | Keywords | Description | Example |
| --- | --- | --- | --- |
| seal_required | seal,ring,fkm | All keywords required for seal classification. | FKM seal ring |
| seal_excluded | elbow,flange,tee,reducer | If any appears, row is not seal class. | Seal ring elbow excluded |
| critical | nipple,nipples,union,connection,connections,welding,adaptor,coupling,threaded | Any keyword marks row as critical after seal. | Threaded union |
| tubi | tubi | Any keyword marks row as TUBI after seal/critical checks. | TUBI |
| nwb | nut,washer | Any keyword marks erection row as NWB class. | Hexagon Nut |
| bolt | bolt | Any keyword marks erection row as bolt class. | Hexagon Bolt |
| gasket | gasket | Any keyword marks erection row as gasket class. | Flat gasket |

## 17. SpareRuleGuide

Reference guide describing formulas used by spare logic. This section is informational and not parsed by the scripts.

| Area | Rule | Condition | Formula | Rounding | Example |
| --- | --- | --- | --- | --- | --- |
| Main/Shop | MAIN#1 | Seal and qty<=seal_qty_threshold | spare=seal_min_spare_low | integer | qty=7 -> spare=5 |
| Main/Shop | MAIN#2 | Seal and qty>seal_qty_threshold | spare=max(seal_min_spare_high, ceil(qty*seal_percent)) | ceil | qty=40 -> max(5,4)=5 |
| Main/Shop | MAIN#3 | Critical and qty<=critical_qty_threshold | spare=max(critical_min_spare, ceil(qty*critical_percent_low)) | ceil | qty=8 -> max(3,4)=4 |
| Main/Shop | MAIN#4 | Critical and qty>critical_qty_threshold | spare=max(critical_min_spare, ceil(qty*critical_percent_high)) | ceil | qty=80 -> max(3,8)=8 |
| Main/Shop | MAIN#T1 | TUBI row | base=ceil(qty*pct); total=round_up_to_multiple(qty+base, tubi_round_multiple); spare=total-qty | ceil + multiple | qty=35,total=42,spare=7 |
| Main/Shop | MAIN#5 | Standard pcs and qty<=threshold_1 | spare=0 | none | qty=4 -> 0 |
| Main/Shop | MAIN#6 | Standard pcs and threshold_1<qty<=threshold_2 | spare=1 | none | qty=12 -> 1 |
| Main/Shop | MAIN#7a | Standard pcs and qty>threshold_2 and size<size_threshold | spare=ceil(qty*percent_pcs_small) | ceil | qty=100,size=25 -> 8 |
| Main/Shop | MAIN#7b | Standard pcs and qty>threshold_2 and size>=size_threshold | spare=ceil(qty*percent_pcs_large) | ceil | qty=100,size=80 -> 5 |
| Main/Shop | MAIN#8a | Standard meter and size<size_threshold | spare=ceil(qty_m*percent_m_small) | ceil | qty_m=30,size=25 -> 5 |
| Main/Shop | MAIN#8b | Standard meter and size>=size_threshold | spare=ceil(qty_m*percent_m_large) | ceil | qty_m=30,size=80 -> 3 |
| Erection | ER#1/2/3/4 | NWB by thresholds | base=ceil(qty*nwb_percent_*); total=round_up_nut_washer(qty+base); spare=total-qty | ceil + stepped | qty=156 -> spare=94 |
| Erection | ER#1B/2B/3B/4B | Bolt by thresholds | base=ceil(qty*nwb_percent_*); total=round_up_to_next_10(qty+base); spare=total-qty | ceil + next10 | qty=108 -> spare=42 |
| Erection | ER#5/6/7 | Gasket by thresholds | low:ceil(qty*gasket_low); mid/high: base+round_to_next10 | ceil + next10 | qty=200 -> spare=100 |
| Erection | ER#8/9 | Standard erection | qty<standard_threshold -> ceil(qty*standard_percent_low), else ceil(qty*standard_percent_high) | ceil | qty=500 -> 100 |
