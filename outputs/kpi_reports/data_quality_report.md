# Data Quality Report — RAW data (synthetic)

Source: `data/raw/shipments_raw.parquet` (304,500 rows) and `data/raw/scan_events_raw.parquet` (1,200,000 rows). All data is synthetic, generated with seed 42.

## Integrity checks

| Check | Count |
|---|---|
| duplicate_shipment_id_rows | 9,000 |
| duplicate_shipment_ids_distinct | 4,500 |
| duplicate_scan_event_id_rows | 0 |
| negative_cost_rows | 2,474 |
| zero_cost_rows | 1,172 |
| invalid_cost_rows | 0 |
| invalid_promised_ts_rows | 3,081 |
| invalid_actual_ts_rows | 2,980 |
| missing_promised_ts_rows | 2,407 |
| missing_actual_ts_rows | 3,601 |
| unrecognised_origin_hub_rows | 0 |
| messy_variant_origin_hub_rows | 91,486 |
| unrecognised_destination_hub_rows | 0 |
| messy_variant_destination_hub_rows | 92,027 |
| invalid_route_ref_rows | 1,548 |
| invalid_vehicle_ref_rows | 1,557 |
| missing_vehicle_rows | 1,455 |
| missing_category_rows | 873 |
| invalid_scan_ts_rows | 0 |
| orphan_scan_rows | 6,000 |
| out_of_order_scan_shipments | 20,914 |
| out_of_order_scan_events | 29,287 |
| negative_dwell_origin_hub_shipments | 39 |
| negative_dwell_dest_hub_shipments | 2,905 |

## Column profiles

Full per-column profiles: `column_profile_shipments.csv` and `column_profile_scans.csv`.

## Key observations

- Duplicate shipment rows: 9,000 (4,500 extra rows beyond the first copy).
- Negative costs: 2,474 rows; zero costs: 1,172 rows — excluded from the cost KPI denominator.
- Invalid promised timestamps: 3,081; invalid actual timestamps: 2,980 — excluded from OTD.
- Orphan scan events: 6,000 rows referencing shipments that do not exist.
- Out-of-order scan events: 29,287 events in 20,914 shipments — re-sequenced in cleaning.
- Negative dwell pairs: origin hub 39, destination hub 2,905 shipments — flagged, excluded from dwell KPI.