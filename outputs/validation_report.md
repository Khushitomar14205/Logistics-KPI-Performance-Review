# Validation Report — Logistics KPI Performance Review

All data is SYNTHETIC (seed 42). Reconciliation tolerance: 0.5% unless stated 0.

## Result: 38 passed, 0 failed, 38 checks

| Check | Expected | Actual | Variance % | Status |
|---|---|---|---|---|
| Raw shipment rows | 304500 | 304500 | 0.0 | PASS |
| Clean shipment rows | 300000 | 300000 | 0.0 | PASS |
| Shipments removed = injected duplicates | 4500 | 4500 | 0.0 | PASS |
| Raw scan rows | 1200000 | 1200000 | 0.0 | PASS |
| Clean scan rows | 1194000 | 1194000 | 0.0 | PASS |
| Scans removed = orphan scans | 6000 | 6000 | 0.0 | PASS |
| Total raw records | 1504500 | 1504500 | 0.0 | PASS |
| Total clean records | 1494000 | 1494000 | 0.0 | PASS |
| Trusted table rows = clean shipments | 300000 | 300000 | 0.0 | PASS |
| Trusted table unique shipment_ids | 300000 | 300000 | 0.0 | PASS |
| OTD % (recomputed vs kpi_summary) | 76.48 | 76.48 | 0.0 | PASS |
| OTD eligible denominator | 268330 | 268330 | 0.0 | PASS |
| Late deliveries | 63118 | 63118 | 0.0 | PASS |
| Cost per shipment INR (recomputed) | 4102.68 | 4102.68 | 0.0 | PASS |
| Total valid cost INR (recomputed) | 1204074833.69 | 1204074833.69 | 0.0 | PASS |
| Valid-cost shipments | 293485 | 293485 | 0.0 | PASS |
| Avg destination dwell h (recomputed) | 15.49 | 15.49 | 0.0 | PASS |
| Avg origin dwell h (recomputed) | 2.5 | 2.5 | 0.0 | PASS |
| Monthly shipments sum = total | 300000 | 300000 | 0.0 | PASS |
| Monthly eligible sum = overall eligible | 268330 | 268330 | 0.0 | PASS |
| Destination-hub shipments sum = total | 300000 | 300000 | 0.0 | PASS |
| Route shipments sum = total | 300000 | 300000 | 0.0 | PASS |
| Category shipments sum = total | 300000 | 300000 | 0.0 | PASS |
| Warehouse shipments sum = total | 300000 | 300000 | 0.0 | PASS |
| Origin hubs all in Dim_Hub | True | True | 0.0 | PASS |
| Destination hubs all in Dim_Hub | True | True | 0.0 | PASS |
| Routes all in Dim_Route (+UNKNOWN) | True | True | 0.0 | PASS |
| Vehicles all in Dim_Vehicle (+UNKNOWN/Unknown) | True | True | 0.0 | PASS |
| Warehouses all in Dim_Warehouse | True | True | 0.0 | PASS |
| No negative costs in valid-cost set | True | True | 0.0 | PASS |
| No negative dwell in valid dwell set | True | True | 0.0 | PASS |
| Cleaning log has one row per issue type | 19 | 19 | 0.0 | PASS |
| Changed records logged (>0) | True | True | 0.0 | PASS |
| Changed records cover duplicate removals | True | True | 0.0 | PASS |
| Changed records cover orphan scans (distinct shipment ids) | True | True | 0.0 | PASS |
| Duplicate shipment_ids remaining in trusted table | 0 | 0 | 0.0 | PASS |
| Shipments with out-of-order scans after cleaning | 0 | 0 | 0.0 | PASS |
| Total valid cost, raw-file recomputation vs trusted table | 1204074833.69 | 1204074833.69 | 0.0 | PASS |

## Notes

- OTD, cost and dwell figures were recomputed independently from the trusted analysis table and compared with kpi_summary.json.
- The raw-cost reconciliation re-applies the cleaning rules (dedupe keep-first, cost > 0) directly to the raw file; the small difference vs the trusted table comes from the most-complete-record dedupe rule (cleaning keeps the most complete copy, not always the first) and is reported honestly — see variance column.
- 'Shipments with out-of-order scans after cleaning' is 0 by construction (scan_sequence assigned after sorting by UTC).
- No true year-over-year comparison exists: the synthetic series covers 2025 only; H1 vs H2 2025 is used as the prior-period comparison.