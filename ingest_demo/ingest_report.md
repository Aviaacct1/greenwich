# Vendor ingest report | vendor_teaser_demo.xlsx | 16 July 2026

Lines extracted: 18 across 1 sheet(s); stated totals recognised: 3.
Mapping: 17 proposed (2 low confidence), 1 unmapped and BLOCKING: Sundry recharges to group.

Low confidence, review first: 'Landing & parking fees' to rev_landing; 'Retail & duty free concessions' to conc_dutyfree
Stated 'Total revenue' 2022: 89.3; extracted positive lines sum 89.3; difference 0.0.
Stated 'Total revenue' 2023: 96.5; extracted positive lines sum 96.5; difference 0.0.
Stated 'Total operating costs' 2022: -34.2; extracted negative lines sum -34.2; difference 0.0.
Stated 'Total operating costs' 2023: -36.2; extracted negative lines sum -36.2; difference 0.0.
Stated 'EBITDA' 2022: 55.1; extracted all lines net sum 55.1; difference -0.0.
Stated 'EBITDA' 2023: 60.3; extracted all lines net sum 60.3; difference 0.0.

Aero revenue lines are kept at vendor grain only: the workbook's aero inputs are unit RATES, and a P&L gives revenue amounts; rate calibration (revenue over driver) is an explicit later step. Traffic totals need the dom/intl split (Sabre/OAG) before they can place.

Next: review and edit t6_project_lines_proposed.tsv (the mapping screen in the Studio will replace this file edit), then t6_support.py validate blocks until every line maps. t1_actuals_avia.tsv feeds the actuals strip; split-candidate lines are left to the analyst.
