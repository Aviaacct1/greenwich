# Aero rate calibration report | 17 July 2026

Each aero line's unit rate is its ingested revenue divided by the matching traffic
driver of the same year. This closes the vendor-grain-only aero rule: the workbook's
aero inputs are rates, and calibration is the named step from a P&L's amounts to those
rates. rate x driver reproduces the revenue by construction (difference column).

code	year	revenue (EUR m)	driver	rate	identity diff (EUR m)
rev_aero_other	2022	3.2	5.0 m pax	0.6400 EUR per pax	0.0
rev_aero_other	2023	3.4	5.5 m pax	0.6182 EUR per pax	0.0
rev_landing	2022	12.4	48.0 k ATMs	258.3333 EUR per ATM	0.0
rev_landing	2023	13.1	52.0 k ATMs	251.9231 EUR per ATM	0.0
rev_psc	2022	38.2	5.0 m pax	7.6400 EUR per pax	0.0
rev_psc	2023	41.5	5.5 m pax	7.5455 EUR per pax	0.0
rev_security	2022	9.1	5.0 m pax	1.8200 EUR per pax	0.0
rev_security	2023	9.8	5.5 m pax	1.7818 EUR per pax	0.0

All aero rate lines calibrated.

Note: the driver per code is read from line_sets tier 1, so calibration divides by exactly what the model multiplies by; the identity check confirms rate x driver returns the revenue. Landing arrives at total grain and uses aggregate ATMs; its segment split (dom/intl/cargo) is a tier-2 step. rev_aero_other is a residual but line_sets drives it per pax, so it calibrates to a blended per-pax rate, flagged as residual rather than a single tariff.
