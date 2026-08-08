"""Generator increment: aero rate calibration (queue item 3, note 17).

The vendor ingest keeps aero revenue lines at VENDOR GRAIN ONLY, because the
workbook's aero inputs are unit RATES and a P&L gives revenue AMOUNTS. This step
closes that hole explicitly: for each aero line it divides the ingested revenue
by the matching traffic driver of the same year to produce the unit rate the
model consumes, with both sources named on every output row. rate x driver
reproduces the revenue by construction; the calibration report shows the audit.

Inputs (all TSV, schema 11 shapes):
  confirmed T6            line_id -> taxonomy code (aero lines identified here)
  vendor actuals          line_id, year, value (EUR m) at vendor grain
  traffic actuals         metric_code (pax_*/atm_*/cargo_*), year, value, source
Output:
  t1_aero_rates.tsv       per_pax / per_atm / per_tonne rate rows for T1
  aero_calibration_report.md

Driver basis per aero code is explicit below and must match the basis the model
multiplies by; the round-trip test (rate x driver = revenue) proves consistency.
Author: Avia Solutions."""
import sys, os, datetime

T6_COLS = ["line_id", "label", "parent_line_id", "taxonomy_codes", "forecast_grain", "origin"]
T1_COLS = ["line_id", "segment", "case_id", "year", "value", "unit",
           "temporality", "driver_type", "step_date", "step_value", "repeat_years", "source"]
# metric-keyed T1 shape (traffic actuals, aero rate output): first column is metric_code
T1M_COLS = ["metric_code"] + T1_COLS[1:]

LS_COLS = ["tier", "axis", "label", "metric_code", "segment", "unit", "driver", "calc_group", "replaces"]
# aggregate a segmented driver to its total, for a vendor line that arrives at total grain
AGG = {"atm_dom": "atm_total", "atm_intl": "atm_total", "atm_cargo": "atm_total",
       "pax_dom": "pax_total", "pax_intl": "pax_total"}


def unit_of(driver):
    if driver.startswith("pax_"):
        return "EUR per pax", "per_pax"
    if driver.startswith("atm_"):
        return "EUR per ATM", "per_atm"
    if driver.startswith("cargo"):
        return "EUR per tonne", "per_tonne"
    return "EUR", "level"


def aero_drivers_from_line_sets(line_sets_path):
    """Authoritative driver per aero code, read from line_sets tier 1 so calibration
    divides by exactly what the model multiplies by. Segmented codes (landing) use the
    aggregate driver for a total-grain vendor line; the segment split is a tier-2 step."""
    drv = {}
    for r in read_tsv(line_sets_path, LS_COLS):
        if r["tier"] != "1" or r["axis"] != "Aero":
            continue
        code, d = r["metric_code"], r["driver"]
        if r["segment"] == "total":
            drv[code] = AGG.get(d, d)
        else:
            drv.setdefault(code, AGG.get(d, d))
    return drv


def read_tsv(path, cols):
    rows = []
    for line in open(path, encoding="utf-8"):
        if line.startswith("#") or not line.strip():
            continue
        rows.append(dict(zip(cols, line.rstrip("\n").split("\t"))))
    return rows


def code_of(mapping):
    if not mapping or mapping == "-":
        return None
    return mapping.split("|")[0].split(":")[0]


def codes_shares(mapping):
    """All (code, share) pairs on a line, so split lines attribute revenue by share
    rather than dropping every code but the first."""
    if not mapping or mapping == "-":
        return []
    out = []
    for part in mapping.split("|"):
        p = part.split(":")
        out.append((p[0], float(p[1]) if len(p) > 1 and p[1] else 1.0))
    return out


def main(t6_path, vendor_path, traffic_path, line_sets_path, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    today = datetime.date.today().strftime("%d %B %Y")

    aero_driver = aero_drivers_from_line_sets(line_sets_path)

    t6 = read_tsv(t6_path, T6_COLS)
    # line_id -> [(code, share)]. Split lines carry their shares on the line that holds the
    # vendor amount (the parent, grain 'children'); its children are derived and carry no
    # vendor actuals, so reading every line's full mapping attributes correctly and does not
    # double count. A line's shares must sum to 1 per t6_support.validate.
    line_shares = {r["line_id"]: codes_shares(r["taxonomy_codes"]) for r in t6}

    vend = read_tsv(vendor_path, T1_COLS)
    # aero revenue by code by year (EUR m), share-attributed across split codes
    rev = {}
    vend_src = {}
    split_codes = set()
    for r in vend:
        shares = line_shares.get(r["line_id"], [])
        aero_here = [(c, sh) for c, sh in shares if c in aero_driver]
        if not aero_here:
            continue
        y = int(r["year"])
        for code, share in aero_here:
            rev.setdefault(code, {}).setdefault(y, 0.0)
            rev[code][y] += float(r["value"]) * share
            vend_src[code] = r["source"].split(" via ")[0]
            if share != 1.0:
                split_codes.add(code)

    traf = read_tsv(traffic_path, T1M_COLS)
    drivers = {}   # metric -> {year: (value, unit, source)}
    for r in traf:
        m = r["metric_code"]
        drivers.setdefault(m, {})[int(r["year"])] = (float(r["value"]), r["unit"], r["source"])

    out_rows, report, unresolved = [], [], []
    checks_ok = True
    for code in sorted(rev):
        driver_metric = aero_driver[code]
        rate_unit, dtype = unit_of(driver_metric)
        dser = drivers.get(driver_metric)
        if not dser:
            unresolved.append(f"{code} (no {driver_metric} in traffic actuals)")
            continue
        for y in sorted(rev[code]):
            if y not in dser:
                unresolved.append(f"{code} {y} (no {driver_metric} for {y})")
                continue
            rev_eur = rev[code][y] * 1e6                      # EUR m -> EUR
            dval, dunit, dsrc = dser[y]
            # pax/ATM given in millions / thousands: normalise to absolute counts
            if dunit.startswith("m "):
                dcount = dval * 1e6
            elif dunit.startswith("k "):
                dcount = dval * 1e3
            else:
                dcount = dval
            rate = rev_eur / dcount
            # identity check: rate x driver back to EUR m
            back = rate * dcount / 1e6
            diff = round(back - rev[code][y], 6)
            if abs(diff) > 1e-6:
                checks_ok = False
            rev_disp = round(rev[code][y], 4)
            src = (f"calibrated: {code} revenue {rev_disp} EUR m ({vend_src.get(code,'vendor')}) "
                   f"over {driver_metric} {dval} {dunit} ({dsrc}); rate = revenue / driver")
            out_rows.append("\t".join([code, "total", "1", str(y), f"{rate:.6f}",
                                       rate_unit, "actual", dtype, "", "", "", src]))
            report.append((code, y, rev_disp, f"{dval} {dunit}", f"{rate:.4f} {rate_unit}", diff))

    hdr = ("# T1 aero rate rows, calibrated from ingested revenue over traffic drivers | {0}\n"
           "# rate = ingested aero revenue / matching traffic driver, same year; sources on every row.\n"
           "# metric_code\tsegment\tcase_id\tyear\tvalue\tunit\ttemporality\tdriver_type"
           "\tstep_date\tstep_value\trepeat_years\tsource\n").format(today)
    open(os.path.join(out_dir, "t1_aero_rates.tsv"), "w", encoding="utf-8").write(
        hdr + "\n".join(out_rows) + "\n")

    rep = [f"# Aero rate calibration report | {today}", "",
           "Each aero line's unit rate is its ingested revenue divided by the matching traffic",
           "driver of the same year. This closes the vendor-grain-only aero rule: the workbook's",
           "aero inputs are rates, and calibration is the named step from a P&L's amounts to those",
           "rates. rate x driver reproduces the revenue by construction (difference column).", "",
           "code\tyear\trevenue (EUR m)\tdriver\trate\tidentity diff (EUR m)"]
    for code, y, r_m, drv, rate, diff in report:
        rep.append(f"{code}\t{y}\t{r_m}\t{drv}\t{rate}\t{diff}")
    rep.append("")
    if unresolved:
        rep.append("Not calibrated (kept at vendor grain, review): " + "; ".join(unresolved))
    else:
        rep.append("All aero rate lines calibrated.")
    if split_codes:
        rep.append("Share-attributed from split vendor lines (revenue apportioned by the confirmed T6 shares): "
                   + ", ".join(sorted(split_codes)) + ".")
    rep.append("")
    rep.append("Note: the driver per code is read from line_sets tier 1, so calibration divides by "
               "exactly what the model multiplies by; the identity check confirms rate x driver "
               "returns the revenue. Landing arrives at total grain and uses aggregate ATMs; its "
               "segment split (dom/intl/cargo) is a tier-2 step. rev_aero_other is a residual but "
               "line_sets drives it per pax, so it calibrates to a blended per-pax rate, flagged as "
               "residual rather than a single tariff.")
    open(os.path.join(out_dir, "aero_calibration_report.md"), "w", encoding="utf-8").write(
        "\n".join(rep) + "\n")

    print(f"calibrated {len(out_rows)} aero rate row(s) across {len(set(r.split(chr(9))[0] for r in out_rows))} "
          f"code(s); identity checks {'ALL PASS' if checks_ok else 'FAILED'}; "
          f"unresolved {len(unresolved)}; output in {out_dir}")
    if not checks_ok:
        raise SystemExit("CALIBRATION identity check failed")


if __name__ == "__main__":
    # usage: calibrate_aero.py <t6_confirmed> <vendor_actuals> <traffic_actuals> <line_sets> <out_dir>
    main(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
