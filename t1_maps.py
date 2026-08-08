"""Shared T1 label maps, built from line_sets.tsv (tier-aware). Single source
of truth for the mapping between taxonomy codes and workbook input labels,
replacing the hard-coded tier-1 dictionaries previously duplicated across the
T1 reader, actuals strip and Checks sheet. Author: Avia Solutions."""

def read_line_sets(path, tier):
    rows = []
    for line in open(path, encoding="utf-8"):
        if line.startswith("#") or not line.strip():
            continue
        p = line.rstrip("\n").split("\t")
        d = dict(zip(["tier", "axis", "label", "metric_code", "segment", "unit",
                      "driver", "calc_group", "replaces"], p))
        d["tier"] = int(d["tier"])
        rows.append(d)
    if tier == 1:
        return [r for r in rows if r["tier"] == 1]
    replaced = {r["replaces"] for r in rows if r["tier"] == 2 and r["replaces"] != "-"}
    return [r for r in rows if not (r["tier"] == 1 and f"{r['metric_code']}/{r['segment']}" in replaced)]

def build_maps(line_sets_path, tier):
    """Returns (by_key, level_names, cats):
    by_key: (metric_code, segment, driver_type) -> Inputs block label;
    level_names: (metric_code, segment) -> label for level/per_pax rows (actuals strip);
    cats: {'na': [(label, code)...], 'op': [...], 'aero': [...]} for stem lists."""
    lines = read_line_sets(line_sets_path, tier)
    by_key, level_names = {}, {}
    cats = {"na": [], "op": [], "aero": [], "traffic": []}
    for r in lines:
        lab, code, seg = r["label"], r["metric_code"], r["segment"]
        if r["axis"] == "Traffic":
            by_key[(code, seg, "level")] = lab
            level_names[(code, seg)] = lab
            cats["traffic"].append((lab, code))
        elif r["axis"] == "Aero":
            by_key[(code, seg, "per_pax")] = f"{lab} - unit rate"
            by_key[(code, seg, "one_off_step")] = f"{lab} - reset uplift"
            level_names[(code, seg)] = f"{lab} - unit rate"
            cats["aero"].append((lab, code, seg))
        elif r["axis"] == "Non-aero":
            by_key[(code, seg, "level")] = f"{lab} - base year revenue"
            by_key[(code, seg, "elasticity_pax")] = f"{lab} - elasticity to pax growth"
            by_key[(code, seg, "elasticity_sqm")] = f"{lab} - elasticity to terminal size"
            by_key[(code, seg, "one_off_step")] = f"{lab} - uplift"
            level_names[(code, seg)] = f"{lab} - base year revenue"
            cats["na"].append((lab, code))
        elif r["axis"] == "Opex":
            by_key[(code, seg, "level")] = f"{lab} - base year"
            by_key[(code, seg, "elasticity_pax")] = f"{lab} - elasticity to pax growth"
            by_key[(code, seg, "one_off_step")] = f"{lab} - step uplift"
            level_names[(code, seg)] = f"{lab} - base year"
            cats["op"].append((lab, code))
    return by_key, level_names, cats
