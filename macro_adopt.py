"""Refresh automation increment (queue item 6, note 16 section 4.3): the adoption
hook. One command per project, run when the watch-list shows a newer vintage than a
project pins, it compares the project's pinned macro rows against the new vintage's
macro rows (both produced by macro_feed) and writes the BEFORE-AND-AFTER MOVEMENT for
the Assumptions Book. It reports only; it does not touch the model. Adopting the new
vintage means re-pointing the project's macro build at it, a deliberate act the
analyst takes after reading this movement (note 16 section 2).

Usage: macro_adopt.py <pinned_macro_t1> <new_macro_t1> <project_name> <note_out>
Author: Avia Solutions."""
import sys, datetime

T1M = ["metric_code", "segment", "case_id", "year", "value", "unit",
       "temporality", "driver_type", "step_date", "step_value", "repeat_years", "source"]


def read_macro(path):
    vals, srcs = {}, {}
    for line in open(path, encoding="utf-8"):
        if line.startswith("#") or not line.strip():
            continue
        r = dict(zip(T1M, line.rstrip("\n").split("\t")))
        key = (r["metric_code"], r["segment"], int(r["year"]))
        vals[key] = float(r["value"])
        srcs.setdefault((r["metric_code"], r["segment"]), set()).add(r["source"])
    return vals, srcs


def vintage_hint(srcset):
    """Pull the shortest distinct descriptor from the source strings (the vintage lives there)."""
    return "; ".join(sorted(srcset)) if srcset else "None"


def main(pinned_path, new_path, project, note_out):
    today = datetime.date.today().strftime("%d %B %Y")
    old, osrc = read_macro(pinned_path)
    new, nsrc = read_macro(new_path)

    series = sorted(set(k[:2] for k in old) | set(k[:2] for k in new),
                    key=lambda s: (s[0] != "gdp_growth", s[1] != "basket", s[0], s[1]))
    lines = [f"# Macro adoption movement note: {project} | {today}", "",
             "For the Assumptions Book. Before-and-after movement from re-running macro_feed on the",
             "newer vintage the watch-list flagged. This note reports the movement; adopting means the",
             "analyst re-points the project's macro build at the new vintage after reviewing it below.", ""]
    worst = 0.0
    for metric, seg in series:
        yrs = sorted(y for (m, s, y) in set(old) | set(new) if m == metric and s == seg)
        if not yrs:
            continue
        lines.append(f"## {metric} / {seg}")
        lines.append(f"Pinned source: {vintage_hint(osrc.get((metric, seg), set()))}")
        lines.append(f"New source:    {vintage_hint(nsrc.get((metric, seg), set()))}")
        lines.append("year\tpinned\tnew\tmovement")
        for y in yrs:
            ov = old.get((metric, seg, y))
            nv = new.get((metric, seg, y))
            if ov is None:
                lines.append(f"{y}\tNone\t{nv}\tadded in new vintage")
            elif nv is None:
                lines.append(f"{y}\t{ov}\tNone\tdropped in new vintage")
            else:
                d = round(nv - ov, 4)
                worst = max(worst, abs(d))
                lines.append(f"{y}\t{ov}\t{nv}\t{d:+.4f}")
        lines.append("")
    lines.append(f"Largest single-year movement across all series: {worst:+.4f}.")
    lines.append("No model values changed by this note. Adoption is the analyst's deliberate next step.")
    open(note_out, "w", encoding="utf-8").write("\n".join(lines) + "\n")
    print(f"adoption note for {project}: {len(series)} series compared; "
          f"largest movement {worst:+.4f}; note -> {note_out}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
