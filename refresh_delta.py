"""Refresh automation increment (queue item 6, note 16 section 4.1 and 4.2): the
scheduled delta-run driver and the vintage watch-list updater.

This is the script a weekly Windows Task Scheduler entry would call on the machine
that holds the Egnyte drive and the Extract mount (that mount is an ops setup task,
not a build task). It inventories a scan directory against a manifest, so it is safe
to run unattended and an empty week costs seconds; it detects new and changed files,
matches them to the named reference sources on the watch-list, updates the watch-list
with the latest vintage seen, and writes a delta report that LEADS with 'new vintages
this week' (note 16 section 4.2). Failures are loud: a missing scan directory exits
non-zero rather than reporting a silent empty week.

Discovery only: this updates the STORE-side inventory and the watch-list. Adoption
into a live project model stays deliberate and is done with macro_adopt.py.
Author: Avia Solutions."""
import sys, os, re, csv, hashlib, datetime

WL_COLS = ["source", "metric_family", "cadence", "filename_pattern", "home",
           "latest_vintage_seen", "last_seen_date", "notes"]
MONTHS = "jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec"
VINTAGE_RES = [re.compile(r"(" + MONTHS + r")[ _-]?(20\d{2})", re.I),
               re.compile(r"(20\d{2})[ _-]?(0[1-9]|1[0-2])"),
               re.compile(r"(20\d{2})[ _-](q[1-4])", re.I),
               re.compile(r"\b(20\d{2})\b")]


def read_tsv(path, cols):
    rows = []
    for line in open(path, encoding="utf-8"):
        if line.startswith("#") or not line.strip():
            continue
        rows.append(dict(zip(cols, line.rstrip("\n").split("\t"))))
    return rows


def read_manifest(path):
    seen = {}
    if os.path.exists(path):
        for line in open(path, encoding="utf-8"):
            if line.startswith("#") or not line.strip():
                continue
            p = line.rstrip("\n").split("\t")
            seen[p[0]] = p[1]  # relpath -> sha1
    return seen


def sha1_of(path):
    h = hashlib.sha1()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def vintage_from(name):
    for rx in VINTAGE_RES:
        m = rx.search(name)
        if m:
            return "".join(g for g in m.groups()).lower()
    return None


def main(scan_dir, manifest_path, watchlist_path, report_path):
    t0 = datetime.datetime.now()
    today = t0.strftime("%d %B %Y")
    if not os.path.isdir(scan_dir):
        sys.exit(f"REFRESH FAILURE: scan directory not found: {scan_dir} (ops: mount the Egnyte drive / Extract)")

    prev = read_manifest(manifest_path)
    cur, new, changed = {}, [], []
    for root, _, files in os.walk(scan_dir):
        for fn in sorted(files):
            fp = os.path.join(root, fn)
            rel = os.path.relpath(fp, scan_dir)
            try:
                digest = sha1_of(fp)
            except OSError as e:
                sys.exit(f"REFRESH FAILURE: cannot read {rel}: {e}")
            cur[rel] = digest
            if rel not in prev:
                new.append(rel)
            elif prev[rel] != digest:
                changed.append(rel)

    wl = read_tsv(watchlist_path, WL_COLS)
    fresh = []  # (source, vintage, relpath)
    for rel in new + changed:
        base = os.path.basename(rel).lower()
        for w in wl:
            pat = w["filename_pattern"].lower()
            if pat and pat in base:
                v = vintage_from(base) or t0.strftime("%Y-%m-%d")
                w["latest_vintage_seen"] = v
                w["last_seen_date"] = t0.strftime("%Y-%m-%d")
                fresh.append((w["source"], v, rel))
                break

    # write manifest (append-only inventory of what is now on disk)
    with open(manifest_path, "w", encoding="utf-8") as fh:
        fh.write(f"# refresh manifest (relpath -> sha1) | updated {t0.isoformat(timespec='seconds')}\n")
        for rel in sorted(cur):
            fh.write(f"{rel}\t{cur[rel]}\n")

    # write watch-list back
    with open(watchlist_path, "w", encoding="utf-8") as fh:
        fh.write("# Macro vintage watch-list (note 16 section 4.2). Named reference sources, cadence, latest vintage seen.\n")
        fh.write("# The scheduled delta run (refresh_delta.py) updates latest_vintage_seen and last_seen_date; a project\n")
        fh.write("# model compares its pinned vintage against latest_vintage_seen to raise the WATCH (note 16 section 2).\n")
        fh.write("# filename_pattern is a case-insensitive substring matched against incoming file names.\n")
        fh.write("# " + "\t".join(WL_COLS) + "\n")
        for w in wl:
            fh.write("\t".join(w[c] for c in WL_COLS) + "\n")

    dur = (datetime.datetime.now() - t0).total_seconds()
    rep = [f"# Data refresh delta report | {today}", ""]
    rep.append("New vintages this week:")
    if fresh:
        for source, v, rel in fresh:
            rep.append(f"  {source}: vintage {v}  ({rel})")
    else:
        rep.append("  none this week.")
    rep.append("")
    rep.append(f"New files: {len(new)}" + ("" if not new else ": " + ", ".join(new)))
    rep.append(f"Changed files: {len(changed)}" + ("" if not changed else ": " + ", ".join(changed)))
    rep.append(f"Unchanged files: {len(cur) - len(new) - len(changed)}")
    rep.append("Failures: none")
    rep.append(f"Run duration: {dur:.2f}s; files inventoried: {len(cur)}.")
    rep.append("")
    rep.append("Adoption is deliberate (note 16 section 2): where a project pins an older vintage than the")
    rep.append("watch-list now shows, run macro_adopt.py for that project and review the movement before adopting.")
    open(report_path, "w", encoding="utf-8").write("\n".join(rep) + "\n")

    print(f"refresh delta: {len(new)} new, {len(changed)} changed, {len(cur)} inventoried; "
          f"new vintages: {len(fresh)}; report -> {report_path}")


if __name__ == "__main__":
    # usage: refresh_delta.py <scan_dir> <manifest> <watchlist> <report_out>
    main(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
