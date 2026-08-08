"""Greenwich host check. Exits non-zero when anything required is missing or broken.

Provisioning a host is five steps and the last is not optional:
  1. git clone
  2. put the data root in place (reference/ holding the client and reference workbooks)
  3. set AVIA_LOCAL_CACHE
  4. <python> -m pip install -r requirements.txt
  5. <python> check_env.py

This file exists because pip reports a broken install as a warning and exits zero, so a
successful-looking install is not evidence that the tool runs. Everything below is checked
by doing it, not by reading a version string.

Author: Avia Solutions."""
import os
import shutil
import subprocess
import sys

FAILS = []
WARNS = []


def check(ok, msg, fatal=True):
    print(("PASS: " if ok else ("FAIL: " if fatal else "WARN: ")) + msg)
    if not ok:
        (FAILS if fatal else WARNS).append(msg)
    return ok


def main():
    print("Greenwich environment check")
    print("=" * 60)

    # --- interpreter -----------------------------------------------------------
    v = sys.version_info
    check(v >= (3, 10), f"Python {v.major}.{v.minor}.{v.micro} (3.10 or later required)")
    print(f"      interpreter: {sys.executable}")

    # --- virtual environment ---------------------------------------------------
    in_venv = sys.prefix != getattr(sys, "base_prefix", sys.prefix)
    check(in_venv,
          "virtualenv. This host may run several Avia tools; installing one tool's "
          "dependencies into a shared interpreter changes every other tool that uses it",
          fatal=False)

    # --- packages, by importing them -------------------------------------------
    for mod, label in (("openpyxl", "openpyxl"), ("matplotlib", "matplotlib")):
        try:
            m = __import__(mod)
            check(True, f"{label} {getattr(m, '__version__', 'imported')}")
        except Exception as e:
            check(False, f"{label} does not import: {e}")

    # --- config resolution -----------------------------------------------------
    try:
        import config
        print("-" * 60)
        print(config.describe())
        print("-" * 60)
        try:
            config.data_root()
            check(True, "data root resolves")
        except config.ConfigError as e:
            check(False, f"data root: {e}")
        try:
            config.find_reference("Plovdiv_FinancialModel_vDraft.xlsm")
            check(True, "Plovdiv reference workbook found (the acceptance oracle's input)")
        except config.ConfigError as e:
            check(False, f"Plovdiv reference workbook: {e}", fatal=False)
    except Exception as e:
        check(False, f"config.py does not import: {e}")

    # --- LibreOffice, needed for the selftest recalculation phase ----------------
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if check(bool(soffice), "LibreOffice on PATH (selftest.py recalculation phase)", fatal=False):
        try:
            r = subprocess.run([soffice, "--version"], capture_output=True, text=True, timeout=90)
            check(r.returncode == 0, f"LibreOffice runs: {r.stdout.strip()[:60]}", fatal=False)
        except Exception as e:
            check(False, f"LibreOffice will not run: {e}", fatal=False)

    # --- smoke test: build a model end to end -----------------------------------
    here = os.path.dirname(os.path.abspath(__file__))
    try:
        import tempfile
        with tempfile.TemporaryDirectory() as td:
            out = os.path.join(td, "smoke.xlsx")
            r = subprocess.run(
                [sys.executable, os.path.join(here, "build_model.py"), here,
                 os.path.join(here, "project_header.tsv"),
                 os.path.join(here, "t1_assumptions.tsv"),
                 os.path.join(here, "t1_actuals.tsv"), out],
                capture_output=True, text=True, timeout=600, cwd=td)
            check(r.returncode == 0 and os.path.isfile(out),
                  "smoke test: build_model.py produces a workbook end to end")
            if r.returncode != 0:
                print(r.stdout[-1500:]); print(r.stderr[-1500:])
    except Exception as e:
        check(False, f"smoke test did not run: {e}")

    print("=" * 60)
    if WARNS:
        print(f"{len(WARNS)} warning(s):")
        for w in WARNS:
            print("  - " + w)
    if FAILS:
        print(f"{len(FAILS)} FAILURE(S):")
        for f in FAILS:
            print("  - " + f)
        print("Host is NOT ready.")
        return 1
    print("Host is ready. Next: run selftest.py and require every assertion green.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
