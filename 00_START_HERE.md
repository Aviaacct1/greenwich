# Greenwich

Airport financial model and business plans. The generator that turns T1 assumption tables and
actuals into an Avia-standard model workbook, plus the Forecast Studio front end that produces
those tables.

Version 1.0 - 8 August 2026 - Avia Solutions

---

## Provisioning a host

Five steps, and the fifth is not optional.

```
git clone https://github.com/Aviaacct1/greenwich.git C:\src\greenwich
cd C:\src\greenwich
rem  2. put the data root in place: E:\Avia\greenwich\reference holds the client and
rem     reference workbooks. They are deliberately outside this repository.
setx AVIA_LOCAL_CACHE "E:\Avia\greenwich"
py -3.12 -m pip install -r requirements.txt
py -3.12 check_env.py
```

`check_env.py` exits non-zero when anything required is missing or broken. It imports every
package rather than reading a version string, resolves the data root, finds the Plovdiv
reference workbook, confirms LibreOffice runs, and builds a model end to end. pip reports a
broken install as a warning and exits zero, so a clean-looking install is not evidence.

One virtual environment per tool on a shared host. This workstation will run four tools;
installing Greenwich's dependencies into a shared interpreter changes every other tool that
uses that interpreter.

## Proving the clone is whole

```
py -3.12 selftest.py . <work_dir> build
py -3.12 selftest.py . <work_dir> verify
```

18 assertions, all must be green. They pin base 2024 revenue at 229.525 and EBITDA at 194.675,
the event resets, the CPI cumulative index at 1.06502, the real-to-nominal identity, the aero
calibration end to end, and the base-year regression guard from review finding C1. The split
into two phases exists only so each fits a short shell timeout; a machine without one runs
`all`.

The acceptance oracle is separate and slower:

```
py -3.12 plovdiv_oracle.py "Aero,Non-aero,Opex,Operations Summary"
```

17,438 cells regenerated from Plovdiv's own cached inputs and diffed against its cached
results. Max absolute difference 0.000e+00 on all four sheets as at 8 August 2026. Twelve
cells use `RRI` and are listed as unsupported, never skipped in silence.

## Paths

Nothing in this tree hardcodes a path. `config.py` resolves every location from
`AVIA_LOCAL_CACHE`, so provisioning a host changes one variable and no code. Protect that:
one hardcoded path breaks it silently on the next host, which is exactly what happened three
times before 8 August 2026.

```
Repository      C:\src\greenwich          code and configuration only
Data root       E:\Avia\greenwich         AVIA_LOCAL_CACHE points here
                  reference\              client and reference workbooks
                  outputs\                builds, created on demand
                  outputs_18Jul\          the 18 workbooks built 15-17 July 2026
```

Find a file by landmark, never by counting folders up from `__file__`. `config.find_reference`
searches and reports every path it tried when it fails.

## The chain

`build_model.py` runs the steps in order and each step is its own module:

`generate_skeleton_v1` to `t1_reader` to `actuals_strip` to `axis_picker` to `t6_support`
(vendor reconciliation, optional) to `capex_block` to `financing_group` to `checks_sheet` to
`space_tables` (T2/T3, optional) to `macro_block` (optional) to `output_suite_v1`.

Optional legs are skipped with a printed line, not silently.

## What is not yet wired

Read `GREENWICH_AUDIT_08Aug2026.md` before extending anything. It records what is built and not
connected, with the check that established each finding. In short: the matplotlib chart factory
is not called by any module, `ingest_vendor.py` is not on the chain, `pack2t1` is named in
`t4_engine_pack.tsv` and does not exist, and the self-test never exercises the vendor
reconciliation or Space and Ops legs even though both run.

A default-off switch is a temporary state with an expiry, not a resting place. The register at
the end of the audit names the test that would close each item.

---

Copyright Avia Solutions Limited. All rights reserved.
