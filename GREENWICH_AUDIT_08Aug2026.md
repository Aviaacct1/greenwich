# Greenwich: pre-git audit and reconciliation position

Version 1.0 - 8 August 2026 - Avia Solutions - working document

Audit only. Nothing was committed, nothing was moved and nothing was edited. Every claim below carries
the check that produced it, and every check was run in this session against the mounted filesystem.

Mounts read: `C:\Users\Carte\OneDrive\Avia`, `E:\Avia`, `C:\Avia`, plus Egnyte `/Shared`.

---

## 1. The headline

Greenwich does not have Meridian's problem. There is no second copy of the generator anywhere. There
is one copy, it is complete, and it runs.

What Greenwich has instead is a **three-way split by artefact type**: the code on OneDrive, the front
end and design set on E:, a mirror of the design set on Egnyte. Plus one **parallel workflow** built
in the last week that reimplements part of the same job and imports Greenwich modules by a dead path.

That is a better starting position than Meridian's five copies, and the git move is correspondingly
cheaper. The expensive work here is not reconciliation. It is the wiring.

---

## 2. Count the copies

| # | Location | Holds | Is it a repo |
|---|---|---|---|
| 1 | `C:\Users\Carte\OneDrive\Avia\Model_refs` | 32 `.py` (23 Greenwich, 9 DDFS), 15 `.tsv`, 4 DDFS `.html`, 26 workbooks, `selftest_fixtures`, `ingest_demo`, `ddfs_bridge_fixtures`, `ddfs_packs`, the state file | No |
| 2 | `E:\Avia\Knowledge Programme` | design notes 00-40, Studio front ends 09 to 09e, `Run Forecast Studio.bat`. No `.py` | No |
| 3 | Egnyte `/Shared/Company Data/14 Avia/AI_System/Knowledge Programme` | 56 files, mirror of copy 2 | Not applicable |
| 4 | `C:\Avia\Neptune` | 29 `.py`, Project Mercury BLQ cargo forecast, 3-7 August | No |

Checked and **not** found: any Greenwich folder under Egnyte `18 Products` (it holds Data, DDFS,
Global Forecast, QSI and Traffic only, so Greenwich is the one tool with no product folder); any
second copy of Model_refs; any `.git` in a Greenwich location. The only repos on the mounted drives
are `C:\Avia\avia_forecast_build` (Atlas), `E:\Avia\Extract`, `E:\Avia\Claude Working\avia-website`
and `E:\Avia\Observatory Website`.

The stray `C:\Avia\qsi-tool` from the Meridian note is still present. Worth confirming the key in it
was rotated and the folder removed.

### Copies 2 and 3 have diverged in both directions

Eight files exist on Egnyte and not on E:: `23 v2`, `25 Fable Strategic Business Review`,
`26 Global Forecast and Cockpit - Work Programme for Jessica Rowden`, `27 QSI Tool - Work Programme
for Nick`, `28 Decision Note - Entity Reference`, `29 Throughput Evidence Log`, `30 Harvest Host
Setup Runbook`, `41 DDFS Unified Tool`. Two files share a name and differ in size: note 35 (E: 4,884
bytes, Egnyte 5,080) and note 37 (E: 5,097 bytes, Egnyte 5,166). The estate index calls E: the
working canon; on this evidence it is not, for those files.

Numbers 23, 25, 26, 27, 28, 29 and 30 each name two different documents, one in the DDFS stream and
one in the strategy stream. That collision is why a reader cannot tell from a number which note is
meant.

---

## 3. What runs today

Copied the 32 `.py`, the 15 `.tsv` and `selftest_fixtures` into a clean empty directory, nothing
else, and ran `selftest.py` in both phases.

**18 assertions, all green.** Base 2024 revenue 229.525 and EBITDA 194.675, the event resets, the
macro CPI index at 1.06502, the real-to-nominal identity, the aero calibration end to end, and the
base-year regression guard from finding C1.

That is the completeness proof for the chain: a clean directory containing only tracked-shape files
reproduces every pinned number. Model_refs is a whole, runnable Greenwich tree.

I then ran the two chain legs the self-test never exercises, because a green suite says nothing about
what it was never asked. Both completed:

- **T2/T3 Space and Ops**: T2 categories `conc_dutyfree`, `conc_fb`, `conc_retail`; R204 uncovered
  none; tier-2 rules appended before Checks row 15.
- **T6 vendor reconciliation**: splits V101 to V103, five codes mapped, flag rows 15 to 17.

Both work. Neither is asserted anywhere.

---

## 4. Built, and not wired

This is where the session's value is. Six items, each verified.

**4.1 The chart factory is orphaned.** No Python file in Model_refs imports `avia_chart_style.py`.
`output_suite_v1.py` builds native Excel charts only, and its own docstring says the matplotlib
factory is the report layer. So Greenwich currently produces no report-layer charts at all. The only
code anywhere that imports `avia_chart_style.py` is in `C:\Avia\Neptune`, by absolute path.

**4.2 The front door is not on the chain.** `ingest_vendor.py`, recorded in the state file's
twenty-sixth update as the pipeline's missing front door, is not called by `build_model.py` and not
called by `selftest.py`. Run with no arguments it raises `IndexError` at line 211 instead of printing
usage.

**4.3 The Plovdiv oracle cannot be rerun.** `plovdiv_oracle.py` line 18 sets
`SRC = "/sessions/inspiring-festive-galileo/mnt/Model_refs/Plovdiv_FinancialModel_vDraft.xlsm"`, a
dead sandbox path from an earlier session. The workbook itself is in Model_refs. The estate index
cites "Plovdiv oracle zero-difference" as Greenwich's verification, and that claim cannot be
reproduced today without editing the file. One-line fix, but it needs making before the claim is used.

**4.4 Two more dead sandbox paths.** `generate_skeleton_v0.py` line 266 and `avia_chart_style.py`
line 78, both pointing at `/sessions/elegant-determined-ramanujan/mnt/...`. Three dead paths from two
different sessions is a pattern, not an accident: absolute output paths get written where a resolved
path belongs.

**4.5 `pack2t1` does not exist.** `t4_engine_pack.tsv` and the state file both name it as the module
that reads a traffic pack into T1 and refuses a `schema_version` it does not know. No file of that
name exists in any mounted location. `t4_engine_pack.tsv` has a comment header and no rows. So the
Atlas-to-Greenwich traffic feed, the link you asked about, is specified, registered and unwritten.
Decision note 14 records the canonical feed as Cockpit output once consistent, with the Studio bridge
as the labelled interim, and the interim is what is live.

**4.6 The Studio launcher opens the wrong front end.** `Run Forecast Studio.bat` opens the most
recently modified file matching `09*Forecast Studio Mockup*.html`. Three files match: `09`, `09b` and
`09c`. The current front end, `09e Forecast Studio - Integrated Ingest, Mapping and Studio v1`, does
not contain the string "Mockup" and never matches. So the launcher opens the 16 July mockup, and
reports nothing amiss. Anyone who has clicked that shortcut since 17 July has been looking at
superseded work.

**One thing that is right, having checked it twice.** 09e does export `project_header_studio.tsv`
alongside `t1_assumptions_studio.tsv`, `t1_actuals_studio.tsv` and `t6_confirmed.tsv`. My first pass
searched for the literal `project_header.tsv` and wrongly concluded the header was missing. The
Studio-to-chain contract is complete.

---

## 5. Greenwich against DDFS in Model_refs

The split is clean and can be made in one pass. Of 32 Python files, 23 are Greenwich-side and nine
carry the `ddfs_` prefix. All four `.html` files are DDFS. **No Python file crosses the boundary in
either direction**: no `ddfs_*` module imports a Greenwich module, and no Greenwich module imports a
`ddfs_*` module. Verified by scanning every file for every other module name.

The one crossing is documentary. `ddfs_front_v1.html` names `avia_chart_style.py` as the canon it
mirrors, which itself mirrors `chart_format.py` in Atlas.

**Chart canon now has three owners.** `chart_format.py` in `C:\Avia\avia_forecast_build\avia_forecast
\outputs`, `avia_chart_style.py` in Model_refs described as its mirror, and a third set restated
inline in `C:\Avia\Neptune\work\23_charts_v2.py` (line 17, `BLUES=[...]`, line 18 `RED="#C00000"`)
under a docstring that says all constants are imported from `avia_chart_style.py`. That file does
both: it imports the module and then hardcodes its own palette. This is the fault Meridian's palette
defect came from, and it is worth fixing before it multiplies further.

---

## 6. What must not enter the repo

Model_refs holds circa 85 MB of workbooks. None of it is committable, and some of it is client
material:

- `Abha_DDFS_Engine_v2.xlsm` 40.9 MB and `DesignDay_Template_v15.xlsm` 41.3 MB (DDFS, not Greenwich)
- Client models: `Plovdiv_FinancialModel_vDraft.xlsm`, `2024.05 - Project East - Financial Model_v24.xlsb`,
  `TNGHIA - Operating Model v9`, `Project Scanner Valuation Model`, `EBITDA FILE - TAS 2023-2024`
- 15 generated `Avia_Model_Skeleton_v*.xlsx` outputs and four demo workbooks
- `__pycache__` with 20 `.pyc`
- The 240 KB session state file

Client-confidential workbooks in a code repository is the item to act on first, ahead of anything
else in this note. `git add -A` would take all of it.

---

## 7. Gaps against the Avia Tool Standard

| Standard | Position | Effect |
|---|---|---|
| 1, git is the single source of truth | No repo exists | Nothing is versioned |
| 3, data on the workstation, never in the repo | Code and 85 MB of workbooks share one folder | See section 6 |
| 4, paths and secrets from config | **Zero** `os.environ` reads across all 23 Greenwich modules. No `AVIA_LOCAL_CACHE` | Three dead absolute paths already, and no way to provision a second host |
| 5, container-ready | Follows from 4 | Not portable |
| 6, edit on the dev PC, run on the workstation | `build_model.py` and `selftest.py` invoke the literal string `"python3"` | Runs on Linux, not on the Dev PC's `py -3.12` |
| Naming register, OneDrive holds documents only | All Greenwich code is on OneDrive | Direct conflict with the settled layout |

Greenwich has no config layer at all. Meridian's `AVIA_LOCAL_CACHE` hinge has no counterpart here, so
this is a build, not a port.

---

## 8. The baseline date, confirmed

The state file's highest block is NINETY-NINTH, dated 24 July, and the file itself was last written
on 24 July. Greenwich's own blocks stop at FORTY-FIFTH, 17 July; everything from FORTY-SIXTH onward
is DDFS. Model_refs' newest Greenwich file is `axis_picker.py` of 18 July.

So the filesystem records no Greenwich work after 18 July. **John confirmed on 8 August 2026 that
this is correct**: the intervening three weeks went to QSI, ask-avia, the Global Forecast, the
harvest and client projects. There is no missing tranche to find, and the filesystem and the record
agree. **The 18 July tree is the baseline.**

This matters for what follows. Greenwich has had the least development of the five tools, so the
tree that goes into git is 23 modules and 15 tables rather than Meridian's 184 files across five
copies. The migration will not be cheaper than it is now.

`C:\Avia\Neptune` is the only recent work of this kind I found: 29 numbered scripts, `01_profile_logs`
through `27_build_v8`, dated 3-7 August. It builds a client workbook, imports Greenwich's chart module
by absolute path, and shares no other code with Greenwich. Whether any of it belongs in the tool
rather than the engagement is a call for you.

---

## 9. Recommended order for the next session

1. **Create `greenwich`, lowercase, empty, first.** Register settled, no discussion needed.
2. **Move the client workbooks out to a data root before anything is staged.** Plovdiv, Project East,
   TNGHIA, Scanner, TAS.
3. **`.gitignore` before `git add`**, then `git add -A -n` and read the list. Exclude `*.xlsx`,
   `*.xlsm`, `*.xls`, `*.xlsb`, `__pycache__/`, `chain_tmp/`, `venv/` and the state file.
4. **Baseline commit and tag**: 23 Greenwich `.py`, 15 `.tsv`, `selftest_fixtures`, `ingest_demo`,
   and 09e as the front end. Prove it with a clone that runs `selftest.py` to 18 green.
5. **Then, in this order**: the config layer and `AVIA_LOCAL_CACHE`; the three dead paths;
   `sys.executable` in place of `"python3"`; the launcher glob; `ingest_vendor` onto the chain or into
   the attic with a reason; the chart factory wired or retired; `pack2t1` written.
6. **DDFS stays where it is** for now, per your instruction, and gets its own session. Note that it
   is the larger half by volume and holds the two 40 MB workbooks.

### Switch and defect register to carry forward

| Item | State | Test that would close it |
|---|---|---|
| `avia_chart_style.py` report layer | Built, not called | An assertion that a build produces the matplotlib chart set |
| `ingest_vendor.py` front door | Built, not on the chain | A self-test case that ingests a vendor pack end to end |
| T6 vendor reconciliation leg | Runs, not asserted | Pin the recon sheet numbers in `selftest.py` |
| T2/T3 Space and Ops leg | Runs, not asserted | Pin the T2 category output in `selftest.py` |
| `plovdiv_oracle.py` | Cannot run, dead `SRC` | Resolve the workbook by search, rerun to zero difference |
| `pack2t1` | Named, does not exist | A pack that loads and a bad `schema_version` that is refused |
| Chart canon | Three owners | One module imported by all three, no restated constants |

---

Avia Solutions Limited. All rights reserved.
