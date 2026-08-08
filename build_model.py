"""One-command chain driver: T1/T6 tables to the finished Avia model workbook.
Runs the nine generator increments in order. The vendor files (T6 project lines
and vendor actuals) are optional: a Studio export with no dataroom skips the
reconciliation step and the Checks sheet reports R108 as not applicable.
Usage: python3 build_model.py <workdir> <header> <t1> <actuals> <out> [t6 vendor_actuals [t2 t3 [tier [macro]]]]
Author: Avia Solutions."""
import sys, subprocess, os
import config

def run(script, *args):
    # config.python_exe() is sys.executable. The literal string "python3" was here, and it
    # has no counterpart on the Dev PC, which runs py -3.12.
    r = subprocess.run([config.python_exe(), script, *args], capture_output=True, text=True)
    print(r.stdout.strip())
    if r.returncode != 0:
        print(r.stderr.strip()); sys.exit(f"chain stopped at {script}")

def main(wd, header, t1, actuals, out, t6=None, vend=None, t2=None, t3=None, tier="1", macro=None):
    t6 = None if t6 in ("-","") else t6; vend = None if vend in ("-","") else vend
    t2 = None if t2 in ("-","") else t2; t3 = None if t3 in ("-","") else t3
    macro = None if macro in ("-","") else macro
    s = lambda n: os.path.join(wd, n)
    tmp = os.path.join(wd, "chain_tmp"); os.makedirs(tmp, exist_ok=True)
    p = lambda n: os.path.join(tmp, n)
    run(s("generate_skeleton_v1.py"), s("line_sets.tsv"), header, tier, p("a.xlsx"))
    run(s("t1_reader.py"), p("a.xlsx"), t1, header, s("line_sets.tsv"), tier, p("b.xlsx"))
    run(s("actuals_strip.py"), p("b.xlsx"), header, actuals, s("line_sets.tsv"), tier, p("c.xlsx"))
    run(s("axis_picker.py"), p("c.xlsx"), p("d.xlsx"))
    if t6 and vend:
        run(s("t6_support.py"), "recon", p("d.xlsx"), header, t6, vend, p("e.xlsx"))
    else:
        os.replace(p("d.xlsx"), p("e.xlsx")); print("no vendor ingest: reconciliation step skipped")
    run(s("capex_block.py"), p("e.xlsx"), header, p("f.xlsx"))
    run(s("financing_group.py"), p("f.xlsx"), header, p("g.xlsx"))
    run(s("checks_sheet.py"), p("g.xlsx"), header, actuals, s("line_sets.tsv"), tier, p("h.xlsx"))
    if t2 and t3:
        run(s("space_tables.py"), p("h.xlsx"), header, t2, t3, t1, p("i.xlsx"))
    else:
        os.replace(p("h.xlsx"), p("i.xlsx")); print("no T2/T3 tables: Space & Ops step skipped, tier-2 rules stay pending")
    last = p("i.xlsx")
    if macro:
        run(s("macro_block.py"), p("i.xlsx"), header, macro, p("j.xlsx")); last = p("j.xlsx")
    else:
        print("no macro file: Macro block skipped, model stays real-only")
    run(s("output_suite_v1.py"), last, header, out)
    print("BUILT:", out)

if __name__ == "__main__":
    main(*sys.argv[1:])
