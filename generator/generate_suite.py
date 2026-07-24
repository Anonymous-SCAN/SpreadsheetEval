"""
generate_suite.py — build the full SpreadsheetEval benchmark: >=5 Generation +
>=5 Debugging tasks, each a distinct random scenario (seed), at high difficulty.

Fully scalable & manual-free (cc.md §1.4 / §9): change TASKS to scale count,
years, difficulty, distractors. Every task's oracle is validated locally right
after emission (recalc reference -> reward must be 1.0), which mirrors
`harbor run -a oracle`.
"""
from __future__ import annotations
import os
import sys
import json

import emit_harbor as E
import recalc
import verifier_lib as V

# (task_id, category, seed, n_years, difficulty, region, n_bugs, n_distractors)
TASKS = [
    # ---- Financial Modeling (Generation) : >=5 ----
    # all at "extreme" (8 stated house conventions) so the model cannot rely on
    # textbook-default formulas; large build regions => big all-or-nothing modset.
    ("fm_01_saas_dcf",       "Generation", 101, 5, "extreme", "downstream", None, 3),
    ("fm_02_forecast",       "Generation", 102, 5, "extreme", "core",       None, 3),
    ("fm_03_valuation",      "Generation", 103, 5, "extreme", "midstream",  None, 3),
    ("fm_04_integration",    "Generation", 104, 5, "extreme", "downstream", None, 3),
    ("fm_06_template",       "Generation", 106, 5, "extreme", "downstream", None, 3),
    # ---- Debugging : >=5 ----  (whole-line-item subtle bugs, no sibling leak,
    #      biased onto twist-carrying formulas so a textbook-restore still fails.
    #      Sized so the model can *answer within the time budget but fail on
    #      merit* — a stronger signal than a timeout.)
    ("db_01_cascade",        "Debugging",  201, 4, "extreme", None, 10, 3),
    ("db_02_reference",      "Debugging",  202, 5, "extreme", None, 11, 3),
    ("db_04_period",         "Debugging",  204, 5, "extreme", None, 11, 3),
    ("db_05_threestmt",      "Debugging",  205, 4, "extreme", None, 10, 2),
    ("db_06_deepchain",      "Debugging",  206, 5, "extreme", None, 11, 3),
]


def oracle_check(rec):
    iv = recalc.recalc(rec["init_path"])
    rv = recalc.recalc(rec["ref_path"])
    r_oracle = V.grade(iv, rv, rv)[0]
    r_noop = V.grade(iv, rv, iv)[0]
    return r_oracle, r_noop


def main(use_llm=True):
    os.makedirs(E.EVAL_DIR, exist_ok=True)
    summary = []
    for (tid, cat, seed, ny, diff, region, nbugs, ndist) in TASKS:
        rec = E.emit_one(tid, cat, seed, n_years=ny, difficulty=diff,
                         region=region or "downstream", n_bugs=nbugs,
                         n_distractors=ndist, use_llm=use_llm)
        r_oracle, r_noop = oracle_check(rec)
        status = "OK" if (r_oracle == 1 and r_noop == 0) else "*** FAIL ***"
        m = rec["meta"]
        summary.append({
            "task_id": tid, "category": cat, "difficulty": diff,
            "industry": m["industry"], "years": f"{m['years'][0]}-{m['years'][-1]}",
            "sheets": len(m["sheets"]), "modset": rec["modset"],
            "n_bugs": m.get("n_bugs"), "oracle": r_oracle, "noop": r_noop,
        })
        print(f"[{status}] {tid:20s} {cat:11s} sheets={len(m['sheets'])} "
              f"modset={rec['modset']:4d} oracle={r_oracle} noop={r_noop}")
    with open(os.path.join(E.ROOT, "suite_summary.json"), "w") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    n_ok = sum(1 for s in summary if s["oracle"] == 1 and s["noop"] == 0)
    print(f"\n{n_ok}/{len(summary)} tasks pass oracle=1 & noop=0")
    print("wrote suite_summary.json")
    return summary


if __name__ == "__main__":
    use_llm = "--no-llm" not in sys.argv
    main(use_llm=use_llm)
