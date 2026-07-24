"""
assemble.py — Task assembler (cc.md §5).

Turns the parametric integrated model into concrete evaluation tasks. Both
categories come from the SAME builder — the only difference is which version of
the target cells is injected into the agent's starting workspace (cc.md §5.2):

  Generation (Financial Modeling):
      reference = all-correct workbook
      initial   = correct workbook with a coherent DOWNSTREAM REGION blanked
                  (agent must build those formulas). Because downstream formulas
                  are intact, filling the blanked roots cascades every dependent
                  cell back to the reference value.

  Debugging:
      reference = all-correct workbook
      initial   = all-correct workbook with a set of target cells replaced by
                  their ERROR VARIANTS (agent must find & fix). Errors cascade
                  down the dependency chain -> "三表不平" symptom.

diff_map is computed by recalc(reference) vs recalc(initial): the exact set of
cells whose correct value differs from the starting workspace = the OJ ground
truth modification set. The verifier recomputes this independently.
"""

from __future__ import annotations
import os
import json
import random
import copy
from typing import List, Dict, Any

import openpyxl

import model_threestatement as M
import build_lib as B
import recalc
import verifier_lib as V


# regions available to blank for Generation, keyed by sheet -> allow whole-sheet
GEN_REGIONS = {
    "downstream": ["CashFlow", "BalanceSheet", "DCF"],
    "midstream": ["WorkingCapital", "CashFlow", "DCF"],
    "valuation": ["DCF"],
    "core": ["IncomeStatement", "CashFlow", "BalanceSheet"],
}


def _all_correct_wb(seed, n_years, difficulty, n_distractors):
    wb, targets, meta = M.build(seed=seed, n_years=n_years,
                                difficulty=difficulty,
                                n_distractors=n_distractors)
    for t in targets:
        wb[t.sheet][t.coord] = t.correct
    return wb, targets, meta


def _clone(wb):
    """Round-trip clone of an openpyxl workbook (formulas preserved)."""
    import io
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return openpyxl.load_workbook(buf, data_only=False)


def build_generation_task(seed, n_years=4, difficulty="high",
                          region="downstream", n_distractors=2):
    rng = random.Random(seed * 31 + 1)
    wb_ref, targets, meta = _all_correct_wb(seed, n_years, difficulty, n_distractors)
    blank_sheets = GEN_REGIONS[region]
    to_build = [t for t in targets if t.sheet in blank_sheets]

    # initial = clone with those cells blanked
    wb_init = _clone(wb_ref)
    for t in to_build:
        wb_init[t.sheet][t.coord] = None

    meta.update({
        "category": "Generation",
        "subtype": "FinancialModeling",
        "seed": seed, "region": region,
        "build_sheets": blank_sheets,
        "n_build_cells": len(to_build),
    })
    return wb_ref, wb_init, targets, to_build, meta


def build_debugging_task(seed, n_years=4, difficulty="high",
                         n_bugs=None, n_distractors=2):
    """Inject bugs at the LINE-ITEM (whole-row) level to kill sibling leakage.

    A naive per-cell bug leaves the correct formula visible in the same row's
    other periods, so a model just copies it. Instead we corrupt EVERY period
    of a chosen (sheet, row) line item with the SAME error variant, and we use
    only SUBTLE error types (off-by-one-period / sign-flip / wrong-cell) that
    still produce plausible-looking numbers — the obvious hardcoded-zero is
    excluded. The agent must reason from the stated conventions + the broken
    balance check to locate the roots, not pattern-match.
    """
    import re
    rng = random.Random(seed * 37 + 5)
    wb_ref, targets, meta = _all_correct_wb(seed, n_years, difficulty, n_distractors)

    # group targets by (sheet, row) so we can corrupt a whole line item
    subtle = {"引用偏移 off-by-one-period", "符号取反 sign-flip",
              "引用错误 wrong-cell"}
    upstream = ["Revenue", "PPE", "IncomeStatement", "WorkingCapital", "DebtSchedule"]
    # kinds whose CORRECT formula encodes a non-standard house convention (twist).
    # Biasing bugs onto these means the fix itself requires reproducing the
    # convention — which the Generation results show the model fails at — so a
    # naive "restore the textbook formula" repair is still wrong.
    twist_kinds = {"depreciation", "capex", "interest", "opex", "cogs",
                   "ar", "inventory", "ap", "delta-nwc"}

    def row_of(coord):
        return int(re.sub(r"[A-Z$]", "", coord))

    groups: Dict[tuple, list] = {}
    for t in targets:
        if t.sheet not in upstream:
            continue
        # only rows that have a subtle error available on the per-period cells
        if not any(e["error_type"] in subtle for e in t.errors):
            continue
        groups.setdefault((t.sheet, t.kind), []).append(t)

    all_keys = list(groups.keys())
    # order: twist-carrying line items first (harder), then the rest
    twisty = [k for k in all_keys if k[1] in twist_kinds]
    plain = [k for k in all_keys if k[1] not in twist_kinds]
    rng.shuffle(twisty)
    rng.shuffle(plain)
    group_keys = twisty + plain
    if n_bugs is None:
        n_bugs = {"low": 3, "medium": 5, "high": 8, "extreme": 12}.get(difficulty, 6)
    # here n_bugs counts line-items (each corrupts multiple period cells)
    n_lines = min(n_bugs, len(group_keys))
    chosen_groups = group_keys[:n_lines]

    wb_init = _clone(wb_ref)
    injected = []
    for gk in chosen_groups:
        cells = groups[gk]
        # pick one subtle error TYPE and apply the corresponding variant to
        # every period cell of this line item
        etype = rng.choice(sorted(
            {e["error_type"] for t in cells for e in t.errors
             if e["error_type"] in subtle}))
        for t in cells:
            ev = next((e for e in t.errors if e["error_type"] == etype), None)
            if ev is None:
                continue
            wb_init[t.sheet][t.coord] = ev["formula"]
            injected.append({"sheet": t.sheet, "coord": t.coord,
                             "correct": t.correct, "wrong": ev["formula"],
                             "error_type": ev["error_type"],
                             "cascade": ev["cascade_effect"], "kind": t.kind})

    meta.update({
        "category": "Debugging",
        "subtype": "ErrorRepair",
        "seed": seed, "n_bugs": len(injected),
        "n_broken_lineitems": len(chosen_groups),
        "injected_bugs": injected,
    })
    return wb_ref, wb_init, targets, injected, meta


def compute_diff_map(ref_path, init_path):
    """Ground-truth modification set = recalc(ref) vs recalc(init)."""
    ref_vals = recalc.recalc(ref_path)
    init_vals = recalc.recalc(init_path)
    diff = V.compute_diff(init_vals, ref_vals)
    return {f"{k[0]}!{k[1]}": v for k, v in diff.items()}, ref_vals, init_vals


if __name__ == "__main__":
    import sys
    kind = sys.argv[1] if len(sys.argv) > 1 else "gen"
    seed = int(sys.argv[2]) if len(sys.argv) > 2 else 3
    if kind == "gen":
        ref, init, targets, tb, meta = build_generation_task(seed)
    else:
        ref, init, targets, tb, meta = build_debugging_task(seed)
    ref.save("/tmp/a_ref.xlsx"); init.save("/tmp/a_init.xlsx")
    dm, rv, iv = compute_diff_map("/tmp/a_ref.xlsx", "/tmp/a_init.xlsx")
    print("category:", meta["category"], "seed:", seed)
    print("modification set size:", len(dm))
    print("sample:", list(dm.items())[:6])
