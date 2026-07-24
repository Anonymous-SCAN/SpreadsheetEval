"""
verifier_lib.py — OJ-style exact-match grading, per cc.md §1.3 / §7.

Rule (SpreadsheetBench v2):
    A task is correct ONLY when
      (a) every cell in the ground-truth modification set is changed to the
          correct computed value, AND
      (b) no cell OUTSIDE that set is modified relative to the initial state.
    Multi-change or missing-change => reward 0. Single task reward is 0 or 1.

We are handed three things:
    initial_values   = recalc(initial.xlsx)     # workspace the agent started from
    reference_values = recalc(reference.xlsx)    # the true correct solution
    output_values    = recalc(output.xlsx)       # what the agent produced

The ground-truth modification set is derived programmatically as the set of
cells where reference differs from initial (this is exactly the `diff_map`
produced at construction time; we recompute it here so the verifier is
self-contained and does not trust a shipped diff blindly).

Comparison is on normalized values (2-decimal numbers, ISO dates, empty==None),
so it is value-based not formula-based.
"""

from __future__ import annotations
from typing import Dict, Tuple, Any, List

CellKey = Tuple[str, str]


def values_equal(a: Any, b: Any) -> bool:
    """Compare two already-normalized values."""
    if a is None and b is None:
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) < 0.01  # tolerance matches instruction spec (<0.01)
    return a == b


def compute_diff(initial: Dict[CellKey, Any],
                 reference: Dict[CellKey, Any]) -> Dict[CellKey, Any]:
    """Cells whose reference value differs from initial => the modification set.
    Value = the required target value from reference (None means 'must be empty').
    """
    diff: Dict[CellKey, Any] = {}
    keys = set(initial) | set(reference)
    for k in keys:
        iv = initial.get(k)
        rv = reference.get(k)
        if not values_equal(iv, rv):
            diff[k] = rv
    return diff


def grade(initial: Dict[CellKey, Any],
          reference: Dict[CellKey, Any],
          output: Dict[CellKey, Any]) -> Tuple[int, Dict[str, Any]]:
    """Return (reward in {0,1}, detail dict).

    Reward is 1 iff:
      * every cell in the modification set has the correct value in output, AND
      * every cell NOT in the modification set is unchanged from initial.
    """
    mod_set = compute_diff(initial, reference)

    wrong_modified: List[Dict[str, Any]] = []   # required cell has wrong value
    unexpected_changes: List[Dict[str, Any]] = []  # untouched cell was altered

    # (a) required modifications correct?
    for k, target in mod_set.items():
        ov = output.get(k)
        if not values_equal(ov, target):
            wrong_modified.append({
                "cell": f"{k[0]}!{k[1]}",
                "expected": target,
                "got": ov,
            })

    # (b) nothing outside the set may change vs initial
    all_keys = set(initial) | set(reference) | set(output)
    for k in all_keys:
        if k in mod_set:
            continue
        iv = initial.get(k)
        ov = output.get(k)
        if not values_equal(iv, ov):
            unexpected_changes.append({
                "cell": f"{k[0]}!{k[1]}",
                "initial": iv,
                "got": ov,
            })

    passed = (len(wrong_modified) == 0 and len(unexpected_changes) == 0)
    detail = {
        "modification_set_size": len(mod_set),
        "wrong_modified_count": len(wrong_modified),
        "unexpected_change_count": len(unexpected_changes),
        # cap lists so logs stay readable
        "wrong_modified": wrong_modified[:40],
        "unexpected_changes": unexpected_changes[:40],
    }
    return (1 if passed else 0), detail
