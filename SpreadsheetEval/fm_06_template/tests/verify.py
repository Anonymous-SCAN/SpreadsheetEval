"""
verify.py — in-container verifier for a SpreadsheetEval task.

Runs inside the Harbor verifier phase. Reads:
    /app/output.xlsx        (agent's submission)
    /tests/reference.xlsx   (programmatically-constructed correct solution)
    /tests/initial.xlsx     (the workspace the agent started from)

Recomputes all three via recalc (LibreOffice if available in the image, else the
pure-python `formulas` fallback) and applies the OJ-style exact-match rule from
verifier_lib. Writes reward (0/1) to /logs/verifier/reward.txt and a detailed
report to /logs/verifier/report.json.

This file is self-contained on the harbor side: recalc.py and verifier_lib.py are
copied alongside it into tests/ by the emitter.
"""
import os
import sys
import json

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import recalc          # noqa: E402
import verifier_lib as V  # noqa: E402


def main():
    out_dir = "/logs/verifier"
    os.makedirs(out_dir, exist_ok=True)

    output = os.environ.get("SP_OUTPUT", "/app/output.xlsx")
    reference = os.path.join(HERE, "reference.xlsx")
    initial = os.path.join(HERE, "initial.xlsx")

    def fail(msg, reward=0):
        print("VERIFY FAIL:", msg)
        with open(os.path.join(out_dir, "reward.txt"), "w") as f:
            f.write(str(reward))
        with open(os.path.join(out_dir, "report.json"), "w") as f:
            json.dump({"reward": reward, "error": msg}, f, ensure_ascii=False, indent=2)
        sys.exit(0)

    if not os.path.exists(output):
        fail(f"output workbook not found at {output}; agent must save output.xlsx")

    print(f"[verify] recalc backend = {recalc.active_backend()}")
    try:
        init_vals = recalc.recalc(initial)
        ref_vals = recalc.recalc(reference)
        out_vals = recalc.recalc(output)
    except Exception as e:
        fail(f"recalc error: {e}")

    reward, detail = V.grade(init_vals, ref_vals, out_vals)
    print(f"[verify] reward={reward} "
          f"modset={detail['modification_set_size']} "
          f"wrong={detail['wrong_modified_count']} "
          f"unexpected={detail['unexpected_change_count']}")

    with open(os.path.join(out_dir, "reward.txt"), "w") as f:
        f.write(str(reward))
    with open(os.path.join(out_dir, "report.json"), "w") as f:
        json.dump({"reward": reward, **detail}, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
