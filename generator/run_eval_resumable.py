"""
run_eval_resumable.py — resumable Pass@1 runner. Appends each task result to
model_eval_partial.jsonl and skips tasks already recorded, so it survives
being killed/restarted. Run repeatedly until all tasks are done.
"""
import os
import sys
import json

import model_agent as MA
import openrouter_client as OR

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EVAL_DIR = MA.EVAL_DIR
PARTIAL = os.path.join(ROOT, "model_eval_partial.jsonl")


def done_ids():
    ids = {}
    if os.path.exists(PARTIAL):
        for line in open(PARTIAL):
            line = line.strip()
            if not line:
                continue
            try:
                r = json.loads(line)
                ids[r["task_id"]] = r
            except Exception:
                pass
    return ids


def main():
    task_ids = sorted(d for d in os.listdir(EVAL_DIR)
                      if os.path.isdir(os.path.join(EVAL_DIR, d)))
    done = done_ids()
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else len(task_ids)
    n = 0
    for tid in task_ids:
        if tid in done:
            continue
        if n >= limit:
            break
        try:
            r = MA.run_task(tid)
        except Exception as e:
            r = {"task_id": tid, "reward": 0, "error": str(e)[:200]}
        with open(PARTIAL, "a") as f:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
        print(f"{tid:20s} reward={r['reward']} edits={r.get('n_edits','-')} "
              f"wrong={r.get('wrong','-')} unexpected={r.get('unexpected','-')} "
              f"{r.get('error','')}", flush=True)
        n += 1
    # summary of everything done so far
    done = done_ids()
    if done:
        tot = sum(v["reward"] for v in done.values())
        print(f"\nDONE {len(done)}/{len(task_ids)}  Pass@1 so far = "
              f"{tot/len(done):.3f} ({tot}/{len(done)})", flush=True)


if __name__ == "__main__":
    main()
