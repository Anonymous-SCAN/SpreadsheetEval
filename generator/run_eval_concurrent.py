"""
run_eval_concurrent.py — concurrent Pass@1 runner.

Each task is independent (own workbook, one model call, own grading) and the
bottleneck is waiting on the model API, so tasks run in a thread pool. Results
are appended to model_eval_partial.jsonl under a lock; already-recorded tasks
are skipped, so this resumes cleanly after interruption.

Concurrency defaults to 6 (kimi-k3 can 429 if we fire all 10 at once). Override
with the first CLI arg.
"""
import os
import sys
import json
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

import model_agent as MA

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EVAL_DIR = MA.EVAL_DIR
PARTIAL = os.path.join(ROOT, "model_eval_partial.jsonl")

_write_lock = threading.Lock()


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


def _run_one(tid):
    try:
        r = MA.run_task(tid)
    except Exception as e:
        r = {"task_id": tid, "reward": 0, "error": str(e)[:200]}
    with _write_lock:
        with open(PARTIAL, "a") as f:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    print(f"{tid:20s} reward={r['reward']} edits={r.get('n_edits','-')} "
          f"wrong={r.get('wrong','-')} unexpected={r.get('unexpected','-')} "
          f"{r.get('error','')}", flush=True)
    return r


def main():
    concurrency = int(sys.argv[1]) if len(sys.argv) > 1 else 6
    task_ids = sorted(d for d in os.listdir(EVAL_DIR)
                      if os.path.isdir(os.path.join(EVAL_DIR, d)))
    done = done_ids()
    todo = [t for t in task_ids if t not in done]
    print(f"total={len(task_ids)} done={len(done)} todo={len(todo)} "
          f"concurrency={concurrency}", flush=True)

    if todo:
        with ThreadPoolExecutor(max_workers=concurrency) as ex:
            futs = {ex.submit(_run_one, t): t for t in todo}
            for _ in as_completed(futs):
                pass

    done = done_ids()
    rows = [done[t] for t in task_ids if t in done]
    tot = sum(r["reward"] for r in rows)
    print(f"\nDONE {len(rows)}/{len(task_ids)}  "
          f"Pass@1 = {tot/len(rows):.3f} ({tot}/{len(rows)})", flush=True)


if __name__ == "__main__":
    main()
