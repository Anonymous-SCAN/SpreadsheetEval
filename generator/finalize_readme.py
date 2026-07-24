"""
finalize_readme.py — fill the score table + failure analysis into README.md
from model_eval.json (or model_eval_partial.jsonl) and suite_summary.json.
"""
import os
import json
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
README = os.path.join(ROOT, "README.md")
EVAL = os.path.join(ROOT, "model_eval.json")
PARTIAL = os.path.join(ROOT, "model_eval_partial.jsonl")
SUITE = os.path.join(ROOT, "suite_summary.json")


def load_results():
    if os.path.exists(EVAL):
        d = json.load(open(EVAL))
        return d["results"], d.get("model", "moonshotai/kimi-k3")
    rows = []
    for line in open(PARTIAL):
        line = line.strip()
        if line:
            rows.append(json.loads(line))
    return rows, "moonshotai/kimi-k3"


def main():
    results, model = load_results()
    suite = {s["task_id"]: s for s in json.load(open(SUITE))}
    by_id = {r["task_id"]: r for r in results}

    order = [s["task_id"] for s in json.load(open(SUITE))]
    rows_md = ["| task_id | 类型 | 行业 | 期间 | sheet | 修改cell | reward |",
               "|---|---|---|---|---|---|---|"]
    n, passed = 0, 0
    fails = []
    for tid in order:
        s = suite[tid]
        r = by_id.get(tid)
        if r is None:
            continue
        n += 1
        rew = r["reward"]
        passed += rew
        rows_md.append(
            f"| {tid} | {s['category']} | {s['industry']} | {s['years']} | "
            f"{s['sheets']} | {s['modset']} | {rew} |")
        if rew == 0:
            fails.append((tid, s, r))
    pass1 = passed / n if n else 0.0
    rows_md.append(f"| **Overall** |  |  |  |  |  | **Pass@1={pass1:.3f} ({passed}/{n})** |")
    table = "\n".join(rows_md)

    # failure analysis
    fa = [f"共 {n} 题，通过 {passed} 题，**Overall Pass@1 = {pass1:.3f}**。\n"]
    fa.append("典型失败模式：\n")
    for tid, s, r in fails[:8]:
        w = r.get("wrong", "?"); u = r.get("unexpected", "?")
        ms = r.get("modset", s["modset"]); ne = r.get("n_edits", "?")
        err = r.get("error")
        if err:
            fa.append(f"- **{tid}**（{s['category']}）：{err}")
        else:
            fa.append(
                f"- **{tid}**（{s['category']}，需改 {ms} 格）：模型提交 {ne} 处编辑，"
                f"其中 **{w} 处应改单元格取值错误**、{u} 处误改。"
                f"在严格 OJ（全对才得分）下判 0——根因多为未遵循 Assumptions 中的"
                f"非标准会计口径（如期初折旧/360天基数/期中折现/退出倍数），"
                f"或未追全跨表级联，导致数百格系统性偏差。")
    fa_txt = "\n".join(fa)

    txt = open(README, encoding="utf-8").read()
    txt = txt.replace(
        "<!-- SCORES_TABLE -->\n*（分数表由评测脚本 `finalize_readme.py` 填充，见 `model_eval.json`）*",
        table)
    txt = txt.replace("<!-- PASS1 -->", f"{pass1:.3f}")
    txt = txt.replace("<!-- FAILURE_ANALYSIS -->", fa_txt)
    open(README, "w", encoding="utf-8").write(txt)
    print(f"README updated. Pass@1={pass1:.3f} ({passed}/{n}), {len(fails)} failures")


if __name__ == "__main__":
    main()
