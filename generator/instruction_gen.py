"""
instruction_gen.py — turn a task's structural facts into a natural-language
business instruction (cc.md §4.6 / Appendix C).

Uses the LLM (Kimi-k3) to phrase the instruction like a real analyst brief, but
NEVER leaks formulas or the test rubric. Falls back to a deterministic template
if the model is unavailable, so the pipeline stays fully automated & offline-safe.
"""
from __future__ import annotations
import json
import textwrap

import openrouter_client as OR


SYS = ("你是一位资深财务系统架构师。把结构化的任务事实改写成给分析师的业务指令。"
       "要求：用业务语言，不出现任何Excel公式语法或单元格坐标清单；说明数据源与产出位置；"
       "标注边界条件与验证标准（如资产负债表须平衡，误差<0.01）；专业、简洁、可操作。"
       "只输出指令正文的markdown，不要额外解释。")


def _gen_facts(meta):
    return {
        "category": "Generation / Financial Modeling",
        "industry": meta["industry"],
        "years": meta["years"],
        "sheets_to_build": meta["build_sheets"],
        "n_cells_to_build": meta["n_build_cells"],
        "headline_outputs": ["每股价值 (DCF!每股价值)", "企业价值"],
    }


def _dbg_facts(meta):
    # NOTE: deliberately does NOT expose the number of bugs, their types, or the
    # exact sheets they live in. Revealing the count turns localization into a
    # bounded search ("find exactly N"); revealing types/sheets narrows it
    # further. The agent must independently trace the broken dependency chain
    # from the symptom (unbalanced statements) — that is the actual skill under
    # test.
    return {
        "category": "Debugging / Error Repair",
        "industry": meta["industry"],
        "years": meta["years"],
        "symptom": "资产负债表不平（Check行≠0），下游估值与净利润失真",
    }


def generate(meta, workspace_file="model.xlsx", use_llm=True):
    if meta["category"] == "Generation":
        facts = _gen_facts(meta)
    else:
        facts = _dbg_facts(meta)
    body = None
    if use_llm:
        try:
            prompt = (f"任务事实：\n{json.dumps(facts, ensure_ascii=False, indent=2)}\n\n"
                      f"工作簿文件名：{workspace_file}（多sheet集成财务模型）。"
                      f"请生成业务指令。")
            body = OR.chat([{"role": "system", "content": SYS},
                            {"role": "user", "content": prompt}],
                           temperature=0.4, max_tokens=1200)
        except Exception as e:
            body = None
    if not body:
        body = _fallback(meta, facts, workspace_file)
    # always append the hard, machine-checkable contract
    body += _contract(meta, workspace_file)
    return body


def _fallback(meta, facts, wf):
    if meta["category"] == "Generation":
        return textwrap.dedent(f"""\
        # 三表联动财务模型补全（{facts['industry']}，{facts['years'][0]}–{facts['years'][-1]}）

        你收到一个多 sheet 集成财务模型 `{wf}`。假设参数（Assumptions）与上游若干
        计算表已经填好，但 **{'、'.join(facts['sheets_to_build'])}** 等下游表格中的公式
        单元格被清空，需要你根据财务勾稽关系补全。

        请完成：
        - 依据间接法编制现金流量表（净利润 + 折旧摊销 − 营运资本变动 = 经营现金流），
          衔接期初/期末现金。
        - 依据勾稽关系补全资产负债表各科目，使 **资产 = 负债 + 所有者权益**。
        - 完成 DCF 估值：无杠杆自由现金流、折现因子、现值、终值、企业价值、股权价值、
          每股价值。

        所有数据来源于同一工作簿的上游表格与 Assumptions 表，请使用跨表引用而非硬编码。
        """)
    else:
        return textwrap.dedent(f"""\
        # 财务模型纠错（{facts['industry']}，{facts['years'][0]}–{facts['years'][-1]}）

        你收到一个多 sheet 集成财务模型 `{wf}`。该模型"能算但结果不对"：
        **资产负债表的平衡校验行（Check）不为 0**，净利润与 DCF 估值明显失真。

        模型中存在若干处公式错误。请你自行排查、定位并修复**所有**错误公式，
        使模型恢复自洽。错误可能出现在任意计算表（收入、固定资产/折旧、利润表、
        营运资本、债务计划、现金流、资产负债表、DCF 等）的任意期间；一处错误可能
        沿依赖链向下游多个科目级联，也可能同一科目的多个期间同时出错。

        请注意 Assumptions 表中列明的**会计口径约定（House conventions）**——
        正确公式必须与这些约定一致。修复时只应更正确实错误的单元格，
        不要改动本来正确的单元格。
        """)


def _contract(meta, wf):
    return textwrap.dedent(f"""

    ## 交付要求（评分口径）
    - 在工作目录下将最终工作簿另存为 **`output.xlsx`**（保留全部 sheet 与结构）。
    - 评分为 OJ 严格匹配：对 `output.xlsx` 用 LibreOffice 重算后，
      **所有应修改单元格的计算值必须正确，且所有不应改动的单元格保持不变**；
      多改、少改均判 0 分（无部分分）。数值比较保留两位小数。
    - 验证标准：资产负债表 `Check` 行各期须为 0（|误差| < 0.01）。
    """)
