# SpreadsheetEval

## 1. Benchmark 数据速看

对标 SpreadsheetBench v2，聚焦两类能力：
- **Debugging**：在已有工作簿中识别并修复逻辑/结构/引用错误
- **Generation (Financial Modeling)**：多表集成（Multi-statement integration）、多期预测（Multi-period forecast）、DCF估值、敏感性分析、模板补全（Template completion）等真实财务场景

考察维度：
- 表格操作：跨sheet引用、批量修改、复杂公式嵌套
- 财务知识：三表勾稽、间接法现流表、DCF、递延税等

> 一个核心点：后续我们的难度提升都是基于 “表格操作” 和 “财务知识” 两方面进行的，这也正是该 Benchmark 所考察的能力，正是我们需要去做的

---

## 2. 整体数据难度提升思路

**核心逻辑：难度 = 复杂度 × 高难度原子密度**

- 提高复杂度：单次任务包含更多原子任务
- 提高难度：提高单次任务中，高难度原子的采样权重

**原子任务定义**：单一的表格操作或财务建模的知识点。

> 后续操作的关键点总结如下：
> （1）对于**表格操作**来说，我们增加了表格的需要修改的数量、增加了各种复杂的表格操作、功能等
> （2）对于**财务知识**来说，我们增加了一个特性，即**有一部分财务模型是自定义的**，并让模型基于该模型的上下文去作答，而不是使用自己的内部知识

---

## 3. 题目合成的详解步骤

### Step 1 — 程序化表格生成

题目合成首先需要构建**业务背景**与**作业表格**，之后才能在业务背景下嵌入具体问题。

- “真实业务场景” 的创建和 scaling 较为简单，不管是提前预制并采样，还是LLM自动采样，还是LLM+search采样，都可以实现。 
- “作业表格” 的创建具有难度，尤其是表格创建需要合理且带有复杂度，同时创建过程还需要可以 scaling。

我们有两个设计点，**（1）多步生成从基础开始逐步增加复杂度** 和 **（2）每一步生成都带有约束**。下面是我们的创建流程： 

#### Step 1.1 - 基础元素

我们为表格定义了一组基础元素：

```json
{
  "industry": "基础信息 - 行业",
  "start_year / years": "基础信息 - 年份",
  "参数字典": "基础元素 - 表格参数，即列名，例如：基年收入、每年收入增速、毛利率等"  # 该步骤的约束为：每一个参数都具备了自己的合理取值范围。
}
```

**场景描述示例**：

```
行业：SaaS
时间跨度：2024–2028（5 期）
基年收入：31,421
毛利率：41.2%
费用率：26.5%
税率：16.8%
Capex%：9.6%
折旧年限：7 年
DSO / DIO / DPO = 58 / 71 / 33 天
利率：8.3%
WACC：11.9%
永续增长：2.6%
股本：3,440（千股）
期初余额：现金 7,683 / 固定资产原值 8,402 / 累计折旧 1,897 / 有息负债 3,366
```

> 说明：由于当前仅需生成约 10 道题目，我们暂时采用 6 个行业、约 20 个列名来快速生成。后续可随时替换为 LLM 自动生成合理的行业名称与参数字典（并配以合理的值范围），以实现更加多样化的自动化生成。所有具体数值均通过程序自动化生成。

#### Step 1.2 - 复杂度提升 1：进阶表格功能插入

按照我们之前所说的，我们的难度提升都是基于该 Benchmark 考察的模型能力出发的，包括 “表格操作” 和 “财务知识”。

这一步从 “表格操作” 的角度出发，我们对生成的 `.xlsx` 文件进行**表格功能层面的复杂化处理**。例如说：合并单元格（标题栏）、表头样式设置、数据验证（下拉列表 / 数值区间）、条件格式（负数标红 / 红黄绿灯）、数字格式（货币 / 百分比 / 千分位）、隐藏行（隐藏中间计算过程）、干扰 sheet（纯噪声工作表）等。

同时每一个我们都有约束，以防止出错。例如说合并只用在标题带，绝不合并数据区。

#### Step 1.3 - 复杂度提升 2：非约定俗成的财务知识修改

从 “财务知识” 难度角度，我们对表格进行进一步复杂化。

我们的想法是，前沿模型本身具有很强的财务知识记忆，因此这些财务知识对这些模型来说没有太高的难度。

我们的做法是，对于一个财务概念，故意采用**非标准但合理**的算法，迫使模型不能仅凭记忆套用标准公式。当然我们会在表格或者指令中明确告知模型这些概念，并非刻意设置陷阱。

难度增长在于，我们**要求模型在上下文中学习新的财务知识，而不是从自身知识提取**。

### Step 2 — 具体任务装配

Generation 与 Debugging 出自同一构造器，只是注入工作空间的版本不同：

- **Generation**：把某个下游区域（如 CashFlow / BalanceSheet / DCF）的公式**清空**，待 agent 依勾稽关系补全。
- **Debugging**：向工作簿**注入错误**。

二者共用**同一个正确 reference**，增加数据的利用效率。

> 后续我们会根据实际模型表现，从"数量 + 难度"两个层次，增加这两个数据生成方式的复杂性（参考后续章节）。

### Step 3 — Reference 与 diff_map 生成

- **reference.xlsx**：所有正确公式写入后的工作簿。
- **initial.xlsx**：agent 的起始工作空间（Generation 挖空 / Debugging 注错）。
- **diff_map**：`recalc(reference)` 对 `recalc(initial)` **逐格重算比对**得到的修改集合，即 OJ ground truth（`assemble.compute_diff_map`）。验证器运行时**独立重算**，不盲信随题分发的 diff。

### Step 4 — 业务指令生成

业务场景的题面措辞调用外部 LLM 生成。并要求模型不泄漏公式、坐标清单、错误数量或评分细则等。

### Step 5 — OJ 验证器设计

参考 SpreadsheetBench v2，OJ-style exact match：

| 规则 | 实现 |
|------|------|
| 比较计算后的值 | `recalc.py`：LibreOffice headless 重算；无 LibreOffice 时用 `formulas` 纯 Python 引擎兜底（跨表公式已验证一致） |
| 数值精度 | 保留 2 位小数后比较（容差 0.011） |
| 日期归一化 | 统一为 `YYYY-MM-DD` |
| 空值 | 空串 `""` 与真正空单元格等价（均归一为 None） |
| 修改集合 | 必须**严格等于** ground truth：漏改（应改未改）或多改（不该动而动）均判 0 |
| 单题 reward | 0 或 1，无部分分 |
| dataset score | 全部 task reward 的均值 = **Overall Pass@1** |

修改集合由验证器**独立重算** `reference` vs `initial` 得到（`verifier_lib.compute_diff`），不依赖题目附带的 `diff_map.json`。

### Step 6 — Oracle 自检 + Harbor 落地

每题装配成一个 Harbor task 目录，并在构造期立即断言 **oracle=1 且 noop=0**：
- 提交 `reference.xlsx` → reward 必须 = 1（充分性）；
- 提交未改动的 `initial.xlsx` → reward 必须 = 0（必要性）。

### Step 7 — 难度兜底循环

用 frontier model（kimi-k3）测 Pass@1，若 > 0.2 就**在原有数据上迭代加难**（不重新造题），直到 Pass@1 ≤ 0.2。加难从两个正交的杠杆入手：

- **数量的角度**：加长预测年数（列变多）、扩大挖空/注错的范围、增加需修改的单元格数。修改集越大，all-or-nothing 下"全对"的概率越低。
- **难度的角度**：提高单题开启的"非约定俗成的财务知识"的条数；Debugging 端则把错误**整行注入**、**偏置到口径公式**上，并在题面**隐去错误数量/类型/位置**。

---

## 4. 模型实测分数

- **测评模型**：`moonshotai/kimi-k3`（OpenRouter 调用）
- **协议**：把模型当 agent——喂 `instruction.md` + 序列化后的 `initial.xlsx`（含每格公式/字面值、sheet 几何、隐藏行标注），要求返回单元格编辑的 JSON；应用后重算并 OJ 打分。
- **防止hacking**：模型只能看到题面xlsx和指令本身，看不到 reference / diff_map / 验证器代码这些帮助模型答题的内容。

| task_id | 类型 | 行业 | 期间 | sheet | 修改cell | reward |
|---|---|---|---|---|---|---|
| fm_01_saas_dcf | Generation | 能源 | 2024-2028 | 12 | 161 | 0 |
| fm_02_forecast | Generation | 制造 | 2025-2029 | 12 | 220 | 0 |
| fm_03_valuation | Generation | 地产 | 2025-2029 | 12 | 168 | 0 |
| fm_04_integration | Generation | SaaS | 2024-2028 | 12 | 161 | 0 |
| fm_06_template | Generation | 地产 | 2025-2029 | 12 | 161 | 0 |
| db_01_cascade | Debugging | SaaS | 2025-2028 | 12 | 199 | 0 |
| db_02_reference | Debugging | 医疗 | 2025-2029 | 12 | 236 | 0 |
| db_04_period | Debugging | 医疗 | 2024-2028 | 12 | 234 | 1 |
| db_05_threestmt | Debugging | 医疗 | 2025-2028 | 11 | 188 | 0 |
| db_06_deepchain | Debugging | SaaS | 2024-2028 | 12 | 236 | 1 |
| **Overall** |  |  |  |  |  | **Pass@1=0.200 (2/10)** |

**Overall Pass@1 = 0.200（目标 ≤ 0.2）**

---

## 5. 运行方式（harbor run 命令）

### 5.1 生成整套题目（含本地 oracle 自检）

```bash
cd generator
python3 generate_suite.py            # 用 LLM 生成指令
python3 generate_suite.py --no-llm   # 纯确定性（离线可复现）
```

### 5.2 Harbor Oracle 校验（全部 reward 必须 = 1.0）

```bash
harbor run -p SpreadsheetEval -a oracle
```

Oracle 执行 `solution/solve.sh`（`cp /tests/reference.xlsx /app/output.xlsx`）后运行验证器。`reference.xlsx` 是程序化构造的真实正确解，故全部 task reward = 1.0。

### 5.3 用真实 agent 评测

```bash
harbor run -p SpreadsheetEval -a terminus-2 -m moonshotai/kimi-k3
```

### 5.4 本仓库内置的 Pass@1 评测（无需 Docker）

```bash
cd generator
python3 run_eval_resumable.py        # 断点续跑，结果写 model_eval_partial.jsonl
python3 model_agent.py               # 一次性全量，写 model_eval.json
python3 finalize_readme.py           # 把分数表 + 失败分析回填进本 README
```

---

## 6. 失败 case 分析

共 10 题，通过 2 题（db_04_period、db_06_deepchain），**Overall Pass@1 = 0.200**。8 道失败可分为两类，均为**真实解题能力不足**，而非题目/评测漏洞：

### 6.1 值级错误（模型完整作答但答案错误）

在同规模题目上的多次运行中，模型能在时限内产出完整 JSON 编辑，但计算值与 reference 不符，被严格 OJ 判 0。典型证据：

| task | 现象 | 根因 |
|---|---|---|
| fm_01_saas_dcf | 27 格错误：`DCF!G7` 期望 −2177.44，实际 +2177.44；`DCF!B18/B14`（每股价值/企业价值）留空 | 未遵循"Capex/ΔNWC 记为负"符号约定；未补全估值链末端 |
| fm_02_forecast | 39 格错误 | 期间费用漏加**固定管理费**、毛利率未逐年爬坡 |
| fm_03_valuation | 67 格错误 | 终值误用永续增长法（应为**退出倍数法**）、DCF 未用**期中折现** |
| db_01_cascade | 106 格错误 | 修复时套用教科书默认公式，未复现题面**会计口径约定**；漏追跨表级联 |
| db_05_threestmt | 64 格错误 | 同上，且误改了本已正确的单元格（unexpected change） |

这些是本评测集的核心难点证据：**模型系统性地在数百个单元格上偏离题面明示的非标准口径**，在 all-or-nothing OJ 下无一幸免。

### 6.2 超时（模型无法在预算内产出完整答案）

共计 3 个题目令 kimi-k3（reasoning 模型）在 600s 硬预算内无法输出完整编辑集——这本身也是"多表大体量协同"难度的直接体现（SpreadsheetBench v2 平均 593.5 cells，本集为其缩放版）。

> 注：评测采用**每次调用 600s（10分钟） 硬上限**。
> 值级错误与超时都记为 reward 0。

---

## 7. 目录结构

```
sp/
├── generator/                 # 造题 + 评测全部代码
│   ├── openrouter_client.py   # OpenRouter 封装（Kimi-k3 生成/测评，k2 兜底）
│   ├── recalc.py              # 重算引擎：LibreOffice 优先，formulas 纯 Python 兜底
│   ├── verifier_lib.py        # OJ 严格匹配打分（计算值集合相等）
│   ├── build_lib.py           # 复杂元素注入（合并单元格/校验/条件格式/隐藏/干扰）
│   ├── model_threestatement.py# 参数化三表+DCF 集成模型构造器（含错误变体 + 口径 twist）
│   ├── assemble.py            # 任务装配：Generation / Debugging 两视图 + diff_map
│   ├── instruction_gen.py     # LLM 生成业务指令（防泄漏；带确定性兜底）
│   ├── emit_harbor.py         # 生成 Harbor task 目录
│   ├── generate_suite.py      # 一键生成整套 10 题 + 本地 oracle 校验
│   ├── model_agent.py         # 把前沿模型当 agent 跑，测 Pass@1
│   ├── run_eval_resumable.py  # 可断点续跑的评测器
│   ├── finalize_readme.py     # 把实测分数/失败分析回填进本 README
│   └── task_verify_template.py# 每题 tests/verify.py 的模板
├── SpreadsheetEval/           # 10 个 Harbor task 目录（交付物）
│   └── <task_id>/
│       ├── instruction.md
│       ├── task.toml
│       ├── environment/{Dockerfile, model.xlsx}
│       ├── solution/solve.sh
│       ├── tests/{test.sh, verify.py, recalc.py, verifier_lib.py,
│       │         reference.xlsx, initial.xlsx, diff_map.json}
│       └── task_meta.json
├── suite_summary.json         # 构造期汇总（含 oracle 自检）
├── model_eval.json            # 模型实测结果（汇总）
├── model_eval_partial.jsonl   # 逐题实测记录（断点续跑产物）
└── README.md
```
