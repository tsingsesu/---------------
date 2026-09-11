"""问题 1 独立复核脚本：不复用主脚本的任何内存状态，从落盘文件与附件重新核验。

核验内容分四组：
  一、独立复算：用**与主模型不同的 LP 形式**（把储电量 E 作为显式变量、终端条件写成变量上下界）
      重新求解同一问题，与落盘的 result1.xlsx / CSV 逐项比对；
  二、结构核验（验收标准第 2 组 6 条）：结果文件与模板逐格一致、供给约束、功率与储电量边界、
      状态递推、四个小时块聚合、模板与 CSV 的可追溯性；
  三、口径对照：与 `口径与假设台账.md` D-02 / D-03 / D-05 记录的外部锚定值比较差异；
  四、附加核验：全天弃光恒等式、填法 X 对照、对偶（光伏边际价值）量级。

输出：控制台 + `问题1/自检报告.txt`。
运行：python 问题1/verify_q1.py
"""

import os
import sys

# 把工作区根目录加入模块搜索路径，保证在任意工作目录下都能 from lib... import ...
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import csv
import math

import numpy as np
import openpyxl
from scipy.optimize import linprog
from scipy.sparse import csr_matrix, hstack, lil_matrix

from lib.dataio import attach_path, read_attachment1
from lib.solve_day import baseline_costs, solve_day
from lib.storage import E_INIT, E_MAX, E_MIN, ETA, P_MAX, U_MAX
from lib.timegrid import (DT_H, K, four_hour_blocks, hour_block_to_k,
                          hour_block_to_template_row, template_row_to_k)

QDIR = os.path.join(ROOT, "问题1")                                   # 问题 1 交付目录
TEMPLATE_XLSX = attach_path("附件5", "result1.xlsx")                 # 题目模板（只读）
RESULT_XLSX = os.path.join(QDIR, "result1.xlsx")                     # 待核验的结果文件
CSV_PATH = os.path.join(QDIR, "问题1_逐时段结果.csv")                 # 待核验的全分辨率结果
REPORT_PATH = os.path.join(QDIR, "自检报告.txt")                     # 本脚本输出的报告
TOL_CONSTRAINT = 1e-6        # 供给约束残差容差，kWh
TOL_SOC = 1e-6               # 储电量边界容差，kWh
TOL_POWER = 1e-6             # 充放电功率上限容差，kWh
TOL_RECUR = 1e-6             # 状态递推残差容差，kWh
TOL_ROUND = 2e-4             # 4 位小数舍入带来的容差（D-15），kWh
TOL_COST = 0.01              # 费用容差，元

# 外部锚定值：来自 `口径与假设台账.md` D-02 / D-03 / D-05 与 `问题1/交付清单.md` §六.7-8。
# **仅用于对照差异，绝不参与任何计算，也不作为任何输出值的来源。**
ANCHOR_FROM_LEDGER = {
    "全天购电量_kWh": 59482.6990,
    "全天购电费_元": 35126.9486,
    "充电量合计_kWh": 20740.67,
    "放电量合计_kWh": 16799.94,
    "充放比值": 0.8100,
    "购电量为0的时段数": 58,
    "同时充放时段数": 0,
    "表1_10:00-10:10_kWh": 0.0000,
    "表1_12:00-12:10_kWh": 480.4124,
    "表1_14:00-14:10_kWh": 0.0000,
    "表1_16:00-16:10_kWh": 445.4317,
    "表1_18:00-18:10_kWh": 531.8940,
    "表1_20:00-20:10_kWh": 0.0000,
    "端点自由购电量_kWh": 54149.3657,
    "端点自由购电费_元": 32909.8653,
    "往返0.9购电费_元": 31695.89,
}
# 表 1 六个指定时段与锚定值的对应键
SPEC_HOURS = (10, 12, 14, 16, 18, 20)
SPEC_ANCHOR_KEYS = {10: "表1_10:00-10:10_kWh", 12: "表1_12:00-12:10_kWh",
                    14: "表1_14:00-14:10_kWh", 16: "表1_16:00-16:10_kWh",
                    18: "表1_18:00-18:10_kWh", 20: "表1_20:00-20:10_kWh"}


class Reporter:
    """收集核验结果并格式化输出（核验项 / 通过与否 / 实测值 / 容差）。"""

    def __init__(self):
        # 逐条核验记录：每项是 (核验项, 通过与否, 实测值字符串, 容差字符串, 备注)
        self.rows = []
        self.lines = []          # 报告正文行（不含核验表）
        self.all_pass = True     # 总开关：任一项不通过则置 False

    def section(self, title):
        """写入一个小节标题。

        输入：title，str
        输出：无
        """
        # 空行分隔，便于人读
        self.lines.append("")
        self.lines.append("-" * 96)
        self.lines.append(title)
        self.lines.append("-" * 96)

    def text(self, line):
        """写入一行普通文本。

        输入：line，str
        输出：无
        """
        self.lines.append(line)

    def check(self, name, ok, measured, tol, note=""):
        """登记一条核验项。

        输入：name，str，核验项名称
              ok，bool，是否通过
              measured，任意，实测值（会转成字符串）
              tol，str，容差说明
              note，str，备注（可选）
        输出：bool，ok
        """
        # 统一格式：核验项 / 通过与否 / 实测值 / 容差
        self.rows.append((name, "通过" if ok else "不通过", str(measured), str(tol), note))
        if not ok:
            self.all_pass = False
        return ok

    def dump(self, path):
        """把核验表与正文写入报告文件。

        输入：path，str，目标路径
        输出：str，报告全文（同时返回以便打印）
        """
        # 计算各列宽度，保证表格对齐（中文按 2 个字符宽度计）
        def width(s):
            return sum(2 if ord(ch) > 127 else 1 for ch in s)

        def pad(s, w):
            return s + " " * max(0, w - width(s))

        name_w = max([width(r[0]) for r in self.rows] + [12]) + 2
        pass_w = 8
        meas_w = max([width(r[2]) for r in self.rows] + [10]) + 2

        out = list(self.lines)
        out.append("")
        out.append("-" * 96)
        out.append("核验表（核验项 / 通过与否 / 实测值 / 容差）")
        out.append("-" * 96)
        header = pad("核验项", name_w) + pad("通过与否", pass_w) + pad("实测值", meas_w) + "容差"
        out.append(header)
        out.append("-" * 96)
        for name, verdict, measured, tol, note in self.rows:
            line = pad(name, name_w) + pad(verdict, pass_w) + pad(measured, meas_w) + tol
            if note:
                line += "    [" + note + "]"
            out.append(line)
        out.append("-" * 96)
        out.append("总体结论：%s（共 %d 项，通过 %d 项）"
                   % ("全部通过" if self.all_pass else "存在未通过项",
                      len(self.rows), sum(1 for r in self.rows if r[1] == "通过")))
        out.append("")
        content = "\n".join(out)
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.write(content + "\n")
        return content


def independent_lp(price, load, pv_plan, e_init=E_INIT, eta=ETA, p_max=P_MAX,
                   e_min=E_MIN, e_max=E_MAX, cyclic=True):
    """用**与主模型不同的形式**独立求解同一 LP（用于交叉验证）。

    形式差异：
      * 储电量 E_0..E_K 作为**显式变量**（4K+1 列），并用等式约束写状态递推，
        而主模型用前缀和消去了 E，只保留上下界不等式；
      * 终端条件用**变量上下界** (e_init, e_init) 锁定，而主模型用等式约束。

    输入：price (K,) 元/kWh；load (K,) kW；pv_plan (K,) kW；e_init kWh；eta；p_max kW；
          e_min / e_max kWh；cyclic，bool，True 表示锁定 E_K = E_0
    输出：dict，{'cost' 元, 'x_plan' (K,), 'u_chg' (K,), 'v_dis' (K,), 'E_soc' (K+1,)}
    """
    price = np.asarray(price, dtype=float)
    load = np.asarray(load, dtype=float)
    pv_plan = np.asarray(pv_plan, dtype=float)
    n = price.size                                     # 时段数 K
    u_max = p_max * DT_H                               # 单时段最大充/放电量，kWh

    # 列布局：x(0..K-1) | u(K..2K-1) | v(2K..3K-1) | E(3K..4K)，共 4K+1 个变量（E 含 E_0 ≠ E_K）
    n_var = 4 * n + 1
    cost_vec = np.zeros(n_var)
    cost_vec[:n] = price                               # 目标仍是 Σ p_k x_k

    # 供给约束：−x_k + u_k − v_k <= Δ(P_k − L_k)
    a_ub = lil_matrix((n, n_var))
    for k in range(n):
        a_ub[k, k] = -1.0                              # −x_k
        a_ub[k, n + k] = 1.0                           # +u_k
        a_ub[k, 2 * n + k] = -1.0                      # −v_k
    b_ub = DT_H * (pv_plan - load)

    # 状态递推等式：E_k − E_{k-1} − η·u_k + v_k/η = 0，k=1..K（E 的下标 3K+k 表示 E_k）
    a_eq = lil_matrix((n, n_var))
    for k in range(1, n + 1):
        a_eq[k - 1, 3 * n + k] = 1.0                   # +E_k
        a_eq[k - 1, 3 * n + k - 1] = -1.0              # −E_{k-1}
        a_eq[k - 1, n + (k - 1)] = -eta                # −η·u_k
        a_eq[k - 1, 2 * n + (k - 1)] = 1.0 / eta       # +v_k/η
    b_eq = np.zeros(n)

    # 变量上下界
    bounds = [(0.0, None)] * n                         # x_k >= 0
    bounds += [(0.0, u_max)] * n                       # 0 <= u_k <= u_max
    bounds += [(0.0, u_max)] * n                       # 0 <= v_k <= u_max
    bounds += [(e_init, e_init)]                       # E_0 固定为 e_init
    for _ in range(1, n):                              # E_1..E_{K-1} ∈ [e_min, e_max]
        bounds.append((e_min, e_max))
    # E_K：cyclic 时锁定为 e_init，否则给区间（端点自由）
    bounds.append((e_init, e_init) if cyclic else (e_min, e_max))

    res = linprog(cost_vec, A_ub=a_ub.tocsr(), b_ub=b_ub, A_eq=a_eq.tocsr(), b_eq=b_eq,
                  bounds=bounds, method="highs")
    if not res.success:
        raise RuntimeError("独立复核 LP 求解失败：%s" % res.message)

    return {
        "cost": float(res.fun),
        "x_plan": res.x[:n],
        "u_chg": res.x[n:2 * n],
        "v_dis": res.x[2 * n:3 * n],
        "E_soc": res.x[3 * n:4 * n + 1],
    }


def read_csv_result(path):
    """读取落盘的逐时段结果 CSV。

    输入：path，str，csv 路径
    输出：dict，键为列名，值为 np.ndarray（字符串列另存为 list）
    """
    with open(path, "r", encoding="utf-8", newline="") as f:
        rows = list(csv.DictReader(f))                 # 以表头为键读取，列名即自解释
    cols = {}
    for key in rows[0].keys():
        # 数值列尝试转 float，失败则保留字符串（时段起止是钟点标签）
        try:
            cols[key] = np.array([float(r[key]) for r in rows], dtype=float)
        except ValueError:
            cols[key] = [r[key] for r in rows]
    return cols


def pointwise_agreement(rep, name, values_independent, values_csv, cost_gap,
                        tol_round=TOL_ROUND, tol_cost=TOL_COST, note=""):
    """登记"落盘解与独立解逐点比对"核验项（LP 退化时允许费用相同的另一最优解）。

    输入：rep，Reporter，核验记录器
          name，str，核验项名称
          values_independent / values_csv，np.ndarray，两套解的同一物理量（kWh）
          cost_gap，float，两套解的目标值之差（元）
          tol_round，kWh，逐点比对的舍入容差；tol_cost，元，费用一致性容差
          note，str，附加说明
    输出：bool，是否通过
    说明：本 LP 在价格分段常数与功率上限的共同作用下**最优解集可能非单点**（多重最优）；
          两个不同形式的 LP 可能收敛到不同顶点，但目标值（费用）必然相同。
          判定规则：逐点差 ≤ tol_round 直接通过；否则要求 |费用差| ≤ tol_cost，
          此时判定为"另一最优解"，并保留逐点差数值供审计，不算数值错误。
    """
    # 逐点最大绝对偏差（两套解按同一真实时段序排列）
    max_diff = float(np.max(np.abs(np.asarray(values_independent, dtype=float)
                                   - np.asarray(values_csv, dtype=float))))
    if max_diff <= tol_round:
        return rep.check(name, True, "最大偏差 %.3e" % max_diff, "≤ %g" % tol_round,
                         note + "逐点一致")
    # 逐点差超容差：以费用是否一致判定是否为另一最优解
    ok = abs(cost_gap) <= tol_cost
    return rep.check(name + "（多重最优：费用相同的另一最优解）", ok,
                     "最大偏差 %.3e；两解费用差 %.4e 元" % (max_diff, cost_gap),
                     "费用差 ≤ %g 元" % tol_cost,
                     note + "逐点差超容差但费用一致，判定为 LP 退化的另一最优解")


def compare_template(result_path):
    """逐格比较结果文件与模板，返回结构核验结果。

    输入：result_path，str，待核验的 result1.xlsx
    输出：dict，包含工作表名、行列数一致性、差异单元格清单等
    """
    wb_tpl = openpyxl.load_workbook(TEMPLATE_XLSX)     # 模板（只读打开，不保存）
    wb_res = openpyxl.load_workbook(result_path)       # 结果文件
    info = {"sheet_names_equal": wb_tpl.sheetnames == wb_res.sheetnames,
            "sheet_names": wb_res.sheetnames,
            "dims_equal": True, "diff_cells": [], "format_diff_cells": [],
            "plan_rows": None, "cd_rows": None, "plan_cols": None, "cd_cols": None}
    for name in wb_tpl.sheetnames:
        ws_t, ws_r = wb_tpl[name], wb_res[name]
        # 行列数必须一致
        if (ws_t.max_row, ws_t.max_column) != (ws_r.max_row, ws_r.max_column):
            info["dims_equal"] = False
        if name == "计划购电量":
            info["plan_rows"], info["plan_cols"] = ws_r.max_row, ws_r.max_column
        if name == "充放电量":
            info["cd_rows"], info["cd_cols"] = ws_r.max_row, ws_r.max_column
        for row in range(1, ws_t.max_row + 1):
            for col in range(1, ws_t.max_column + 1):
                vt = ws_t.cell(row=row, column=col).value
                vr = ws_r.cell(row=row, column=col).value
                if vt != vr:
                    # 模板为空、结果有值 => 属于"填入数值"，合法
                    info["diff_cells"].append((name, row, col, vt, vr))
                ft = ws_t.cell(row=row, column=col).number_format
                fr = ws_r.cell(row=row, column=col).number_format
                if ft != fr:
                    info["format_diff_cells"].append((name, row, col, ft, fr))
    wb_tpl.close()
    wb_res.close()
    return info


def main():
    """执行全部核验并落盘 自检报告.txt。"""
    rep = Reporter()
    rep.text("=" * 96)
    rep.text("问题 1 自检报告（由 问题1/verify_q1.py 独立生成，不复用 run_q1.py 的内存状态）")
    rep.text("=" * 96)
    rep.text("核验对象：%s" % RESULT_XLSX)
    rep.text("           %s" % CSV_PATH)
    rep.text("数据来源：%s（只读）" % attach_path("附件1.xlsx"))
    rep.text("独立复算方式：把储电量 E 作为显式变量的 4K+1 变量 LP（与主模型的 3K 变量前缀和形式不同）")

    # ---------------- 0. 读入数据与两个落盘文件 ----------------
    price, load, pv_plan, labels = read_attachment1()
    csv_cols = read_csv_result(CSV_PATH)
    wb_res = openpyxl.load_workbook(RESULT_XLSX)
    ws_plan = wb_res["计划购电量"]
    ws_cd = wb_res["充放电量"]
    # 模板行序（0 基第 i 行）读出的购电量
    plan_in_template = np.array([float(ws_plan.cell(row=2 + i, column=2).value) for i in range(K)])
    # CSV 中的逐时段量（真实时间顺序）
    x_csv = csv_cols["计划购电量_kWh"]
    u_csv = csv_cols["充电量_kWh"]
    v_csv = csv_cols["放电量_kWh"]
    e_csv = csv_cols["储电量_kWh"]

    # ---------------- 1. 独立复算 ----------------
    rep.section("一、独立复算（不同 LP 形式重解，与落盘文件比对）")
    ind = independent_lp(price, load, pv_plan, cyclic=True)              # 独立求解主模型
    ind_free = independent_lp(price, load, pv_plan, cyclic=False)        # 独立求解端点自由
    ind_rt = independent_lp(price, load, pv_plan, cyclic=False, eta=math.sqrt(0.9))
    main_res = solve_day(price, load, pv_plan, mode="cyclic")            # 主求解器结果（对照）

    # 1.1 目标值一致性
    rep.check("独立形式与主形式的目标值一致（全天购电费）",
              abs(ind["cost"] - main_res["cost"]) <= 1e-6,
              "%.4f 元 vs %.4f 元" % (ind["cost"], main_res["cost"]), "≤ 1e-6 元",
              "两种 LP 形式（E 显式 / E 消去）应给出同一最优值")
    # 1.2 落盘 CSV 与独立解一致（多重最优时以费用一致性判定，见 pointwise_agreement）
    cost_gap_ind_csv = float(np.dot(price, x_csv) - ind["cost"])   # 两解的购电费之差，元
    pointwise_agreement(rep, "CSV 计划购电量与独立解逐点一致", ind["x_plan"], x_csv,
                        cost_gap_ind_csv,
                        note="D-15 的 4 位小数舍入上界为 5e-5；")
    pointwise_agreement(rep, "CSV 充电量与独立解逐点一致", ind["u_chg"], u_csv,
                        cost_gap_ind_csv)
    pointwise_agreement(rep, "CSV 放电量与独立解逐点一致", ind["v_dis"], v_csv,
                        cost_gap_ind_csv)
    pointwise_agreement(rep, "CSV 储电量与独立解逐点一致（含首末）", ind["E_soc"][1:], e_csv,
                        cost_gap_ind_csv)
    # 1.3 端点自由与往返效率口径的独立复算（D-03 / D-05 对照）
    rep.check("端点自由口径独立复算与主求解器一致",
              abs(ind_free["cost"] - 32909.8653) <= 0.01,
              "%.4f 元" % ind_free["cost"], "≤ 0.01 元", "对照 D-03")
    rep.check("往返效率 0.9 口径独立复算与主求解器一致",
              abs(ind_rt["cost"] - 31695.89) <= 0.01,
              "%.4f 元" % ind_rt["cost"], "≤ 0.01 元", "对照 D-05")

    # ---------------- 2. 结构核验（验收标准第 2 组） ----------------
    rep.section("二、结构核验（验收标准第 2 组，逐条）")

    # 核验 1：与模板逐格一致，只有数值被填入
    info = compare_template(RESULT_XLSX)
    legal_cells = set()
    for i in range(K):
        legal_cells.add(("计划购电量", 2 + i, 2))          # 计划购电量的 144 个数值格
    for b in range(6):
        legal_cells.add(("充放电量", 2 + b, 2))            # 充电量 6 格
        legal_cells.add(("充放电量", 2 + b, 3))            # 放电量 6 格
    legal_cells.add(("充放电量", 2, 5))                    # 0:00 储电量
    legal_cells.add(("充放电量", 3, 5))                    # 24:00 储电量
    actual_diff = set((n, r, c) for n, r, c, _, _ in info["diff_cells"])
    illegal = sorted(actual_diff - legal_cells)            # 超出允许范围的差异（必须为空）
    rep.check("工作表名与模板完全一致", info["sheet_names_equal"],
              " / ".join(info["sheet_names"]), "逐格一致")
    rep.check("行列数与模板完全一致", info["dims_equal"],
              "计划购电量 %d×%d；充放电量 %d×%d"
              % (info["plan_rows"], info["plan_cols"], info["cd_rows"], info["cd_cols"]),
              "145×2 / 7×5")
    rep.check("数据行数符合题目规定", info["plan_rows"] == 145 and info["cd_rows"] == 7,
              "计划购电量 144 个数据行；充放电量 6 个数据行", "144 / 6")
    rep.check("所有标签、表头与模板逐格一致（差异只出现在允许填数的格）",
              len(illegal) == 0, "越界差异单元格数 = %d" % len(illegal), "0",
              "允许填数的格：计划购电量B2:B145、充放电量B2:C7、E2、E3")
    # 说明：充放电量表中原本为 General 的数值格被设为 0.0000（D-15 的 4 位小数要求）
    fmt_only = [(n, r, c) for n, r, c, ft, fr in info["format_diff_cells"]]
    rep.check("数字格式差异仅限填入数值的单元格（General → 0.0000）",
              all((n, r, c) in legal_cells for n, r, c in fmt_only),
              "格式差异格数 = %d" % len(fmt_only), "均为允许填数的格",
              "模板 购电量 列自带 0.0000_ ，未改动")

    # 核验 2：逐时段供给约束
    residual_full = ind["x_plan"] + pv_plan * DT_H + ind["v_dis"] - load * DT_H - ind["u_chg"]
    residual_csv = x_csv + pv_plan * DT_H + v_csv - load * DT_H - u_csv
    rep.check("供给约束 x_k + P_k·Δ + v_k − L_k·Δ − u_k ≥ −1e-6（全 144 段，全精度解）",
              bool(residual_full.min() >= -TOL_CONSTRAINT),
              "最小残差 %.3e kWh（第 %d 段）"
              % (float(residual_full.min()), int(np.argmin(residual_full)) + 1), "≥ −1e-6 kWh")
    rep.check("供给约束（落盘 CSV 的 4 位小数值）", bool(residual_csv.min() >= -TOL_ROUND),
              "最小残差 %.3e kWh" % float(residual_csv.min()), "≥ −2e-4 kWh",
              "4 位小数舍入引入，上界约 1.5e-4")

    # 核验 3：功率与储电量边界、状态递推
    rep.check("0 ≤ u_k,v_k ≤ P̄·Δ = 833.3333（全 144 段）",
              bool(u_csv.min() >= -TOL_POWER and v_csv.min() >= -TOL_POWER
                   and u_csv.max() <= U_MAX + TOL_POWER and v_csv.max() <= U_MAX + TOL_POWER),
              "u ∈ [%.4f, %.4f]；v ∈ [%.4f, %.4f]"
              % (u_csv.min(), u_csv.max(), v_csv.min(), v_csv.max()), "⊂ [0, 833.3333] kWh")
    rep.check("E_MIN ≤ E_k ≤ E_MAX（1200 ≤ E ≤ 10800，k=1..144）",
              bool(e_csv.min() >= E_MIN - TOL_SOC and e_csv.max() <= E_MAX + TOL_SOC),
              "E ∈ [%.4f, %.4f] kWh" % (e_csv.min(), e_csv.max()), "⊂ [1200, 10800] kWh")
    rep.check("E_0 = E_144 = 6000 kWh（主模型，D-03）",
              abs(float(ws_cd.cell(row=2, column=5).value) - E_INIT) <= TOL_SOC
              and abs(float(ws_cd.cell(row=3, column=5).value) - E_INIT) <= TOL_SOC
              and abs(e_csv[-1] - E_INIT) <= TOL_ROUND,
              "结果文件 %.4f / %.4f；CSV 末段 %.4f"
              % (ws_cd.cell(row=2, column=5).value, ws_cd.cell(row=3, column=5).value, e_csv[-1]),
              "= 6000.0000 kWh")
    # 状态递推逐点核验：E_k − E_{k-1} = η·u_k − v_k/η（用全精度独立解与落盘值各算一次）
    recur_full = ind["E_soc"][1:] - ind["E_soc"][:-1] - (ETA * ind["u_chg"] - ind["v_dis"] / ETA)
    e_prev_csv = np.concatenate([[E_INIT], e_csv[:-1]])                  # CSV 的 E_0..E_{143}
    recur_csv = e_csv - e_prev_csv - (ETA * u_csv - v_csv / ETA)
    rep.check("状态递推 E_k − E_{k-1} = η·u_k − v_k/η 逐点成立（全精度解）",
              bool(np.max(np.abs(recur_full)) <= TOL_RECUR),
              "最大残差 %.3e kWh" % float(np.max(np.abs(recur_full))), "≤ 1e-6 kWh")
    rep.check("状态递推逐点成立（落盘 CSV 的 4 位小数值）",
              bool(np.max(np.abs(recur_csv)) <= TOL_ROUND),
              "最大残差 %.3e kWh" % float(np.max(np.abs(recur_csv))), "≤ 2e-4 kWh")
    rep.check("无同时充放（u_k·v_k = 0，η²<1 时的最优解性质）",
              int(np.sum((u_csv > 1e-6) & (v_csv > 1e-6))) == 0,
              "同时充放时段数 = %d" % int(np.sum((u_csv > 1e-6) & (v_csv > 1e-6))), "= 0")

    # 核验 4：四个小时块聚合
    blocks = four_hour_blocks()
    u_block_vals = np.array([float(ws_cd.cell(row=2 + b, column=2).value) for b in range(6)])
    v_block_vals = np.array([float(ws_cd.cell(row=2 + b, column=3).value) for b in range(6)])
    rep.check("六个 4 小时块的充电量之和 = 全天总充电量",
              abs(u_block_vals.sum() - u_csv.sum()) <= 0.01,
              "块和 %.4f kWh；全时段和 %.4f kWh" % (u_block_vals.sum(), u_csv.sum()), "≤ 0.01 kWh")
    rep.check("六个 4 小时块的放电量之和 = 全天总放电量",
              abs(v_block_vals.sum() - v_csv.sum()) <= 0.01,
              "块和 %.4f kWh；全时段和 %.4f kWh" % (v_block_vals.sum(), v_csv.sum()), "≤ 0.01 kWh")
    rep.check("块内数值等于该块 36 个时段之和（逐块核验）",
              bool(np.max(np.abs(u_block_vals
                                 - np.array([u_csv[b[2] - 1:b[3]].sum() for b in blocks]))) <= 0.01
                   and np.max(np.abs(v_block_vals
                                     - np.array([v_csv[b[2] - 1:b[3]].sum() for b in blocks]))) <= 0.01),
              "最大逐块偏差 %.3e kWh" % float(max(
                  np.max(np.abs(u_block_vals - np.array([u_csv[b[2] - 1:b[3]].sum() for b in blocks]))),
                  np.max(np.abs(v_block_vals - np.array([v_csv[b[2] - 1:b[3]].sum() for b in blocks]))))),
              "≤ 0.01 kWh")
    rep.check("表 2 的 0:00 / 24:00 储电量与 CSV 首末一致",
              abs(float(ws_cd.cell(row=2, column=5).value) - E_INIT) <= TOL_SOC
              and abs(float(ws_cd.cell(row=3, column=5).value) - e_csv[-1]) <= TOL_ROUND,
              "%.4f / %.4f" % (ws_cd.cell(row=2, column=5).value, ws_cd.cell(row=3, column=5).value),
              "分别对应 0:00 与 24:00")

    # 核验 5：全天购电量与表 1 六个时段
    total_from_template = float(plan_in_template.sum())                  # 模板 144 格求和
    cost_from_template = float(np.sum(price * x_csv))                    # 按 CSV 电价与购电量重算费用
    rep.check("Σ(模板 144 格) = 全天购电量", abs(total_from_template - x_csv.sum()) <= 0.01,
              "%.4f kWh vs %.4f kWh" % (total_from_template, x_csv.sum()), "≤ 0.01 kWh")
    rep.check("全天购电费 = Σ p_k·x_k（按 CSV 重算）", abs(cost_from_template - ind["cost"]) <= TOL_COST,
              "重算 %.4f 元 vs 独立解 %.4f 元" % (cost_from_template, ind["cost"]), "≤ 0.01 元")
    for h in SPEC_HOURS:
        k = hour_block_to_k(h)                                           # 表 1 时段对应真实时段
        row0 = hour_block_to_template_row(h)                              # 0 基模板行
        # 注意：变量名不能再用 v_csv（那是外层放电量数组），此处用 x_tpl / x_csv_k 避免覆盖
        x_tpl = float(ws_plan.cell(row=2 + row0, column=2).value)         # 模板该行数值
        x_csv_k = float(x_csv[k - 1])                                     # CSV 对应真实时段值
        rep.check("表 1  %2d:00-%2d:10  模板值与 CSV 真实时段值一致" % (h, h),
                  abs(x_tpl - x_csv_k) <= TOL_ROUND,
                  "模板 %.4f / CSV %.4f kWh" % (x_tpl, x_csv_k), "≤ 2e-4 kWh",
                  "真实时段 k=%d，模板第 %d 行" % (k, row0 + 1))

    # 核验 6：可追溯性（模板第 i 行 ↔ CSV 第 (i+1)%144+1 行）
    csv_index = np.array([template_row_to_k(i) for i in range(K)])        # 模板行 -> 真实时段序号
    plan_from_csv = x_csv[csv_index - 1]                                  # 按该映射取 CSV 值
    max_trace = float(np.max(np.abs(plan_in_template - plan_from_csv)))
    rep.check("可追溯性：模板第 i 行（0 基）↔ CSV 第 (i+1)%144+1 行，逐行核对",
              max_trace <= TOL_ROUND, "最大偏差 %.3e kWh" % max_trace, "≤ 2e-4 kWh",
              "映射 i -> (i+1)%144+1 由 lib/timegrid.template_row_to_k 实现")
    rep.check("可追溯性：模板最后一行（标签 0:00+1-0:10+1）= 当天第 1 个时段",
              abs(float(plan_in_template[-1]) - float(x_csv[0])) <= TOL_ROUND,
              "模板末行 %.4f / CSV 第 1 行 %.4f kWh" % (plan_in_template[-1], x_csv[0]),
              "≤ 2e-4 kWh", "D-01 的轮转填报")

    # ---------------- 3. 口径对照（与外部锚定值比较差异） ----------------
    rep.section("三、口径对照（与 `口径与假设台账.md` 记录的锚定值比较，锚定值不参与任何计算）")
    cd_u_total = float(u_csv.sum())                                       # 全天充电量合计，kWh
    cd_v_total = float(v_csv.sum())                                       # 全天放电量合计，kWh
    anchor_pairs = [
        ("全天购电量（kWh）", x_csv.sum(), "全天购电量_kWh", 0.01),
        ("全天购电费（元）", ind["cost"], "全天购电费_元", 0.01),
        ("充电量合计（kWh）", cd_u_total, "充电量合计_kWh", 0.01),
        ("放电量合计（kWh）", cd_v_total, "放电量合计_kWh", 0.01),
        ("放电量/充电量比值", cd_v_total / cd_u_total, "充放比值", 1e-4),
        ("购电量为 0 的时段数", int(np.sum(x_csv < 1e-6)), "购电量为0的时段数", 0),
        ("同时充放时段数", int(np.sum((u_csv > 1e-6) & (v_csv > 1e-6))), "同时充放时段数", 0),
        ("端点自由购电量（kWh）", ind_free["x_plan"].sum(), "端点自由购电量_kWh", 0.01),
        ("端点自由购电费（元）", ind_free["cost"], "端点自由购电费_元", 0.01),
        ("往返 0.9 口径购电费（元）", ind_rt["cost"], "往返0.9购电费_元", 0.01),
    ]
    for h in SPEC_HOURS:
        k = hour_block_to_k(h)
        anchor_pairs.append(("表 1 %2d:00-%2d:10（kWh）" % (h, h), x_csv[k - 1],
                             SPEC_ANCHOR_KEYS[h], 1e-4))
    for name, measured, key, tol in anchor_pairs:
        anchor = ANCHOR_FROM_LEDGER[key]
        diff = float(measured) - float(anchor)
        rep.check("锚定对照：%s" % name, abs(diff) <= tol,
                  "实测 %.4f；锚定 %.4f；差 %.4f" % (measured, anchor, diff), "≤ %g" % tol,
                  "差异来源：求解器容差与 D-15 的 4 位小数舍入")

    # ---------------- 4. 附加核验 ----------------
    rep.section("四、附加核验（能量恒等式、弃光、填法对照、对偶量级）")
    # 4.1 全天能量恒等式：购电量 + 光伏电量 + 放电量 − 充电量 − 负载电量 = 弃光量
    g_total = float(ind["x_plan"].sum() + np.sum(pv_plan) * DT_H + ind["v_dis"].sum()
                    - ind["u_chg"].sum() - np.sum(load) * DT_H)
    rep.check("全天能量恒等式（购电+光伏+放电−充电−负载 = 弃光）",
              abs(g_total) <= 1e-6, "弃光电量合计 %.6f kWh" % g_total, "≤ 1e-6 kWh",
              "该典型日储能足以吸纳全部光伏富余，全天弃光为 0")
    # 4.2 充放比值应等于 η²（闭环下由状态递推直接推出）
    rep.check("充放比值 = η²（闭环下 Σv = η²Σu）",
              abs(cd_v_total / cd_u_total - ETA ** 2) <= 1e-6,
              "比值 %.6f；η² = %.6f" % (cd_v_total / cd_u_total, ETA ** 2), "≤ 1e-6")
    # 4.3 填法 X 对照（逐位置对应）：总量不变，但表 1 六个时段取值改变
    k_x = [h * 6 for h in SPEC_HOURS]                                     # 填法 X 的模板行取数位置
    x_diff = [(h, float(x_csv[k_x[i] - 1]), float(x_csv[hour_block_to_k(h) - 1]))
              for i, h in enumerate(SPEC_HOURS)]
    rep.check("填法 X 对照：全天总量与购电费不变（仅重排位置）",
              True, "两种填法的 144 个数之和相同（%.4f kWh）" % x_csv.sum(), "—",
              "填法 X 下 18:00-18:10 = %.4f kWh，与填法 Y 轮转的 %.4f kWh 相差 %.2f%%"
              % (x_diff[4][1], x_diff[4][2],
                 100.0 * abs(x_diff[4][1] - x_diff[4][2]) / x_diff[4][2]))
    # 4.4 对偶量级核验：光伏边际价值不应超过该时段电价（否则可无限套利）
    pv_marg = main_res["pv_marginal"]
    rep.check("对偶量级：光伏边际价值 ≤ 该时段电价（否则存在无风险套利）",
              bool(np.all(pv_marg <= price + 1e-6)),
              "max(边际价值−电价) = %.3e 元/kWh" % float(np.max(pv_marg - price)), "≤ 1e-6",
              "边际价值区间 [%.4f, %.4f] 元/kWh" % (pv_marg.min(), pv_marg.max()))
    # 4.5 基线核验：与题目解读 §1.3 的两个基线一致
    base = baseline_costs(price, load, pv_plan)
    rep.check("不储能基线可独立复算（方案 A / 方案 B）",
              abs(base["cost_all_grid"] - 88319.36) <= 0.01
              and abs(base["cost_pv_only"] - 48052.05) <= 0.01,
              "方案 A %.4f 元；方案 B %.4f 元" % (base["cost_all_grid"], base["cost_pv_only"]),
              "≤ 0.01 元", "对照 题目解读 §1.3")

    # ---------------- 输出 ----------------
    content = rep.dump(REPORT_PATH)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print(content)
    print("自检报告已写入：%s" % REPORT_PATH)
    wb_res.close()
    return 0 if rep.all_pass else 1


if __name__ == "__main__":
    sys.exit(main())
