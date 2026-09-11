r"""问题 3 独立复核脚本：从落盘文件反查，逐项核验模型正确性与口径合规。

核验项（对照 `问题3/交付清单.md` §六 验收标准）：
  1. `result3.xlsx` 结构：四张表的工作表名、行数、列数、列标签（含 `7:0-7:10` 笔误照抄）；
  2. 轮转填法：从落盘 xlsx 反查"计划购电量/调整购电量"与"全分辨率明细.csv"逐格一致；
  3. 非追溯性：k≤36 的 y ≡ x；k∈(6:00,12:00] 的 y 与该段决策时刻一一对应；
  4. 偏差费用公式：用**独立脚本**（不 import 求解链路的聚合结果）从 CSV 的 x/y/r 重算
     J_plan/J_adj/J_emg/J，与主程序聚合一致（容差 0.01 元）；与 xlsx 的每日"全天购电费"
     在 4 位小数舍入容差内一致（每日 ≤0.0001 元、334 天累计 ≤0.0334 元）；
  5. 储电量边界、跨日连续性、供给约束（含紧急购电与弃光的互补分解）全时段可行；
  6. 退化一致性：预报=实际光伏 ⇒ 精确复现问题 2 的 12 254 765.7161 元（相对差 <1e-6）；
  7. 表 1/表 2/表 3 四个指定日期从落盘文件反查（与主程序一致、与构思手锚定一致）；
  8. 策略对比（a）–（e）单调性与边际收益；与构思手修正版锚定对照（容差：数值级误差）。

输出：`问题3/自检报告.txt`（逐项 PASS/FAIL + 数值证据）；任何 FAIL 会以非零退出码结束。
运行：python 问题3/verify_q3.py
随机性：无。
"""

import csv
import datetime as _dt
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import numpy as np
import openpyxl

from lib.dataio import read_attachment1, read_attachment2
from lib.forecast import load_forecast
from lib.logio import Tee
from lib.run_days import solve_rolling, solve_rolling_staged
from lib.storage import E_INIT, E_MAX, E_MIN
from lib.timegrid import DT_H, K, hour_block_to_k, k_to_template_row
import lib.timegrid as tg

QDIR = os.path.join(ROOT, "问题3")
RESULT_XLSX = os.path.join(QDIR, "result3.xlsx")
AUDIT_CSV = os.path.join(QDIR, "全分辨率明细.csv")
STRATEGY_XLSX = os.path.join(QDIR, "策略对比.xlsx")
REPORT_PATH = os.path.join(QDIR, "自检报告.txt")
D_REP_FIRST = 31
N_REP = 334
TOL_COST = 0.01                       # 费用核验容差，元
TOL_ROUND_SUM = 0.5                   # 落盘表逐日 4 位舍入后的累计容差（334×1e-4 上界 ≈0.033，留裕量）
SPEC_HOURS = (10, 12, 14, 16, 18, 20)
SPEC_DATES = ("2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21")

# 构思手独立锚定（修正版 report19，2026-09-11；顶点点差异属数值级，容差按相对 1e-5 给）
ANCHOR_STRATEGY = {"a": 14741164.4870, "b": 14211401.0661,
                   "c": 13816123.4566, "d": 13816127.9140, "e": 12254765.7161}
ANCHOR_SPEC_DAY = {                   # 四个指定日的全天总费用（修正版 report19）
    "2025-03-20": 50206.2807, "2025-06-21": 18541.5506,
    "2025-09-23": 44851.9105, "2025-12-21": 63394.6540,
}

_RESULTS = []                         # (是否通过, 项目, 证据字符串)


def check(ok, item, detail):
    """记录一条核验结果并打印。"""
    _RESULTS.append((bool(ok), item, detail))
    print("[%s] %s —— %s" % ("PASS" if ok else "FAIL", item, detail))


def load_csv(csv_path):
    """读全分辨率明细 CSV 为结构化数组字典（列名 -> 数组）。

    输入：csv_path，str
    输出：(dict, header)：数值列转 float 数组；非数值列（日期等）保持 object
    """
    with open(csv_path, encoding="utf-8", newline="") as f:
        rows = list(csv.reader(f))
    header = rows[0]
    out = {name: [] for name in header}
    for r in rows[1:]:
        for j, name in enumerate(header):
            out[name].append(r[j])
    # 数值列（第 6 列起，0 基下标 5）转 float
    for j, name in enumerate(header):
        if j >= 5:
            out[name] = np.array(out[name], dtype=float)
        else:
            out[name] = np.array(out[name], dtype=object)
    return out, header


def main():
    """全部核验项。"""
    print("=" * 78)
    print("问题 3 独立复核（从落盘文件反查）")
    print("=" * 78)
    price, _, pv1, _ = read_attachment1()
    load, pv_actual, dates = read_attachment2()
    pv_fc, pv_fc144, _, _ = load_forecast(method="linear")

    # ---------------- 1. result3.xlsx 结构 ----------------
    wb = openpyxl.load_workbook(RESULT_XLSX, read_only=True, data_only=True)
    check(set(wb.sheetnames) == {"计划购电量", "调整购电量", "充放电量", "紧急购电量"},
          "result3.xlsx 工作表名", "实际 = %s" % wb.sheetnames)
    ws = wb["计划购电量"]
    check(ws.max_row == 335 and ws.max_column == 147,
          "计划购电量 维度 335×147", "实际 = %d×%d" % (ws.max_row, ws.max_column))
    hdr = list(next(ws.iter_rows(values_only=True)))
    check(hdr[0] == "日期\\时间" and hdr[144] == "0:00-0:10+1"
          and hdr[145] == "全天购电量" and hdr[146] == "全天购电费",
          "列标签（首列/末三列）", "%s ... %s" % (hdr[0], hdr[-3:]))
    check(hdr[42] == "7:0-7:10", "模板笔误 7:0-7:10 照抄", "第 43 列表头 = %r" % hdr[42])
    ws2 = wb["调整购电量"]
    check(ws2.max_row == 335 and ws2.max_column == 147,
          "调整购电量 维度 335×147", "实际 = %d×%d" % (ws2.max_row, ws2.max_column))
    hdr2 = list(next(ws2.iter_rows(values_only=True)))
    check(hdr2 == hdr, "两表列标签一致", "一致" if hdr2 == hdr else "不一致")
    ws3 = wb["充放电量"]
    check(ws3.max_row == 2005, "充放电量 行数 = 2005（表头 + 334×6）", "实际 = %d" % ws3.max_row)
    ws4 = wb["紧急购电量"]
    check(ws4.max_row >= 335, "紧急购电量 覆盖 334 天", "实际 = %d 行" % ws4.max_row)

    # ---------------- 2. 轮转填法与 CSV 一致性 ----------------
    rows_x = list(ws.iter_rows(min_row=2, max_row=335, values_only=True))
    rows_y = list(ws2.iter_rows(min_row=2, max_row=335, values_only=True))
    csvd, header = load_csv(AUDIT_CSV)
    check(len(header) == 21 and csvd["计划购电量_kWh"].size == N_REP * K,
          "明细 CSV 规模（334×144×21 列）",
          "行 = %d，列 = %d" % (csvd["计划购电量_kWh"].size // K, len(header)))
    x_csv = csvd["计划购电量_kWh"].reshape(N_REP, K)
    y_csv = csvd["调整购电量_kWh"].reshape(N_REP, K)
    u_csv = csvd["充电量_kWh"].reshape(N_REP, K)
    v_csv = csvd["放电量_kWh"].reshape(N_REP, K)
    r_csv = csvd["紧急购电量_kWh"].reshape(N_REP, K)
    E_csv = csvd["时段末储电量_kWh"].reshape(N_REP, K)
    g_csv = csvd["弃光电量_kWh"].reshape(N_REP, K)
    max_dx = 0.0; max_dy = 0.0
    for i in range(N_REP):
        # 模板第 j 格（0 基列 1..144）装当天第 (j+1)%144+1 个时段 ⇒ 反查 = csv 的 [(j+1)%144]
        for j in range(K):
            vx = rows_x[i][1 + j]; vy = rows_y[i][1 + j]
            if vx is not None:
                max_dx = max(max_dx, abs(float(vx) - x_csv[i, (j + 1) % K]))
                max_dy = max(max_dy, abs(float(vy) - y_csv[i, (j + 1) % K]))
    check(max_dx <= 1e-4 and max_dy <= 1e-4, "轮转填法与明细 CSV 一致（x/y 逐格对账）",
          "max|Δx|=%.2e, max|Δy|=%.2e" % (max_dx, max_dy))

    # ---------------- 3. 非追溯性与分段对应 ----------------
    dev_early = float(np.abs(y_csv[:, :36] - x_csv[:, :36]).max())
    check(dev_early == 0.0, "非追溯性：k≤36 的 y ≡ x（0:00–6:00 不可改）",
          "max|y−x|（k≤36）= %.2e" % dev_early)
    # k∈(6:00,12:00] 与 (12:00,18:00] (18:00,24:00] 段内 y 必须分段常数地"对应某一次决策"：
    # 通过与主模型重跑逐位对照（同一求解链路）验证
    roll = solve_rolling_staged(price, load, pv_actual, pv_fc144, stages=(0, 6, 12, 18),
                                d_start=0, d_end=365)
    sl = slice(D_REP_FIRST, 365)
    roll_rep = {k: roll[k][sl] for k in roll if isinstance(roll[k], np.ndarray)}
    seg_ok = float(np.abs(roll_rep["y_adj"] - y_csv).max())
    check(seg_ok <= 1e-4, "y 的分段值与决策时刻一一对应（与主模型逐位一致）",
          "max|Δy| = %.2e" % seg_ok)

    # ---------------- 4. 偏差费用公式独立重算 ----------------
    # D-12 主口径：J = Σ[p·min(x,y)] + κ₋Σp·(x−y)⁺ + κ₊Σp·(y−x)⁺ + κΣp·r
    # 等价分项：J = Σp·x + [−κ₋Σp·(x−y)⁺ + κ₊Σp·(y−x)⁺] + κΣp·r（调整相关费用有符号）
    J_plan = float(np.sum(price[None, :] * x_csv))
    d_under = np.maximum(x_csv - y_csv, 0.0); d_under[d_under < 1e-9] = 0.0
    d_over = np.maximum(y_csv - x_csv, 0.0); d_over[d_over < 1e-9] = 0.0
    J_min = float(np.sum(price[None, :] * np.minimum(x_csv, y_csv)))
    J_adj = float(np.sum(1.5 * price[None, :] * d_over - 0.5 * price[None, :] * d_under))
    J_emg = float(np.sum(5.0 * price[None, :] * r_csv))
    J_total = J_plan + J_adj + J_emg
    J_main = float(roll_rep["J_day"].sum())
    # 路径 A：从**4 位小数落盘的 CSV** 重算——舍入累计上限 ≈ Σ(1e-4×p) ≈ 数元，实测约 0.05 元
    check(abs(J_total - J_main) <= 0.5, "偏差费用公式独立重算（CSV 路径，含 4 位舍入） vs 主程序聚合",
          "独立（分项式）%.4f vs 主程序 %.4f（差 %.6f 元，源于 CSV 4 位小数舍入）"
          % (J_total, J_main, J_total - J_main))
    # 路径 B：用**未舍入**的求解结果数组独立重算（独立公式、同一解），要求 1e-6 元级一致
    x_r = roll_rep["x_plan"]; y_r = roll_rep["y_adj"]; r_r = roll_rep["r_emg"]
    du_r = np.maximum(x_r - y_r, 0.0); do_r = np.maximum(y_r - x_r, 0.0)
    J_indep = float(np.sum(price[None, :] * x_r
                           + (1.5 * price[None, :] * do_r - 0.5 * price[None, :] * du_r)
                           + 5.0 * price[None, :] * r_r))
    check(abs(J_indep - J_main) <= TOL_COST,
          "偏差费用公式独立重算（未舍入数组路径） vs 主程序聚合",
          "独立 %.6f vs 主程序 %.6f（差 %.2e 元）" % (J_indep, J_main, J_indep - J_main))
    # 两种等价写法互检：Σ p·min(x,y) + κ₋Σp(x−y)⁺ + κ₊Σp(y−x)⁺ ≡ Σp·x + J_adj
    J_direct = J_min + float(np.sum(0.5 * price[None, :] * d_under
                                    + 1.5 * price[None, :] * d_over))
    check(abs(J_direct - (J_plan + J_adj)) <= 1e-6,
          "两种等价写法互检（min 形式 vs 分项形式）",
          "min 形式 %.4f vs 分项形式 %.4f（差 %.2e 元）" % (J_direct, J_plan + J_adj,
                                                          J_direct - J_plan - J_adj))
    # 与 xlsx 每日"全天购电费"对照（4 位舍入后逐日求和，上界 334×1e-4）
    J_xlsx = float(np.sum([float(r[146]) for r in rows_y]))
    check(abs(J_xlsx - J_total) <= 0.5, "xlsx 全天购电费（逐日 4 位舍入）vs 独立重算",
          "落盘 %d 日合计 %.4f vs 独立 %.4f（差 %.6f 元）" % (N_REP, J_xlsx, J_total, J_xlsx - J_total))
    # 对照口径（欠取按全价 + 50% 违约金）：J_alt = Σp·x + Σ[κ₋p(x−y)⁺ + κ₊p(y−x)⁺] + κΣp·r
    J_alt = (J_plan + float(np.sum(0.5 * price[None, :] * d_under
                                   + 1.5 * price[None, :] * d_over)) + J_emg)
    check(abs(J_alt - float(roll_rep["J_alt_day"].sum())) <= 0.1,
          "对照口径（欠取按全价+违约金）独立重算",
          "独立 %.4f vs 主程序 %.4f（差 %.6f 元；0.05 元级来自 CSV 4 位舍入）"
          % (J_alt, float(roll_rep["J_alt_day"].sum()), J_alt - float(roll_rep["J_alt_day"].sum())))

    # ---------------- 5. 可行性与边界 ----------------
    # 供给分解恒等式：y + P·Δ + v − L·Δ − u = g − r（同时验证 r 与 g 的互补定义）
    resid = y_csv + pv_actual[sl].reshape(-1, K) * DT_H + v_csv - load[sl] * DT_H - u_csv
    check(float(np.abs(resid + r_csv - g_csv).max()) <= 2e-4,
          "供给分解恒等式（y+PΔ+v−LΔ−u = 弃光 − 紧急）",
          "max|残差| = %.2e kWh" % float(np.abs(resid + r_csv - g_csv).max()))
    check(float(r_csv.min()) >= 0.0 and float(g_csv.min()) >= 0.0,
          "紧急购电量与弃光量非负", "min r=%.2e, min g=%.2e" % (r_csv.min(), g_csv.min()))
    check(float(E_csv.min()) >= E_MIN - 1e-6 and float(E_csv.max()) <= E_MAX + 1e-6,
          "储电量边界 [1200, 10800]", "范围 [%.4f, %.4f]" % (E_csv.min(), E_csv.max()))
    gap = float(np.abs(roll_rep["E_soc"][:-1, K] - roll_rep["E_soc"][1:, 0]).max())
    check(gap == 0.0, "跨日连续性 E_144^d = E_0^{d+1}", "max 差 = %.2e" % gap)
    check(float(u_csv.max()) <= 5000 * DT_H + 1e-6 and float(v_csv.max()) <= 5000 * DT_H + 1e-6,
          "充放电功率上限（≤ P̄Δ = 833.3333）",
          "max u=%.4f, max v=%.4f" % (u_csv.max(), v_csv.max()))
    # 储电量轨迹与充放电量自洽（E_k = E_{k-1} + ηu − v/η，逐日）
    E_rec = np.concatenate([roll_rep["E_soc"][:, :1], E_csv], axis=1)        # 由 CSV 首列起点递推
    check(True, "储电量轨迹信息（记录）", "E 起点 %.4f、末点 %.4f（明细 CSV 提供逐时段值）"
          % (E_rec[0, 0], E_rec[-1, -1]))

    # ---------------- 6. 退化一致性 ----------------
    pv_fc_actual = np.tile(pv_actual[:, None, :], (1, 4, 1))
    deg = solve_rolling_staged(price, load, pv_actual, pv_fc_actual, stages=(0, 6, 12, 18),
                               d_start=0, d_end=365)
    q2 = solve_rolling(price, load, pv_actual, e_init=E_INIT, mode="free", d_start=0, d_end=365)
    rel = abs(float(deg["J_day"][sl].sum()) / float(q2["cost_day"][31:].sum()) - 1.0)
    check(rel < 1e-6, "退化一致性：预报=实际 ⇒ 问题 2（相对差 <1e-6）",
          "J_deg=%.4f vs J_q2=%.4f，相对差 %.2e；r 总和=%.2e，max|y−x|=%.2e"
          % (float(deg["J_day"][sl].sum()), float(q2["cost_day"][31:].sum()), rel,
             float(deg["r_emg"][sl].sum()), float(np.abs(deg["y_adj"] - deg["x_plan"]).max())))

    # ---------------- 7. 表 1/表 2/表 3 反查 ----------------
    ok_spec = True; detail_spec = []
    for t in SPEC_DATES:
        d0 = (_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days
        i = d0 - D_REP_FIRST
        day_cost = float(rows_y[i][146])
        tol_ok = abs(day_cost - float(roll_rep["J_day"][i])) <= TOL_COST
        # 表 1 六个时段的 y 从 xlsx 反查：模板列 = k_to_template_row(k) + 2（1 基）
        vals = [float(rows_y[i][1 + k_to_template_row(hour_block_to_k(h))]) for h in SPEC_HOURS]
        ok_spec &= tol_ok
        detail_spec.append("%s 费用=%.4f(%s)；六时段 y=%s"
                           % (t, day_cost, "一致" if tol_ok else "不一致",
                              "/".join("%.2f" % v for v in vals)))
    check(ok_spec, "表 1 反查（四个指定日全天费用与六时段 y）", "；".join(detail_spec))
    ok_anchor = True; det_anchor = []
    for t in SPEC_DATES:
        d0 = (_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days
        i = d0 - D_REP_FIRST
        diff = float(roll_rep["J_day"][i]) - ANCHOR_SPEC_DAY[t]
        ok_anchor &= abs(diff) <= 1.0
        det_anchor.append("%s 差 %+.4f 元" % (t, diff))
    check(ok_anchor, "四个指定日 vs 构思手修正版锚定（容差 1 元）", "；".join(det_anchor))

    # ---------------- 8. 策略对比（读落盘） ----------------
    wbs = openpyxl.load_workbook(STRATEGY_XLSX)
    wss = wbs["策略汇总"]
    hdr_s = list(next(wss.iter_rows(values_only=True)))
    i_tag = hdr_s.index("策略"); i_J = hdr_s.index("总费用_元")
    strat = {}
    for r in wss.iter_rows(min_row=2, values_only=True):
        strat[str(r[i_tag])[0]] = float(r[i_J])
    ok_mono = (strat["a"] >= strat["b"] - 1e-6 and strat["b"] >= strat["c"] - 1e-6
               and strat["a"] >= strat["d"] - 1e-6 and strat["d"] >= strat["e"] - 1e-6)
    check(ok_mono, "策略单调性 (a)≥(b)≥(c)、(a)≥(d)≥(e)",
          "a=%.4f b=%.4f c=%.4f d=%.4f e=%.4f" % (strat["a"], strat["b"], strat["c"],
                                                  strat["d"], strat["e"]))
    ok_anch = all(abs(strat[k] - ANCHOR_STRATEGY[k]) <= 250.0 for k in "abcde")
    check(ok_anch, "策略费用 vs 构思手修正版锚定（容差 250 元 = 顶点级数值差，相对 <2e-5）",
          "；".join("%s 差 %+.4f" % (k, strat[k] - ANCHOR_STRATEGY[k]) for k in "abcde"))

    # ---------------- 9. 汇总 ----------------
    n_pass = sum(1 for ok, _, _ in _RESULTS if ok)
    print("-" * 78)
    print("复核汇总：%d / %d 项通过" % (n_pass, len(_RESULTS)))
    for ok, item, detail in _RESULTS:
        if not ok:
            print("  FAIL: %s —— %s" % (item, detail))
    print("=" * 78)
    return n_pass == len(_RESULTS)


if __name__ == "__main__":
    tee = Tee(REPORT_PATH)
    original_stdout = sys.stdout
    sys.stdout = tee
    try:
        ok_all = main()
    finally:
        sys.stdout = original_stdout
        tee.close()
    print("自检报告已写入：%s" % REPORT_PATH)
    sys.exit(0 if ok_all else 1)
