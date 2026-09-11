"""问题 2 全年联合 LP（完全预见下界，对照用，非交付文件）。

目的（`问题2/交付清单.md` §三）：给出全年费用的**下界**，量化逐日滚动策略的最优性差距。
模型（与 `_phase0/probe13_q2_joint.py` 同构，本文件用 lib/ 公共模块重写并做独立可行性复核）：

    变量 [x(T) | u(T) | v(T) | E(T)]，T = 365×144 = 52560，共 210240 个变量
      supply_t:  x_t + v_t − u_t ≥ (L_t − P_t)·Δ          （供给不低于净负荷）
      recur_t:   E_t − E_{t−1} − η·u_t + v_t/η = 0        （E_{−1} = E_0 = 6000）
      bounds:    x_t ≥ 0；0 ≤ u_t, v_t ≤ P̄·Δ；1200 ≤ E_t ≤ 10800
      min  Σ p_t·x_t

口径：仿真 2025-01-01 起（E=6000）；"填报区间"= 2025-02-01 至 12-31 共 334 天，
      其联合最优费用 = 联合解在区间内的 Σ p·x（构思手锚定 12 227 243.6427 元）。

输出文件（均在 问题2/ 下）：
  联合LP下界_对照.xlsx   汇总对照 + 联合解逐日购电量/费用
  联合LP下界运行日志.txt  控制台记录

运行：python 问题2/run_q2_joint.py（约数秒；建议先跑 run_q2.py）
依赖：numpy、scipy、openpyxl；lib/ 公共模块。
随机性：无（确定性 LP）。
"""

import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import numpy as np
import openpyxl
from scipy.optimize import linprog
from scipy.sparse import coo_matrix

from lib.dataio import read_attachment1, read_attachment2
from lib.storage import E_INIT, E_MAX, E_MIN, ETA, P_MAX, U_MAX
from lib.timegrid import DT_H, K

QDIR = os.path.join(ROOT, "问题2")
XLSX_PATH = os.path.join(QDIR, "联合LP下界_对照.xlsx")
LOG_PATH = os.path.join(QDIR, "联合LP下界运行日志.txt")

ND = 4
D_REP_FIRST = 31                                        # 填报区间首日 2025-02-01（0 基 31）

# 锚定值（`_phase0/报告18_问题2联合LP下界.txt`）——只用于对照打印
ANCHOR = {
    "joint_rep": 12227243.6427,      # 填报区间联合最优费用，元
    "joint_rep_mean": 36608.5139,    # 填报区间日均，元
    "roll_rep": 12254765.7161,       # 逐日滚动填报区间总费用，元
    "gap": 27522.0734,               # 滚动 − 下界，元
    "gap_pct": 0.2251,               # 相对差，%
}


class Tee:
    """把控制台输出同时写入日志文件（与 run_q2.py 同型）。"""

    def __init__(self, path):
        self.stdout = sys.stdout
        try:
            self.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
        self.file = open(path, "w", encoding="utf-8")

    def write(self, text):
        # 双通道写入，日志与回显一致
        self.stdout.write(text)
        self.file.write(text)

    def flush(self):
        self.stdout.flush()
        self.file.flush()

    def close(self):
        self.file.close()


def build_joint_lp(price_t, load_t, pv_t):
    """构造全年联合 LP 的稀疏矩阵。

    输入：price_t / load_t / pv_t，np.ndarray (T,)，元/kWh、kW、kW；T = 52560
    输出：(c, A_ub, b_ub, A_eq, b_eq, bounds, index)——索引 dict 给出各变量块起点
    """
    T = price_t.size
    n = 4 * T                                           # 变量总数 = [x|u|v|E]
    idx = {"x": 0, "u": T, "v": 2 * T, "E": 3 * T}
    c = np.zeros(n)
    c[idx["x"]:idx["x"] + T] = price_t                  # 目标只含购电费

    # ---- 供给约束（T 行）：−x_t + u_t − v_t ≤ Δ(P_t − L_t) ----
    r = np.arange(T)
    rows = np.concatenate([r, r, r])
    cols = np.concatenate([idx["x"] + r, idx["u"] + r, idx["v"] + r])
    vals = np.concatenate([-np.ones(T), np.ones(T), -np.ones(T)])
    A_ub = coo_matrix((vals, (rows, cols)), shape=(T, n)).tocsr()
    b_ub = -(load_t - pv_t) * DT_H                      # 右端 = Δ(P−L)

    # ---- 状态递推（T 行等式）：E_t − E_{t−1} − η·u_t + (1/η)·v_t = 0 ----
    rows3 = np.concatenate([r, r[1:], r, r])
    cols3 = np.concatenate([idx["E"] + r, idx["E"] + r[:-1], idx["u"] + r, idx["v"] + r])
    vals3 = np.concatenate([np.ones(T), -np.ones(T - 1), -ETA * np.ones(T), (1.0 / ETA) * np.ones(T)])
    A_eq = coo_matrix((vals3, (rows3, cols3)), shape=(T, n)).tocsr()
    b_eq = np.zeros(T)
    b_eq[0] = E_INIT                                    # E_{−1} = 6000 移到右端

    # ---- 变量上下界 ----
    bounds = ([(0.0, None)] * T                          # 计划购电量 x ≥ 0
              + [(0.0, U_MAX)] * T                       # 充电量 u ∈ [0, P̄Δ]
              + [(0.0, U_MAX)] * T                       # 放电量 v ∈ [0, P̄Δ]
              + [(E_MIN, E_MAX)] * T)                    # 储电量 E ∈ [1200, 10800]
    return c, A_ub, b_ub, A_eq, b_eq, bounds, idx


def verify_joint_solution(x, u, v, E, price_t, load_t, pv_t):
    """对联合解做独立可行性复核（从解出发重算全部约束残差）。

    输入：x/u/v/E，np.ndarray (T,)，联合解各变量块；price_t/load_t/pv_t (T,)
    输出：dict，各约束的最大违反量（正数表示违反）
    """
    supply = x + v - u - (load_t - pv_t) * DT_H          # 供给残差（应 ≥ 0）
    e_prev = np.concatenate([[E_INIT], E[:-1]])
    recur = E - e_prev - ETA * u + v / ETA               # 递推残差（应 = 0）
    return {
        "supply_min": float(supply.min()),               # 最小供给余量
        "recur_max_abs": float(np.abs(recur).max()),     # 递推最大绝对残差
        "u_violation": float(max(0.0, u.max() - U_MAX)), # 充电功率越限量
        "v_violation": float(max(0.0, v.max() - U_MAX)), # 放电功率越限量
        "E_violation": float(max(0.0, E.max() - E_MAX, E_MIN - E.min())),  # 储电量越限量
        "x_negative": float(max(0.0, -x.min())),         # 负购电量（应为 0）
    }


def main():
    """联合 LP 主流程：构造 → 求解 → 复核 → 落盘 → 打印对照。"""
    price, _, _, _ = read_attachment1()
    load2, pv_actual, dates = read_attachment2()
    D = load2.shape[0]
    T = D * K

    # 平铺成全年时间串：t = d*144 + k（真实时间顺序）
    price_t = np.tile(price, D)                          # 电价（附件1 广播到 365 天）
    load_t = load2.reshape(-1)
    pv_t = pv_actual.reshape(-1)

    print("=" * 78)
    print("问题 2 全年联合 LP（完全预见下界，对照用）")
    print("=" * 78)
    print("变量数 = %d（4×%d），约束 = %d 不等式 + %d 等式" % (4 * T, T, T, T))
    t0 = time.time()
    c, A_ub, b_ub, A_eq, b_eq, bounds, idx = build_joint_lp(price_t, load_t, pv_t)
    print("构造耗时 %.1f s，开始求解……" % (time.time() - t0))
    res = linprog(c, A_ub=A_ub, b_ub=b_ub, A_eq=A_eq, b_eq=b_eq, bounds=bounds, method="highs")
    print("求解状态：status=%d（0=最优），总耗时 %.1f s" % (res.status, time.time() - t0))
    if not res.success:
        raise RuntimeError("联合 LP 求解失败：%s" % res.message)

    # ---- 拆解变量块并独立复核 ----
    x = res.x[idx["x"]:idx["x"] + T]
    u = res.x[idx["u"]:idx["u"] + T]
    v = res.x[idx["v"]:idx["v"] + T]
    E = res.x[idx["E"]:idx["E"] + T]
    chk = verify_joint_solution(x, u, v, E, price_t, load_t, pv_t)

    # ---- 费用口径 ----
    cost_all = float(np.dot(price_t, x))                 # 全年（1.1–12.31）联合最优
    seg = slice(D_REP_FIRST * K, T)                      # 填报区间切片（2.1 起）
    cost_rep = float(np.dot(price_t[seg], x[seg]))       # 填报区间联合最优（= 滚动下界）
    x_rep = float(x[seg].sum())                          # 填报区间购电量
    # 联合解的紧急购电验证（完全预见下不应有缺口）：r = [L·Δ + u − x − P·Δ − v]^+
    r_joint = np.maximum((load_t - pv_t) * DT_H + u - v - x, 0.0)

    print("-" * 78)
    print("【独立可行性复核（从解重算全部约束）】")
    print("供给残差最小值 = %.3e kWh（应 ≥ 0）" % chk["supply_min"])
    print("状态递推残差最大值 = %.3e kWh（应 ≈ 0）" % chk["recur_max_abs"])
    print("功率/储电量越限量 = %.3e / %.3e / %.3e kWh；负购电量 = %.3e kWh"
          % (chk["u_violation"], chk["v_violation"], chk["E_violation"], chk["x_negative"]))
    print("联合解紧急购电量 max r = %.6f kWh（完全预见应恒为 0）" % float(r_joint.max()))
    print("联合解储电量轨迹：最小 %.4f，最大 %.4f；期末 E(12.31 24:00) = %.4f kWh"
          % (E.min(), E.max(), E[-1]))

    print("-" * 78)
    print("【费用对照（填报区间 334 天）】")
    print("全年（1.1–12.31）联合最优总费用 = %.4f 元；日均 = %.4f 元" % (cost_all, cost_all / D))
    print("填报区间联合最优费用 = %.4f 元（锚定 %.4f，差 %+.4f 元）"
          % (cost_rep, ANCHOR["joint_rep"], cost_rep - ANCHOR["joint_rep"]))
    print("填报区间日均 = %.4f 元（锚定 %.4f）" % (cost_rep / 334.0, ANCHOR["joint_rep_mean"]))
    print("填报区间购电量 = %.4f kWh" % x_rep)
    gap = ANCHOR["roll_rep"] - cost_rep
    print("逐日滚动（锚定） − 联合下界 = %.4f 元（%.4f%%）"
          % (gap, 100.0 * gap / cost_rep))
    print("⇒ 滚动策略距完全预见下界仅 %.4f%%，逐日滚动近优；该差距即跨日协调可再挖掘的全部空间"
          % (100.0 * gap / cost_rep))

    # ---- 落盘：汇总 + 联合解逐日明细 ----
    daily_rows = [["日期", "购电量_kWh", "购电费_元", "24:00储电量_kWh"]]
    for d in range(D):
        s = slice(d * K, (d + 1) * K)
        daily_rows.append([dates[d].isoformat(),
                           round(float(x[s].sum()), ND),
                           round(float(np.dot(price_t[s], x[s])), ND),
                           round(float(E[(d + 1) * K - 1]), ND)])
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "汇总对照"
    for row in [
        ["指标", "数值", "单位", "说明"],
        ["全年联合最优总费用（1.1–12.31）", round(cost_all, ND), "元", "210240 变量单一 LP"],
        ["填报区间联合最优费用（2.1–12.31）", round(cost_rep, ND), "元", "逐日滚动的完全预见下界"],
        ["填报区间日均费用", round(cost_rep / 334.0, ND), "元", "334 天均值"],
        ["填报区间购电量", round(x_rep, ND), "kWh", "联合解在区间内的 Σx"],
        ["逐日滚动总费用（主模型）", round(ANCHOR["roll_rep"], ND), "元", "来自 run_q2.py"],
        ["滚动 − 下界", round(gap, ND), "元", "跨日协调可再挖掘的全部空间"],
        ["相对差", round(100.0 * gap / cost_rep, ND), "%", "相对下界"],
        ["联合解紧急购电量最大值", round(float(r_joint.max()), 6), "kWh", "完全预见应恒为 0"],
        ["联合解期末储电量", round(float(E[-1]), ND), "kWh", "12.31 24:00"],
        ["可行性复核：供给残差最小", "%.3e" % chk["supply_min"], "kWh", "应 ≥ 0"],
        ["可行性复核：递推残差最大", "%.3e" % chk["recur_max_abs"], "kWh", "应 ≈ 0"],
    ]:
        ws.append(row)
    ws2 = wb.create_sheet("联合解逐日明细")
    for row in daily_rows:
        ws2.append(row)
    wb.save(XLSX_PATH)
    print("-" * 78)
    print("已落盘：%s" % XLSX_PATH)
    print("=" * 78)


if __name__ == "__main__":
    tee = Tee(LOG_PATH)
    original_stdout = sys.stdout
    sys.stdout = tee
    try:
        main()
    finally:
        sys.stdout = original_stdout
        tee.close()
        print("运行日志已写入：%s" % LOG_PATH)
