"""问题 4 全年联合 LP（完全预见下界，对照用，非交付文件）。

目的（`问题4/交付清单.md` §二.4 与 §六.5）：把价格换成附件4 后，逐日滚动策略（4-2 链）相对
**完全预见下界**的差距必须实测。模型与 `问题2/run_q2_joint.py` 完全同构，唯一差异是价格：

    变量 [x(T) | u(T) | v(T) | E(T)]，T = 365×144 = 52560，共 210240 个变量
      supply_t:  x_t + v_t − u_t ≥ (L_t − P_t)·Δ          （供给不低于净负荷）
      recur_t:   E_t − E_{t−1} − η·u_t + v_t/η = 0        （E_{−1} = E_0 = 6000）
      bounds:    x_t ≥ 0；0 ≤ u_t, v_t ≤ P̄·Δ；1200 ≤ E_t ≤ 10800
      min  Σ p_t·x_t          （p = 附件4 逐日逐时段电价，铺平成 52560 维）

口径：仿真 2025-01-01 起（E=6000）；"填报区间"= 2025-02-01 至 12-31 共 334 天，
      其联合最优费用 = 联合解在区间内的 Σ p·x（与问题 2 的对照口径一致）。
      "逐日滚动 − 联合下界"的差距即**跨天协调（跨天套利）可再挖掘的全部空间**（D-09 的实测）。

输出文件（均在 问题4/ 下）：
  联合LP下界_对照.xlsx   汇总对照 + 联合解逐日购电量/费用
  联合LP下界运行日志.txt  控制台记录
  全年联合LP下界预览.png  逐日购电量对比（联合解 vs 逐日滚动解）与逐日费用差

运行：python 问题4/run_q4_joint.py（约数秒；建议先跑 run_q4.py 生成 4-2 明细）
依赖：numpy、scipy、openpyxl、matplotlib；lib/ 公共模块。
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

from lib.dataio import read_attachment2, read_attachment4
from lib.logio import Tee
from lib.storage import E_INIT, E_MAX, E_MIN, ETA, P_MAX, U_MAX
from lib.timegrid import DT_H, K

QDIR = os.path.join(ROOT, "问题4")
FIGDIR = os.path.join(ROOT, "图片", "问题4")                        # 图片统一目录（2026-09-11 起，用户要求）
os.makedirs(FIGDIR, exist_ok=True)                                # 确保目录存在（重跑时自动建）
XLSX_PATH = os.path.join(QDIR, "联合LP下界_对照.xlsx")
LOG_PATH = os.path.join(QDIR, "联合LP下界运行日志.txt")
PNG_PREVIEW = os.path.join(FIGDIR, "全年联合LP下界预览.png")
CSV_42 = os.path.join(QDIR, "全分辨率明细_4-2.csv")     # 逐日滚动解（run_q4.py 落盘）

D_REP_FIRST = 31                                        # 填报区间首日 2025-02-01（0 基 31）
N_REP = 334

ND = 4                                                  # 小数位数（D-15）

# 主模型锚定（`问题4/主模型运行日志.txt`，run_q4.py 落盘）——只用于对照打印
ANCHOR = {
    "roll_42": 12815460.2655,       # 4-2 逐日滚动实际缴费，元
    "roll_43": 14450082.3797,       # 4-3 多阶段滚动总费用 J，元
    "resid_42": 469.0667,           # 4-2 期末残值（仅期末计一次），元
}


def build_joint_lp(price_t, load_t, pv_t):
    """构造全年联合 LP 的稀疏矩阵（与问题 2 的 run_q2_joint.py 完全同构）。

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


def verify_joint_solution(x, u, v, E, load_t, pv_t):
    """对联合解做独立可行性复核（从解出发重算全部约束残差）。

    输入：x/u/v/E，np.ndarray (T,)，联合解各变量块；load_t/pv_t (T,)，kW
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


def read_roll_daily():
    """从 4-2 全分辨率明细 CSV 汇总逐日购电量与购电费（供图与对照表）。

    输入：无（读 `问题4/全分辨率明细_4-2.csv`，由 run_q4.py 生成）
    输出：(days, x_day, cost_day)——日期 list[str]、逐日购电量 (334,) kWh、逐日购电费 (334,) 元
    """
    import csv
    days, x_acc, c_acc = [], [], []
    cur = None
    x_sum = c_sum = 0.0
    with open(CSV_42, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            d = row["日期"]
            if d != cur:
                if cur is not None:
                    days.append(cur); x_acc.append(x_sum); c_acc.append(c_sum)
                cur = d; x_sum = c_sum = 0.0
            x_sum += float(row["计划购电量_kWh"])          # 逐时段求和（单位 kWh）
            c_sum += float(row["计划购电量_kWh"]) * float(row["电价_元每kWh"])   # p·x，元
    days.append(cur); x_acc.append(x_sum); c_acc.append(c_sum)
    return days, np.array(x_acc), np.array(c_acc)


def draw_preview(dates, x_joint_day, cost_joint_day, x_roll_day, cost_roll_day):
    """绘制逐日购电量与费用对照图（联合解 vs 逐日滚动）。

    输入：dates list[date]；x_joint_day/x_roll_day (334,) kWh；cost_joint_day/cost_roll_day (334,) 元
    输出：str，PNG 路径
    """
    import matplotlib.pyplot as plt
    from lib.plotstyle import (COLOR_BUY, COLOR_PRICE, COLOR_REF, FIGSIZE_TALL,
                               apply_chinese_style, save_figure)
    apply_chinese_style()
    fig, axes = plt.subplots(2, 1, figsize=FIGSIZE_TALL)
    x = np.arange(len(dates))
    ax = axes[0]
    ax.plot(x, x_roll_day, color=COLOR_BUY, linewidth=0.9, label="逐日滚动（4-2 链）")
    ax.plot(x, x_joint_day, color=COLOR_PRICE, linewidth=0.9, linestyle="--",
            label="全年联合 LP（完全预见）")
    ticks = [0, 58, 120, 181, 242, 303, 333]
    ax.set_xticks(ticks, [dates[t] for t in ticks], rotation=30, fontsize=9)
    ax.set_xlabel("日期（2025 年，填报区间）")
    ax.set_ylabel("逐日购电量（kWh）")
    ax.set_title("逐日购电量：联合解 vs 逐日滚动（均值 %.0f vs %.0f kWh）"
                 % (x_joint_day.mean(), x_roll_day.mean()))
    ax.legend(fontsize=10)
    ax2 = axes[1]
    ax2.plot(x, cost_roll_day - cost_joint_day, color=COLOR_PRICE, linewidth=0.9,
             label="逐日购电费差（滚动 − 联合）")
    ax2.axhline(0, color=COLOR_REF, linewidth=1.0)
    ax2.set_xticks(ticks, [dates[t] for t in ticks], rotation=30, fontsize=9)
    ax2.set_xlabel("日期（2025 年，填报区间）")
    ax2.set_ylabel("费用差（元）")
    ax2.set_title("逐日费用差（滚动 − 联合）：总差 %.2f 元 = 跨天协调的全部可挖掘空间"
                  % float((cost_roll_day - cost_joint_day).sum()))
    ax2.legend(fontsize=10)
    fig.suptitle("附件4 波动电价下：逐日滚动策略与完全预见下界的逐日对照", fontsize=14)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    return save_figure(fig, PNG_PREVIEW)


def main():
    """联合 LP 主流程：构造 → 求解 → 复核 → 对照（4-2/4-3）→ 落盘 → 打印。"""
    load2, pv_actual, dates = read_attachment2()
    pr4, dates4 = read_attachment4()
    D = load2.shape[0]
    T = D * K

    # 平铺成全年时间串：t = d*144 + k（真实时间顺序）；价格 = 附件4（唯一输入差异）
    price_t = pr4.reshape(-1)
    load_t = load2.reshape(-1)
    pv_t = pv_actual.reshape(-1)

    print("=" * 78)
    print("问题 4 全年联合 LP（附件4 波动电价；完全预见下界，对照用）")
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
    chk = verify_joint_solution(x, u, v, E, load_t, pv_t)

    # ---- 费用口径 ----
    cost_all = float(np.dot(price_t, x))                 # 全年（1.1–12.31）联合最优，元
    seg = slice(D_REP_FIRST * K, T)                      # 填报区间切片（2.1 起）
    cost_rep = float(np.dot(price_t[seg], x[seg]))       # 填报区间联合最优（= 滚动下界），元
    x_rep = float(x[seg].sum())                          # 填报区间购电量，kWh
    r_joint = np.maximum((load_t - pv_t) * DT_H + u - v - x, 0.0)   # 完全预见下应恒为 0

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
    print("【费用对照（填报区间 334 天；4-2 与 4-3 相对下界的差距）】")
    print("全年（1.1–12.31）联合最优总费用 = %.4f 元；日均 = %.4f 元"
          % (cost_all, cost_all / D))
    print("填报区间联合最优费用 = %.4f 元；日均 = %.4f 元；购电量 = %.4f kWh"
          % (cost_rep, cost_rep / N_REP, x_rep))
    gap42 = ANCHOR["roll_42"] - cost_rep                  # 滚动缴费 − 联合下界（应 ≥ 0）
    print("4-2 逐日滚动实际缴费 = %.4f 元；滚动 − 联合下界 = %+.4f 元（%+.4f%%）"
          % (ANCHOR["roll_42"], gap42, 100.0 * gap42 / cost_rep))
    print("   全期真实成本口径（缴费 − 期末残值 %.4f）= %.4f 元；相对下界 %+.4f 元（%+.4f%%）"
          % (ANCHOR["resid_42"], ANCHOR["roll_42"] - ANCHOR["resid_42"],
             ANCHOR["roll_42"] - ANCHOR["resid_42"] - cost_rep,
             100.0 * (ANCHOR["roll_42"] - ANCHOR["resid_42"] - cost_rep) / cost_rep))
    gap43 = ANCHOR["roll_43"] - cost_rep                  # 4-3 总费用（含调整与紧急）− 联合下界
    print("4-3 多阶段滚动总费用 J = %.4f 元；J − 联合下界 = %+.4f 元（%+.4f%%）"
          % (ANCHOR["roll_43"], gap43, 100.0 * gap43 / cost_rep))
    print("⇒ 该差距即跨天协调（跨天套利）可再挖掘的全部空间；D-09 论证的实测证据。")

    # ---- 落盘：汇总 + 联合解逐日明细 + 预览图 ----
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
        ["填报区间日均费用", round(cost_rep / N_REP, ND), "元", "334 天均值"],
        ["填报区间购电量", round(x_rep, ND), "kWh", "联合解在区间内的 Σx"],
        ["4-2 逐日滚动实际缴费（主模型）", round(ANCHOR["roll_42"], ND), "元", "来自 run_q4.py"],
        ["4-2 滚动 − 联合下界", round(gap42, ND), "元", "跨天协调可再挖掘的全部空间"],
        ["4-2 相对差", round(100.0 * gap42 / cost_rep, ND), "%", "相对下界"],
        ["4-2 全期真实成本（缴费 − 期末残值）", round(ANCHOR["roll_42"] - ANCHOR["resid_42"], ND),
         "元", "期末残值仅计一次（D-10）"],
        ["4-3 多阶段滚动总费用 J（主模型）", round(ANCHOR["roll_43"], ND), "元", "来自 run_q4.py"],
        ["4-3 J − 联合下界", round(gap43, ND), "元", "含预报误差+调整机制的代价"],
        ["4-3 相对差", round(100.0 * gap43 / cost_rep, ND), "%", "相对下界"],
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

    # 预览图：与 4-2 逐日滚动解对照（若明细 CSV 存在）
    if os.path.exists(CSV_42):
        days_r, x_roll_day, cost_roll_day = read_roll_daily()
        x_joint_day = np.array([float(x[d * K:(d + 1) * K].sum()) for d in range(D)])[D_REP_FIRST:]
        cost_joint_day = np.array([float(np.dot(price_t[d * K:(d + 1) * K], x[d * K:(d + 1) * K]))
                                   for d in range(D)])[D_REP_FIRST:]
        print("已落盘：%s" % draw_preview(days_r, x_joint_day, cost_joint_day,
                                        x_roll_day, cost_roll_day))
    else:
        print("未找到 %s，跳过预览图（先运行 run_q4.py）" % CSV_42)
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
