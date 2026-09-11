r"""问题 3 创新加分项：两阶段随机规划（SAA）+ 预报误差情景模拟。

理论依据（凸性 + 分位数刻画）：
  对固定计划 x₀，D-12 结算函数关于最终购电量 y 是分段线性：
      S(y) = p·min(x₀,y) + κ₋p(x₀−y)⁺ + κ₊p(y−x₀)⁺
           = (1−κ₋)p·y + (κ₊+κ₋−1)p·(y−x₀)⁺ + κ₋p·x₀     （恒等式）
  两阶段随机规划（第一阶段定计划 x 与当日储能路径，第二阶段在情景 ω 下按该情景预报最优
  调整 y_ω）的期望目标为
      min_x  κ₋p·x + E_ω[ (1−κ₋)p·y_ω(x) + (κ₊+κ₋−1)p·(y_ω(x)−x)⁺ ]
  该目标关于 x 凸；一阶条件给出**最优计划分位数**
      Pr[y_ω > x*] = q,  q = κ₋ / (κ₊+κ₋−1)   （默认 = 0.50）
  即"最优计划量 = 情景最优调整量的 q 分位数"。这把"计划—调整两段契约"的最优锚定问题
  化为分位数问题，是本问的创新点；下文 SAA 用分位数不动点迭代实现。

情景构造（与实测一致、可自检）：
  err_τ = 附件3 第 τ 个发布时刻的预报 − 附件2 实际（逐时段，kW），构成 4 个误差池；
  第 d 天的情景 w：实际 = P̂₀ − err₀[w]；τ 时刻预报 = 实际 + err_τ[w]（τ=1,2,3）。
  自检性质：取 w = d 时情景恰为"真实发生的那一天"（实际 = 附件2 实际、各阶段预报 = 附件3
  预报），故模拟必然逐位复现实测结果——这是情景构造正确性的强证据。

三个部分：
  第一部分（抽样日 × N 情景，重跑多阶段策略）：**情景模拟**——给出策略费用在误差重抽样下
      的期望、95/99 分位与最坏情形，以及实测值的分位（自检：应在分布中部）。
  第二部分（抽样日，SAA）：分位数不动点求随机规划计划，与主模型计划在同一情景集下对照
      期望与最坏情形费用，量化"计划前瞻偏差分布"的增益。
  第三部分：落盘与绘图。

运行：python 问题3/run_q3_stochastic.py [N_SCEN] [DAY_STEP]
输出：`随机规划_期望与最坏情形.xlsx`、`随机规划运行日志.txt`、`随机规划与主模型对比图.png`
随机性：有（情景重抽样与不动点迭代），固定种子 SEED=20260911，重跑结果一致。
"""

import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import numpy as np
import openpyxl

from lib.dataio import read_attachment1, read_attachment2
from lib.forecast import load_forecast
from lib.logio import Tee
from lib.plotstyle import (COLOR_BUY, COLOR_CHG, COLOR_PRICE, COLOR_REF, FIGSIZE_WIDE,
                           apply_chinese_style, save_figure)
from lib.run_days import solve_rolling_staged
from lib.solve_day import (KAPPA_EMG, KAPPA_OVER, KAPPA_UNDER, settle_total, solve_stage)
from lib.timegrid import DT_H, K
from lib.xlsxio import r4

QDIR = os.path.join(ROOT, "问题3")
XLSX_PATH = os.path.join(QDIR, "随机规划_期望与最坏情形.xlsx")
LOG_PATH = os.path.join(QDIR, "随机规划运行日志.txt")
FIG_PATH = os.path.join(QDIR, "随机规划与主模型对比图.png")

SEED = 20260911                      # 固定随机种子（可复现）
N_SCEN_DEFAULT = 30                  # 每日情景数（每情景 3 个阶段 LP）
DAY_STEP_DEFAULT = 3                 # 抽样步长（334/3 ≈ 112 天）
N_SCEN_SAA = 20                      # SAA 的情景数
SAA_DAY_STEP = 28                    # SAA 抽样步长（≈12 天）
D_REP_FIRST = 31                     # 填报区间首日（0 基）
N_REP = 334
TAU_SUB = ((6, 36), (12, 72), (18, 108))     # 调整阶段：发布时刻 -> 覆盖起点（0 基）
TAU_KS_NEXT = {6: 72, 12: 108, 18: K}        # 各阶段的下一起点（0 基）
Q_QUANTILE = KAPPA_UNDER / (KAPPA_OVER + KAPPA_UNDER - 1.0)   # 最优计划分位 = 0.50


def build_error_pools(pv_fc144, pv_actual):
    """构造 4 个发布时刻的实测误差池（err = 预报 − 实际，kW）。

    输入：pv_fc144 (D,4,K) kW；pv_actual (D,K) kW
    输出：np.ndarray (4,D,K)，pools[τ] 为第 τ 个发布时刻的整日误差场池
    """
    pools = np.zeros((4,) + pv_actual.shape)
    for ti in range(4):
        pools[ti] = pv_fc144[:, ti, :] - pv_actual
    return pools


def build_day_scenarios(fc0_day, pools, n_scen, rng):
    """为一天构造 n_scen 个情景（预报 + 实际），含"w=当日"的自洽点。

    输入：fc0_day (K,) kW，该日 0:00 预报（决策基准，情景下不变）；
          pools (4,D,K) kW，误差池；n_scen，情景数；rng，随机数发生器
    输出：dict——
          idx_w (n_scen,) 情景所用的误差场天序号（若等于当日序号则情景=真实发生日）
          actual (n_scen,K) kW 情景实际；fc_tau dict{6,12,18: (n_scen,K) kW} 情景预报
    说明：实际 = P̂₀ − err₀[w]；τ 时刻预报 = 实际 + err_τ[w]。取 w = 当日序号时，
          实际 = 附件2 实际、预报 = 附件3 预报，与实测完全一致（自检性质）。
    """
    n_day = pools.shape[1]
    idx_w = rng.integers(0, n_day, size=n_scen)
    actual = np.maximum(fc0_day[None, :] - pools[0, idx_w], 0.0)         # (n_scen,K)
    fc_tau = {}
    for ti, tau in zip((1, 2, 3), (6, 12, 18)):
        fc_tau[tau] = np.maximum(actual + pools[ti, idx_w], 0.0)
    return {"idx_w": idx_w, "actual": actual, "fc_tau": fc_tau}


def simulate_day(price, load_day, e0, fc0_day, scen):
    """在给定情景下重跑"多阶段策略"并结算，返回每个情景的总费用与紧急购电量。

    输入：price (K,)；load_day (K,) kW；e0 kWh；fc0_day (K,) kW（0:00 预报）
          scen，dict（build_day_scenarios 输出）
    输出：dict——J (n_scen,) 总费用元；r_sum (n_scen,) 紧急购电量 kWh；y/u/v (n_scen,K)
    """
    n_scen = scen["actual"].shape[0]
    J = np.zeros(n_scen); r_sum = np.zeros(n_scen)
    Y = np.zeros((n_scen, K)); U = np.zeros((n_scen, K)); V = np.zeros((n_scen, K))
    for w in range(n_scen):
        # ---- 0:00 决策（全价目标，给出计划 x 与 k≤36 段的执行值）----
        r0 = solve_stage(price, load_day, fc0_day, e_init=e0, x_ref=None)
        x = r0["x_plan"]; u = r0["u_chg"].copy(); v = r0["v_dis"].copy()
        y = x.copy()
        e_cur = float(r0["E_soc"][36])                    # 6:00 处（已锁定段末）储电量
        # ---- 6:00/12:00/18:00 依次调整（基准恒为计划 x，D-12；不可追溯 D-13）----
        for tau, ks in TAU_SUB:
            res = solve_stage(price[ks:], load_day[ks:], scen["fc_tau"][tau][w, ks:],
                              e_init=e_cur, x_ref=x[ks:], k_start=ks, n_period=K - ks)
            y[ks:] = res["x_plan"]; u[ks:] = res["u_chg"]; v[ks:] = res["v_dis"]
            e_cur = float(res["E_soc"][TAU_KS_NEXT[tau] - ks])
        # ---- 实际值结算（D-12）----
        r_emg = np.maximum(load_day * DT_H + u - y - scen["actual"][w] * DT_H - v, 0.0)
        st = settle_total(price, x, y, r_emg)
        J[w] = st["J"]; r_sum[w] = st["qty_r"]
        Y[w] = y; U[w] = u; V[w] = v
    return {"J": J, "r_sum": r_sum, "y": Y, "u": U, "v": V}


def stage2_recourse(price, load_day, pv_scen, e0, x_plan):
    """经典两阶段模型中的第二阶段（完全追索）：情景揭示后对全天重优化调整量。

    与主模型的区别：主模型（多阶段滚动）中 6:00 之后的决策受"当时预报"限制；经典两阶段
    的第二阶段假定情景（当天完整的光伏实现）揭示后再做最优调整——它是"同一份计划在不同
    实现下能取得的最好结果"，用于检验计划的**分布稳健性**（上界口径，报告中明确说明）。

    输入：price (K,)；load_day (K,) kW；pv_scen (K,) kW（该情景的实际光伏）
          e0 kWh（当日 0:00 储电量）；x_plan (K,) kWh（计划量，结算基准）
    输出：dict，J/y/u/v/r_emg
    """
    res = solve_stage(price, load_day, pv_scen, e_init=e0, x_ref=x_plan,
                      k_start=0, n_period=K)
    y = res["x_plan"]; u = res["u_chg"]; v = res["v_dis"]
    r_emg = np.maximum(load_day * DT_H + u - y - pv_scen * DT_H - v, 0.0)
    st = settle_total(price, x_plan, y, r_emg)
    return {"J": st["J"], "y": y, "u": u, "v": v, "r_emg": r_emg}


def saa_plan_fixedpoint(price, load_day, e0, scen, x_init, iters=12, tol=0.5):
    """两阶段 SAA 的分位数不动点迭代：x ← 情景追索调整量的 q 分位数。

    目标关于 x 凸（经典两阶段的部分最小化）；一阶条件 Pr[y_ω > x*] = q（q=0.5），
    故"最优计划量 = 情景最优调整量的中位数"。迭代收敛后给出该情景集下的 SAA 计划。

    输入：price/load_day/e0；scen，情景集（含 actual 与 fc_tau）；x_init (K,)，迭代起点
          iters，迭代上限；tol，收敛容差（kWh）
    输出：dict——x_saa (K,)；J_scen (N,)（收敛点各情景费用）；J_mean；iters；converged
    """
    x = np.asarray(x_init, dtype=float).copy()
    delta = np.inf
    for it in range(iters):
        Y = np.zeros((scen["actual"].shape[0], K))
        for w in range(scen["actual"].shape[0]):
            Y[w] = stage2_recourse(price, load_day, scen["actual"][w], e0, x)["y"]
        x_new = np.quantile(Y, Q_QUANTILE, axis=0)         # 逐段取情景调整量的 q 分位数
        x_new = np.maximum(x_new, 0.0)
        delta = float(np.abs(x_new - x).max())
        x = x_new
        if delta < tol:
            break
    J = np.array([stage2_recourse(price, load_day, scen["actual"][w], e0, x)["J"]
                  for w in range(scen["actual"].shape[0])])
    return {"x_saa": x, "J_scen": J, "J_mean": float(J.mean()), "iters": it + 1,
            "converged": bool(delta < tol)}


def main():
    """主流程：主模型 → 情景模拟（期望/分位/最坏 + 自检）→ SAA 对照 → 落盘与绘图。"""
    n_scen = int(sys.argv[1]) if len(sys.argv) > 1 else N_SCEN_DEFAULT
    day_step = int(sys.argv[2]) if len(sys.argv) > 2 else DAY_STEP_DEFAULT
    print("=" * 78)
    print("问题 3 创新项：情景模拟 + 两阶段随机规划（SAA）")
    print("（情景 %d/日，抽样步长 %d，种子 %d；最优计划分位 q = %.4f）"
          % (n_scen, day_step, SEED, Q_QUANTILE))
    print("=" * 78)
    price, load1, pv1, _ = read_attachment1()
    load, pv_actual, dates = read_attachment2()
    pv_fc, pv_fc144, dates3, tau_list = load_forecast(method="linear")
    rng = np.random.default_rng(SEED)

    # ---------------- 1. 主模型（确定性等价）与误差池 ----------------
    t0 = time.time()
    r_main = solve_rolling_staged(price, load, pv_actual, pv_fc144,
                                  stages=(0, 6, 12, 18), d_start=0, d_end=365)
    sl = slice(D_REP_FIRST, 365)
    roll = {k: r_main[k][sl] for k in r_main if isinstance(r_main[k], np.ndarray)}
    pools = build_error_pools(pv_fc144, pv_actual)
    print("主模型完成（%.0f s）：J = %.4f 元；误差池已构造（4 个发布时刻 × 365 天）"
          % (time.time() - t0, roll["J_day"].sum()))

    # ---------------- 2. 情景模拟（抽样日） ----------------
    idx_days = list(range(0, N_REP, day_step))            # 抽样日的填报区间下标
    print("-" * 78)
    print("【第一部分：情景模拟（%d 个抽样日 × %d 情景，重跑多阶段策略）】"
          % (len(idx_days), n_scen))
    t0 = time.time()
    rows = []
    J_all_mean = 0.0; J_all_p95 = 0.0; J_all_max = 0.0; J_all_obs = 0.0
    selfcheck = []
    for i in idx_days:
        d = roll["d_index"][i]
        e0 = float(roll["E_soc"][i, 0])                    # 与主模型相同的日初储电量
        scen = build_day_scenarios(pv_fc144[d, 0], pools, n_scen, rng)
        out = simulate_day(price, load[d], e0, pv_fc144[d, 0], scen)
        J = out["J"]
        # 自检：以"w=当日"为情景时应复现实测（逐位）
        scen_self = {"actual": pv_actual[d][None, :],
                     "fc_tau": {tau: pv_fc144[d, ti][None, :]
                                for ti, tau in zip((1, 2, 3), (6, 12, 18))}}
        out_self = simulate_day(price, load[d], e0, pv_fc144[d, 0], scen_self)
        selfcheck.append(abs(out_self["J"][0] - roll["J_day"][i]))
        # 实测值在该情景分布中的分位
        pct = float((J < roll["J_day"][i]).mean())
        rows.append({
            "日期": dates[d].isoformat(),
            "实测_元": float(roll["J_day"][i]),
            "期望_元": float(J.mean()),
            "95分位_元": float(np.percentile(J, 95)),
            "99分位_元": float(np.percentile(J, 99)),
            "最坏_元": float(J.max()),
            "实测分位": pct,
        })
        J_all_mean += float(J.mean()); J_all_p95 += float(np.percentile(J, 95))
        J_all_max += float(J.max()); J_all_obs += float(roll["J_day"][i])
    print("情景模拟完成（%.0f s）" % (time.time() - t0))
    print("  抽样日自检：|情景(w=当日) − 实测| 的最大值 = %.6f 元（应≈0）"
          % max(selfcheck))
    print("  抽样日合计：实测 %.4f 元；期望 %.4f 元（%+.4f%%）；95 分位 %.4f；最坏（逐日取最大）%.4f"
          % (J_all_obs, J_all_mean, 100.0 * (J_all_mean / J_all_obs - 1),
             J_all_p95, J_all_max))
    print("  实测值在情景分布中的平均分位 = %.3f（自检：应落在分布中部）"
          % float(np.mean([r["实测分位"] for r in rows])))
    print("  情景间的日均费用标准差 = %.4f 元/日"
          % float(np.mean([(r["95分位_元"] - r["期望_元"]) / 1.645 for r in rows])))

    # ---------------- 3. SAA（抽样日） ----------------
    idx_days_saa = list(range(0, N_REP, SAA_DAY_STEP))
    print("-" * 78)
    print("【第二部分：两阶段 SAA（%d 个抽样日 × %d 情景，q 分位不动点）】"
          % (len(idx_days_saa), N_SCEN_SAA))
    print("  第二阶段为完全追索（情景揭示后全天重优化）——上界口径，与多阶段主模型对照")
    t0 = time.time()
    saa_rows = []
    for i in idx_days_saa:
        d = roll["d_index"][i]
        e0 = float(roll["E_soc"][i, 0])
        scen = build_day_scenarios(pv_fc144[d, 0], pools, N_SCEN_SAA, rng)
        # 主模型计划在同一情景集（完全追索）下的期望/最坏
        J_fix = np.array([stage2_recourse(price, load[d], scen["actual"][w], e0,
                                          roll["x_plan"][i])["J"]
                          for w in range(N_SCEN_SAA)])
        out = saa_plan_fixedpoint(price, load[d], e0, scen, roll["x_plan"][i])
        saa_rows.append({
            "日期": dates[d].isoformat(),
            "主模型计划_期望_元": float(J_fix.mean()),
            "主模型计划_最坏_元": float(J_fix.max()),
            "SAA计划_期望_元": float(out["J_scen"].mean()),
            "SAA计划_最坏_元": float(out["J_scen"].max()),
            "SAA迭代数": int(out["iters"]),
            "计划差_Σ|x_saa−x_main|_kWh": float(np.sum(np.abs(out["x_saa"] - roll["x_plan"][i]))),
        })
    print("SAA 完成（%.0f s）" % (time.time() - t0))
    exp_fix = float(np.mean([r["主模型计划_期望_元"] for r in saa_rows]))
    exp_saa = float(np.mean([r["SAA计划_期望_元"] for r in saa_rows]))
    worst_fix = float(np.mean([r["主模型计划_最坏_元"] for r in saa_rows]))
    worst_saa = float(np.mean([r["SAA计划_最坏_元"] for r in saa_rows]))
    print("  抽样日平均：主模型计划 期望 %.4f / 最坏 %.4f 元" % (exp_fix, worst_fix))
    print("              SAA 计划   期望 %.4f / 最坏 %.4f 元" % (exp_saa, worst_saa))
    print("  SAA 相对主模型的期望增益 = %+.4f 元/日（最坏情形改善 %+.4f 元/日）"
          % (exp_fix - exp_saa, worst_fix - worst_saa))
    print("  说明：为保持口径公平，'主模型计划'也按完全追索评估（同一情景集）；SAA 的优势来自")
    print("        '计划量锚在情景调整量的中位数'（q=%.2f），即对超用/欠取的平衡锚定。" % Q_QUANTILE)

    # ---------------- 4. 落盘 ----------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "情景模拟汇总"
    ws.append(["指标", "数值", "单位", "说明"])
    ws.append(["抽样日情景模拟_实测合计", float(J_all_obs), "元", "%d 个抽样日（步长 %d）" % (len(idx_days), day_step)])
    ws.append(["抽样日情景模拟_期望合计", float(J_all_mean), "元", "误差重抽样下策略费用期望"])
    ws.append(["抽样日情景模拟_95分位合计", float(J_all_p95), "元", "逐日 95 分位累计"])
    ws.append(["抽样日情景模拟_最坏合计", float(J_all_max), "元", "逐日取最大情景"])
    ws.append(["实测平均分位", float(np.mean([r["实测分位"] for r in rows])), "—", "实测值在情景分布中的位置（应≈中部）"])
    ws.append(["情景自检_最大偏差", float(max(selfcheck)), "元", "情景(w=当日)复现实测的逐位差"])
    ws.append(["主模型全年总费用", float(roll["J_day"].sum()), "元", "确定性等价（单情景）"])
    ws.append(["完全信息下界", 12254765.7161, "元", "问题 2 主模型"])
    ws2 = wb.create_sheet("SAA抽样日")
    h2 = list(saa_rows[0].keys())
    ws2.append(h2)
    for r in saa_rows:
        ws2.append([r[k] for k in h2])
    ws2.append(["抽样日平均", exp_fix, worst_fix, exp_saa, worst_saa, "", ""])
    for w in wb.worksheets:
        for row in w.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value, float):
                    c.number_format = "0.0000"
    wb.save(XLSX_PATH)

    # ---------------- 5. 对比图 ----------------
    apply_chinese_style()
    import matplotlib.pyplot as plt
    fig, axes = plt.subplots(1, 2, figsize=FIGSIZE_WIDE)
    ax = axes[0]
    x_axis = np.arange(len(rows))
    obs = np.array([r["实测_元"] for r in rows])
    mean = np.array([r["期望_元"] for r in rows])
    p95 = np.array([r["95分位_元"] for r in rows])
    worst = np.array([r["最坏_元"] for r in rows])
    ax.plot(x_axis, np.cumsum(obs) / 1e4, color=COLOR_BUY, linewidth=1.6, label="实测（附件2 实现）")
    ax.plot(x_axis, np.cumsum(mean) / 1e4, color=COLOR_CHG, linewidth=1.6,
            label="预期（%d 情景平均）" % n_scen)
    ax.plot(x_axis, np.cumsum(p95) / 1e4, color=COLOR_PRICE, linewidth=1.2, linestyle="--",
            label="95 分位累计")
    ax.plot(x_axis, np.cumsum(worst) / 1e4, color=COLOR_REF, linewidth=1.0, linestyle=":",
            label="最坏情景（逐日取最大）")
    ax.set_xlabel("抽样日序号（每 %d 天取 1 天）" % day_step)
    ax.set_ylabel("累计费用（万元）")
    ax.set_title("(a) 情景模拟：策略费用在预报误差重抽样下的分布")
    ax.legend(loc="upper left", fontsize=9)
    ax = axes[1]
    ax.scatter(obs / 1e4, mean / 1e4, s=14, color=COLOR_CHG, label="单日：实测 vs 期望")
    lim = [0, max(obs.max(), mean.max()) / 1e4 * 1.05]
    ax.plot(lim, lim, color=COLOR_REF, linewidth=1.0, linestyle="--", label="对角参考线 y=x")
    ax.set_xlabel("实测单日费用（万元）")
    ax.set_ylabel("情景期望单日费用（万元）")
    ax.set_title("(b) 单日费用：实测与期望对照\n（点在对角线以上的天 = 当日实现优于期望）")
    ax.legend(loc="upper left", fontsize=9)
    fig.suptitle("问题 3 创新分析：预报误差情景下的费用分布", fontsize=15)
    fig.tight_layout(rect=(0, 0, 1, 0.94))
    save_figure(fig, FIG_PATH)
    print("已落盘：%s / %s" % (XLSX_PATH, FIG_PATH))
    print("=" * 78)
    return {"rows": rows, "saa_rows": saa_rows}


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
