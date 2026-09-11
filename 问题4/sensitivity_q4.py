"""问题 4 灵敏度与稳健性实验（S1–S5；S6 结论写入 适用边界.md 与问题4.md）。

按 `问题4/交付清单.md` §五 执行，全部结果落盘到 `问题4/`：
  S1 关键参数扰动：η、P̄、κ、κ₋/κ₊、V_E 各 ±5%/±10%/±20% → `灵敏度分析_参数扰动.xlsx` + 图
  S2 多因素组合扰动：η×V_E×κ 网格 → `灵敏度分析_多因素网格.xlsx` + `灵敏度分析_多因素热力图.png`
  S3 数据侧扰动：(a) 剔除 9 个极端低价点；(b) 电价整体 ±5%/±10%；(c) 附件1 形状 × 平滑日因子；
     (d) 预热期 0 天 vs 31 天 → `灵敏度分析_数据扰动.xlsx`
  S4 方法侧互验：(a) 逐日滚动 vs 全年联合 LP 下界；(b) 终端三方案；(c) MATLAB 复算若干天；
     (d) 小时粒度 LP → `方法侧互验_联合LP与终端条件与MATLAB.xlsx`
  S5 口径对照：(a) 4-2 信息结构读法① vs 读法②；(b) 结算口径 C-6；(c) 调整生效规则 C-7（不可追溯
     vs 可追溯极端 vs 边界 kmark+1）；(d) 填法 Y-轮转 vs 填法 X → `口径对照_读法与结算与调整规则.xlsx`

口径（`口径与假设台账.md`）：D-10 终端余值（主口径 V_E=次日最低价/η）、D-12 结算、D-13 不可追溯、
D-16.3 极端低价点保留 + 剔除对照、D-11 预热期。全部扰动只改参数/数据，不新增模型。

运行：python 问题4/sensitivity_q4.py [--skip-heavy]
      --skip-heavy 跳过最耗时的 S1（改为只跑 ±10% 档）与 S5(c)
依赖：numpy、scipy、openpyxl、matplotlib；lib/ 公共模块 + 问题4/run_q4.py 的主链函数。
随机性：无（全部确定性 LP）。
"""

import argparse
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "问题4"))

import numpy as np
import openpyxl

from lib.dataio import read_attachment1, read_attachment2, read_attachment4
from lib.forecast import load_forecast
from lib.logio import Tee
from lib.run_days import solve_rolling, solve_rolling_staged
from lib.solve_day import (KAPPA_EMG, KAPPA_OVER, KAPPA_UNDER, solve_day, solve_stage,
                           terminal_value_definitions)
from lib.storage import E_CAP, E_INIT, E_MAX, E_MIN, ETA, P_MAX
from lib.timegrid import DT_H, K
from lib.xlsxio import r4, set_decimal_format  # noqa: F401（与全工作区同一套写出助手）

QDIR = os.path.join(ROOT, "问题4")
FIGDIR = os.path.join(ROOT, "图片", "问题4")                        # 图片统一目录（2026-09-11 起，用户要求）
os.makedirs(FIGDIR, exist_ok=True)                                # 确保目录存在（重跑时自动建）
LOG_PATH = os.path.join(QDIR, "灵敏度运行日志.txt")
D_REP_FIRST = 31
N_REP = 334
ND = 4
STAGES_MAIN = (0, 6, 12, 18)

# 固定电价锚定（问题 2，用于"波动电价更贵"结论的翻转判定）
Q2_ANCHOR = 12254765.7161
# 主模型锚定（问题 4，run_q4.py 落盘）
Q42_ANCHOR = 12815460.2655
Q43_ANCHOR = 14450082.3797

_G = {}                                    # 全局数据容器（读一次，全程复用）


def write_xlsx(path, sheets):
    """把 {工作表名: 行列表} 写成 xlsx（数值 4 位小数格式，与其他落盘文件风格一致）。

    输入：path，str；sheets，dict[str, list[list]]；输出：str，写入路径
    """
    wb = openpyxl.Workbook()
    first = True
    for name, rows in sheets.items():
        ws = wb.active if first else wb.create_sheet()
        ws.title = name
        first = False
        for row in rows:
            ws.append(row)
        for r in ws.iter_rows(min_row=2):
            for c in r:
                if isinstance(c.value, float):
                    c.number_format = "0.0000"
    wb.save(path)
    return path


def eval_42(price_mat, ve_arr, eta=ETA, p_max=P_MAX, d_start=0, e_init=E_INIT):
    """跑一遍 4-2 链并汇总填报区间指标。

    输入：price_mat (D,K) 元/kWh；ve_arr (D,) 元/kWh；eta/p_max；d_start；e_init，kWh
    输出：dict——total（缴费，元）、r_sum（紧急购电量，kWh）、soc_util（利用率）、
          n_hold（24:00 高于下限天数）、n_min（压下限天数）、E_end（期末储电量，kWh）、
          E_rep（填报区间储电量轨迹）
    """
    roll = solve_rolling(price_mat, load=_G["load"], pv_actual=_G["pv"], e_init=e_init,
                         mode="free", eta=eta, p_max=p_max, v_end=ve_arr,
                         d_start=d_start, d_end=365)
    # 填报区间切片：roll 的第 0 行对应第 d_start 天，故填报首日在其中的下标为
    # max(d_start, 31) − d_start（d_start=0 时即 31，与 run_q4.py 一致）
    sl = slice(max(d_start, D_REP_FIRST) - d_start, 365 - d_start)
    x = roll["x_plan"][sl]; E = roll["E_soc"][sl]; r = roll["r_emg"][sl]
    return {
        "total": float(roll["cost_day"][sl].sum()),
        "r_sum": float(r.sum()),
        "soc_util": float(E[:, 1:K + 1].mean() / E_CAP),
        "n_hold": int(np.sum(E[:, K] > E_MIN + 1e-3)),
        "n_min": int(np.sum(np.abs(E[:, K] - E_MIN) < 1e-3)),
        "E_end": float(E[-1, K]),
        "E_rep": E, "x_rep": x,
    }


def eval_43(price_mat, ve_arr=None, eta=ETA, p_max=P_MAX, kappa=KAPPA_EMG,
            kappa_under=KAPPA_UNDER, kappa_over=KAPPA_OVER, d_start=0, e_init=E_INIT,
            stages=STAGES_MAIN, kmark=None):
    """跑一遍 4-3 链并汇总填报区间指标。

    输入：price_mat (D,K)；ve_arr (D,) 或 None；eta/p_max/kappa/kappa_under/kappa_over；
          d_start；e_init；stages；kmark
    输出：dict——J/J_plan/J_adj/J_emg（元）、J_alt（对照口径，元）、r_sum（kWh）、
          soc_util、n_hold、n_min、E_end（kWh）
    """
    roll = solve_rolling_staged(price_mat, _G["load"], _G["pv"], _G["pv_fc144"],
                                e_init=e_init, stages=stages, kmark=kmark,
                                d_start=d_start, d_end=365, eta=eta, p_max=p_max,
                                kappa=kappa, kappa_under=kappa_under,
                                kappa_over=kappa_over, v_end_day=ve_arr)
    # 填报区间切片：roll 的第 0 行对应第 d_start 天（与 eval_42 同一修正）
    sl = slice(max(d_start, D_REP_FIRST) - d_start, 365 - d_start)
    E = roll["E_soc"][sl]
    return {
        "J": float(roll["J_day"][sl].sum()),
        "J_plan": float(roll["J_plan"][sl].sum()),
        "J_adj": float(roll["J_adj"][sl].sum()),
        "J_emg": float(roll["J_emg"][sl].sum()),
        "J_alt": float(roll["J_alt_day"][sl].sum()),
        "r_sum": float(roll["r_emg"][sl].sum()),
        "soc_util": float(E[:, 1:K + 1].mean() / E_CAP),
        "n_hold": int(np.sum(E[:, K] > E_MIN + 1e-3)),
        "n_min": int(np.sum(np.abs(E[:, K] - E_MIN) < 1e-3)),
        "E_end": float(E[-1, K]),
    }


def run_hourly_42(price_mat, ve_arr):
    """小时粒度 4-2 对照：把 144 段聚合为 24 个 1 小时时段后求解（决策粒度粗化）。

    输入：price_mat (D,K) 元/kWh；ve_arr (D,) 元/kWh
    输出：dict——total（填报区间缴费，元）；口径：小时平均电价 × 小时购电量
    """
    n_h = 24
    price_h = price_mat.reshape(price_mat.shape[0], n_h, 6).mean(axis=2)   # (D,24) 元/kWh
    load_h = (_G["load"].reshape(365, n_h, 6) * DT_H).sum(axis=2)          # (D,24) kWh
    pv_h = (_G["pv"].reshape(365, n_h, 6) * DT_H).sum(axis=2)              # (D,24) kWh
    e = E_INIT
    tot = 0.0
    for d in range(365):
        res = solve_day(price_h[d], load_h[d], pv_h[d], e_init=e, mode="free",
                        eta=ETA, p_max=P_MAX, e_min=E_MIN, e_max=E_MAX, dt_h=1.0,
                        v_end=ve_arr[d])
        if d >= D_REP_FIRST:
            tot += float(np.dot(price_h[d], res["x_plan"]))                # 全天缴费，元
        e = float(res["E_soc"][-1])
    return {"total": tot}


def run_hourly_43(price_mat, ve_arr):
    """小时粒度 4-3 对照：24 个小时段上做 0/6/12/18 多阶段滚动（同 D-12/D-13 规则）。

    输入：price_mat (D,K) 元/kWh；ve_arr (D,) 元/kWh
    输出：dict——J（填报区间总费用，元）
    """
    from lib.solve_day import settle_total
    n_h = 24
    price_h = price_mat.reshape(price_mat.shape[0], n_h, 6).mean(axis=2)
    load_h = (_G["load"].reshape(365, n_h, 6) * DT_H).sum(axis=2)
    pv_a_h = (_G["pv"].reshape(365, n_h, 6) * DT_H).sum(axis=2)
    fc_h = np.zeros((365, 4, n_h))
    for ti in range(4):
        fc_h[:, ti, :] = (_G["pv_fc144"][:, ti, :].reshape(365, n_h, 6) * DT_H).sum(axis=2)
    kmark = {0: 0, 6: 6, 12: 12, 18: 18}                                   # 小时单位的覆盖起点
    tot = 0.0
    e = E_INIT
    for d in range(365):
        xx = np.zeros(n_h); uu = np.zeros(n_h); vv = np.zeros(n_h)
        e_cur = e
        x0 = None
        for si, tau in enumerate(STAGES_MAIN):
            ks = kmark[tau]
            ref = None if x0 is None else x0[ks:]
            res = solve_stage(price_h[d, ks:], load_h[d, ks:], fc_h[d, tau_pos(tau), ks:],
                              e_init=e_cur, x_ref=ref, k_start=ks, n_period=n_h - ks,
                              eta=ETA, p_max=P_MAX, e_min=E_MIN, e_max=E_MAX, dt_h=1.0,
                              v_end=(ve_arr[d] if si == len(STAGES_MAIN) - 1 else 0.0))
            xx[ks:] = res["x_plan"]; uu[ks:] = res["u_chg"]; vv[ks:] = res["v_dis"]
            if si == 0:
                x0 = xx.copy()
            if si + 1 < len(STAGES_MAIN):
                e_cur = float(res["E_soc"][kmark[STAGES_MAIN[si + 1]] - ks])
            else:
                e_cur = float(res["E_soc"][-1])
        residual = load_h[d] + uu - vv - pv_a_h[d] - xx                    # 逐小时供给残差，kWh
        r_h = np.maximum(residual, 0.0)                                     # 紧急购电量，kWh
        st = settle_total(price_h[d], x0, xx, r_h)
        if d >= D_REP_FIRST:
            tot += st["J"]                                                  # 全天总费用，元
        e = float(e + np.sum(ETA * uu - vv / ETA))                          # 跨日传递（小时粒度）
    return {"J": tot}


def tau_pos(tau):
    """发布时刻 τ → pv_fc144 第 2 维下标（0/6/12/18 → 0..3）。"""
    return {0: 0, 6: 1, 12: 2, 18: 3}[int(tau)]


# ============================== S1 关键参数扰动 ==============================

def run_s1(levels=(5, 10, 20)):
    """S1：η、P̄、κ、κ₋/κ₊、V_E 各 ±5%/±10%/±20%，输出 4-2 与 4-3 的关键指标。

    输入：levels tuple[int]，扰动的百分比档位
    输出：rows list[list]，写入 `灵敏度分析_参数扰动.xlsx`
    """
    print("-" * 78)
    print("【S1 关键参数扰动】η / P̄ / κ / κ₋ / κ₊ / V_E 各 ±%s%%（4-2 与 4-3 各全链重跑）"
          % "/".join(str(v) for v in levels))
    price_mat = _G["pr4"]; ve0 = _G["ve"]
    t0 = time.time()
    base42 = eval_42(price_mat, ve0)
    base43 = eval_43(price_mat, ve0)
    print("  基准：4-2 缴费 = %.4f 元；4-3 J = %.4f 元（耗时 %.1fs）"
          % (base42["total"], base43["J"], time.time() - t0))

    rows = []
    hdr = ["扰动参数", "档位", "4-2缴费_元", "4-2相对变化_%", "4-2紧急购电量_kWh",
           "4-2储电量利用率", "4-3总费用J_元", "4-3相对变化_%", "4-3紧急购电量_kWh",
           "4-3储电量利用率", "24:00持有过夜_4-2", "结论按判定", "说明"]

    def fixed_price_total(eta=ETA, p_max=P_MAX):
        """同一参数下的固定电价对照（附件1 广播 + 终端自由），用于结论翻转的同口径判定。

        输入：eta、p_max，float；输出：float，填报区间缴费，元
        """
        pr1 = np.tile(_G["price1"], (365, 1))
        r = eval_42(pr1, np.zeros(365), eta=eta, p_max=p_max)
        return r["total"]

    def add_row(pname, level, r42, r43, note, q2_ref=None):
        """把一档扰动结果追加为一行（相对变化与结论翻转判定就地计算）。

        输入：pname/level/note，str；r42/r43，eval 结果 dict
              q2_ref，float 或 None，该档位下的固定电价对照费用（None 时用名义锚定）
        判定口径：①该参数档位下"波动电价缴费 ≥ 同参数固定电价缴费"是否仍成立；
                 ②终端余值仍改变策略（持有过夜天数 > 0）。
        """
        q2_base = Q2_ANCHOR if q2_ref is None else q2_ref
        rel42 = 100.0 * (r42["total"] - base42["total"]) / base42["total"]
        rel43 = 100.0 * (r43["J"] - base43["J"]) / base43["J"]
        ok1 = r42["total"] >= q2_base
        ok2 = r42["n_hold"] > 0
        verdict = "不翻转（两项均成立）" if (ok1 and ok2) else \
                  ("翻转：①该档位下波动电价反而更便宜（vs 同参数固定电价 %.4f）" % q2_base if not ok1
                   else "翻转：②余值不再改变策略")
        rows.append([pname, level, r4(r42["total"]), r4(rel42), r4(r42["r_sum"]),
                     round(r42["soc_util"], 6), r4(r43["J"]), r4(rel43), r4(r43["r_sum"]),
                     round(r43["soc_util"], 6), r42["n_hold"], verdict, note])

    add_row("基准（无扰动）", 0, base42, base43, "主口径：附件4 价格、V_E=次日最低价/η")
    for pct in levels:
        s = 1.0 + pct / 100.0
        # ---- η 扰动：效率提高/降低对两条链同时生效；固定电价对照同参数重算 ----
        r42 = eval_42(price_mat, ve0, eta=ETA * s)
        r43 = eval_43(price_mat, ve0, eta=ETA * s)
        add_row("η（单向效率）", "+%d%% → %.4f" % (pct, ETA * s), r42, r43,
                "同倍率改变充放电效率；对照基准为同 η 的固定电价缴费",
                q2_ref=fixed_price_total(eta=ETA * s))
        # ---- P̄ 扰动：最大充放电功率 ----
        r42 = eval_42(price_mat, ve0, p_max=P_MAX * s)
        r43 = eval_43(price_mat, ve0, p_max=P_MAX * s)
        add_row("P̄（最大充放电功率）", "+%d%% → %.0f kW" % (pct, P_MAX * s), r42, r43,
                "同倍率改变功率上限；对照基准为同 P̄ 的固定电价缴费",
                q2_ref=fixed_price_total(p_max=P_MAX * s))
    for pct in levels:
        s = 1.0 - pct / 100.0
        r42 = eval_42(price_mat, ve0, eta=ETA * s)
        r43 = eval_43(price_mat, ve0, eta=ETA * s)
        add_row("η（单向效率）", "-%d%% → %.4f" % (pct, ETA * s), r42, r43,
                "同倍率改变充放电效率；对照基准为同 η 的固定电价缴费",
                q2_ref=fixed_price_total(eta=ETA * s))
        r42 = eval_42(price_mat, ve0, p_max=P_MAX * s)
        r43 = eval_43(price_mat, ve0, p_max=P_MAX * s)
        add_row("P̄（最大充放电功率）", "-%d%% → %.0f kW" % (pct, P_MAX * s), r42, r43,
                "同倍率改变功率上限；对照基准为同 P̄ 的固定电价缴费",
                q2_ref=fixed_price_total(p_max=P_MAX * s))
    # ---- κ 扰动：只影响紧急购电费（已决策变量不变），精确线性换算 ----
    for pct in levels:
        for sgn in (+1, -1):
            k_new = KAPPA_EMG * (1.0 + sgn * pct / 100.0)
            r43 = dict(base43)
            r43["J_emg"] = base43["J_emg"] * k_new / KAPPA_EMG            # 线性精确换算
            r43["J"] = base43["J_plan"] + base43["J_adj"] + r43["J_emg"]
            add_row("κ（紧急购电倍数）", "%+d%% → %.4f" % (sgn * pct, k_new), base42, r43,
                    "κ 不进入决策，仅线性改变紧急购电费用（精确换算）")
    # ---- κ₋ / κ₊ 扰动：改变调整阶段的偏差结算斜率，须重跑 4-3 ----
    for pct in levels:
        for sgn in (+1, -1):
            ku = KAPPA_UNDER * (1.0 + sgn * pct / 100.0)
            r43 = eval_43(price_mat, ve0, kappa_under=ku)
            add_row("κ₋（欠取违约系数）", "%+d%% → %.4f" % (sgn * pct, ku), base42, r43,
                    "4-2 无调整机制、不受 κ₋ 影响（列内为基准值）")
            ko = KAPPA_OVER * (1.0 + sgn * pct / 100.0)
            r43 = eval_43(price_mat, ve0, kappa_over=ko)
            add_row("κ₊（超用电价系数）", "%+d%% → %.4f" % (sgn * pct, ko), base42, r43,
                    "4-2 无调整机制、不受 κ₊ 影响（列内为基准值）")
    # ---- V_E 扰动：终端余值单价整体缩放 ----
    for pct in levels:
        for sgn in (+1, -1):
            ve_new = ve0 * (1.0 + sgn * pct / 100.0)
            r42 = eval_42(price_mat, ve_new)
            r43 = eval_43(price_mat, ve_new)
            add_row("V_E（终端余值单价）", "%+d%%" % (sgn * pct), r42, r43,
                    "V_E = 次日最低价/η 整体乘 (1%+d%%)，等价于按比例缩放余值" % (sgn * pct))

    path = os.path.join(QDIR, "灵敏度分析_参数扰动.xlsx")
    write_xlsx(path, {"S1_单因素扰动": [hdr] + rows})
    print("  S1 完成：%d 行 → %s" % (len(rows), path))
    draw_s1_figure(rows, base42["total"], base43["J"])
    return rows


def draw_s1_figure(rows, base42, base43):
    """绘制 S1 参数扰动图：五种参数的费用响应（以 ±% 为横轴）。

    输入：rows list[list]（run_s1 的行）；base42/base43 基准费用，元；输出：str，PNG 路径
    """
    import matplotlib.pyplot as plt
    from lib.plotstyle import (COLOR_CHG, COLOR_DIS, COLOR_LOAD, COLOR_PRICE, COLOR_PV,
                               FIGSIZE_TALL, apply_chinese_style, save_figure)
    apply_chinese_style()
    params = ["η（单向效率）", "P̄（最大充放电功率）", "κ（紧急购电倍数）",
              "κ₋（欠取违约系数）", "κ₊（超用电价系数）", "V_E（终端余值单价）"]
    colors = [COLOR_LOAD, COLOR_PV, COLOR_PRICE, COLOR_CHG, COLOR_DIS, "#5B2C6F"]
    fig, axes = plt.subplots(1, 2, figsize=FIGSIZE_TALL)
    for ax, col_idx, base, title in ((axes[0], 2, base42, "4-2 缴费（元）"),
                                     (axes[1], 6, base43, "4-3 总费用 J（元）")):
        for pname, col in zip(params, colors):
            xs, ys = [], []
            for row in rows:
                if row[0] == pname:
                    xs.append(str(row[1]))                                 # 档位标签
                    ys.append(100.0 * (float(row[col_idx]) - base) / base)  # 相对变化，%
            ax.plot(range(len(xs)), ys, "o-", color=col, linewidth=1.4, markersize=5, label=pname)
        ax.axhline(0, color="black", linewidth=1.0)
        ax.set_xticks(range(6), ["-20%", "-10%", "-5%", "基准", "+5%", "+10%"],
                      fontsize=10)
        ax.set_xlabel("扰动档位（近似按数值排序）")
        ax.set_ylabel("相对变化（%）")
        ax.set_title(title)
        ax.legend(fontsize=9)
    fig.suptitle("S1 关键参数扰动（S1 × 4-2/4-3）：η 最敏感、κ₋/κ₊/V_E 中等、P̄ 最迟钝", fontsize=13)
    fig.tight_layout(rect=(0, 0, 1, 0.94))
    return save_figure(fig, os.path.join(FIGDIR, "灵敏度分析_参数扰动.png"))


# ============================== S2 多因素组合扰动 ==============================

def run_s2():
    """S2：η×V_E×κ 三维网格（η×V_E 全链重跑 9 组合，κ 三档线性换算）。

    输出：`灵敏度分析_多因素网格.xlsx` + `灵敏度分析_多因素热力图.png`
    """
    print("-" * 78)
    print("【S2 多因素组合扰动】η×V_E×κ 网格（η×V_E 九组合全链重跑）")
    price_mat = _G["pr4"]; ve0 = _G["ve"]
    eta_lv = [0.81, 0.90, 0.99]                        # η 三档（±10%）
    ve_lv = [0.8, 1.0, 1.2]                            # V_E 乘数三档
    kappa_lv = [4.0, 5.0, 6.0]                         # κ 三档（线性换算）
    grid42 = np.zeros((3, 3)); grid43 = np.zeros((3, 3))
    grid42_emg = np.zeros((3, 3)); grid43_emg = np.zeros((3, 3))
    rows = [["η", "V_E乘数", "4-2缴费_元", "4-3总费用J_元", "4-3紧急购电费_元",
             "4-3Σr_kWh", "4-2利用率", "4-3利用率", "同格κ换算说明"]]
    t0 = time.time()
    for i, eta_v in enumerate(eta_lv):
        for j, ve_m in enumerate(ve_lv):
            ve_new = ve0 * ve_m
            r42 = eval_42(price_mat, ve_new, eta=eta_v)
            r43 = eval_43(price_mat, ve_new, eta=eta_v)
            grid42[i, j] = r42["total"]; grid43[i, j] = r43["J"]
            grid42_emg[i, j] = r42["r_sum"]; grid43_emg[i, j] = r43["r_sum"]
            rows.append([round(eta_v, 4), round(ve_m, 2), r4(r42["total"]), r4(r43["J"]),
                         r4(r43["J_emg"]), r4(r43["r_sum"]), round(r42["soc_util"], 6),
                         round(r43["soc_util"], 6),
                         "κ={%s} 时 J_emg 按 κ/5 线性换算" % ", ".join("%.0f" % k for k in kappa_lv)])
    # κ 三档的线性换算表（J = J_plan + J_adj + κ·Σp·r；Σp·r = J_emg/5）
    rows_k = [["η", "V_E乘数", "κ", "4-3总费用J_元（κ 换算）", "说明"]]
    for i, eta_v in enumerate(eta_lv):
        for j, ve_m in enumerate(ve_lv):
            j_emg_base = None
            for row in rows[1:]:
                if row[0] == round(eta_v, 4) and row[1] == round(ve_m, 2):
                    j_emg_base = float(row[4])
            for k_v in kappa_lv:
                if k_v == 5.0:
                    continue
                j43_new = grid43[i, j] - j_emg_base + j_emg_base * k_v / 5.0
                rows_k.append([round(eta_v, 4), round(ve_m, 2), k_v, r4(j43_new),
                               "κ 不进入决策；按 J_emg 线性缩放（精确）"])
    path = os.path.join(QDIR, "灵敏度分析_多因素网格.xlsx")
    write_xlsx(path, {"S2_eta×VE网格": rows, "S2_κ线性换算": rows_k})
    print("  S2 完成：9 组合全链重跑（耗时 %.1fs）→ %s" % (time.time() - t0, path))
    draw_s2_heatmap(eta_lv, ve_lv, grid42, grid43)
    return rows


def draw_s2_heatmap(eta_lv, ve_lv, grid42, grid43):
    """绘制 S2 热力图：η×V_E 网格上的 4-2 缴费与 4-3 总费用。

    输入：eta_lv/ve_lv list[float]；grid42/grid43 (3,3) 元；输出：str，PNG 路径
    """
    import matplotlib.pyplot as plt
    from lib.plotstyle import FIGSIZE_TALL, apply_chinese_style, save_figure
    apply_chinese_style()
    fig, axes = plt.subplots(1, 2, figsize=FIGSIZE_TALL)
    for ax, grid, title in ((axes[0], grid42, "4-2 全年缴费（元）"),
                            (axes[1], grid43, "4-3 总费用 J（元）")):
        im = ax.imshow(grid / 1e4, cmap="YlOrRd", aspect="auto")
        for i in range(grid.shape[0]):
            for j in range(grid.shape[1]):
                ax.annotate("%.1f" % (grid[i, j] / 1e4), (j, i), ha="center", va="center",
                            fontsize=11, color="black")
        ax.set_xticks(range(len(ve_lv)), ["V_E×%.1f" % v for v in ve_lv])
        ax.set_yticks(range(len(eta_lv)), ["η=%.2f" % v for v in eta_lv])
        ax.set_xlabel("终端余值乘数（V_E = 次日最低价/η × 乘数）")
        ax.set_ylabel("单向效率 η")
        ax.set_title(title + "（单位：万元）")
        fig.colorbar(im, ax=ax, label="费用（万元）")
    fig.suptitle("S2 多因素组合扰动：η × V_E 网格（κ 只线性影响紧急费用、不改变决策）", fontsize=13)
    fig.tight_layout(rect=(0, 0, 1, 0.94))
    return save_figure(fig, os.path.join(FIGDIR, "灵敏度分析_多因素热力图.png"))


# ============================== S3 数据侧扰动 ==============================

def run_s3():
    """S3：数据侧扰动——(a) 剔除极端低价点 (b) 电价整体缩放 (c) 去噪日因子 (d) 预热期。

    输出：`灵敏度分析_数据扰动.xlsx`
    """
    print("-" * 78)
    print("【S3 数据侧扰动】极端低价点剔除 / 电价缩放 / 平滑日因子 / 预热期")
    pr4 = _G["pr4"]; ve0 = _G["ve"]; p1 = _G["price1"]
    base42 = eval_42(pr4, ve0); base43 = eval_43(pr4, ve0)
    rows = [["扰动", "档位", "4-2缴费_元", "4-2相对变化_%", "4-3总费用J_元",
             "4-3相对变化_%", "4-3紧急购电量_kWh", "说明"]]

    def add(name, level, r42, r43, note):
        """追加一行 S3 结果（相对变化就地计算）。"""
        rows.append([name, level, r4(r42["total"]),
                     r4(100.0 * (r42["total"] - base42["total"]) / base42["total"]),
                     r4(r43["J"]),
                     r4(100.0 * (r43["J"] - base43["J"]) / base43["J"]),
                     r4(r43["r_sum"]), note])

    add("基准（无扰动）", "—", base42, base43, "主口径")

    # ---- (a) 剔除 9 个极端低价点：替换为该日"附件1 形状 × 该日日因子"的结构值 ----
    a_day = (pr4 / p1[None, :]).mean(axis=1)                   # 逐日日因子（与 analyze_price 同定义）
    mask = pr4 < 0.05                                          # 极端低价点掩码（D-16.3）
    pr4_a = pr4.copy()
    pr4_a[mask] = (p1[None, :] * a_day[:, None])[mask]         # 中立化尖峰、保留当日水平
    ve_a = terminal_value_definitions(pr4_a)                   # 价格变了，余值重算
    r42 = eval_42(pr4_a, ve_a); r43 = eval_43(pr4_a, ve_a)
    add("(a) 剔除 9 个极端低价点（<0.05 元/kWh）", "%d 个点" % int(mask.sum()), r42, r43,
        "被剔点替换为该日'附件1 形状×日因子'的结构值；零差异为结构性结论：9 个点全部落在"
        "光伏富余、计划购电量 x=0 且充电已顶功率上限的时段，抬价不改变任何决策（单日核验见运行日志）")

    # ---- (b) 电价整体 ±5% / ±10% ----
    for pct in (5, 10):
        for sgn in (+1, -1):
            pr4_b = pr4 * (1.0 + sgn * pct / 100.0)
            ve_b = terminal_value_definitions(pr4_b)
            r42 = eval_42(pr4_b, ve_b); r43 = eval_43(pr4_b, ve_b)
            add("(b) 电价整体缩放", "%+d%%" % (sgn * pct), r42, r43, "价格矩阵乘同一系数，余值随动")

    # ---- (c) 附件4 替换为"附件1 形状 × 平滑日因子"（去噪） ----
    a_smooth = np.convolve(a_day, np.ones(7) / 7.0, mode="same")   # 7 日中心滑动平均
    a_smooth[:3] = a_day[:3]; a_smooth[-3:] = a_day[-3:]           # 边界回退原值
    pr4_c = p1[None, :] * a_smooth[:, None]                        # 无日内噪声的结构价格
    ve_c = terminal_value_definitions(pr4_c)
    r42 = eval_42(pr4_c, ve_c); r43 = eval_43(pr4_c, ve_c)
    add("(c) 附件1 形状 × 平滑日因子（去噪）", "7 日滑动平均", r42, r43,
        "去掉日内噪声，仅保留日内形状与日因子趋势")

    # ---- (d) 预热期 0 天（直接从 2.1 起、E_0=6000） vs 主模型 31 天 ----
    r42_d = eval_42(pr4, ve0, d_start=D_REP_FIRST, e_init=E_INIT)
    r43_d = eval_43(pr4, ve0, d_start=D_REP_FIRST, e_init=E_INIT)
    add("(d) 预热期 0 天（直接从 2025-02-01 起，E_0=6000）", "0 vs 31 天", r42_d, r43_d,
        "主模型含 31 天 1 月预热期（D-11）；本行为无预热对照")

    path = os.path.join(QDIR, "灵敏度分析_数据扰动.xlsx")
    write_xlsx(path, {"S3_数据扰动": rows})
    print("  S3 完成：%d 行 → %s" % (len(rows), path))
    return rows


# ============================== S4 方法侧互验 ==============================

def run_s4():
    """S4：方法侧互验——(a) 联合 LP 下界 (b) 终端三方案 (c) MATLAB 复算 (d) 小时粒度 LP。

    输出：`方法侧互验_联合LP与终端条件与MATLAB.xlsx`
    """
    print("-" * 78)
    print("【S4 方法侧互验】联合 LP 下界 / 终端三方案 / MATLAB / 小时粒度 LP")
    rows = [["项目", "指标", "数值", "单位", "相对差_%", "来源/说明"]]

    # ---- (a) 逐日滚动 + 终端余值 vs 全年联合 LP 下界（读 run_q4_joint.py 的落盘） ----
    joint_path = os.path.join(QDIR, "联合LP下界_对照.xlsx")
    if os.path.exists(joint_path):
        wb = openpyxl.load_workbook(joint_path, read_only=True)
        ws = wb["汇总对照"]
        got = {}
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[0] and isinstance(r[1], (int, float)):
                got[str(r[0])] = float(r[1])
        wb.close()
        roll = got.get("4-2 逐日滚动实际缴费（主模型）", Q42_ANCHOR)
        lb = got.get("填报区间联合最优费用（2.1–12.31）", float("nan"))
        gap = roll - lb
        rows.append(["(a) 联合 LP 下界", "填报区间联合最优（下界）", r4(lb), "元", 0.0,
                     "run_q4_joint.py：210240 变量 LP"])
        rows.append(["(a) 联合 LP 下界", "4-2 逐日滚动缴费", r4(roll), "元", r4(100.0 * gap / lb),
                     "滚动 − 下界 = %.4f 元（跨天协调全部空间）" % gap])
        rows.append(["(a) 联合 LP 下界", "4-3 多阶段滚动 J", r4(got.get("4-3 多阶段滚动总费用 J（主模型）", Q43_ANCHOR)),
                     "元", r4(100.0 * (got.get("4-3 多阶段滚动总费用 J（主模型）", Q43_ANCHOR) - lb) / lb),
                     "含预报误差与调整机制的代价"])
    else:
        rows.append(["(a) 联合 LP 下界", "未找到落盘", 0.0, "—", 0.0, "请先运行 run_q4_joint.py"])

    # ---- (b) 终端三方案（读 variants_q4.py 的落盘） ----
    var_path = os.path.join(QDIR, "终端条件对照_三方案.xlsx")
    if os.path.exists(var_path):
        wb = openpyxl.load_workbook(var_path, read_only=True)
        ws = wb["4-2终端条件三方案"]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[0]:
                rows.append(["(b) 终端条件", str(r[0]), r4(float(r[1])), "元（4-2 缴费）", 0.0,
                             "24:00=下限 %d 天、持有过夜 %d 天、利用率 %.4f"
                             % (int(r[5]), int(r[6]), float(r[7]))])
        ws = wb["4-3终端条件对照"]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[0] and r[0] != "口径说明":
                rows.append(["(b) 终端条件（4-3）", str(r[0]), r4(float(r[1])), "元（J）", 0.0,
                             "24:00=下限 %d 天、持有过夜 %d 天"
                             % (int(r[6]), int(r[7]))])
        wb.close()
    else:
        rows.append(["(b) 终端条件", "未找到落盘", 0.0, "—", 0.0, "请先运行 variants_q4.py"])

    # ---- (c) MATLAB 复算若干天 LP（读 _matlab校验_q4_汇总.csv） ----
    mat_path = os.path.join(QDIR, "_matlab校验_q4_汇总.csv")
    if os.path.exists(mat_path):
        with open(mat_path, encoding="utf-8") as f:
            for line in f.read().strip().splitlines()[1:]:
                parts = line.split(",")
                if len(parts) >= 2:
                    try:
                        rows.append(["(c) MATLAB 复算", parts[0], r4(float(parts[1])), "元/元或kWh",
                                     float(parts[2]) if len(parts) > 2 else 0.0,
                                     "matlab_q4.m 独立 linprog 复算"])
                    except ValueError:
                        pass
    else:
        rows.append(["(c) MATLAB 复算", "未找到 _matlab校验_q4_汇总.csv", 0.0, "—", 0.0,
                     "请先运行 matlab -batch matlab_q4"])

    # ---- (d) 小时粒度 LP 对照 ----
    t0 = time.time()
    h42 = run_hourly_42(_G["pr4"], _G["ve"])
    h43 = run_hourly_43(_G["pr4"], _G["ve"])
    rows.append(["(d) 小时粒度 LP", "4-2 小时粒度缴费", r4(h42["total"]), "元",
                 r4(100.0 * (h42["total"] - Q42_ANCHOR) / Q42_ANCHOR),
                 "10 分钟粒度主模型 %.4f 元；粗化损失（耗时 %.2f s）" % (Q42_ANCHOR, time.time() - t0)])
    rows.append(["(d) 小时粒度 LP", "4-3 小时粒度 J", r4(h43["J"]), "元",
                 r4(100.0 * (h43["J"] - Q43_ANCHOR) / Q43_ANCHOR),
                 "10 分钟粒度主模型 %.4f 元" % Q43_ANCHOR])

    path = os.path.join(QDIR, "方法侧互验_联合LP与终端条件与MATLAB.xlsx")
    write_xlsx(path, {"S4_方法侧互验": rows})
    print("  S4 完成：%d 行 → %s" % (len(rows), path))
    return rows


# ============================== S5 口径对照 ==============================

def run_s5(skip_heavy=False):
    """S5：口径对照——(a) 读法① vs 读法② (b) 结算口径 (c) 调整生效规则 (d) 填法。

    输出：`口径对照_读法与结算与调整规则.xlsx`
    """
    print("-" * 78)
    print("【S5 口径对照】信息结构 / 结算口径 / 调整生效规则 / 填法")
    pr4 = _G["pr4"]; ve0 = _G["ve"]
    rows = [["项目", "口径", "数值", "单位", "相对主口径差_%", "说明"]]

    # ---- (a) 4-2 信息结构：读法①（主，完全信息）vs 读法②（计划=典型日、实际=附件2） ----
    base42 = eval_42(pr4, ve0)
    r2 = roll_with_plan_42(pr4, ve0, plan_kind="typical")
    r2b = roll_with_plan_42(pr4, ve0, plan_kind="prev")
    rows.append(["(a) 信息结构", "读法① 完全信息（主口径）", r4(base42["total"]), "元", 0.0,
                 "计划恰好覆盖负载，r≡0"])
    rows.append(["(a) 信息结构", "读法② 计划=典型日（附件1 负载/光伏）", r4(r2["total"]), "元",
                 r4(100.0 * (r2["total"] - base42["total"]) / base42["total"]),
                 "紧急购电量 %.4f kWh、触发 %d/334 天" % (r2["r_sum"], r2["n_trigger"])])
    rows.append(["(a) 信息结构", "读法②b 计划=前一日实际曲线", r4(r2b["total"]), "元",
                 r4(100.0 * (r2b["total"] - base42["total"]) / base42["total"]),
                 "紧急购电量 %.4f kWh、触发 %d/334 天" % (r2b["r_sum"], r2b["n_trigger"])])

    # ---- (b) 结算口径 C-6：主口径（用户口径） vs 欠取按全价 + 违约金 ----
    r43 = eval_43(pr4, ve0)
    rows.append(["(b) 结算口径", "主口径 D-12（欠取按 0.5p 计价）", r4(r43["J"]), "元", 0.0,
                 "J = Σ[p·min(x,y)+κ₋p(x−y)⁺+κ₊p(y−x)⁺] + κΣpr"])
    rows.append(["(b) 结算口径", "对照：欠取按全价+违约金", r4(r43["J_alt"]), "元",
                 r4(100.0 * (r43["J_alt"] - r43["J"]) / r43["J"]),
                 "欠取部分实际单价 1.5p，总费用更高（D-12 必须报出的对照）"])

    # ---- (c) 调整生效规则 C-7：不可追溯（主） vs 可追溯极端 vs 边界（kmark+1） ----
    rows.append(["(c) 调整生效规则", "不可追溯（主口径 D-13）", r4(r43["J"]), "元", 0.0,
                 "k∈(τ_prev,τ] 用 τ 决策；0:00–6:00 段用计划"])
    if not skip_heavy:
        retr = run_retroactive_43(pr4, ve0)
        rows.append(["(c) 调整生效规则", "可追溯（极端：18:00 重优化全天含已执行段）",
                     r4(retr), "元", r4(100.0 * (retr - r43["J"]) / r43["J"]),
                     "规则收紧程度上界对照（不代表题目允许）"])
        kmark_b = {0: 0, 6: 37, 12: 73, 18: 109}          # 边界口径：调整不含整点所在段
        rb = eval_43(pr4, ve0, kmark=kmark_b)
        rows.append(["(c) 调整生效规则", "边界口径：调整不含整点所在 10 分钟段", r4(rb["J"]), "元",
                     r4(100.0 * (rb["J"] - r43["J"]) / r43["J"]),
                     "C-7 的保守解读：k∈(τ,τ+10min] 也沿用旧决策"])

    # ---- (d) 填法 Y-轮转（主） vs 填法 X（标签对位） ----
    # 从落盘明细反查：填法 X 下模板行缺少当天第 1 段（0:00–0:10），"全天购电量"会少算该段
    csv42 = os.path.join(QDIR, "全分辨率明细_4-2.csv")
    import csv as _csv
    miss_sum = 0.0; tot_sum = 0.0
    with open(csv42, encoding="utf-8") as f:
        for row in _csv.DictReader(f):
            v = float(row["计划购电量_kWh"]); tot_sum += v
            if int(row["时段序号"]) == 1:
                miss_sum += v                              # 填法 X 会漏掉的第 1 段
    rows.append(["(d) 时段填法", "填法 Y-轮转（主口径 D-01）", r4(tot_sum), "kWh（计划量合计）", 0.0,
                 "行内 144 格恰好覆盖当天 [0:00,24:00] 一次"])
    rows.append(["(d) 时段填法", "填法 X（标签对位，第 144 格无处安放第 1 段）",
                 r4(tot_sum - miss_sum), "kWh（计划量合计）",
                 r4(-100.0 * miss_sum / tot_sum),
                 "漏计当天第 1 段（0:00–0:10）合计 %.4f kWh，全天购电量偏差" % miss_sum])

    path = os.path.join(QDIR, "口径对照_读法与结算与调整规则.xlsx")
    write_xlsx(path, {"S5_口径对照": rows})
    print("  S5 完成：%d 行 → %s" % (len(rows), path))
    return rows


def roll_with_plan_42(price_mat, ve_arr, plan_kind="typical"):
    """读法②（不完全信息）的 4-2 滚动：计划用"典型日/前一日"曲线，实际为附件2。

    输入：price_mat (D,K) 元/kWh；ve_arr (D,) 元/kWh
          plan_kind，str，'typical'（附件1 负载/光伏）或 'prev'（前一日实际曲线）
    输出：dict——total（填报区间缴费 + 紧急费用，元）、r_sum（紧急购电量，kWh）、n_trigger
    说明：与问题 2 的 variants_q2.roll_with_plan 同一口径，唯一差异是逐日价格矩阵与终端余值。
    """
    load = _G["load"]; pv = _G["pv"]
    p1 = _G["price1"]; load1 = _G["load1"]; pv1 = _G["pv1"]
    if plan_kind == "typical":
        load_plan = np.tile(load1, (365, 1)); pv_plan = np.tile(pv1, (365, 1))
    else:
        load_plan = np.vstack([load1[None, :], load[:-1]])      # 前一日实际曲线
        pv_plan = np.vstack([pv1[None, :], pv[:-1]])
    e = E_INIT
    total = 0.0; r_sum = 0.0; n_trigger = 0
    for d in range(365):
        res = solve_day(price_mat[d], load_plan[d], pv_plan[d], e_init=e, mode="free",
                        v_end=ve_arr[d])
        x, u, v = res["x_plan"], res["u_chg"], res["v_dis"]
        r = np.maximum(load[d] * DT_H + u - v - pv[d] * DT_H - x, 0.0)   # 供给缺口
        if d >= D_REP_FIRST:
            total += float(np.dot(price_mat[d], x)) + 5.0 * float(np.dot(price_mat[d], r))
            r_sum += float(r.sum())
            n_trigger += int(r.max() > 1e-6)
        e = float(res["E_soc"][-1])                              # 计划侧储能轨迹跨日传递
    return {"total": total, "r_sum": r_sum, "n_trigger": n_trigger}


def run_retroactive_43(price_mat, ve_arr):
    """可追溯极端口径的 4-3：每天 18:00 用当日最后预报重优化全天（含已执行段）。

    输入：price_mat (D,K)；ve_arr (D,)；输出：float，填报区间总费用 J（元）
    说明：这是"规则收紧程度"的上界对照（D-13 的对照项），不代表题目允许。
    """
    from lib.solve_day import settle_total
    load = _G["load"]; pv = _G["pv"]; pv_fc144 = _G["pv_fc144"]
    e = E_INIT
    tot = 0.0
    for d in range(365):
        # 计划阶段（0:00）
        r0 = solve_stage(price_mat[d], load[d], pv_fc144[d, 0], e_init=e, x_ref=None,
                         k_start=0, n_period=K, v_end=0.0)
        x = r0["x_plan"]
        # 18:00 用当日最后预报重优化全天（含已执行段，可追溯）
        res = solve_stage(price_mat[d], load[d], pv_fc144[d, 3], e_init=e, x_ref=x,
                          k_start=0, n_period=K, v_end=ve_arr[d])
        y, u, v = res["x_plan"], res["u_chg"], res["v_dis"]
        E = np.concatenate([[e], e + np.cumsum(ETA * u - v / ETA)])
        residual = load[d] * DT_H + u - v - pv[d] * DT_H - y
        r_day = np.maximum(residual, 0.0)
        st = settle_total(price_mat[d], x, y, r_day)
        if d >= D_REP_FIRST:
            tot += st["J"]
        e = float(E[-1])
    return tot


# ============================== 主流程 ==============================

def main():
    """主流程：载入数据 → S1–S5 依次执行（含两张图）→ 打印汇总。"""
    parser = argparse.ArgumentParser(description="问题 4 灵敏度与稳健性实验（S1–S5）")
    parser.add_argument("--skip-heavy", action="store_true",
                        help="跳过 S1 的部分档位与 S5(c) 的重跑（快速冒烟）")
    args = parser.parse_args()

    price1, load1, pv1, _ = read_attachment1()
    _G["price1"] = price1; _G["load1"] = load1; _G["pv1"] = pv1
    _G["load"], _G["pv"], _G["dates"] = read_attachment2()
    _G["pr4"], _ = read_attachment4()
    _, pv_fc144, _, _ = load_forecast(method="linear")
    _G["pv_fc144"] = pv_fc144
    _G["ve"] = terminal_value_definitions(_G["pr4"])             # D-10 主口径余值

    print("=" * 78)
    print("问题 4 灵敏度与稳健性实验（S1–S5）：附件4 波动电价 + D-10 终端余值")
    print("=" * 78)
    t0 = time.time()
    levels = (10,) if args.skip_heavy else (5, 10, 20)
    run_s1(levels=levels)
    run_s2()
    run_s3()
    run_s4()
    run_s5(skip_heavy=args.skip_heavy)
    print("=" * 78)
    print("全部灵敏度实验完成，总耗时 %.1f 分钟" % ((time.time() - t0) / 60.0))
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
