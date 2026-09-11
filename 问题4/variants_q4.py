"""问题 4 终端条件对照（D-10）：终端自由 / 终端=当日初始 / 终端加余值 V_E。

目的（`问题4/交付清单.md` §三、§五 S4b 与台账 D-10）：
  给出三种终端条件下 4-2 与 4-3 链的全年费用与储电量利用情况，量化"终端自由低估储能价值"
  的幅度、以及"终端=当日初始"的对照水平。主口径（D-10）为终端加余值 V_E = 次日最低价/η。
  附加口径（台账 C-9-细分）：V_E = 次日 24h 均价/η，用于对照余值取法。

三个方案的实现（全部复用既有代码链，不新写模型）：
  * 终端自由：`solve_rolling(..., v_end=0)`（4-2）/ `solve_rolling_staged(..., v_end_day=None)`（4-3）；
  * 终端=当日初始：`solve_rolling(..., mode='cyclic')`（每日 E_144 = E_0，跨日仍连续）；
  * 终端加余值：`v_end = terminal_value_definitions(...)`（D-10 主口径）。

输出文件（均在 `问题4/` 下）：
  终端条件对照_三方案.xlsx   4-2 与 4-3 各方案的费用、期末储电量、24:00 分布、利用率
  终端条件对照图.png         三种终端条件下全年费用与储电量利用率对比

运行：python 问题4/variants_q4.py（约 2 分钟）
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

from lib.dataio import read_attachment2, read_attachment4
from lib.forecast import load_forecast
from lib.logio import Tee
from lib.run_days import solve_rolling, solve_rolling_staged
from lib.solve_day import terminal_value_definitions
from lib.storage import E_CAP, E_INIT, E_MAX, E_MIN, ETA
from lib.timegrid import K

QDIR = os.path.join(ROOT, "问题4")
FIGDIR = os.path.join(ROOT, "图片", "问题4")                        # 图片统一目录（2026-09-11 起，用户要求）
os.makedirs(FIGDIR, exist_ok=True)                                # 确保目录存在（重跑时自动建）
XLSX_PATH = os.path.join(QDIR, "终端条件对照_三方案.xlsx")
PNG_PATH = os.path.join(FIGDIR, "终端条件对照图.png")
LOG_PATH = os.path.join(QDIR, "终端条件对照运行日志.txt")

D_REP_FIRST = 31                          # 填报区间首日 2025-02-01（0 基 31）
N_REP = 334
ND = 4                                    # 小数位数（D-15）
STAGES_MAIN = (0, 6, 12, 18)              # 4-3 主模型全调整


def run_scheme_42(price_mat, load, pv_actual, scheme, ve_arr=None):
    """跑 4-2 链的一种终端条件方案。

    输入：price_mat (D,K) 元/kWh；load/pv_actual (D,K) kW
          scheme，str，'free'（终端自由）/ 'cyclic'（终端=当日初始）/ 'residual'（终端余值）
          ve_arr，(D,) 元/kWh；scheme='residual' 时必须给出
    输出：dict，含填报区间费用的统计（元/kWh 量纲见键名注释）
    """
    if scheme == "cyclic":
        roll = solve_rolling(price_mat, load, pv_actual, e_init=E_INIT, mode="cyclic")
    else:
        ve = np.zeros(365) if scheme == "free" else ve_arr
        roll = solve_rolling(price_mat, load, pv_actual, e_init=E_INIT, mode="free", v_end=ve)
    sl = slice(D_REP_FIRST, 365)
    x = roll["x_plan"][sl]; u = roll["u_chg"][sl]; v = roll["v_dis"][sl]
    E = roll["E_soc"][sl]; cost = roll["cost_day"][sl]
    return {
        "total": float(cost.sum()),                                  # 填报区间缴费 Σp·x，元
        "E0_rep": float(E[0, 0]),                                    # 期初（2.1 0:00）储电量，kWh
        "E_end": float(E[-1, K]),                                    # 期末（12.31 24:00）储电量，kWh
        "n_hold": int(np.sum(E[:, K] > E_MIN + 1e-3)),               # 24:00 高于下限的天数
        "n_min": int(np.sum(np.abs(E[:, K] - E_MIN) < 1e-3)),        # 24:00 落于下限的天数
        "u_sum": float(u.sum()), "v_sum": float(v.sum()),
        "soc_util": float(E[:, 1:K + 1].mean() / E_CAP),             # 储电量利用率 = 平均储电量/额定容量
        "date_tot": E, "cost_day": cost,                             # 供逐日图/复核使用
    }


def run_scheme_43(price_mat, load, pv_actual, pv_fc144, scheme, ve_arr=None):
    """跑 4-3 链的一种终端条件方案（多阶段；终端=当日初始用 v_end 缺失时的默认自由。

    说明：多阶段链的"终端=当日初始"没有直接开关（计划与调整阶段串联），因此 4-3 只报
    终端自由与终端余值两种；"终端=当日初始"的效应由 4-2 链给出（两链同价格、同储能模型）。

    输入：price_mat (D,K) 元/kWh；load/pv_actual (D,K) kW；pv_fc144 (D,4,K) kW
          scheme，str，'free' 或 'residual'；ve_arr，(D,) 元/kWh
    输出：dict，含填报区间 D-12 汇总（元/kWh 量纲见键名注释）
    """
    ve_day = None if scheme == "free" else ve_arr
    roll = solve_rolling_staged(price_mat, load, pv_actual, pv_fc144,
                                stages=STAGES_MAIN, d_start=0, d_end=365, v_end_day=ve_day)
    sl = slice(D_REP_FIRST, 365)
    E = roll["E_soc"][sl]
    return {
        "J": float(roll["J_day"][sl].sum()),                         # 主口径总费用，元
        "J_plan": float(roll["J_plan"][sl].sum()),                   # 计划购电费用，元
        "J_adj": float(roll["J_adj"][sl].sum()),                     # 调整相关费用，元
        "J_emg": float(roll["J_emg"][sl].sum()),                     # 紧急购电费用，元
        "E_end": float(E[-1, K]),                                    # 期末储电量，kWh
        "n_hold": int(np.sum(E[:, K] > E_MIN + 1e-3)),               # 24:00 高于下限的天数
        "n_min": int(np.sum(np.abs(E[:, K] - E_MIN) < 1e-3)),        # 24:00 落于下限的天数
        "soc_util": float(E[:, 1:K + 1].mean() / E_CAP),             # 储电量利用率
        "r_sum": float(roll["r_emg"][sl].sum()),                     # 紧急购电量，kWh
    }


def write_xlsx(path, sheets):
    """把 {工作表名: 行列表} 写成 xlsx（数值 4 位小数格式，与问题 2/3 风格一致）。

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


def draw_figure(names, cost_42, util_42, cost_43, util_43, path):
    """绘制"终端条件对照图"：三方案的费用与储电量利用率对比。

    输入：names list[str] 方案名；cost_42/cost_43 (m,) 元；util_42/util_43 (m,) 无量纲；path，str
    输出：str，PNG 路径
    """
    import matplotlib.pyplot as plt
    from lib.plotstyle import (COLOR_BUY, COLOR_CHG, COLOR_PRICE, FIGSIZE_TALL,
                               apply_chinese_style, save_figure)
    apply_chinese_style()
    fig, axes = plt.subplots(2, 1, figsize=FIGSIZE_TALL)
    m = len(names)
    xs = np.arange(m)
    ax = axes[0]
    w = 0.36
    b1 = ax.bar(xs - w / 2, cost_42, width=w, color=COLOR_BUY, alpha=0.9, label="4-2 全年缴费（元）")
    ax2 = ax.twinx()
    b2 = ax2.bar(xs + w / 2, cost_43, width=w, color=COLOR_PRICE, alpha=0.9, label="4-3 全年总费用 J（元）")
    for b, v in zip(b1, cost_42):
        if np.isfinite(v):
            ax.annotate("%.1f 万" % (v / 1e4), (b.get_x() + b.get_width() / 2, v), ha="center",
                        va="bottom", fontsize=9, textcoords="offset points", xytext=(0, 2))
    for b, v in zip(b2, cost_43):
        if np.isfinite(v):
            ax2.annotate("%.1f 万" % (v / 1e4), (b.get_x() + b.get_width() / 2, v), ha="center",
                         va="bottom", fontsize=9, textcoords="offset points", xytext=(0, 2))
    ax.set_xticks(xs, names, fontsize=10)
    ax.set_ylabel("4-2 全年缴费（元）")
    ax2.set_ylabel("4-3 全年总费用 J（元）")
    ax.set_title("三种终端条件下的全年费用（附件4 波动电价）")
    h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, fontsize=10, loc="upper left")
    ax3 = axes[1]
    ax3.bar(xs - w / 2, np.asarray(util_42) * 100, width=w, color=COLOR_CHG, alpha=0.9,
            label="4-2 储电量利用率（%）")
    ax3.bar(xs + w / 2, np.asarray(util_43) * 100, width=w, color="#7F8C8D", alpha=0.9,
            label="4-3 储电量利用率（%）")
    for x, v in zip(xs - w / 2, util_42):
        if np.isfinite(v):
            ax3.annotate("%.2f%%" % (v * 100), (x, v * 100), ha="center", va="bottom", fontsize=9,
                         textcoords="offset points", xytext=(0, 2))
    for x, v in zip(xs + w / 2, util_43):
        if np.isfinite(v):
            ax3.annotate("%.2f%%" % (v * 100), (x, v * 100), ha="center", va="bottom", fontsize=9,
                         textcoords="offset points", xytext=(0, 2))
    ax3.set_xticks(xs, names, fontsize=10)
    ax3.set_ylabel("平均储电量 / 额定容量（%）")
    ax3.set_title("储电量利用率：终端余值使储能在日末保留存量（持有过夜），利用率上升")
    ax3.legend(fontsize=10)
    fig.suptitle("问题 4 终端条件三方案对照（D-10）：终端自由 / 终端=当日初始 / 终端加余值 "
                 "$V_E$ = 次日最低价/η", fontsize=14)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    return save_figure(fig, path)


def main():
    """主流程：读数据 → 三方案 4-2 + 两方案 4-3 → 落盘 xlsx 与图 → 打印对照。"""
    load, pv_actual, dates = read_attachment2()
    pr4, dates4 = read_attachment4()
    pv_fc, pv_fc144, dates3, tau_list = load_forecast(method="linear")
    ve_min = terminal_value_definitions(pr4)                      # D-10 主口径：次日最低价/η
    p_min_avg = pr4.mean(axis=1)                                  # 次日 24h 均价口径（C-9-细分）
    ve_avg = np.concatenate([p_min_avg[1:], p_min_avg[-1:]]) / ETA

    print("=" * 78)
    print("问题 4 终端条件对照：终端自由 / 终端=当日初始 / 终端加余值（D-10）")
    print("=" * 78)

    # ---- 4-2 链三方案（+ 均价余值附加口径） ----
    names = ["终端自由\n($V_E$=0)", "终端=当日初始\n(cyclic)", "终端加余值\n$V_E$=次日最低价/η",
             "附加：终端加余值\n$V_E$=次日均价/η"]
    schemes = ["free", "cyclic", "residual", "residual_avg"]
    r42 = []
    for sch in schemes:
        t0 = time.time()
        if sch == "residual_avg":
            res = run_scheme_42(pr4, load, pv_actual, "residual", ve_arr=ve_avg)
        else:
            res = run_scheme_42(pr4, load, pv_actual, sch, ve_arr=ve_min)
        r42.append(res)
        print("4-2 [%s] 缴费 = %.4f 元；期末E = %.4f kWh；24:00=下限 %d 天；利用率 %.4f（%.1fs）"
              % (names[schemes.index(sch)].replace("\n", ""), res["total"], res["E_end"],
                 res["n_min"], res["soc_util"], time.time() - t0))

    # ---- 4-3 链两方案（自由 / 余值） ----
    r43_free = run_scheme_43(pr4, load, pv_actual, pv_fc144, "free")
    print("4-3 [终端自由] J = %.4f 元；期末E = %.4f kWh；24:00=下限 %d 天；利用率 %.4f"
          % (r43_free["J"], r43_free["E_end"], r43_free["n_min"], r43_free["soc_util"]))
    r43_res = run_scheme_43(pr4, load, pv_actual, pv_fc144, "residual", ve_arr=ve_min)
    print("4-3 [终端余值] J = %.4f 元；期末E = %.4f kWh；24:00=下限 %d 天；利用率 %.4f"
          % (r43_res["J"], r43_res["E_end"], r43_res["n_min"], r43_res["soc_util"]))

    # ---- 落盘 ----
    rows42 = [["终端方案", "全年缴费_元", "日均_元", "期末储电量_kWh", "期初储电量_kWh",
               "24:00落于下限天数", "持有过夜天数", "储电量利用率", "总充电量_kWh", "总放电量_kWh"],
              ["终端自由（V_E=0）", round(r42[0]["total"], ND), round(r42[0]["total"] / N_REP, ND),
               round(r42[0]["E_end"], ND), round(r42[0]["E0_rep"], ND), r42[0]["n_min"],
               r42[0]["n_hold"], round(r42[0]["soc_util"], 6), round(r42[0]["u_sum"], ND),
               round(r42[0]["v_sum"], ND)],
              ["终端=当日初始（cyclic）", round(r42[1]["total"], ND), round(r42[1]["total"] / N_REP, ND),
               round(r42[1]["E_end"], ND), round(r42[1]["E0_rep"], ND), r42[1]["n_min"],
               r42[1]["n_hold"], round(r42[1]["soc_util"], 6), round(r42[1]["u_sum"], ND),
               round(r42[1]["v_sum"], ND)],
              ["终端加余值 V_E=次日最低价/η（D-10 主口径）", round(r42[2]["total"], ND),
               round(r42[2]["total"] / N_REP, ND), round(r42[2]["E_end"], ND),
               round(r42[2]["E0_rep"], ND), r42[2]["n_min"], r42[2]["n_hold"],
               round(r42[2]["soc_util"], 6), round(r42[2]["u_sum"], ND), round(r42[2]["v_sum"], ND)],
              ["附加：终端加余值 V_E=次日 24h 均价/η（C-9 对照）", round(r42[3]["total"], ND),
               round(r42[3]["total"] / N_REP, ND), round(r42[3]["E_end"], ND),
               round(r42[3]["E0_rep"], ND), r42[3]["n_min"], r42[3]["n_hold"],
               round(r42[3]["soc_util"], 6), round(r42[3]["u_sum"], ND), round(r42[3]["v_sum"], ND)],
              ]
    rows43 = [["终端方案", "总费用J_元", "计划购电费_元", "调整相关费用_元", "紧急购电费_元",
               "期末储电量_kWh", "24:00落于下限天数", "持有过夜天数", "储电量利用率", "紧急购电量_kWh"],
              ["终端自由（V_E=0）", round(r43_free["J"], ND), round(r43_free["J_plan"], ND),
               round(r43_free["J_adj"], ND), round(r43_free["J_emg"], ND),
               round(r43_free["E_end"], ND), r43_free["n_min"], r43_free["n_hold"],
               round(r43_free["soc_util"], 6), round(r43_free["r_sum"], ND)],
              ["终端加余值 V_E=次日最低价/η（D-10 主口径）", round(r43_res["J"], ND),
               round(r43_res["J_plan"], ND), round(r43_res["J_adj"], ND), round(r43_res["J_emg"], ND),
               round(r43_res["E_end"], ND), r43_res["n_min"], r43_res["n_hold"],
               round(r43_res["soc_util"], 6), round(r43_res["r_sum"], ND)],
              ["口径说明", "4-3 多阶段链的终端=当日初始无直接开关（计划+调整串联），其效应由 4-2 链给出",
               "", "", "", "", "", "", "", ""],
              ]
    write_xlsx(XLSX_PATH, {"4-2终端条件三方案": rows42, "4-3终端条件对照": rows43})
    print("已落盘：%s" % XLSX_PATH)
    # 4-3 无"终端=当日初始"与"均价余值"两档，用 nan 占位（画图时跳过）
    nan4 = [float("nan")] * 4
    cost43_plot = [r43_free["J"], nan4[0], r43_res["J"], nan4[1]]
    util43_plot = [r43_free["soc_util"], nan4[2], r43_res["soc_util"], nan4[3]]
    print("已落盘：%s" % draw_figure(names,
                                    [r42[0]["total"], r42[1]["total"], r42[2]["total"], r42[3]["total"]],
                                    [r42[0]["soc_util"], r42[1]["soc_util"], r42[2]["soc_util"],
                                     r42[3]["soc_util"]],
                                    cost43_plot, util43_plot, PNG_PATH))
    # 关键结论打印
    d13 = r42[2]["total"] - r42[0]["total"]                # 余值方案 − 终端自由（缴费），元
    print("-" * 78)
    print("【关键对照】")
    print("终端余值 vs 终端自由：缴费 %+.4f 元（%+.4f%%）；24:00=下限 %d → %d 天；"
          "持有过夜 %d → %d 天；利用率 %.4f → %.4f"
          % (d13, 100.0 * d13 / r42[0]["total"], r42[0]["n_min"], r42[2]["n_min"],
             r42[0]["n_hold"], r42[2]["n_hold"], r42[0]["soc_util"], r42[2]["soc_util"]))
    print("终端=当日初始 vs 终端自由：缴费 %+.4f 元（%+.4f%%）"
          % (r42[1]["total"] - r42[0]["total"],
             100.0 * (r42[1]["total"] - r42[0]["total"]) / r42[0]["total"]))
    print("4-3：终端余值 vs 终端自由 J %+.4f 元（%+.4f%%）"
          % (r43_res["J"] - r43_free["J"], 100.0 * (r43_res["J"] - r43_free["J"]) / r43_free["J"]))
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
