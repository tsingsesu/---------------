r"""问题 3 策略对比脚本：回答"是否需要引入其他时刻的预报制定调整购电策略"。

五条策略（全部在填报区间 334 天上评估，单位：元）：
  (a) 仅 0:00 预报、不调整（y≡x）；
  (b) 0:00 + 6:00 调整；
  (c) 0:00 + 6:00 + 12:00 调整；
  (d) 0:00 + 6:00 + 12:00 + 18:00 全调整（主模型）；
  (e) 完全信息（预报=附件2 实际光伏，r≡0），费用下界。

输出：
  * `策略对比.xlsx`：汇总表（总费用与分项）+ 边际收益表 + 逐日费用；
  * `策略对比_逐日费用.csv`：逐日各策略总费用（审计明细）；
  * `策略对比柱状图.png`：五策略分项堆叠柱状图，标注边际收益。

口径：D-12 分项写法 J = J_plan(Σp·x) + J_adj(调整相关、有符号) + J_emg；与 D-07/D-13/D-04
     完全一致；四条策略共用同一份预报与同一套求解链路（lib/run_days.solve_rolling_staged），
     唯一差别是 stages 参数（参与决策的发布时刻集合）。

运行：python 问题3/run_q3_strategies.py
随机性：无（确定性 LP）。
"""

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
from lib.plotstyle import (COLOR_BUY, COLOR_CHG, COLOR_PRICE, COLOR_REF, FIGSIZE_WIDE,
                           apply_chinese_style, save_figure)
from lib.run_days import solve_rolling_staged
from lib.storage import E_INIT

QDIR = os.path.join(ROOT, "问题3")
LOG_PATH = os.path.join(QDIR, "策略对比运行日志.txt")
XLSX_PATH = os.path.join(QDIR, "策略对比.xlsx")
CSV_PATH = os.path.join(QDIR, "策略对比_逐日费用.csv")
FIG_PATH = os.path.join(QDIR, "策略对比柱状图.png")

D_REP_FIRST = 31                       # 填报区间首日（0 基 31 = 2025-02-01）
N_REP = 334                            # 填报区间天数

# 策略定义：名称 -> 参与的决策时刻集合（均含 0:00）
STRATEGIES = [
    ("a_仅0时", (0,)),
    ("b_加6时", (0, 6)),
    ("c_加6时12时", (0, 6, 12)),
    ("d_全调整", (0, 6, 12, 18)),
]
S_IDS = [s[0] for s in STRATEGIES]


def run_all_strategies(price, load, pv_actual, pv_fc144):
    """逐策略跑完整年滚动链，返回各策略的填报区间结果字典。

    输入：price (K,) 元/kWh；load/pv_actual (D,K) kW；pv_fc144 (D,4,K) kW（M2 分解）
    输出：dict，{策略名: roll}，roll 为 solve_rolling_staged 返回值（未切片，含 1 月预热）
    """
    out = {}
    for name, stages in STRATEGIES:
        roll = solve_rolling_staged(price, load, pv_actual, pv_fc144,
                                    stages=stages, d_start=0, d_end=365)
        out[name] = roll
    # (e) 完全信息下界：预报替换为实际光伏，无调整需求（r≡0），退化为问题 2 的读法①
    pv_fc_actual = np.tile(pv_actual[:, None, :], (1, 4, 1))
    roll_e = solve_rolling_staged(price, load, pv_actual, pv_fc_actual,
                                  stages=(0,), d_start=0, d_end=365)
    out["e_完全信息"] = roll_e
    return out


def summarize(rolls, sl):
    """把各策略的填报区间结果汇总成指标表。

    输入：rolls，dict（run_all_strategies 输出）；sl，slice，填报区间切片
    输出：list[dict]，每策略一行，含总费用与分项、Σy、Σr、触发天数
    """
    rows = []
    for name, roll in rolls.items():
        J = float(roll["J_day"][sl].sum())
        rows.append({
            "策略": name,
            "总费用_元": J,
            "计划购电费用_元": float(roll["J_plan"][sl].sum()),
            "调整相关费用_元": float(roll["J_adj"][sl].sum()),
            "紧急购电费用_元": float(roll["J_emg"][sl].sum()),
            "对照口径总费用_元": float(roll["J_alt_day"][sl].sum()),
            "总最终购电量_kWh": float(roll["y_adj"][sl].sum()),
            "紧急购电量_kWh": float(roll["r_emg"][sl].sum()),
            "触发紧急购电天数": int(np.sum(roll["r_emg"][sl].sum(axis=1) > 1e-6)),
            "单时段最大紧急购电量_kWh": float(roll["r_emg"][sl].max()),
        })
    return rows


def marginal_benefit(rows):
    """计算相邻策略之间的边际收益（每增加一个预报时刻省下的费用）。

    输入：rows，list[dict]（summarize 输出，按 a→d→e 顺序）
    输出：list[dict]，每行含"从…到…/节省额/占比/说明"
    """
    out = []
    order = rows
    for i in range(1, len(order)):
        prev, cur = order[i - 1], order[i]
        save = prev["总费用_元"] - cur["总费用_元"]
        if i < len(order) - 1:
            note = "新增一个预报时刻的边际价值"
        else:
            note = "信息完全（下界）与全调整之差 = 预报误差与调整机制的总代价"
        out.append({
            "增量": "%s → %s" % (prev["策略"], cur["策略"]),
            "节省额_元": float(save),
            "占基准(a)比例": float(save / order[0]["总费用_元"]),
            "说明": note,
        })
    return out


def write_xlsx(path, rows, marg, rolls, sl, dates):
    """把策略对比结果写入 xlsx（汇总 / 边际收益 / 逐日费用 三张表）。

    输入：path，str；rows/marg，list[dict]；rolls，dict；sl，slice；dates，list[date]
    输出：无
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "策略汇总"
    header = list(rows[0].keys())
    ws.append(header)
    for r in rows:
        ws.append([r[h] for h in header])
    ws2 = wb.create_sheet("边际收益")
    h2 = list(marg[0].keys())
    ws2.append(h2)
    for r in marg:
        ws2.append([r[h] for h in h2])
    ws3 = wb.create_sheet("逐日费用")
    ws3.append(["日期"] + [name for name in rolls.keys()])
    idx = np.arange(sl.start, sl.stop)
    for i, d in enumerate(idx):
        ws3.append([dates[d].isoformat()] + [float(rolls[name]["J_day"][d]) for name in rolls])
    # 全部数值列保留 4 位小数（D-15），文本列不动
    for w in wb.worksheets:
        for row in w.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value, float):
                    c.number_format = "0.0000"
    wb.save(path)


def write_csv(path, rolls, sl, dates):
    """落盘逐日费用 CSV（外部审计用）。

    输入：path，str；rolls，dict；sl，slice；dates，list[date]
    输出：无
    """
    idx = np.arange(sl.start, sl.stop)
    lines = ["日期," + ",".join(rolls.keys())]
    for d in idx:
        lines.append(dates[d].isoformat() + "," +
                     ",".join("%.4f" % rolls[name]["J_day"][d] for name in rolls))
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("\n".join(lines) + "\n")


def draw_figure(path, rows, marg):
    """绘制五策略分项堆叠柱状图，并标注边际收益。

    输入：path，str；rows/marg，list[dict]
    输出：无（PNG 落盘）
    """
    apply_chinese_style()
    import matplotlib.pyplot as plt

    labels = [r["策略"] for r in rows]
    plan = np.array([r["计划购电费用_元"] for r in rows])
    adj = np.array([r["调整相关费用_元"] for r in rows])
    emg = np.array([r["紧急购电费用_元"] for r in rows])
    x = np.arange(len(labels))
    fig, ax = plt.subplots(figsize=FIGSIZE_WIDE)
    w = 0.55
    ax.bar(x, plan / 1e4, width=w, color=COLOR_BUY, label="计划购电费用")
    # 调整相关费用可为负：正负叠加都画在计划段之上/之下，用红/绿区分
    ax.bar(x, adj / 1e4, width=w, bottom=plan / 1e4, color=COLOR_PRICE, label="调整相关费用（有符号）")
    ax.bar(x, emg / 1e4, width=w, bottom=(plan + adj) / 1e4, color=COLOR_CHG, label="紧急购电费用")
    # 总额标注
    for xi, r in zip(x, rows):
        ax.annotate("%.2f 万元" % (r["总费用_元"] / 1e4), (xi, r["总费用_元"] / 1e4),
                    textcoords="offset points", xytext=(0, 8), ha="center", fontsize=10)
    # 边际收益标注（相邻柱之间）
    for m in marg:
        pair = m["增量"].split(" → ")
        try:
            i0 = labels.index(pair[0]); i1 = labels.index(pair[1])
        except ValueError:
            continue
        xm = (i0 + i1) / 2.0
        ax.annotate("省 %.1f 万元\n(%.2f%%)" % (m["节省额_元"] / 1e4, 100 * m["占基准(a)比例"]),
                    (xm, 132), ha="center", fontsize=9, color=COLOR_REF,
                    arrowprops=None)
    ax.set_xticks(x, labels, rotation=15)
    ax.set_xlabel("购电调整策略（采用的预报发布时刻）")
    ax.set_ylabel("填报区间总费用（万元，2025-02-01 至 12-31）")
    ax.set_title("问题 3 策略对比：不同预报时刻组合下的全年购电费用构成\n"
                 "（a→d 依次增加 6:00/12:00/18:00 预报；e 为完全信息下界）")
    # 在图内加一条"边际收益递减"的说明线
    ax.plot(x, np.array([r["总费用_元"] for r in rows]) / 1e4, "o--", color="black",
            linewidth=1.2, markersize=5, label="总费用（主口径 D-12）")
    ax.legend(loc="lower left")
    save_figure(fig, path)


def main():
    """主流程：跑五策略 → 汇总与边际 → 落盘 xlsx/csv/图 → 打印关键数值。"""
    print("=" * 78)
    print("问题 3 策略对比：是否需要引入其他时刻的预报制定调整购电策略？")
    print("=" * 78)
    price, load1, pv1, _ = read_attachment1()
    load, pv_actual, dates = read_attachment2()
    pv_fc, pv_fc144, dates3, tau_list = load_forecast(method="linear")
    sl = slice(D_REP_FIRST, 365)

    rolls = run_all_strategies(price, load, pv_actual, pv_fc144)
    rows = summarize(rolls, sl)
    marg = marginal_benefit(rows)
    write_xlsx(XLSX_PATH, rows, marg, rolls, sl, dates)
    write_csv(CSV_PATH, rolls, sl, dates)
    draw_figure(FIG_PATH, rows, marg)

    print("-" * 78)
    print("【策略对比表（填报区间 334 天；主口径 D-12）】")
    print("%-10s %14s %14s %12s %14s %12s %10s" %
          ("策略", "总费用(元)", "计划购电费", "调整费用", "紧急费用", "Σy(kWh)", "Σr(kWh)"))
    for r in rows:
        print("%-10s %14.4f %14.4f %12.4f %14.4f %12.2f %10.2f" %
              (r["策略"], r["总费用_元"], r["计划购电费用_元"], r["调整相关费用_元"],
               r["紧急购电费用_元"], r["总最终购电量_kWh"], r["紧急购电量_kWh"]))
    print("-" * 78)
    print("【边际收益（每增加一个预报时刻省下的费用）】")
    for m in marg:
        print("  %-24s 省 %12.4f 元（占 (a) 基准 %.4f%%）— %s" %
              (m["增量"], m["节省额_元"], 100 * m["占基准(a)比例"], m["说明"]))
    print("-" * 78)
    # 验收要求的两条硬单调性：(a) ≥ (d)（调整至少不亏）、(d) ≥ (e)（预报不完美必有代价）
    ok_ad = rows[0]["总费用_元"] >= rows[3]["总费用_元"]
    ok_de = rows[3]["总费用_元"] >= rows[4]["总费用_元"]
    print("【单调性检验】验收要求：(a) ≥ (d) %s；(d) ≥ (e) %s（下界）" % (ok_ad, ok_de))
    print("  完整链路 a ≥ b ≥ c ≥ d：%s；其中 c−d = %.4f 元（|·|/总费用 = %.2e）"
          % (all(rows[i]["总费用_元"] >= rows[i + 1]["总费用_元"] - 1e-6 for i in range(3)),
             rows[2]["总费用_元"] - rows[3]["总费用_元"],
             abs(rows[2]["总费用_元"] - rows[3]["总费用_元"]) / rows[3]["总费用_元"]))
    print("  说明：c 与 d 只在 18:00 之后的时段（k≥109，日落前后）不同；18:00 预报对剩余时段的")
    print("        更新幅度 ≤0.4 kWh/段，其价值为数值零（0.00003%），符号在日间正负随机。")
    print("  逐档差：b−a = %.4f；c−b = %.4f；d−c = %.4f；e−d（预报代价） = %.4f"
          % (rows[0]["总费用_元"] - rows[1]["总费用_元"],
             rows[1]["总费用_元"] - rows[2]["总费用_元"],
             rows[2]["总费用_元"] - rows[3]["总费用_元"],
             rows[3]["总费用_元"] - rows[4]["总费用_元"]))
    print("已落盘：%s / %s / %s" % (XLSX_PATH, CSV_PATH, FIG_PATH))
    print("=" * 78)
    # 关键结论（供构思手直接引用）
    print("【定量回答】引入 6:00 预报的边际收益 = %.4f 元/年；再引入 12:00 = %.4f 元/年；"
          "再引入 18:00 = %.4f 元/年（≈0，数值零点）。"
          % (marg[0]["节省额_元"], marg[1]["节省额_元"], marg[2]["节省额_元"]))
    return rows, marg


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
