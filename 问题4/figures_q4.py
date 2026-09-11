"""问题 4 主图绘制（从落盘文件出图，保证"图表数字 = 交付文件数字"）。

按 `问题4/交付清单.md` §四 绘制四张图（全部 ≥300 dpi、中文标注）：
  波动电价三维热力图.png       横轴时刻（144）、纵轴日期（365）、颜色电价；叠加四个指定日期参考线
  固定电价与波动电价结果对比.png 四个指定日期：问题2 vs 4-2、问题3 vs 4-3 的购电量与费用（分组柱）
  波动电价下计划与调整购电量图.png 四个指定日期：计划量、最终量（调整后）叠加与偏差阴影
  波动电价下储电量轨迹图.png    全年储电量热力图 + 四个指定日期的储电量轨迹

数据来源（只读，均为已交付/已落盘文件）：
  问题4/全分辨率明细_4-2.csv、全分辨率明细_4-3.csv（4-2 与 4-3 的全分辨率明细）
  附件/附件4.xlsx（电价热力图）
  问题2/逐日结果.csv、问题3/全分辨率明细.csv（固定电价对照值；只读，不修改）

运行：python 问题4/figures_q4.py（须先跑 run_q4.py 与 run_q2.py / run_q3.py）
依赖：numpy、matplotlib；lib/ 公共模块。
随机性：无。
"""

import csv
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import numpy as np

from lib.dataio import read_attachment4
from lib.logio import Tee
from lib.plotstyle import (COLOR_BUY, COLOR_CHG, COLOR_DIS, COLOR_LOAD, COLOR_PRICE,
                           COLOR_REF, COLOR_SOC, FIGSIZE_TALL, FIGSIZE_WIDE,
                           apply_chinese_style, save_figure)
from lib.timegrid import DT_H, K

QDIR = os.path.join(ROOT, "问题4")
CSV_42 = os.path.join(QDIR, "全分辨率明细_4-2.csv")
CSV_43 = os.path.join(QDIR, "全分辨率明细_4-3.csv")
Q2_CSV = os.path.join(ROOT, "问题2", "逐日结果.csv")       # 固定电价 4-2 对照（只读）
Q3_CSV = os.path.join(ROOT, "问题3", "全分辨率明细.csv")   # 固定电价 4-3 对照（只读）

D_REP_FIRST = 31
N_REP = 334
SPEC_DATES = ("2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21")


def load_matrix(path, cols):
    """读全分辨率明细 CSV 的指定列为 (334,144) 数组 + 日期列表。

    输入：path，str；cols，list[str]，列名
    输出：(data, dates)——data dict{列名: (334,144) float}；dates list[str]（334 个）
    """
    with open(path, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    data = {c: np.zeros((N_REP, K)) for c in cols}
    dates = []
    for idx, r in enumerate(rows):
        i, k = divmod(idx, K)
        if k == 0:
            dates.append(r["日期"])
        for c in cols:
            data[c][i, k] = float(r[c])
    return data, dates


def main():
    """主流程：读四份落盘文件 → 依次绘制四张图。"""
    apply_chinese_style()
    d42, dates = load_matrix(CSV_42, ["电价_元每kWh", "计划购电量_kWh", "充电量_kWh",
                                      "放电量_kWh", "时段末储电量_kWh"])
    d43, _ = load_matrix(CSV_43, ["电价_元每kWh", "计划购电量_kWh", "调整购电量_kWh",
                                  "充电量_kWh", "放电量_kWh", "紧急购电量_kWh",
                                  "时段末储电量_kWh"])
    pr4, dates4 = read_attachment4()
    import datetime as _dt
    spec_idx = [(_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days - D_REP_FIRST
                for t in SPEC_DATES]
    hours = np.arange(1, K + 1) / 6.0 - 1.0 / 12.0              # 各时段中心钟点，h
    n_day = len(dates)
    print("已读入：4-2 明细 %d 天、4-3 明细 %d 天、附件4 电价 %s" % (n_day, n_day, pr4.shape))

    # ============ 图 1：波动电价三维热力图 ============
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=FIGSIZE_WIDE)
    im = ax.pcolormesh(np.arange(K + 1), np.arange(365 + 1), pr4, cmap="viridis",
                       shading="flat")
    cb = fig.colorbar(im, ax=ax, label="电价（元/kWh）")
    for i, t in zip(spec_idx, SPEC_DATES):
        # 四个指定日期在全年中的位置（填报区间下标 + 31 = 全年 0 基天序号）
        y = i + D_REP_FIRST + 0.5
        ax.axhline(y, color="white", linewidth=1.4, linestyle="--")
        ax.annotate(t, (2, y + 3), color="white", fontsize=10)
    ax.set_xticks(np.arange(0, K + 1, 12), ["%d:00" % h for h in range(0, 25, 2)], fontsize=8)
    ax.set_xlabel("时刻（h）")
    ax.set_ylabel("日期（2025 年，1 月 1 日 → 12 月 31 日）")
    ax.set_title("附件4 逐日逐时段电价热力图（365 天 × 144 时段，%.4f–%.4f 元/kWh）\n"
                 "白色虚线为四个指定日期；日内形状逐日相同、跨天差异为乘性日因子"
                 % (pr4.min(), pr4.max()))
    p1 = os.path.join(QDIR, "波动电价三维热力图.png")
    save_figure(fig, p1)
    print("已落盘：%s" % p1)

    # ============ 图 2：固定电价与波动电价结果对比（四个指定日期） ============
    # 固定电价对照：问题2（逐日结果.csv，13 列）与问题3（全分辨率明细.csv，21 列）
    q2 = {c: np.zeros((N_REP, K)) for c in ["计划购电量_kWh"]}
    q2_price = np.zeros((N_REP, K))
    with open(Q2_CSV, encoding="utf-8") as f:
        for idx, r in enumerate(csv.DictReader(f)):
            i, k = divmod(idx, K)
            q2["计划购电量_kWh"][i, k] = float(r["计划购电量_kWh"])
            q2_price[i, k] = float(r["电价_元每kWh"])
    q3y = np.zeros((N_REP, K)); q3x = np.zeros((N_REP, K)); q3p = np.zeros((N_REP, K))
    with open(Q3_CSV, encoding="utf-8") as f:
        for idx, r in enumerate(csv.DictReader(f)):
            i, k = divmod(idx, K)
            q3y[i, k] = float(r["调整购电量_kWh"])
            q3x[i, k] = float(r["计划购电量_kWh"])
            q3p[i, k] = float(r["电价_元每kWh"])
    x_labels = ["计划购电量\n(4-2)", "计划购电量\n(4-3)", "全天购电费\n(问题2/4-2)", "全天总费用\n(问题3/4-3)"]
    fig, axes = plt.subplots(1, 2, figsize=FIGSIZE_TALL)
    xpos = np.arange(len(SPEC_DATES))
    w = 0.38
    # 左：购电量对比
    ax = axes[0]
    v_fix = [float(q2["计划购电量_kWh"][i].sum()) for i in spec_idx]
    v_42 = [float(d42["计划购电量_kWh"][i].sum()) for i in spec_idx]
    ax.bar(xpos - w / 2, v_fix, width=w, color=COLOR_LOAD, alpha=0.9, label="固定电价：问题2 计划量")
    ax.bar(xpos + w / 2, v_42, width=w, color=COLOR_BUY, alpha=0.9, label="波动电价：4-2 计划量")
    for x, a, b in zip(xpos, v_fix, v_42):
        ax.annotate("%.2f%%" % (100.0 * (b - a) / a), (x, max(a, b)), ha="center", va="bottom",
                    fontsize=10, textcoords="offset points", xytext=(0, 4))
    ax.set_xticks(xpos, SPEC_DATES, fontsize=10)
    ax.set_ylabel("全天计划购电量（kWh）")
    ax.set_title("四个指定日期的全天计划购电量\n（标注：波动电价相对固定电价的变化率）")
    ax.legend(fontsize=10)
    # 右：费用对比
    ax2 = axes[1]
    c_fix2 = [float((q2_price[i] * q2["计划购电量_kWh"][i]).sum()) for i in spec_idx]
    c_42 = [float((d42["电价_元每kWh"][i] * d42["计划购电量_kWh"][i]).sum()) for i in spec_idx]
    c_fix3 = [float((q3p[i] * np.minimum(q3x[i], q3y[i])
                     + 0.5 * q3p[i] * np.maximum(q3x[i] - q3y[i], 0)
                     + 1.5 * q3p[i] * np.maximum(q3y[i] - q3x[i], 0)).sum()) for i in spec_idx]
    c_43 = [float((d43["电价_元每kWh"][i] * d43["计划购电量_kWh"][i]
                   + 0.5 * d43["电价_元每kWh"][i] * np.maximum(d43["计划购电量_kWh"][i] - d43["调整购电量_kWh"][i], 0)
                   + 1.5 * d43["电价_元每kWh"][i] * np.maximum(d43["调整购电量_kWh"][i] - d43["计划购电量_kWh"][i], 0)
                   + 5.0 * d43["电价_元每kWh"][i] * d43["紧急购电量_kWh"][i]).sum()) for i in spec_idx]
    b1 = ax2.bar(xpos - w / 2, c_fix2, width=w, color=COLOR_LOAD, alpha=0.9,
                 label="固定电价：问题2 购电费")
    b2 = ax2.bar(xpos + w / 2, c_42, width=w, color=COLOR_BUY, alpha=0.9,
                 label="波动电价：4-2 购电费")
    for x, a, b in zip(xpos, c_fix2, c_42):
        ax2.annotate("%+.1f%%" % (100.0 * (b - a) / a), (x, max(a, b)), ha="center", va="bottom",
                     fontsize=10, textcoords="offset points", xytext=(0, 4))
    ax2.set_xticks(xpos, SPEC_DATES, fontsize=10)
    ax2.set_ylabel("全天购电费（元）")
    ax2.set_title("四个指定日期的全天费用\n（问题2 vs 4-2；问题3 的 J=%.0f/%.0f/%.0f/%.0f vs 4-3 %.0f/%.0f/%.0f/%.0f）"
                  % (*c_fix3, *c_43))
    ax2.legend(fontsize=10)
    fig.suptitle("固定电价（附件1）与波动电价（附件4）结果对比：四个指定日期", fontsize=14)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    p2 = os.path.join(QDIR, "固定电价与波动电价结果对比.png")
    save_figure(fig, p2)
    print("已落盘：%s" % p2)

    # ============ 图 3：波动电价下计划与调整购电量图（四个指定日期） ============
    fig, axes = plt.subplots(2, 2, figsize=FIGSIZE_TALL)
    for ax, i, t in zip(axes.ravel(), spec_idx, SPEC_DATES):
        x_plan = d43["计划购电量_kWh"][i]
        y_adj = d43["调整购电量_kWh"][i]
        ax.step(hours, x_plan / DT_H, where="post", color=COLOR_LOAD, linewidth=1.3,
                label="计划购电量 $x$（折合功率）")
        ax.step(hours, y_adj / DT_H, where="post", color=COLOR_BUY, linewidth=1.3,
                label="最终购电量 $y$（调整后）")
        # 偏差阴影：欠取（y<x，蓝）/ 超用（y>x，红）
        ax.fill_between(hours, x_plan / DT_H, y_adj / DT_H, where=(y_adj < x_plan),
                        step="post", color=COLOR_CHG, alpha=0.30, label="欠取 $x-y$")
        ax.fill_between(hours, x_plan / DT_H, y_adj / DT_H, where=(y_adj >= x_plan),
                        step="post", color=COLOR_DIS, alpha=0.30, label="超用 $y-x$")
        for h0 in (6, 12, 18):
            ax.axvline(h0, color=COLOR_REF, linestyle=":", linewidth=0.9)
        ax.set_xlim(0, 24)
        ax.set_xlabel("时刻（h）")
        ax.set_ylabel("购电量（折合功率，kW）")
        ax.set_title("%s（Σx=%.0f、Σy=%.0f kWh；Σ|y−x|=%.1f kWh）"
                     % (t, x_plan.sum(), y_adj.sum(),
                        np.abs(y_adj - x_plan).sum()))
        ax.legend(fontsize=8, loc="upper left")
    fig.suptitle("波动电价下 4-3 的计划购电量 $x$ 与最终购电量 $y$（四个指定日期；虚线为 6/12/18 调整时刻）",
                 fontsize=14)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    p3 = os.path.join(QDIR, "波动电价下计划与调整购电量图.png")
    save_figure(fig, p3)
    print("已落盘：%s" % p3)

    # ============ 图 4：波动电价下储电量轨迹图（全年热力图 + 四日期轨迹） ============
    from lib.storage import E_MAX, E_MIN
    fig = plt.figure(figsize=FIGSIZE_TALL)
    ax1 = fig.add_subplot(2, 1, 1)
    E_mat = d42["时段末储电量_kWh"]
    im = ax1.pcolormesh(np.arange(K + 1), np.arange(N_REP + 1), E_mat, cmap="magma",
                        shading="flat", vmin=E_MIN, vmax=E_MAX)
    cb = fig.colorbar(im, ax=ax1, label="储电量（kWh）")
    cb.set_ticks([E_MIN, 3000, 6000, 9000, E_MAX])
    for i, t in zip(spec_idx, SPEC_DATES):
        ax1.axhline(i + 0.5, color="cyan", linewidth=1.2, linestyle="--")
        ax1.annotate(t, (2, i + 3), color="cyan", fontsize=10)
    ax1.set_xticks(np.arange(0, K + 1, 12), ["%d:00" % h for h in range(0, 25, 2)], fontsize=8)
    ticks = [0, 58, 120, 181, 242, 303, 333]
    ax1.set_yticks([t + 0.5 for t in ticks], [dates[t] for t in ticks], fontsize=9)
    ax1.set_xlabel("时刻（h）")
    ax1.set_ylabel("日期（2025 年，填报区间）")
    ax1.set_title("4-2 储电量年度轨迹热力图（334 天 × 144 时段；终端余值使 60 天持有过夜）")
    ax2 = fig.add_subplot(2, 1, 2)
    # CSV 的储电量列为各时段**末**（k=1..144 ↔ 钟点 1/6..24 h），共 144 个点
    x_hour = np.arange(1, K + 1) / 6.0
    for i, t in zip(spec_idx, SPEC_DATES):
        ax2.plot(x_hour, d42["时段末储电量_kWh"][i], linewidth=1.4, label=t)
    ax2.axhline(E_MIN, color=COLOR_REF, linewidth=1.2, linestyle="--", label="下限 1200 kWh")
    ax2.axhline(E_MAX, color=COLOR_REF, linewidth=1.2, linestyle=":", label="上限 10800 kWh")
    ax2.set_xlabel("时刻（h）")
    ax2.set_ylabel("储电量（kWh）")
    ax2.set_xlim(0, 24)
    ax2.set_title("四个指定日期的储电量轨迹（4-2 主口径；12.21 由 10200 kWh 高位放至 1200 kWh）")
    ax2.legend(fontsize=10)
    fig.suptitle("波动电价下的储能运行（附件4 + D-10 终端余值）", fontsize=14)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    p4 = os.path.join(QDIR, "波动电价下储电量轨迹图.png")
    save_figure(fig, p4)
    print("已落盘：%s" % p4)
    print("四张主图绘制完成。")


if __name__ == "__main__":
    tee = Tee(os.path.join(QDIR, "绘图运行日志.txt"))
    original_stdout = sys.stdout
    sys.stdout = tee
    try:
        main()
    finally:
        sys.stdout = original_stdout
        tee.close()
    print("绘图运行日志已写入：%s" % os.path.join(QDIR, "绘图运行日志.txt"))
