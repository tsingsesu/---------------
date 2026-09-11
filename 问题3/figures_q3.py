r"""问题 3 图表生成脚本：从**落盘交付物**（全分辨率明细.csv）出图，不重跑求解。

产出（均在 `问题3/`，中文名、≥300 dpi；交付清单 §四）：
  * `预报与实际光伏对比图.png`：四个指定日期，各发布时刻整点预报 + 分解后 10 分钟预报 + 实际；
  * `预报误差分布图.png`：预报误差（预报−实际）按发布时刻与提前期的分布（小提琴图）；
  * `计划与调整购电量对比图.png`：四个指定日期的计划量 vs 最终量，标出欠取/超用方向；
  * `紧急购电时段分布图.png`：全年紧急购电的（时段 × 日期）热力图与触发计数；
  * `储电量轨迹与预报误差叠加.png`：四个指定日期的储电量轨迹 + 预报误差（双轴）。

为什么从 CSV 出图：图的每个数据点都可在 `全分辨率明细.csv` 中逐条查到（可外部审计），
且图表与交付结果文件严格同源，避免"图与表不一致"。

运行：python 问题3/figures_q3.py
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

from lib.dataio import read_attachment1
from lib.forecast import load_forecast
from lib.plotstyle import (COLOR_BUY, COLOR_CHG, COLOR_DIS, COLOR_LOAD, COLOR_PRICE,
                           COLOR_PV, COLOR_REF, COLOR_SOC, FIGSIZE_TALL, FIGSIZE_WIDE,
                           apply_chinese_style, hour_axis_ticks, save_figure)
from lib.timegrid import K, hour_block_to_k, four_hour_blocks

QDIR = os.path.join(ROOT, "问题3")
FIGDIR = os.path.join(ROOT, "图片", "问题3")                        # 图片统一目录（2026-09-11 起，用户要求）
os.makedirs(FIGDIR, exist_ok=True)                                # 确保目录存在（重跑时自动建）
AUDIT_CSV = os.path.join(QDIR, "全分辨率明细.csv")
D_REP_FIRST = 31
N_REP = 334
SPEC_DATES = ("2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21")
SPEC_HOURS = (10, 12, 14, 16, 18, 20)


def load_audit():
    """读全分辨率明细 CSV 为字典（列名 -> (334,144) 数组）。"""
    with open(AUDIT_CSV, encoding="utf-8", newline="") as f:
        rows = list(csv.reader(f))
    hdr = rows[0]
    out = {}
    for j, name in enumerate(hdr):
        if j >= 5:
            out[name] = np.array([float(r[j]) for r in rows[1:]]).reshape(N_REP, K)
        else:
            out[name] = np.array([r[j] for r in rows[1:]])
    return out


def spec_index(t):
    """指定日期 -> 填报区间内的 0 基下标。"""
    d0 = (_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days
    return d0 - D_REP_FIRST


def fig_forecast_vs_actual(d, pv_fc144):
    """图 1：预报与实际光伏对比（四个指定日期）。"""
    import matplotlib.pyplot as plt
    apply_chinese_style()
    hours = np.arange(1, K + 1) / 6.0 - 1.0 / 12.0        # 每段中心钟点
    node_hours = np.arange(1, 25)                          # 整点
    fig, axes = plt.subplots(2, 2, figsize=FIGSIZE_TALL)
    for ax, t in zip(axes.ravel(), SPEC_DATES):
        i = spec_index(t)
        d0 = i + D_REP_FIRST
        # 分解后的 0:00 预报（10 分钟）
        ax.plot(hours, pv_fc144[d0, 0], color=COLOR_CHG, linewidth=1.4, alpha=0.8,
                label="0:00 预报（分解到 10 分钟）")
        # 18:00 预报（分解后）
        ax.plot(hours, pv_fc144[d0, 3], color=COLOR_DIS, linewidth=1.1, alpha=0.8,
                linestyle="-.", label="18:00 预报（分解后）")
        # 原始整点预报（0:00 发布）：节点圆点
        ax.plot(node_hours, pv_fc144[d0, 0, 5::6], "o", color=COLOR_CHG, markersize=4,
                markerfacecolor="white", label="0:00 整点预报值（附件3 原始）")
        # 实际
        ax.plot(hours, d["光伏实际_kW"][i], color=COLOR_PV, linewidth=1.6, label="实际光伏（附件2）")
        ax.axhline(0, color=COLOR_REF, linewidth=0.8)
        for h in SPEC_HOURS:
            ax.axvline(h, color=COLOR_REF, linestyle=":", linewidth=0.6, alpha=0.5)
        tick_val = (np.abs(pv_fc144[d0, 0] - d["光伏实际_kW"][i]) * (hours > 6) * (hours < 19)).sum()
        ax.set_xlabel("时刻（h）")
        ax.set_ylabel("功率（kW）")
        ax.set_title("%s（白天绝对误差之和 %.0f kW·段）" % (t, tick_val))
        ax.legend(fontsize=8, loc="upper left")
        ax.set_xlim(0, 24)
    fig.suptitle("问题 3 预报与实际光伏对比：整点预报、分解后的 10 分钟预报与实际（四个指定日期）",
                 fontsize=15)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    save_figure(fig, os.path.join(FIGDIR, "预报与实际光伏对比图.png"))


def fig_error_distribution(d, pv_fc144):
    """图 2：预报误差分布（按发布时刻与提前期分组的小提琴图 + 直方图）。"""
    import matplotlib.pyplot as plt
    apply_chinese_style()
    from lib.dataio import read_attachment2
    load, pv_actual, dates = read_attachment2()
    fig, axes = plt.subplots(1, 2, figsize=FIGSIZE_WIDE)
    # 面板 (a)：四个发布时刻的误差分布（白天掩码：预报 > 100 kW）
    errs = []
    labels = []
    for ti, tau in enumerate((0, 6, 12, 18)):
        e = pv_fc144[:, ti, :] - pv_actual
        mask = pv_fc144[:, ti, :] > 100
        errs.append(e[mask])
        labels.append("τ=%d:00" % tau)
    parts = axes[0].violinplot(errs, showmedians=True, widths=0.8)
    for pc, color in zip(parts["bodies"], (COLOR_CHG, COLOR_BUY, COLOR_DIS, COLOR_PRICE)):
        pc.set_facecolor(color); pc.set_alpha(0.55)
    axes[0].set_xticks(range(1, 5), labels)
    axes[0].axhline(0, color=COLOR_REF, linewidth=1.0, linestyle="--")
    axes[0].set_ylabel("预报误差（预报 − 实际，kW）")
    axes[0].set_xlabel("预报发布时刻")
    axes[0].set_title("(a) 按发布时刻（白天样本；中位线接近 0，尾部误差大）")
    # 面板 (b)：按提前期 m 的 RMSE（0:00 预报）
    rmse_m = []
    for m in range(1, 25):
        e = pv_fc144[:, 0, 6 * m - 1] - pv_actual[:, 6 * m - 1]
        rmse_m.append(float(np.sqrt(np.mean(e ** 2))))
    axes[1].bar(np.arange(1, 25), rmse_m, color=COLOR_PV, alpha=0.85)
    axes[1].set_xlabel("预报提前期 m（小时）")
    axes[1].set_ylabel("整点 RMSE（kW）")
    axes[1].set_title("(b) 0:00 预报的整点 RMSE 随提前期变化\n（白天时段误差大，晨昏与时滞相关）")
    fig.suptitle("问题 3 预报误差分布（预报 − 实际）：按发布时刻与提前期", fontsize=15)
    fig.tight_layout(rect=(0, 0, 1, 0.93))
    save_figure(fig, os.path.join(FIGDIR, "预报误差分布图.png"))


def fig_plan_vs_adjust(d):
    """图 3：计划与调整购电量对比（四个指定日期；标出欠取/超用方向）。"""
    import matplotlib.pyplot as plt
    apply_chinese_style()
    hours = np.arange(1, K + 1) / 6.0 - 1.0 / 12.0
    fig, axes = plt.subplots(2, 2, figsize=FIGSIZE_TALL)
    for ax, t in zip(axes.ravel(), SPEC_DATES):
        i = spec_index(t)
        x = d["计划购电量_kWh"][i]; y = d["调整购电量_kWh"][i]
        du = d["欠取量_kWh"][i]; do = d["超用量_kWh"][i]
        ax.step(hours, x / (1 / 6.0), where="mid", color=COLOR_BUY, linewidth=1.4,
                label="计划量 x（0:00 制定）")
        ax.step(hours, y / (1 / 6.0), where="mid", color=COLOR_PRICE, linewidth=1.2,
                linestyle="--", label="最终量 y（调整后生效）")
        ax.fill_between(hours, x / (1 / 6.0), y / (1 / 6.0), where=(du > 1e-6),
                        color=COLOR_LOAD, alpha=0.35, label="欠取（按 0.5p）")
        ax.fill_between(hours, x / (1 / 6.0), y / (1 / 6.0), where=(do > 1e-6),
                        color=COLOR_DIS, alpha=0.35, label="超用（按 1.5p）")
        ax.axhline(0, color=COLOR_REF, linewidth=0.8)
        # 标出四个决策时刻的分界
        for hh in (6, 12, 18):
            ax.axvline(hh, color=COLOR_REF, linestyle=":", linewidth=0.8, alpha=0.7)
        ax.set_xlabel("时刻（h）")
        ax.set_ylabel("购电量折合功率（kW）")
        ax.set_title("%s（Σ欠取 %.1f / Σ超用 %.1f kWh）" % (t, du.sum(), do.sum()))
        ax.legend(fontsize=8, loc="upper left")
        ax.set_xlim(0, 24)
    fig.suptitle("问题 3 计划与调整购电量对比（0:00 计划 vs 最终生效；虚线为 6:00/12:00/18:00 分界）",
                 fontsize=15)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    save_figure(fig, os.path.join(FIGDIR, "计划与调整购电量对比图.png"))


def fig_emergency_map(d):
    """图 4：紧急购电的（时段 × 日期）热力图 + 逐小时触发计数。"""
    import matplotlib.pyplot as plt
    apply_chinese_style()
    r = d["紧急购电量_kWh"]                                # (334,144)
    r_hour = r.reshape(N_REP, 24, 6).sum(axis=2)           # 按小时聚合（kWh）
    fig, axes = plt.subplots(2, 1, figsize=FIGSIZE_TALL, gridspec_kw={"height_ratios": [3, 1]})
    ax = axes[0]
    im = ax.pcolormesh(np.arange(25), np.arange(N_REP + 1), r_hour, cmap="YlOrRd",
                       shading="flat")
    cb = fig.colorbar(im, ax=ax, label="紧急购电量（kWh/小时）")
    ticks = [0, 58, 120, 181, 242, 303, 333]
    ax.set_yticks([t + 0.5 for t in ticks],
                  ["%s" % (_dt.date(2025, 2, 1) + _dt.timedelta(days=int(t))).isoformat()
                   for t in ticks], fontsize=9)
    ax.set_xticks(np.arange(0, 25, 2), ["%d:00" % h for h in range(0, 25, 2)], fontsize=8)
    ax.set_xlabel("时刻（h）")
    ax.set_ylabel("日期（2025 年）")
    ax.set_title("问题 3 全年紧急购电（时段 × 日期）热力图：缺口集中于昼间光伏波动时段与晚峰")
    # 下：逐小时触发计数与平均电量
    ax2 = axes[1]
    cnt_hour = (r_hour > 1e-6).sum(axis=0)                 # 该小时在 334 天中被触发的天数
    mean_hour = r_hour.mean(axis=0)
    ax2.bar(np.arange(24), cnt_hour, color=COLOR_PRICE, alpha=0.7, label="触发天数（334 天中）")
    axb = ax2.twinx()
    axb.plot(np.arange(24), mean_hour, "o-", color=COLOR_CHG, linewidth=1.4, markersize=5,
             label="日均紧急购电量（kWh）")
    ax2.set_xticks(np.arange(0, 24, 2), ["%d:00" % h for h in range(0, 24, 2)])
    ax2.set_xlabel("时刻（h）")
    ax2.set_ylabel("触发天数（天）", color=COLOR_PRICE)
    axb.set_ylabel("日均紧急购电量（kWh）", color=COLOR_CHG)
    h1, l1 = ax2.get_legend_handles_labels(); h2, l2 = axb.get_legend_handles_labels()
    ax2.legend(h1 + h2, l1 + l2, loc="upper left", fontsize=9)
    fig.tight_layout()
    save_figure(fig, os.path.join(FIGDIR, "紧急购电时段分布图.png"))


def fig_soc_vs_error(d, pv_fc144):
    """图 5：储电量轨迹与预报误差叠加（四个指定日期，双轴）。"""
    import matplotlib.pyplot as plt
    apply_chinese_style()
    hours = np.arange(0, K + 1) / 6.0                      # 储电量轨迹的钟点（含 0:00 与 24:00）
    hours_mid = np.arange(1, K + 1) / 6.0 - 1.0 / 12.0
    fig, axes = plt.subplots(2, 2, figsize=FIGSIZE_TALL)
    for ax, t in zip(axes.ravel(), SPEC_DATES):
        i = spec_index(t)
        d0 = i + D_REP_FIRST
        # 储电量：由当日 0:00 值与逐时段末储电量拼出完整轨迹
        e_prev = float(d["时段末储电量_kWh"][i - 1][K - 1]) if i > 0 else 1200.0
        e_traj = np.concatenate([[e_prev], d["时段末储电量_kWh"][i]])
        ax.plot(hours, e_traj, color=COLOR_SOC, linewidth=1.6, label="储电量 E（kWh）")
        ax.axhline(1200, color=COLOR_REF, linestyle=":", linewidth=0.9, label="下限 1200")
        ax.axhline(10800, color=COLOR_REF, linestyle=":", linewidth=0.9, label="上限 10800")
        ax.set_xlabel("时刻（h）")
        ax.set_ylabel("储电量（kWh）", color=COLOR_SOC)
        ax.set_ylim(0, 12000)
        axb = ax.twinx()
        err = pv_fc144[d0, 0] - d["光伏实际_kW"][i]
        axb.bar(hours_mid, err, width=1 / 6.0 * 0.9, color=COLOR_PRICE, alpha=0.38,
                label="预报误差（预报−实际，kW）")
        axb.set_ylabel("预报误差（kW）", color=COLOR_PRICE)
        axb.axhline(0, color=COLOR_REF, linewidth=0.7)
        ax.set_title("%s（误差正=高报；储能按预报排程时的高报风险直接反映在缺口上）" % t)
        h1, l1 = ax.get_legend_handles_labels(); h2, l2 = axb.get_legend_handles_labels()
        ax.legend(h1 + h2, l1 + l2, loc="upper left", fontsize=8)
        ax.set_xlim(0, 24)
    fig.suptitle("问题 3 储电量轨迹与预报误差叠加（四个指定日期）", fontsize=15)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    save_figure(fig, os.path.join(FIGDIR, "储电量轨迹与预报误差叠加.png"))


def main():
    """生成全部图表。"""
    print("=" * 78)
    print("问题 3 图表生成（从落盘的全分辨率明细.csv 出图）")
    print("=" * 78)
    d = load_audit()
    _, _, _, _ = read_attachment1()
    _, pv_fc144, _, _ = load_forecast(method="linear")
    fig_forecast_vs_actual(d, pv_fc144); print("已出图：预报与实际光伏对比图.png")
    fig_error_distribution(d, pv_fc144); print("已出图：预报误差分布图.png")
    fig_plan_vs_adjust(d); print("已出图：计划与调整购电量对比图.png")
    fig_emergency_map(d); print("已出图：紧急购电时段分布图.png")
    fig_soc_vs_error(d, pv_fc144); print("已出图：储电量轨迹与预报误差叠加.png")
    print("=" * 78)


if __name__ == "__main__":
    main()
