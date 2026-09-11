"""附件4 电价结构分析（问题 4 §4.4）：逐列比值、日因子、日内价差 vs 跨天价差。

产出（均在 `问题4/` 下）：
  电价日因子分布图.png          逐日乘性因子 a_d 的时间序列 + 直方图（标注均值 1.0 与标准差）
  日内价差与跨天价差对比图.png   日内峰谷价差（逐日均值）vs 跨天同刻价差（逐时刻均值）柱状对比，
                               附往返效率套利门槛 1/η² 参考线
  电价结构统计.xlsx            上述统计的落盘表格（供编写手引用与外部审计）

口径（`口径与假设台账.md` D-09）：附件4 电价 = 附件1 典型日形状 × 逐日乘性因子 + 日内噪声；
  日因子定义 a_d = mean_k(p_k^{4,d} / p_k^{1})（与构思手 probe04/报告04 的"逐日比值的日均值"一致）；
  日内峰谷价差 = 逐日 max/min；跨天同刻价差 = 每个时刻 k 的 max_d p / min_d p。

运行：python 问题4/analyze_price.py
依赖：numpy、openpyxl、matplotlib；lib/ 公共模块（dataio/plotstyle/logio）。
随机性：无。
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import numpy as np
import openpyxl

from lib.dataio import read_attachment1, read_attachment4
from lib.logio import Tee
from lib.storage import ETA
from lib.timegrid import K

QDIR = os.path.join(ROOT, "问题4")
FIGDIR = os.path.join(ROOT, "图片", "问题4")                        # 图片统一目录（2026-09-11 起，用户要求）
os.makedirs(FIGDIR, exist_ok=True)                                # 确保目录存在（重跑时自动建）
XLSX_PATH = os.path.join(QDIR, "电价结构统计.xlsx")
LOG_PATH = os.path.join(QDIR, "电价结构分析运行日志.txt")
FIG_FACTOR = os.path.join(FIGDIR, "电价日因子分布图.png")
FIG_SPREAD = os.path.join(FIGDIR, "日内价差与跨天价差对比图.png")

THR_ARB = 1.0 / ETA ** 2        # 往返效率套利门槛 1/η² = 1.2346


def write_xlsx(path, sheets):
    """把 {工作表名: 行列表} 写成 xlsx（与问题 2/3 的落盘风格一致：数值 4 位小数格式）。

    输入：path，str；sheets，dict[str, list[list]]
    输出：str，写入路径
    """
    wb = openpyxl.Workbook()
    first = True
    for name, rows in sheets.items():
        ws = wb.active if first else wb.create_sheet()
        ws.title = name
        first = False
        for row in rows:
            ws.append(row)
        for r in ws.iter_rows(min_row=2):                    # 数值列统一 4 位小数显示
            for c in r:
                if isinstance(c.value, float):
                    c.number_format = "0.0000"
    wb.save(path)
    return path


def analyze(price1, pr4, dates):
    """计算附件4 电价的全部结构统计量。

    输入：price1，np.ndarray (K,)，附件1 典型日电价，元/kWh
          pr4，np.ndarray (D,K)，附件4 电价，元/kWh
          dates，list[date]，365 天日期
    输出：dict，含逐列比值、日因子、价差、极端点等统计量（元/kWh 或无量纲）
    """
    d_all = pr4.shape[0]
    ratio = pr4 / price1[None, :]                        # 逐点比值 p4/p1，无量纲
    col_mean = ratio.mean(axis=0)                        # 逐列（逐时刻）比值均值（D-09 应为 1.0000）
    a_day = ratio.mean(axis=1)                           # 逐日乘性因子 a_d（与报告04 同定义）
    resid = pr4 - price1[None, :] * a_day[:, None]       # 乘性日因子模型的残差（噪声部分）
    rel_resid = np.abs(resid / pr4).mean()               # 相对残差均值（报告04：0.0957）
    ln_ratio = np.log(ratio)                             # 逐点对数比值（D-09 的 0.20197 口径）

    # 日内峰谷价差：逐日 max/min（无量纲）；跨天同刻价差：每个时刻 k 的 max_d/min_d
    within = pr4.max(axis=1) / pr4.min(axis=1)           # (D,)，逐日峰谷比
    across = pr4.max(axis=0) / pr4.min(axis=0)           # (K,)，同刻跨天比
    # 极端低价点（D-16.3）：< 0.05 元/kWh 的格点
    lo_mask = pr4 < 0.05
    lo_days = sorted({dates[d].isoformat() for d in np.where(lo_mask.any(axis=1))[0]})

    return {
        "ratio_col_mean_min": float(col_mean.min()),      # 逐列比值均值范围（下）
        "ratio_col_mean_max": float(col_mean.max()),      # 逐列比值均值范围（上）
        "a_mean": float(a_day.mean()),                    # 日因子均值（应 1.0000）
        "a_std": float(a_day.std()),                      # 日因子标准差（锚定 0.1401）
        "a_min": float(a_day.min()),                      # 日因子最小（锚定 0.6876）
        "a_max": float(a_day.max()),                      # 日因子最大（锚定 1.2132）
        "a_ratio_extreme": float(a_day.max() / a_day.min()),   # 日因子极值比（锚定 ~1.75）
        "ln_ratio_mean": float(ln_ratio.mean()),          # 逐点对数比值均值（D-09：−0.01843）
        "ln_ratio_std": float(ln_ratio.std()),            # 逐点对数比值标准差（D-09：0.20197）
        "rel_resid": float(rel_resid),                    # 乘性模型相对残差均值（锚定 ~0.096）
        "within_mean": float(within.mean()),              # 日内峰谷比均值（附件4 逐日）
        "within_median": float(np.median(within)),
        "within_p1": float(price1.max() / price1.min()),  # 附件1 典型日日内峰谷比（D-09：3.758）
        "across_mean": float(across.mean()),              # 同刻跨天比均值
        "across_max": float(across.max()),                # 同刻跨天比最大（受极端低价点影响）
        "across_at": int(np.argmax(across)) + 1,          # 跨天比最大的时刻（1 基 k）
        "thr_arb": float(THR_ARB),                        # 套利门槛 1/η²
        "n_low_points": int(lo_mask.sum()),               # 极端低价点个数（锚定 9）
        "low_days": lo_days,                              # 出现极端低价点的日期
        "p_min": float(pr4.min()),                        # 全年最低价（锚定 0.0076）
        "p_max": float(pr4.max()),                        # 全年最高价（锚定 1.7936）
    }


def draw_factor_figure(a_day, dates, st, path):
    """绘制"电价日因子分布图"：时间序列 + 直方图，标注均值与±1σ。

    输入：a_day，np.ndarray (D,)；dates list[date]；st，analyze 的统计 dict；path，str
    输出：str，PNG 路径
    """
    import matplotlib.pyplot as plt
    from lib.plotstyle import (COLOR_PRICE, COLOR_REF, FIGSIZE_TALL, apply_chinese_style,
                               save_figure)
    apply_chinese_style()
    fig, axes = plt.subplots(2, 1, figsize=FIGSIZE_TALL)
    # 上：时间序列（标注 ±1σ 带与四个指定日期）
    ax = axes[0]
    x = np.arange(len(a_day))
    ax.plot(x, a_day, color=COLOR_PRICE, linewidth=0.9, label="逐日乘性因子 $a_d$")
    ax.axhline(1.0, color=COLOR_REF, linewidth=1.2, linestyle="--", label="均值 1.0")
    ax.axhspan(1.0 - st["a_std"], 1.0 + st["a_std"], color=COLOR_REF, alpha=0.15,
               label="均值 ±1 标准差（%.4f）" % st["a_std"])
    ticks = [0, 58, 120, 181, 242, 303, 364]
    ax.set_xticks(ticks, [dates[t].isoformat() for t in ticks], rotation=30, fontsize=9)
    ax.set_xlabel("日期（2025 年）")
    ax.set_ylabel("日因子 $a_d$（无量纲）")
    ax.set_title("附件4 电价逐日乘性因子时间序列（范围 %.4f–%.4f）" % (st["a_min"], st["a_max"]))
    ax.legend(fontsize=10)
    # 下：直方图（标注均值、标准差与极值）
    ax2 = axes[1]
    ax2.hist(a_day, bins=32, color=COLOR_PRICE, alpha=0.85, edgecolor="white")
    ax2.axvline(st["a_mean"], color=COLOR_REF, linewidth=1.4, linestyle="--",
                label="均值 %.4f" % st["a_mean"])
    ax2.axvline(st["a_min"], color="black", linewidth=1.0, linestyle=":",
                label="最小 %.4f" % st["a_min"])
    ax2.axvline(st["a_max"], color="black", linewidth=1.0, linestyle=":",
                label="最大 %.4f" % st["a_max"])
    ax2.set_xlabel("日因子 $a_d$（无量纲）")
    ax2.set_ylabel("天数（天）")
    ax2.set_title("日因子分布：标准差 %.4f；极值比 a_max/a_min = %.4f；乘性模型相对残差 %.4f"
                  % (st["a_std"], st["a_ratio_extreme"], st["rel_resid"]))
    ax2.legend(fontsize=10)
    fig.suptitle("附件4 电价结构：日内形状逐日相同、跨天差异为乘性日因子 + 日内噪声（D-09）",
                 fontsize=14)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    return save_figure(fig, path)


def draw_spread_figure(st, path):
    """绘制"日内价差与跨天价差对比图"：柱状对比 + 套利门槛参考线。

    输入：st，analyze 的统计 dict；path，str
    输出：str，PNG 路径
    """
    import matplotlib.pyplot as plt
    from lib.plotstyle import (COLOR_CHG, COLOR_DIS, COLOR_PRICE, COLOR_REF, FIGSIZE_WIDE,
                               apply_chinese_style, save_figure)
    apply_chinese_style()
    fig, ax = plt.subplots(figsize=FIGSIZE_WIDE)
    names = ["日内峰谷价差\n附件1 典型日", "日内峰谷价差\n附件4 逐日均值",
             "跨天同刻价差\n附件4 均值（同刻跨天）", "跨天同刻价差\n附件4 最大（受极端低价点放大）"]
    vals = [st["within_p1"], st["within_mean"], st["across_mean"], st["across_max"]]
    colors = [COLOR_REF, COLOR_CHG, COLOR_DIS, COLOR_PRICE]
    bars = ax.bar(names, vals, width=0.55, color=colors, alpha=0.9)
    ax.axhline(st["thr_arb"], color="black", linewidth=1.6, linestyle="--",
               label="往返效率套利门槛 $1/\\eta^2$ = %.4f" % st["thr_arb"])
    for b, v in zip(bars, vals):
        ax.annotate("%.4f" % v, (b.get_x() + b.get_width() / 2, v), textcoords="offset points",
                    xytext=(0, 6), ha="center", fontsize=11)
    ax.set_ylabel("价差比（无量纲，峰值价/谷值价）")
    ax.set_title("日内套利空间 vs 跨天套利空间（附件4 电价）\n"
                 "日内均值 %.4f > 门槛 %.4f；跨天均值 %.4f < 门槛、跨天最大值 %.4f 由极端低价日拉高\n"
                 "⇒ 跨天平均套利空间远小于日内（D-09），逐日滚动是合理近似"
                 % (st["within_mean"], st["thr_arb"], st["across_mean"], st["across_max"]))
    ax.legend(fontsize=11)
    return save_figure(fig, path)


def figure_main():
    """主流程：读数据 → 统计 → 落盘 xlsx 与两张图 → 打印统计量。"""
    price1, _, _, _ = read_attachment1()                 # 附件1 典型日电价（144,）
    pr4, dates = read_attachment4()                      # 附件4 电价（365,144）
    print("=" * 78)
    print("附件4 电价结构分析（逐列比值 / 日因子 / 日内与跨天价差）")
    print("=" * 78)
    st = analyze(price1, pr4, dates)
    a_day = (pr4 / price1[None, :]).mean(axis=1)         # 与 analyze 同定义（绘图用）

    print("逐列比值均值范围 = [%.4f, %.4f]（D-09：应为 1.0000 附近）"
          % (st["ratio_col_mean_min"], st["ratio_col_mean_max"]))
    print("日因子 a_d：均值 %.4f、标准差 %.4f、最小 %.4f、最大 %.4f（极值比 %.4f）"
          % (st["a_mean"], st["a_std"], st["a_min"], st["a_max"], st["a_ratio_extreme"]))
    print("逐点对数比值：均值 %.5f、标准差 %.5f（D-09：−0.01843 / 0.20197）"
          % (st["ln_ratio_mean"], st["ln_ratio_std"]))
    print("乘性模型相对残差均值 = %.4f（日内噪声量级）" % st["rel_resid"])
    print("日内峰谷价差：附件1 典型日 %.4f；附件4 逐日均值 %.4f（中位 %.4f）"
          % (st["within_p1"], st["within_mean"], st["within_median"]))
    print("跨天同刻价差：均值 %.4f、最大 %.4f（第 %d 个时段，受极端低价点放大）；套利门槛 1/η² = %.4f"
          % (st["across_mean"], st["across_max"], st["across_at"], st["thr_arb"]))
    print("全年电价范围 %.4f–%.4f 元/kWh；极端低价点（<0.05 元/kWh）= %d 个，出现在 %s（D-16.3 保留）"
          % (st["p_min"], st["p_max"], st["n_low_points"], "、".join(st["low_days"])))

    # 落盘统计表（结构与 xlsx 供编写手引用；图另存）
    rows = [
        ["统计量", "数值", "单位", "说明"],
        ["逐列比值均值范围下界", st["ratio_col_mean_min"], "—", "p4/p1 逐时刻均值的下界（D-09）"],
        ["逐列比值均值范围上界", st["ratio_col_mean_max"], "—", "p4/p1 逐时刻均值的上界（D-09）"],
        ["日因子均值", st["a_mean"], "—", "a_d = mean_k(p4/p1) 的均值"],
        ["日因子标准差", st["a_std"], "—", "跨天波动强度"],
        ["日因子最小值", st["a_min"], "—", "全年最低的日因子"],
        ["日因子最大值", st["a_max"], "—", "全年最高的日因子"],
        ["日因子极值比（max/min）", st["a_ratio_extreme"], "—", "跨天套利空间的上界指示"],
        ["逐点对数比值均值", st["ln_ratio_mean"], "—", "ln(p4/p1) 的均值（D-09）"],
        ["逐点对数比值标准差", st["ln_ratio_std"], "—", "ln(p4/p1) 的标准差（D-09：0.20197）"],
        ["乘性模型相对残差均值", st["rel_resid"], "—", "日内噪声量级（|残差/p4| 均值）"],
        ["附件1 日内峰谷价差", st["within_p1"], "—", "典型日 max/min（D-09：3.758）"],
        ["附件4 日内峰谷价差（均值）", st["within_mean"], "—", "逐日 max/min 的均值"],
        ["跨天同刻价差（均值）", st["across_mean"], "—", "每个时刻跨天 max/min 的均值"],
        ["跨天同刻价差（最大）", st["across_max"], "—", "对应第 %d 个时段（由极端低价点放大）" % st["across_at"]],
        ["套利门槛 1/η²", st["thr_arb"], "—", "往返效率 0.81 对应的最低有利价差比"],
        ["全年最低价", st["p_min"], "元/kWh", "极端低价点（D-16.3 保留）"],
        ["全年最高价", st["p_max"], "元/kWh", "全年最大值"],
        ["极端低价点个数（<0.05）", st["n_low_points"], "个", "出现日期：" + "、".join(st["low_days"])],
    ]
    write_xlsx(XLSX_PATH, {"电价结构统计": rows})
    print("已落盘：%s" % XLSX_PATH)
    print("已落盘：%s" % draw_factor_figure(a_day, dates, st, FIG_FACTOR))
    print("已落盘：%s" % draw_spread_figure(st, FIG_SPREAD))
    print("=" * 78)
    return st


if __name__ == "__main__":
    tee = Tee(LOG_PATH)
    original_stdout = sys.stdout
    sys.stdout = tee
    try:
        figure_main()
    finally:
        sys.stdout = original_stdout
        tee.close()
    print("运行日志已写入：%s" % LOG_PATH)
