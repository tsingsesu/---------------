"""绘图样式（全工作区共用）：中文字体、统一配色与落盘规范。

规范来源：`roles/编程手.md` §6 —— 图片必须中文正常显示、有标题/坐标轴标签与单位/图例，
分辨率不低于 300 dpi，文件名用中文且自解释。本机 matplotlib 已实测可用 SimHei。
"""

import matplotlib
matplotlib.use("Agg")            # 无界面后端，保证脚本在命令行下也能出图
import matplotlib.pyplot as plt

# 图幅与分辨率：论文插图统一 12×6 英寸、300 dpi 以上
FIGSIZE_WIDE = (12.0, 6.0)       # 单面板宽图，英寸
FIGSIZE_TALL = (12.0, 9.0)       # 多面板图，英寸
DPI = 300                        # 落盘分辨率，dpi

# 统一配色（中文图例用，全工作区一致）
COLOR_PRICE = "#B22222"          # 电价曲线：深红
COLOR_LOAD = "#1F4E79"           # 负载曲线：深蓝
COLOR_PV = "#E8A33D"             # 光伏曲线：橙
COLOR_BUY = "#2E7D32"            # 计划购电量：绿
COLOR_CHG = "#0072B2"            # 充电：蓝
COLOR_DIS = "#D55E00"            # 放电：橙红
COLOR_SOC = "#5B2C6F"            # 储电量：紫
COLOR_REF = "#808080"            # 参考线：灰


def apply_chinese_style():
    """设置全局绘图样式，使中文与负号正常显示。

    输入：无
    输出：无（直接修改 matplotlib 全局 rcParams）
    """
    # 指定中文字体候选，避免出现方块乱码（本机实测 SimHei 可用）
    plt.rcParams["font.sans-serif"] = ["SimHei", "Microsoft YaHei"]
    # 解决坐标轴负号显示为方块的问题
    plt.rcParams["axes.unicode_minus"] = False
    # 统一的字号体系，保证论文缩放后仍清晰
    plt.rcParams["font.size"] = 12
    plt.rcParams["axes.titlesize"] = 15
    plt.rcParams["axes.labelsize"] = 13
    plt.rcParams["legend.fontsize"] = 11
    # 坐标轴刻度朝内、带次刻度，接近论文插图习惯
    plt.rcParams["xtick.direction"] = "in"
    plt.rcParams["ytick.direction"] = "in"
    plt.rcParams["axes.grid"] = True
    plt.rcParams["grid.alpha"] = 0.3
    plt.rcParams["grid.linestyle"] = "--"
    plt.rcParams["figure.autolayout"] = False


def save_figure(fig, path):
    """按统一规范保存图片。

    输入：fig，matplotlib Figure；path，str，目标 png 绝对路径
    输出：str，实际写入的路径
    """
    # bbox_inches='tight' 防止中文标签或图例被裁掉；dpi 固定为 300
    fig.savefig(path, dpi=DPI, bbox_inches="tight", facecolor="white")
    plt.close(fig)               # 关闭图对象，避免批量出图时内存堆积
    return path


def hour_axis_ticks():
    """返回 24 小时图的横轴刻度位置与标签（横轴以"小时"为单位）。

    输入：无
    输出：(ticks, labels) —— ticks list[float] 小时刻度，labels list[str]
    """
    ticks = list(range(0, 25, 2))          # 每 2 小时一个刻度
    labels = ["%d:00" % h for h in ticks]  # 标签形如 0:00、2:00 ...
    return ticks, labels
