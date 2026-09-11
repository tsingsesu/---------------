"""问题 2 主脚本：全年逐日滚动线性规划（LP），读法① 完全信息。

流程：读附件1 电价 + 附件2 负载与光伏 → 从 2025-01-01（E_0=6000 kWh）逐日滚动求解 365 天
      → 落盘 `逐日结果.csv`（全分辨率审计明细）与 `result2.xlsx` 三张表 → 打印关键数值。

口径（全部来自 `口径与假设台账.md`，不得自行更改；详见文件头注释与 §口径说明）：
  * D-01 填法 Y-轮转：模板第 i 列（0 基）装当天第 (i+1)%144+1 个时段；第 144 列在时间上
         是当天最早的一段；一切聚合与绘图按真实时段序 k=1..144；
  * D-04 终端储电量自由 + 跨日传递 E_{144}^{d}=E_{0}^{d+1}（E_0^{1}=6000）；每日 24:00
         被压到下限 1200 kWh 是价格结构下的经济结论（须在论文中解释）；
  * D-06 读法① 完全信息：0:00 已知当天负载与光伏，最优计划恰好覆盖负载，r ≡ 0；
  * D-11 从 2025-01-01 仿真、1 月预热、结果文件只填 2025-02-01 至 12-31 共 334 天；
  * D-15 统一保留 4 位小数；D-16.8/9/10 模板扩表、紧急购电填法与 `7:0-7:10` 笔误照抄。

运行：python 问题2/run_q2.py
依赖：numpy、scipy、openpyxl（均本机已装）；lib/ 公共模块（含 lib/solve_day.py）。
随机性：无（全流程确定性 LP，不需要随机种子）。
"""

import csv
import os
import shutil
import sys

# 把工作区根目录加入模块搜索路径，保证在任意工作目录下都能 from lib... import ...
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import numpy as np
import openpyxl

from lib.dataio import attach_path, read_attachment1, read_attachment2, template_rows_for_plan
from lib.run_days import solve_rolling
from lib.storage import E_INIT, E_MAX, E_MIN, ETA, P_MAX
from lib.timegrid import (DT_H, K, four_hour_blocks, hour_block_to_k,
                          hour_block_to_template_row, k_to_label)

QDIR = os.path.join(ROOT, "问题2")                                   # 问题 2 交付目录
TEMPLATE_XLSX = attach_path("附件5", "result2.xlsx")                 # 题目模板（只读）
RESULT_XLSX = os.path.join(QDIR, "result2.xlsx")                     # 交付结果文件
AUDIT_CSV = os.path.join(QDIR, "逐日结果.csv")                        # 全分辨率审计明细（334 天 × 全天 × 144 时段）
LOG_PATH = os.path.join(QDIR, "主模型运行日志.txt")                   # 运行日志（数值证据）

ND = 4                                                               # 统一小数位数（D-15）
D_REP_FIRST = 31                                                     # 填报区间首日：2025-02-01（0 基 31）
N_REP = 334                                                          # 填报区间天数（2025-02-01 至 12-31）
E_MAT_MAX = 10800.0                                                  # 储电量热力图的显示上限（= 储能上限 Ē）
SPEC_HOURS = (10, 12, 14, 16, 18, 20)                                # 题目表 1 的六个指定时段
SPEC_DATES = ("2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21")  # 题目表 1/2/3 的四个指定日期

# 构思手独立锚定值（`_phase0/报告17_问题2量级体检.txt`）——只用于日志对照打印，不参与计算
ANCHOR = {
    "cost_rep": 12254765.7161,        # 填报区间总计划购电费，元
    "cost_day_mean": 36690.9153,      # 日均费用，元
    "x_rep": 20189815.7979,           # 填报区间总购电量，kWh
    "u_rep": 6635686.0657,            # 填报区间总充电量，kWh
    "v_rep": 5374905.7132,            # 填报区间总放电量，kWh
    "joint_lb": 12227243.6427,        # 填报区间全年联合 LP 下界（报告18），元
}

# 逐日结果 CSV 的列定义（全分辨率、自解释，供外部审计逐条核验）
CSV_HEADER = ["日期", "天序号", "时段序号", "时段起", "时段止", "电价_元每kWh",
              "负载_kW", "光伏_kW", "计划购电量_kWh", "充电量_kWh", "放电量_kWh",
              "紧急购电量_kWh", "时段末储电量_kWh"]


class Tee:
    """把控制台输出同时写入日志文件，保证"落盘"不依赖人工复制。

    输入：path，str，日志文件路径
    输出：无（作为上下文管理器使用）
    """

    def __init__(self, path):
        # 控制台默认可能是 GBK 代码页，先切到 utf-8，避免中文与数学符号报编码错
        self.stdout = sys.stdout
        try:
            self.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
        # 以 utf-8 打开日志文件（python 文件本身不写编码注释，编码在 open 时指定）
        self.file = open(path, "w", encoding="utf-8")

    def write(self, text):
        # 同时写日志与回显，二者内容完全一致
        self.stdout.write(text)
        self.file.write(text)

    def flush(self):
        # 刷新两个通道，避免异常退出时日志残缺
        self.stdout.flush()
        self.file.flush()

    def close(self):
        # 关闭日志文件句柄
        self.file.close()


def r4(values, tol=1e-6):
    """把数组/标量先"容差归零"再四舍五入到 4 位小数（D-15 的统一精度口径）。

    输入：values，array_like，任意数值
          tol，float，归零阈值；|value|<tol 时归零
    输出：np.ndarray 或 float，保留 4 位小数
    说明：LP 的数值残差常在 1e-13 量级，且退化最优解中可能出现 -1e-14 的"负零"，
          直接格式化会得到 '-0.0000' 这种会被外部审计误读的显示；物理上这些量为 0，
          故在**输出格式化的最后一步**统一按 tol=1e-9 归零（计算本身不受影响）。
    """
    arr = np.asarray(values, dtype=float)
    arr = np.where(np.abs(arr) < tol, 0.0, arr)        # 容差量级一律记 0
    return np.round(arr, ND)


def set_decimal_format(cell):
    """仅当单元格原格式为 General 时才设为 4 位小数，避免覆盖模板自带的数字格式。

    输入：cell，openpyxl 单元格对象
    输出：无（就地修改 cell.number_format）
    """
    # 模板中日期/时刻/文本单元格都有各自格式，只对空白的数值格改格式
    if cell.number_format in ("General", "general"):
        cell.number_format = "0.0000"


def write_audit_csv(path, dates, price, load, pv_actual, roll):
    """落盘全分辨率审计明细 CSV：填报区间 334 天 × 144 时段逐条记录。

    输入：path，str，目标 csv 路径
          dates，list[datetime.date]，365 天日期
          price，np.ndarray (144,)，电价，元/kWh
          load / pv_actual，np.ndarray (365,144)，负载/光伏实际功率，kW
          roll，dict，solve_rolling 的返回值（填报区间部分）
    输出：str，写入的路径
    """
    lines = [",".join(CSV_HEADER)]
    # 逐日、逐时段按真实时间顺序拼行：时段起止标签由 timegrid 统一生成
    for i, d in enumerate(roll["d_index"]):
        for k in range(1, K + 1):
            start_label, end_label = k_to_label(k)          # 该时段的（起, 止）钟点
            row = [
                dates[d].isoformat(),                       # 日期，如 2025-02-01
                str(d + 1),                                 # 天序号（1 基，2025-01-01 记 1）
                str(k),                                     # 时段序号 1..144（真实时间序）
                start_label, end_label,                     # 时段起 / 止
                "%.4f" % price[k - 1],                      # 电价，元/kWh
                "%.4f" % load[d, k - 1],                    # 负载功率，kW
                "%.4f" % pv_actual[d, k - 1],               # 光伏功率，kW
                "%.4f" % r4(roll["x_plan"][i, k - 1]),      # 计划购电量，kWh
                "%.4f" % r4(roll["u_chg"][i, k - 1]),       # 充电量，kWh
                "%.4f" % r4(roll["v_dis"][i, k - 1]),       # 放电量，kWh
                "%.4f" % r4(roll["r_emg"][i, k - 1]),       # 紧急购电量，kWh
                "%.4f" % r4(roll["E_soc"][i, k]),           # 时段 k **末**储电量，kWh
            ]
            lines.append(",".join(row))
    # 写文件时显式指定 utf-8 与换行，保证 Excel 与 pandas 都能正确读取
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("\n".join(lines) + "\n")
    return path


def _emergency_text(r_day):
    """把一天的紧急购电量聚合成连续时段文本（按表 4 示例的写法）。

    输入：r_day，np.ndarray (K,)，当日的逐时段紧急购电量，kWh
    输出：(time_text, energy_text)，均为 str：
          time_text  形如 '13:00-13:30 14:40-15:50'（连续时段已合并，空格分隔）；
          energy_text 形如 '100 400'（与 time_text 的段一一对应，按时序）。
          无紧急购电时返回 ('', '0')（交付清单 §一 规定：时间段列留空、购电量列填 0）。
    """
    active = np.asarray(r_day, dtype=float) > 1e-6      # 布尔掩码：该时段是否需要紧急购电
    if not active.any():
        return "", "0"
    idx = np.where(active)[0]                           # 需要紧急购电的时段下标（0 基）
    segments = []                                       # 连续段列表，每项为 [起下标, 止下标]
    seg_start = idx[0]
    for j in range(1, idx.size + 1):
        # 下标不连续（或已到末尾）说明当前连续段结束
        if j == idx.size or idx[j] != idx[j - 1] + 1:
            segments.append((seg_start, idx[j - 1]))
            if j < idx.size:
                seg_start = idx[j]
    time_parts, energy_parts = [], []
    for a, b in segments:
        # 连续区段 [a, b] 的真实起止标签：a 段起点 → b 段终点
        start_label = k_to_label(a + 1)[0]              # 第 a+1 个时段的起点钟点
        end_label = k_to_label(b + 1)[1]                # 第 b+1 个时段的终点钟点
        time_parts.append("%s-%s" % (start_label, end_label))
        # 段内电量求和（保留 4 位小数后转字符串，整数则去掉尾随 0）
        seg_energy = r4(float(np.sum(r_day[a:b + 1])))
        energy_parts.append(("%.4f" % seg_energy).rstrip("0").rstrip("."))
    return " ".join(time_parts), " ".join(energy_parts)


def fill_result_xlsx(path, dates, roll, price):
    """把滚动解按模板规格填入 result2.xlsx（先复制模板，绝不改写附件）。

    输入：path，str，目标 result2.xlsx 路径
          dates，list[datetime.date]，365 天日期
          roll，dict，solve_rolling 的返回值（填报区间部分）
          price，np.ndarray (144,)，电价，元/kWh
    输出：dict，写入摘要（四个指定日期表 1 值、紧急购电触发统计等）
    """
    # 先复制模板：附件目录只读，一切填写都在问题2/ 下的副本上完成（D-16.6）
    shutil.copyfile(TEMPLATE_XLSX, path)
    wb = openpyxl.load_workbook(path)                   # 载入副本（保留模板全部标签）

    # ---------------- 工作表 1：计划购电量（334 天 × 144 时段 + 全天购电量/购电费） ----------------
    ws = wb["计划购电量"]
    for i, d in enumerate(roll["d_index"]):
        row = 2 + i                                     # 第 1 行是表头，数据从第 2 行起
        assert ws.cell(row=row, column=1).value.date() == dates[d], "模板日期列与填报区间不一致"
        # 填法 Y-轮转：模板第 j 列装当天第 (j+1)%144+1 个时段（dataio 已实现该重排）
        plan_in_template_order = template_rows_for_plan(roll["x_plan"][i])
        for j in range(K):
            cell = ws.cell(row=row, column=2 + j)       # 第 2..145 列是 144 个时段
            cell.value = float(r4(plan_in_template_order[j]))
            set_decimal_format(cell)                    # 模板该区域原为 General，统一设 4 位小数
        # 全天购电量 = 当天 144 个时段之和（kWh）；全天购电费 = 计划购电费 + 紧急购电费（元）
        c_x = ws.cell(row=row, column=146); c_x.value = float(r4(roll["x_plan"][i].sum()))
        set_decimal_format(c_x)
        c_j = ws.cell(row=row, column=147); c_j.value = float(r4(roll["cost_day"][i]))
        set_decimal_format(c_j)

    # ---------------- 工作表 2：充放电量（334 天 × 6 个 4 小时块 = 2004 行） ----------------
    ws2 = wb["充放电量"]
    blocks = four_hour_blocks()                         # 六个 4 小时块（真实时间顺序）
    assert len(blocks) == 6, "表 2 必须是 6 个 4 小时块"
    for i, d in enumerate(roll["d_index"]):
        for b, (start_label, end_label, k_first, k_last) in enumerate(blocks):
            row = 2 + i * 6 + b                         # 每天 6 行，从第 2 行起连续排布
            # 日期只在该天第 1 行填写（与模板示例一致：后续行留空）
            if b == 0:
                date_cell = ws2.cell(row=row, column=1)
                date_cell.value = wb_date(dates[d], ws2.cell(row=2, column=1))   # 与模板同型日期对象
                date_cell.number_format = ws2.cell(row=2, column=1).number_format
            # 时间段标签：延用模板第 2..7 行的六个块标签（从示例行取，保证逐字一致）
            ws2.cell(row=row, column=2).value = ws2.cell(row=2 + b, column=2).value
            u_block = float(np.sum(roll["u_chg"][i, k_first - 1:k_last]))    # 该块充电量，kWh
            v_block = float(np.sum(roll["v_dis"][i, k_first - 1:k_last]))    # 该块放电量，kWh
            c_u = ws2.cell(row=row, column=3); c_u.value = float(r4(u_block)); set_decimal_format(c_u)
            c_v = ws2.cell(row=row, column=4); c_v.value = float(r4(v_block)); set_decimal_format(c_v)
            # 时刻列：每天第 1 行填 0:00、第 2 行填 24:00（其余留空），与模板一致
            c_t = ws2.cell(row=row, column=5)
            if b == 0:
                c_t.value = ws2.cell(row=2, column=5).value          # datetime.time(0, 0)
            elif b == 1:
                c_t.value = ws2.cell(row=3, column=5).value          # 字符串 '24:00'
            # 储电量列：第 1 行填 E_0、第 2 行填 E_144（其余留空）
            c_s = ws2.cell(row=row, column=6)
            if b == 0:
                c_s.value = float(r4(roll["E_soc"][i, 0])); set_decimal_format(c_s)
            elif b == 1:
                c_s.value = float(r4(roll["E_soc"][i, K])); set_decimal_format(c_s)

    # ---------------- 工作表 3：紧急购电量（334 天，每天一行） ----------------
    ws3 = wb["紧急购电量"]
    n_trigger = 0                                       # 触发紧急购电的天数（读法① 应为 0）
    spec_emg = {}                                       # 四个指定日期的（时间段, 电量）文本
    for i, d in enumerate(roll["d_index"]):
        row = 2 + i                                     # 覆盖模板示例行（第 2..11 行）并向下扩展
        date_cell = ws3.cell(row=row, column=1)
        date_cell.value = wb_date(dates[d], ws3.cell(row=2, column=1))
        date_cell.number_format = ws3.cell(row=2, column=1).number_format
        time_text, energy_text = _emergency_text(roll["r_emg"][i])
        ws3.cell(row=row, column=2).value = time_text
        # 无紧急购电的日期：时间段列留空、购电量列填数值 0（交付清单 §一 规定）
        ws3.cell(row=row, column=3).value = energy_text if time_text != "" else 0
        if time_text != "":
            n_trigger += 1
        if dates[d].isoformat() in SPEC_DATES:
            spec_emg[dates[d].isoformat()] = (time_text, energy_text)

    wb.save(path)                                       # 保存副本，附件模板不受影响
    return {"n_trigger": n_trigger, "spec_emg": spec_emg}


def wb_date(d, template_cell):
    """把 datetime.date 转成与模板单元格同类型的日期对象（datetime.datetime）。

    输入：d，datetime.date；template_cell，openpyxl 单元格（提供类型与样式参照）
    输出：datetime.datetime 或 datetime.date（取决于模板单元格的值类型）
    """
    # openpyxl 读回的日期是 datetime.datetime；保持同型可让 Excel 的 mm-dd-yy 格式正常工作
    import datetime as _dt
    if template_cell.value is not None and isinstance(template_cell.value, _dt.datetime):
        return _dt.datetime(d.year, d.month, d.day)
    return d


def draw_main_figures(price, load, pv_actual, dates, roll_rep):
    """绘制四张主图（交付清单 §四 前四张），均 ≥300 dpi、中文标注。

    输入：price (K,) 元/kWh；load/pv_actual (D,K) kW；dates list；roll_rep，填报区间滚动结果
    输出：无（四张 PNG 落盘到 问题2/）
    """
    import datetime as _dt
    import matplotlib.pyplot as plt
    from lib.plotstyle import (COLOR_BUY, COLOR_CHG, COLOR_DIS, COLOR_LOAD, COLOR_PRICE,
                               COLOR_PV, COLOR_REF, COLOR_SOC, FIGSIZE_TALL, FIGSIZE_WIDE,
                               apply_chinese_style, hour_axis_ticks, save_figure)
    apply_chinese_style()
    sl = lambda i: i - D_REP_FIRST                      # 日期 0 基 -> 填报区间下标
    day_list = [dates[i] for i in roll_rep["d_index"]]   # 334 天的日期
    x_day = roll_rep["x_plan"].sum(axis=1)               # 逐日购电量，kWh
    cost_day = roll_rep["cost_day"]                      # 逐日购电费，元
    # 不储能基线（按日）：Σ p·max(L−P,0)Δ
    base_day = (price[None, :] * np.maximum(
        load[D_REP_FIRST:] * DT_H - pv_actual[D_REP_FIRST:] * DT_H, 0.0)).sum(axis=1)
    save_day = base_day - cost_day                       # 逐日节省额，元
    # 四个指定日期在填报区间中的位置
    spec_pos = [sl((_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days) for t in SPEC_DATES]

    # ---------------- 图 1：全年逐日购电费与节省额 ----------------
    fig, ax = plt.subplots(figsize=FIGSIZE_WIDE)
    idx = np.arange(N_REP)                               # 334 天的横轴位置
    ax.bar(idx, cost_day, width=0.8, color=COLOR_BUY, alpha=0.85, label="逐日购电费（元）")
    ax2 = ax.twinx()
    ax2.plot(idx, save_day, color=COLOR_PRICE, linewidth=1.2,
             label="相对不储能基线的节省额（元）")
    for p, t in zip(spec_pos, SPEC_DATES):
        ax.axvline(p, color=COLOR_REF, linestyle="--", linewidth=0.9, alpha=0.8)
        ax.annotate(t[5:], (p, ax.get_ylim()[1] * 0.98), rotation=0, ha="center", va="top",
                    fontsize=9, color=COLOR_REF)
    ticks = [0, 58, 120, 181, 242, 303, 333]             # 约每两个月一个刻度
    ax.set_xticks(ticks, [day_list[i].isoformat() for i in ticks], rotation=30, fontsize=9)
    ax.set_xlabel("日期（2025 年）")
    ax.set_ylabel("逐日购电费（元）", color=COLOR_BUY)
    ax2.set_ylabel("节省额（元）", color=COLOR_PRICE)
    ax.set_title("问题 2 全年逐日购电费与相对不储能基线的节省额\n"
                 "（日均节省 %.2f 元；总节省 %.2f 元，%.2f%%）"
                 % (save_day.mean(), save_day.sum(), 100.0 * save_day.sum() / base_day.sum()))
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="upper left")
    save_figure(fig, os.path.join(QDIR, "全年逐日购电费与节省额.png"))

    # ---------------- 图 2：四季典型日购电量与负载对比 ----------------
    fig, axes = plt.subplots(2, 2, figsize=FIGSIZE_TALL)
    hours = np.arange(1, K + 1) / 6.0 - 1.0 / 12.0       # 每个时段中心的钟点（h）
    for ax, t in zip(axes.ravel(), SPEC_DATES):
        i = sl((_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days)
        d0 = i + D_REP_FIRST
        # 购电量换算成平均功率（kW）= x/Δ，与负载/光伏同量纲可比
        ax.plot(hours, roll_rep["x_plan"][i] / DT_H, color=COLOR_BUY, linewidth=1.3,
                label="计划购电量（折合功率，kW）")
        ax.plot(hours, load[d0], color=COLOR_LOAD, linewidth=1.3, label="小区负载（kW）")
        ax.plot(hours, pv_actual[d0], color=COLOR_PV, linewidth=1.3, label="光伏实际（kW）")
        ax.axhline(0, color=COLOR_REF, linewidth=0.8)
        for h in SPEC_HOURS:
            ax.axvline(h, color=COLOR_REF, linestyle=":", linewidth=0.6, alpha=0.5)
        ax.set_xlabel("时刻（h）")
        ax.set_ylabel("功率（kW）")
        ax.set_title("%s（全天购电量 %.0f kWh，购电费 %.0f 元）" % (t, x_day[i], cost_day[i]))
        ax.legend(fontsize=8, loc="upper left")
        ax.set_xlim(0, 24)
    fig.suptitle("问题 2 四季典型日：计划购电量、负载与光伏对比（购电量按 Δ 折合为功率）", fontsize=15)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    save_figure(fig, os.path.join(QDIR, "四季典型日购电量与负载对比.png"))

    # ---------------- 图 3：储电量年度轨迹热力图 ----------------
    fig, ax = plt.subplots(figsize=FIGSIZE_WIDE)
    E_mat = roll_rep["E_soc"][:, :K]                     # (334,144)：各时段末储电量，kWh
    im = ax.pcolormesh(np.arange(K + 1), np.arange(N_REP + 1), E_mat,
                       cmap="magma", shading="flat", vmin=E_MIN, vmax=E_MAT_MAX)
    cb = fig.colorbar(im, ax=ax, label="储电量（kWh）")
    cb.set_ticks([E_MIN, 3000, 6000, 9000, E_MAT_MAX])
    ticks = [0, 58, 120, 181, 242, 303, 333]
    ax.set_yticks([t + 0.5 for t in ticks], [day_list[t].isoformat() for t in ticks], fontsize=9)
    ax.set_xticks(np.arange(0, K + 1, 12), ["%d:00" % h for h in range(0, 25, 2)], fontsize=8)
    ax.set_xlabel("时刻（h）")
    ax.set_ylabel("日期（2025 年）")
    ax.set_title("问题 2 储电量年度轨迹热力图（334 天 × 144 时段）\n"
                 "颜色越亮储电量越高；上下限参考值 %d / %d kWh；多数时段处于下限附近、"
                 "充电集中在日内低价窗" % (E_MIN, E_MAT_MAX))
    save_figure(fig, os.path.join(QDIR, "储电量年度轨迹热力图.png"))

    # ---------------- 图 4：全年购电费构成堆叠图（按月） ----------------
    months = [d.month for d in day_list]
    m_idx, m_cost, m_emg, m_base = [], [], [], []
    for m in range(2, 13):
        sel = np.array([mm == m for mm in months])
        m_idx.append("%d月" % m)
        m_cost.append(float(cost_day[sel].sum()))
        m_emg.append(float(roll_rep["cost_emg_day"][sel].sum()))
        m_base.append(float(base_day[sel].sum()))
    m_cost, m_emg, m_base = np.array(m_cost), np.array(m_emg), np.array(m_base)
    fig, ax = plt.subplots(figsize=FIGSIZE_WIDE)
    xs = np.arange(len(m_idx))
    ax.bar(xs, m_cost, width=0.55, color=COLOR_BUY, label="计划购电费用（元）")
    ax.bar(xs, m_emg, width=0.55, bottom=m_cost, color=COLOR_PRICE,
           label="紧急购电费用（元；读法① 为 0）")
    ax.plot(xs, m_base, "o--", color=COLOR_REF, linewidth=1.4, markersize=5,
            label="不储能基线（元）")
    ax.plot(xs, m_base - m_cost, "s-", color=COLOR_CHG, linewidth=1.4, markersize=5,
            label="储能节省额（元）")
    for x, (c, b) in enumerate(zip(m_cost, m_base)):
        ax.annotate("省 %.1f 万" % ((b - c) / 1e4), (x, b - c),
                    textcoords="offset points", xytext=(0, 7), ha="center", fontsize=9)
    ax.set_xticks(xs, m_idx)
    ax.set_xlabel("月份（2025 年，填报区间 2–12 月）")
    ax.set_ylabel("费用（元）")
    ax.set_title("问题 2 全年购电费构成与不储能基线对照（按月汇总）\n"
                 "总计：优化 %.2f 万元 vs 基线 %.2f 万元，节省 %.2f 万元（%.2f%%）"
                 % (cost_day.sum() / 1e4, base_day.sum() / 1e4,
                    (base_day.sum() - cost_day.sum()) / 1e4,
                    100.0 * (base_day.sum() - cost_day.sum()) / base_day.sum()))
    ax.legend(loc="upper left")
    save_figure(fig, os.path.join(QDIR, "全年购电费构成堆叠图.png"))


def main():
    """问题 2 主流程：载入数据 → 全年逐日滚动 → 落盘 CSV 与 result2.xlsx → 打印关键数值。"""
    print("=" * 78)
    print("问题 2：全年逐日滚动 LP（读法① 完全信息；终端自由 + 跨日传递，E_0=%.0f kWh）" % E_INIT)
    print("=" * 78)

    # ---------------- 1. 读附件1（电价）与附件2（全年负载/光伏实际） ----------------
    price, load1, pv1, _ = read_attachment1()           # 附件1：问题 2 只用其电价列（D-08）
    load2, pv_actual, dates = read_attachment2()        # 附件2：365 天逐日负载与光伏实际功率
    assert load2.shape == (365, K) and pv_actual.shape == (365, K), "附件2 形状异常"
    print("附件1 读入电价：%d 个时段，%.4f–%.4f 元/kWh；附件2 读入 %d 天 × %d 时段（%s 至 %s）"
          % (K, price.min(), price.max(), load2.shape[0], K, dates[0], dates[-1]))

    # ---------------- 2. 全年逐日滚动求解（含 1 月预热期） ----------------
    roll_all = solve_rolling(price, load2, pv_actual, e_init=E_INIT, mode="free", d_start=0, d_end=365)
    # 填报区间 = 2025-02-01 至 12-31（0 基第 31..364 天），与结果文件一一对应
    sl = slice(D_REP_FIRST, 365)
    roll_rep = {key: roll_all[key][sl] for key in
                ("d_index", "x_plan", "u_chg", "v_dis", "r_emg", "E_soc",
                 "cost_day", "cost_plan_day", "cost_emg_day", "status_day")}
    print("滚动求解完成：365 天全部 status=0（最优）=%s" % bool(np.all(roll_all["status_day"] == 0)))

    # ---------------- 3. 落盘审计 CSV 与 result2.xlsx ----------------
    write_audit_csv(AUDIT_CSV, dates, price, load2, pv_actual, roll_rep)
    summary = fill_result_xlsx(RESULT_XLSX, dates, roll_rep, price)
    print("已落盘：%s（334 天 × 144 时段全分辨率明细）" % AUDIT_CSV)
    print("已落盘：%s（由附件5 模板复制后填写三张表，附件未改动）" % RESULT_XLSX)

    # ---------------- 4. 关键数值打印（供构思手与外部审计对照） ----------------
    tot_cost = float(roll_rep["cost_day"].sum())        # 填报区间总费用（计划 + 紧急），元
    tot_plan = float(roll_rep["cost_plan_day"].sum())   # 计划购电费合计，元
    tot_emg = float(roll_rep["cost_emg_day"].sum())     # 紧急购电费合计，元
    base_pv_only = float(np.sum(price[None, :] * np.maximum(
        load2[D_REP_FIRST:] * DT_H - pv_actual[D_REP_FIRST:] * DT_H, 0.0)))  # 不储能基线

    print("-" * 78)
    print("【全年汇总（填报区间 334 天，2025-02-01 至 12-31）】")
    print("总计划购电费 = %.4f 元（锚定 %.4f，差 %+.4f 元）"
          % (tot_plan, ANCHOR["cost_rep"], tot_plan - ANCHOR["cost_rep"]))
    print("总紧急购电费 = %.4f 元（读法① 应为 0）" % tot_emg)
    print("总购电费用（计划+紧急）= %.4f 元；日均 = %.4f 元（锚定 %.4f）"
          % (tot_cost, tot_cost / 334.0, ANCHOR["cost_day_mean"]))
    print("总购电量 = %.4f kWh（锚定 %.4f）；总充电量 = %.4f kWh（锚定 %.4f）；总放电量 = %.4f kWh（锚定 %.4f）"
          % (float(roll_rep["x_plan"].sum()), ANCHOR["x_rep"],
             float(roll_rep["u_chg"].sum()), ANCHOR["u_rep"],
             float(roll_rep["v_dis"].sum()), ANCHOR["v_rep"]))
    print("填报区间期初 E_0(2.1 0:00) = %.4f kWh；期末 E(12.31 24:00) = %.4f kWh"
          % (roll_rep["E_soc"][0, 0], roll_rep["E_soc"][-1, K]))
    print("每日 24:00 储电量 = 下限 1200 的天数 = %d / 334"
          % int(np.sum(np.abs(roll_rep["E_soc"][:, K] - E_MIN) < 1e-3)))
    print("紧急购电量 max r = %.6f kWh（读法① 命题：应恒为 0）；触发紧急购电的天数 = %d"
          % (float(roll_rep["r_emg"].max()), summary["n_trigger"]))
    print("与不储能基线（方案 B，光伏自用不储能）对比：基线 %.4f 元，节省 %.4f 元（%.4f%%）"
          % (base_pv_only, base_pv_only - tot_cost, 100.0 * (base_pv_only - tot_cost) / base_pv_only))
    print("与全年联合 LP 下界对照：下界 %.4f 元，滚动超过下界 %.4f 元（%.4f%%）"
          % (ANCHOR["joint_lb"], tot_cost - ANCHOR["joint_lb"],
             100.0 * (tot_cost - ANCHOR["joint_lb"]) / ANCHOR["joint_lb"]))

    print("-" * 78)
    print("【题目表 1/表 2：四个指定日期（真实时段序 k=1..144；H:00-H:10 ↔ 第 6H+1 个时段 ↔ 模板第 6H 列）】")
    import datetime as _dt
    for t in SPEC_DATES:
        d0 = (_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days
        i = d0 - D_REP_FIRST                             # 在填报区间数组中的下标
        print("  --- %s（天序号 %d）---" % (t, d0 + 1))
        for h in SPEC_HOURS:
            k = hour_block_to_k(h)                       # 真实时段序号
            j0 = hour_block_to_template_row(h)           # 0 基模板列号
            print("    %2d:00-%2d:10  真实时段 %3d  模板第 %3d 列   x = %10.4f kWh"
                  % (h, h, k, j0 + 1, roll_rep["x_plan"][i, k - 1]))
        print("    全天购电量 = %.4f kWh；全天购电费 = %.4f 元"
              % (float(roll_rep["x_plan"][i].sum()), float(roll_rep["cost_day"][i])))
        for start_label, end_label, k_first, k_last in four_hour_blocks():
            print("    %s-%s  充 = %10.4f  放 = %10.4f"
                  % (start_label, end_label,
                     float(np.sum(roll_rep["u_chg"][i, k_first - 1:k_last])),
                     float(np.sum(roll_rep["v_dis"][i, k_first - 1:k_last]))))
        print("    0:00 储电量 = %.4f kWh；24:00 储电量 = %.4f kWh"
              % (roll_rep["E_soc"][i, 0], roll_rep["E_soc"][i, K]))
        print("    表 3（紧急购电）：时间段='%s'，购电量='%s'"
              % summary["spec_emg"].get(t, ("", "0")))

    print("-" * 78)
    print("【与问题 1 典型日对照】")
    print("问题 1 主模型（端点锁定）= 35126.9486 元；问题 2 日均 = %.4f 元，相对 +%.4f%%"
          "（逐日波动下的凸性效应，属正常）"
          % (tot_cost / 334.0, 100.0 * (tot_cost / 334.0 / 35126.9486 - 1.0)))

    # ---------------- 5. 四张主图 ----------------
    draw_main_figures(price, load2, pv_actual, dates, roll_rep)
    print("-" * 78)
    print("已落盘四张主图：全年逐日购电费与节省额 / 四季典型日购电量与负载对比 / "
          "储电量年度轨迹热力图 / 全年购电费构成堆叠图")
    print("=" * 78)


if __name__ == "__main__":
    # 把标准输出同时写入日志，保证报告中出现的每个数字都能在文件里找到出处
    tee = Tee(LOG_PATH)
    original_stdout = sys.stdout
    sys.stdout = tee
    try:
        main()
    finally:
        sys.stdout = original_stdout
        tee.close()
    print("运行日志已写入：%s" % LOG_PATH)
