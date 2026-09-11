"""结果工作簿读写助手（问题 3/4 共用）：模板复制、轮转填报重排、日期与紧急购电文本。

本模块把"往附件5 模板副本里填数"所需的通用操作集中在一处：
  1. `r4`：统一 4 位小数（D-15），并先把 1e-9 以下的求解器噪声归零；
  2. `set_decimal_format`：只改空白数值格的格式，不动模板自带格式；
  3. `wb_date`：生成与模板同型的日期单元格值；
  4. `fill_plan_like_sheet`：把"按真实时段序排列的一天 144 个值"按填法 Y-轮转写入一行；
  5. `emergency_text`：把一天的紧急购电量聚合成表 4 示例的"时间段+电量"文本。

约定来源：`口径与假设台账.md` D-01（轮转填报）、D-15（4 位小数）、D-16.8/9/10
（扩表、紧急购电填法、`7:0-7:10` 笔误照抄——笔误在模板里，本模块只复制模板标签，不生成标签）。

随机性：无（纯格式化与搬运，确定性）。
"""

import datetime as _dt

import numpy as np

ND = 4                      # 统一小数位数（D-15）
TOL_ZERO = 1e-9             # 容差归零阈值（数值噪声上界）


def r4(values, tol=TOL_ZERO):
    """把数组/标量先"容差归零"再四舍五入到 4 位小数（D-15 统一精度）。

    输入：values，array_like，任意数值；tol，float，归零阈值
    输出：np.ndarray 或 float，保留 4 位小数
    """
    arr = np.asarray(values, dtype=float)
    arr = np.where(np.abs(arr) < tol, 0.0, arr)            # 容差量级一律记 0（避免 '-0.0000'）
    return np.round(arr, ND)


def set_decimal_format(cell):
    """仅当单元格原格式为 General 时才设为 4 位小数，避免覆盖模板自带的数字格式。

    输入：cell，openpyxl 单元格对象
    输出：无（就地修改 cell.number_format）
    """
    if cell.number_format in ("General", "general"):
        cell.number_format = "0.0000"


def wb_date(d, template_cell):
    """把 datetime.date 转成与模板单元格同类型的日期对象（datetime.datetime）。

    输入：d，datetime.date；template_cell，openpyxl 单元格（提供类型参照）
    输出：datetime.datetime 或 datetime.date（取决于模板单元格的值类型）
    """
    if template_cell.value is not None and isinstance(template_cell.value, _dt.datetime):
        return _dt.datetime(d.year, d.month, d.day)
    return d


def template_order(values):
    """按填法 Y-轮转把"真实时段序"数组重排成模板行序（D-01）。

    输入：values，np.ndarray (144,)，按真实时段序 k=1..144 排列
    输出：np.ndarray (144,)，第 i 个元素是应填入模板第 i 个（0 基）时段格的值
    说明：模板第 i 格装当天第 (i+1)%144+1 个时段 ⇒ 等价于对原数组循环左移一位。
    """
    arr = np.asarray(values, dtype=float)
    return arr[(np.arange(arr.size) + 1) % arr.size]


def fill_row_periods(ws, row, values_k, col0=2):
    """把一天 144 个值（真实时段序）按轮转填法写入工作表一行。

    输入：ws，openpyxl 工作表；row，int，目标行号（1 基）
          values_k，np.ndarray (144,)，kWh 或任意逐时段量
          col0，int，首个时段列（1 基，默认 2 = 第 2 列）
    输出：无
    """
    values_t = template_order(r4(values_k))               # 轮转 + 4 位小数
    for j in range(values_t.size):
        cell = ws.cell(row=row, column=col0 + j)
        cell.value = float(values_t[j])
        set_decimal_format(cell)


def emergency_text(r_day, k_to_label):
    """把一天的紧急购电量聚合成连续时段文本（按表 4 示例写法，D-16.9）。

    输入：r_day，np.ndarray (144,)，逐时段紧急购电量，kWh
          k_to_label，callable，k→(起标签, 止标签)（来自 lib.timegrid.k_to_label）
    输出：(time_text, energy_text)，str；无紧急购电时返回 ("", "0")
    """
    active = np.asarray(r_day, dtype=float) > 1e-6        # 布尔掩码：该时段是否需要紧急购电
    if not active.any():
        return "", "0"
    idx = np.where(active)[0]                             # 需要紧急购电的时段下标（0 基）
    segments = []                                         # 连续段列表，每项为 [起下标, 止下标]
    seg_start = idx[0]
    for j in range(1, idx.size + 1):
        if j == idx.size or idx[j] != idx[j - 1] + 1:     # 断开或到末尾 ⇒ 当前段结束
            segments.append((seg_start, idx[j - 1]))
            if j < idx.size:
                seg_start = idx[j]
    time_parts, energy_parts = [], []
    for a, b in segments:
        time_parts.append("%s-%s" % (k_to_label(a + 1)[0], k_to_label(b + 1)[1]))
        seg_energy = r4(float(np.sum(r_day[a:b + 1])))    # 段内电量，4 位小数
        energy_parts.append(("%.4f" % seg_energy).rstrip("0").rstrip("."))
    return " ".join(time_parts), " ".join(energy_parts)
