"""数据读取层（全工作区共用）：把 `附件/` 下的四份数据统一读成 numpy 数组。

统一形态（与 `符号表.md` §五.3 的变量名对照一致）：
  * 附件1（典型日）    -> price (144,) 元/kWh、load (144,) kW、pv_plan (144,) kW、时间标签
  * 附件2（全年实际）  -> load (365,144) kW、pv_actual (365,144) kW、日期序列
  * 附件3（光伏预报）  -> pv_fc (365,4,24) kW、预报发布时刻
  * 附件4（全年电价）  -> price (365,144) 元/kWh

口径提醒（`口径与假设台账.md` D-01）：附件 1/2/4 的列标签是**时段右端点**，
第 j 列（0 基）代表当天第 j+1 个时段 (10j min, 10(j+1) min] 的量，144 列恰好铺满 [0:00, 24:00]。

`附件/` 全程只读：本模块只以只读方式打开工作簿，不做任何写回。
"""

import datetime as _dt
import os

import numpy as np
import openpyxl

# 工作区根目录：本文件位于 <根>/lib/dataio.py，故根目录是上一级
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# 附件目录（只读）
ATTACH_DIR = os.path.join(ROOT, "附件")

K = 144                  # 一天的 10 分钟时段数
D = 365                  # 全年天数（2025-01-01 至 2025-12-31）
TAU_LIST = (0, 6, 12, 18)  # 附件3 的四个预报发布时刻，h


def attach_path(*parts):
    """拼出 `附件/` 下的绝对路径，避免各脚本各写一套相对路径。

    输入：parts，str 片段，如 '附件1.xlsx' 或 '附件5', 'result1.xlsx'
    输出：str，绝对路径
    """
    # 统一在附件目录下拼接，保证"只读"入口唯一
    return os.path.join(ATTACH_DIR, *parts)


def _read_rows(path, sheet=None):
    """以只读方式读取工作表全部单元格（含表头）。

    输入：path，str，xlsx 路径
          sheet，str 或 None，工作表名；None 表示第一张表
    输出：list[tuple]，逐行的单元格值
    """
    # read_only=True 降低内存，data_only=True 取缓存值而非公式
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet is not None else wb.worksheets[0]
    rows = [tuple(r) for r in ws.iter_rows(values_only=True)]   # 逐行取出，转成元组
    wb.close()
    return rows


def read_attachment1(path=None):
    """读附件1（典型日：电价 / 小区负载 / 光伏发电预测功率）。

    输入：path，str 或 None，默认 `附件/附件1.xlsx`
    输出：(price, load, pv_plan, time_labels) ——
          price      np.ndarray (144,)，元/kWh
          load       np.ndarray (144,)，kW
          pv_plan    np.ndarray (144,)，kW
          time_labels list[str]，144 个原始时间标签（'0:10' .. '0:00+1'），仅供核对
    """
    path = path or attach_path("附件1.xlsx")
    rows = _read_rows(path)
    # 第 1 行是表头，其后 144 行是数据；第 A 列是时间标签，B/C/D 列是电价/负载/光伏
    data = rows[1:1 + K]
    price = np.array([float(r[1]) for r in data], dtype=float)      # 电价，元/kWh
    load = np.array([float(r[2]) for r in data], dtype=float)       # 小区负载功率，kW
    pv_plan = np.array([float(r[3]) for r in data], dtype=float)    # 光伏预测功率，kW
    time_labels = [_label_text(r[0]) for r in data]                 # 原始标签，便于人工核对
    return price, load, pv_plan, time_labels


def _label_text(value):
    """把附件里的时间单元格（可能是 time 或 str）统一成字符串标签。

    输入：value，datetime.time 或 str
    输出：str，如 '0:10'、'0:00+1'
    """
    # time 对象按 H:MM 输出；字符串直接原样返回（附件1 末行是 '0:00+1'）
    if isinstance(value, _dt.time):
        return "%d:%02d" % (value.hour, value.minute)
    return str(value)


def _read_wide_sheet(path, sheet):
    """读"日期 × 144 时段"的宽表工作表（附件2 的两张表、附件4）。

    输入：path，str，xlsx 路径
          sheet，str，工作表名
    输出：(values, dates) —— values np.ndarray (365,144)；dates list[datetime.date]
    """
    rows = _read_rows(path, sheet)
    body = rows[1:1 + D]                                            # 跳过表头，取 365 行
    # 第 1 列是日期，第 2..145 列是 144 个时段的值
    values = np.array([[float(v) for v in r[1:1 + K]] for r in body], dtype=float)
    dates = [r[0].date() if isinstance(r[0], _dt.datetime) else r[0] for r in body]
    return values, dates


def read_attachment2(path=None):
    """读附件2（全年小区负载与光伏实际功率，两张工作表）。

    输入：path，str 或 None，默认 `附件/附件2.xlsx`
    输出：(load, pv_actual, dates) ——
          load       np.ndarray (365,144)，kW
          pv_actual  np.ndarray (365,144)，kW
          dates      list[datetime.date]，长度 365
    """
    path = path or attach_path("附件2.xlsx")
    load, dates = _read_wide_sheet(path, "小区负载")           # 工作表名见附件2
    pv_actual, _ = _read_wide_sheet(path, "光伏发电实际功率")   # 同名日期列，只需一次
    return load, pv_actual, dates


def read_attachment3(path=None):
    """读附件3（整点光伏预报长表）。

    输入：path，str 或 None，默认 `附件/附件3.xlsx`
    输出：(pv_fc, dates, tau_list) ——
          pv_fc    np.ndarray (365,4,24)，kW；pv_fc[d,i,m-1] 为第 d 天、第 i 个发布时刻、
                   提前期 m 小时（即整点 (tau+m):00 处）的预报值
          dates    list[datetime.date]，长度 365
          tau_list tuple[int,...]，(0,6,12,18)
    """
    path = path or attach_path("附件3.xlsx")
    rows = _read_rows(path)                                    # 第 1 行表头，其后 4×365 行
    pv_fc = np.zeros((D, len(TAU_LIST), 24), dtype=float)      # 结果：预报立方
    current_date = None                                        # 附件3 的日期列按天合并，需向下填充
    tau_index = {t: i for i, t in enumerate(TAU_LIST)}         # 发布时刻 -> 第几层
    for r in rows[1:]:
        if r[0] not in (None, ""):                             # 出现新日期时更新当前日期
            current_date = r[0]
        # "预报时刻"列写成 '0:00' / '6:00' / '12:00' / '18:00'
        tau_hour = int(str(r[1]).split(":")[0])
        # 第 3..26 列是预报1小时..预报24小时
        values = [float(v) for v in r[2:2 + 24]]
        # 日期 -> 行号：附件3 从 2025-1-1 起按天排列；这里用日期构造行号
        day_index = _day_index(current_date)
        pv_fc[day_index, tau_index[tau_hour], :] = values
    dates = [_dt.date(2025, 1, 1) + _dt.timedelta(days=i) for i in range(D)]
    return pv_fc, dates, TAU_LIST


def _day_index(value):
    """把附件3 的日期（'2025-1-1' 字符串或 datetime）转成 0 基天数。

    输入：value，str 或 datetime.datetime/datetime.date
    输出：int，0..364（2025-01-01 为 0）
    """
    if isinstance(value, _dt.datetime):
        day = value.date()
    elif isinstance(value, _dt.date):
        day = value
    else:
        # 字符串形如 '2025-1-1'
        parts = [int(x) for x in str(value).split("-")]
        day = _dt.date(parts[0], parts[1], parts[2])
    # 与基准日相减得到 0 基天数
    return (day - _dt.date(2025, 1, 1)).days


def read_attachment4(path=None):
    """读附件4（全年逐日逐时段电价）。

    输入：path，str 或 None，默认 `附件/附件4.xlsx`
    输出：(price, dates) —— price np.ndarray (365,144)，元/kWh；dates list[datetime.date]
    """
    path = path or attach_path("附件4.xlsx")
    price, dates = _read_wide_sheet(path, "Sheet1")   # 附件4 只有一张 Sheet1
    return price, dates


def template_rows_for_plan(price_day):
    """按"填法 Y-轮转"把逐时段解（真实时间序）重排成模板行序。

    输入：price_day，np.ndarray (144,)，任一按真实时段序 k=1..144 排列的逐时段量
    输出：np.ndarray (144,)，第 i 个元素是应填入模板第 i 行（0 基）的值
    说明：模板第 i 行装当天第 (i+1)%144+1 个时段，故重排索引 = (i+1)%144（0 基）。
    """
    # 0 基索引下的真实时段 k-1 = (i+1) % 144，等价于对原数组做一次循环左移一位
    values = np.asarray(price_day, dtype=float)
    index = (np.arange(K) + 1) % K          # 0,1,...,143 -> 1,2,...,143,0
    return values[index]
