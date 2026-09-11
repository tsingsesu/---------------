"""时间网格与时段口径（全工作区共用，权威口径见 `口径与假设台账.md` D-01）。

本模块把"10 分钟时段"的一切换算集中在一处，四问共用，避免各问各写一套：
  1. 真实时段序号 k（1..144）与钟点区间 (10(k-1) min, 10k min] 的互转；
  2. 附件 1/2/4 的数据列（0 基）与 k 的对应：第 j 列（0 基 j）就是当天第 j+1 个时段；
  3. 附件5 模板行（0 基）与 k 的"轮转"对应（填法 Y-轮转，D-01）；
  4. 题目表 1 的指定时段 ``H:00-H:10`` 与 k、模板行的对应；
  5. 题目表 2 的六个 4 小时块（真实时间顺序）所含时段集合。

约定与提醒（D-01，务必遵守）：
  * 模板第 0 基第 i 行装的是**当天第 (i+1) % 144 + 1 个时段**，
    因此模板最后一行（标签 `0:00+1-0:10+1`）在时间上是**当天最早的一段**；
  * 一切聚合、统计、绘图必须按**真实时段序 k=1..144**，不得按模板行序。
"""

# ============================== 模块级常量 ==============================

K = 144                     # 一天的 10 分钟时段数，K=144
DT_H = 1.0 / 6.0            # 单个时段长度，h（10 分钟 = 1/6 h）
MINUTES_PER_DAY = 24 * 60   # 一天的总分钟数，1440 min
HOURS_PER_DAY = 24          # 一天的小时数，24 h
BLOCK_HOURS = 4             # 表 2 的聚合块长度，4 h
BLOCK_COUNT = 6             # 表 2 的聚合块个数，6 块（4 h × 6 = 24 h）
K_PER_BLOCK = BLOCK_HOURS * 60 // 10   # 每个 4 小时块含 36 个 10 分钟时段


def k_to_minutes(k):
    """真实时段序号 → 该时段的（起始分钟, 结束分钟）。

    输入：k，int，真实时段序号，取值 1..144（第 k 个时段为 (10(k-1) min, 10k min]）
    输出：(start_min, end_min)，tuple[int, int]，单位 min，范围 (0, 0) .. (1430, 1440)
    """
    # 起点为 10(k-1) 分钟，终点为 10k 分钟：与 D-01 中"标签为右端点"的口径一致
    return (10 * (k - 1), 10 * k)


def minute_to_label(minute):
    """分钟刻度 → 钟点标签字符串（24:00 单独处理，与表 2 的写法一致）。

    输入：minute，int，当天 0:00 起的分钟数，0..1440
    输出：str，如 0 -> '0:00'、610 -> '10:10'、1440 -> '24:00'
    """
    # 1440 min 即当天 24:00，直接返回固定写法，避免出现 '0:00' 造成歧义
    if minute >= MINUTES_PER_DAY:
        return "24:00"
    # 整点部分与分钟部分分别格式化，分钟固定两位（10 分钟粒度下只可能是 00/10/..50）
    return "%d:%02d" % (minute // 60, minute % 60)


def k_to_label(k):
    """真实时段序号 → （起点标签, 终点标签）。

    输入：k，int，1..144
    输出：(起, 止)，tuple[str, str]，如 k=1 -> ('0:00', '0:10')、k=144 -> ('23:50', '24:00')
    """
    start_min, end_min = k_to_minutes(k)          # 取出该时段的分钟区间
    return (minute_to_label(start_min), minute_to_label(end_min))


def k_to_data_col(k):
    """真实时段序号 → 附件 1/2/4 的数据列号（0 基）。

    输入：k，int，1..144
    输出：int，0 基列号；附件第 j 列（0 基）的时间标签为 10(j+1) min，即当天第 j+1 个时段
    """
    # 附件数据第 k 个值就是当天第 k 个时段，故 0 基列号 = k-1
    return k - 1


def data_col_to_k(col):
    """附件 1/2/4 的数据列号（0 基）→ 真实时段序号 k。

    输入：col，int，0..143
    输出：int，真实时段序号 1..144
    """
    return col + 1


def template_row_to_k(i):
    """附件5 模板数据行（0 基）→ 该行应填的**真实时段序号** k（填法 Y-轮转，D-01）。

    输入：i，int，0..143（0 基；模板第 1 个数据行是 i=0，标签 '0:10-0:20'）
    输出：int，真实时段序号，1..144
    说明：i=0..142 填当天第 i+2 个时段；i=143（标签 '0:00+1-0:10+1'）填当天第 1 个时段。
    """
    # 等价写法：(i+1) % 144 + 1 —— i=143 时回绕到 1，其余为 i+2
    return (i + 1) % K + 1


def k_to_template_row(k):
    """真实时段序号 k → 应填入的附件5 模板数据行（0 基）。

    输入：k，int，1..144
    输出：int，0 基模板行号 0..143；与 template_row_to_k 互为逆映射
    """
    # 逆映射：(k-2) mod 144；k=1 -> 143（模板最后一行），k=2 -> 0
    return (k - 2) % K


def hour_block_to_k(hour):
    """题目表 1 的指定时段 ``H:00-H:10`` → 真实时段序号 k。

    输入：hour，int，整点小时 H（题目取 10/12/14/16/18/20）
    输出：int，真实时段序号 6H+1（第 6H+1 个时段恰为 (H:00, H:10]）
    """
    # 每 6 个时段恰为 1 小时，故 H:00-H:10 是第 6H+1 个时段
    return 6 * hour + 1


def hour_block_to_template_row(hour):
    """题目表 1 的指定时段 ``H:00-H:10`` → 附件5 模板数据行（0 基）。

    输入：hour，int，整点小时 H
    输出：int，0 基模板行 6H-1（模板第 6H 个数据行，1 基）
    """
    # 先得到真实时段序号，再经轮转映射回模板行；等价于 6H-1
    return k_to_template_row(hour_block_to_k(hour))


def k_to_hour_labels():
    """全部 144 个时段的（起点标签, 终点标签）列表，按真实时间顺序。

    输入：无
    输出：list[tuple[str, str]]，长度 144
    """
    # 逐个时段生成标签，供 CSV 落盘与绘图使用
    return [k_to_label(k) for k in range(1, K + 1)]


def four_hour_blocks():
    """题目表 2 的六个 4 小时块（按真实时间顺序）。

    输入：无
    输出：list[tuple[str, str, int, int]]，每项为 (块起标签, 块止标签, 首时段 k, 末时段 k)
    """
    blocks = []                                      # 结果容器：六个块的元信息
    for b in range(BLOCK_COUNT):                     # 依次生成 0:00-4:00 .. 20:00-24:00
        k_first = b * K_PER_BLOCK + 1                # 该块第一个时段：1, 37, 73, 109, 145 之外
        k_last = (b + 1) * K_PER_BLOCK               # 该块最后一个时段
        start_label = minute_to_label(b * BLOCK_HOURS * 60)                 # 块起钟点
        end_label = minute_to_label((b + 1) * BLOCK_HOURS * 60)             # 块止钟点
        blocks.append((start_label, end_label, k_first, k_last))
    return blocks


def block_sums_by_k(values):
    """把一个"按真实时段序排列"的长度 144 数组聚合成六个 4 小时块的和。

    输入：values，np.ndarray，shape (144,)，按真实时段序 k=1..144 排列
    输出：np.ndarray，shape (6,)，六个 4 小时块（真实时间顺序）的和
    """
    # 仅依赖 numpy 的 reshape：36 个时段一组，共 6 组；调用方保证顺序为真实时段序
    import numpy as np
    return np.asarray(values, dtype=float).reshape(BLOCK_COUNT, K_PER_BLOCK).sum(axis=1)


def k_of_hour(hour):
    """附件 1/2/4 中标签为 ``H:00`` 的那一列（0 基）对应的真实时段序号。

    输入：hour，int，整点小时 H（1..24；H=24 表示标签 '0:00+1'）
    输出：int，真实时段序号 6H（标签 H:00 是第 6H 个时段的右端点）
    """
    # 标签 H:00 对应 60H 分钟，是第 60H/10 = 6H 个时段的右端点
    return 6 * hour
