r"""光伏预报处理（问题 3/4-3 共用）：整点对齐与"整点预报 → 144 个 10 分钟值"的分解。

口径来源：`口径与假设台账.md` D-07（已实测裁决）与 `_phase0/报告12_整点预报分解判定.txt`：
  * 附件3 的"预报 m 小时"= 整点 (τ+m):00 处的光伏功率值（τ 为发布时刻 0/6/12/18）；
  * 主口径 **M2（linear）**：把 25 个整点结点（0:00, 1:00, …, 24:00）线性插值，
    第 k 个 10 分钟区间的取值 = 结点折线在**该区间右端点**（钟点 k/6 h）处的值
    —— 与 D-07"整点所在的那一段直接取预报值"等价（右端点恰为整点）；
  * 当天 0:00 结点取自**前一日 0:00 预报的第 24 个值**（D-07）；2025-01-01 无前一日，取 0；
  * τ>0 发布的预报没有 τ:00 处的自身预报值，故 τ:00 结点取**上一次发布**对 τ:00 的预报
    （τ=6:00 取 0:00 预报的"预报 6 小时"值；τ=12:00 取 6:00 预报的"预报 6 小时"值；
    τ=18:00 取 12:00 预报的"预报 6 小时"值）——这是"决策时点可获得的最新估计"，本次新增口径，
    见 `问题3/编程手汇报.md` 假设说明。
  * 对照口径（灵敏度 S4a 用）：M1 整点常数（freq="hour_const"）、
    M3 历史小时内平均形状（freq="shape_hist"）、M4 典型日小时内形状（freq="shape_typ"）。

变量名（与 `符号表.md` §五.3 一致，新增者见编程手汇报的符号登记申请）：
  pv_fc      (365,4,24) 附件3 原始整点预报，kW（符号 $\hat P_m^{(d,\tau)}$）
  pv_fc144   (365,4,144) 分解到 10 分钟粒度后的预报，kW（符号 $\hat P_k^{d,\tau}$）
  tau_list   (0,6,12,18) 发布时刻，h

随机性：本模块的确定性函数无随机性；重抽样函数一律显式传入随机种子。
"""

import numpy as np

from lib.dataio import read_attachment3
from lib.timegrid import DT_H, K

# 发布时刻（h）与 0 基层号的对应，供各脚本统一引用（避免各写一套映射）
TAU_LIST = (0, 6, 12, 18)
TAU_INDEX = {tau: i for i, tau in enumerate(TAU_LIST)}
# 各发布时刻覆盖的"决策时段"的 0 基起点：τ 时刻之后（不含正在执行的 10 分钟）的时段
# D-13：k ∈ (τ, τ+10min] 的第一个时段是第 6τ+1 个（1 基），对应 0 基下标 6τ
TAU_K_START = {0: 0, 6: 36, 12: 72, 18: 108}
# 结点时刻（h），0:00..24:00 共 25 个
NODE_HOURS = np.arange(0, 25, dtype=float)


def release_nodes(pv_fc, d, tau):
    """取第 d 天、发布时刻 tau 的"整点结点序列"（0:00..24:00 共 25 个，未覆盖处为 nan）。

    输入：pv_fc，np.ndarray (365,4,24)，附件3 原始预报，kW
          d，int，天序号（0 基，2025-01-01 为 0）
          tau，int，发布时刻，取 0/6/12/18
    输出：np.ndarray (25,)，该发布对当天 0:00..24:00 每个整点的"最新可得估计"，kW
    说明：每个整点 h 取**覆盖它的最新一次发布**（发布时刻 τ′ < h 的最大 τ′）给出的预报；
          0:00 结点取自前一日 0:00 预报的第 24 个值（D-07）；2025-01-01 取 0（夜间）。
          该规则保证：τ 越大，结点被更新的整点越多（信息集 F_τ 单调增大），且每个结点
          都严格是"决策时刻之前已发布"的信息，不需要未来数据。
    """
    nodes = np.full(25, np.nan)                 # 25 个结点，先置 nan
    for h in range(0, 25):
        if h == 0:
            # 当天 0:00 结点：前一日 0:00 预报的第 24 个值（D-07）；2025-01-01 取 0
            nodes[h] = pv_fc[d - 1, 0, 23] if d > 0 else 0.0
            continue
        # 覆盖 h 的最新发布：满足 τ′ < h 且 τ′ ≤ τ（不能使用"未来才发布"的信息）的最大 τ′
        i_star = -1
        for i2, t2 in enumerate(TAU_LIST):
            if h > t2 and t2 <= tau:
                i_star = i2
        assert i_star >= 0, "整点 %d 无法被任何发布覆盖" % h
        nodes[h] = pv_fc[d, i_star, h - TAU_LIST[i_star] - 1]
    return nodes


def decompose_linear(nodes):
    """M2 主口径：结点折线在每段右端点取值，得到 144 个 10 分钟区间的预报值。

    输入：nodes，np.ndarray (25,)，0:00..24:00 的整点结点值，kW
    输出：np.ndarray (144,)，逐时段（真实时间序 k=1..144）的预报功率，kW
    说明：第 k 段右端点钟点为 k/6 h，落在结点区间 [m−1, m]（m=ceil(k/6)）内；
          整点段（k=6m）恰好取结点值 nodes[m]，与 D-07 一致。
    """
    t = np.arange(1, K + 1) / 6.0               # 各段右端点钟点，h：1/6, 2/6, …, 24
    m = np.ceil(t).astype(int)                  # 所在结点区间的右端结点号（1..24）
    frac = m - t                                # 距右端结点的比例（0 表示正好在结点上）
    # 线性插值：val = nodes[m-1]·frac + nodes[m]·(1-frac)
    # 推导：右端点 t 位于 [m-1, m]，距 m-1 的距离为 t-(m-1) = 1-frac，故 m-1 的权重是 frac
    return nodes[m - 1] * frac + nodes[m] * (1.0 - frac)


def decompose_hour_const(nodes):
    """M1 对照口径：整点值覆盖其后整小时（分段常数）。

    输入：nodes，np.ndarray (25,)，整点结点值，kW
    输出：np.ndarray (144,)，逐时段预报功率，kW
    """
    frame = np.asarray(nodes, dtype=float)[1:25]           # 1:00..24:00 共 24 个整点值
    return np.repeat(frame, 6)                             # 每个整点值重复 6 段（=1 小时）


def within_hour_shape(pv_ref, thr=500.0):
    """统计"小时内相对形状"：6 个 10 分钟位置相对该小时整点值的平均比例。

    输入：pv_ref，np.ndarray (n,144)，参考光伏序列（附件2 实际或附件1 典型日），kW
          thr，float，参与统计的最小整点值（kW），避免夜间近零值放大比值噪声
    输出：np.ndarray (6,)，相对形状，末位归一化为 1（与 `_phase0/probe07` 的 S3/S4 同法）
    """
    pv_ref = np.asarray(pv_ref, dtype=float)
    shape_num = np.zeros(6)                                # 6 个位置的比值累加器
    shape_cnt = np.zeros(6)                                # 6 个位置的计数（各位置相同，见下）
    for h in range(24):                                    # 逐小时统计
        c0 = 6 * h                                         # 该小时首段的 0 基列号
        end_v = pv_ref[:, c0 + 5]                          # 该小时"整点列"（小时末段）的值
        ok = end_v > thr                                   # 只统计白天、值足够大的小时
        n_ok = float(np.count_nonzero(ok))                 # 本小时参与统计的样本数
        if n_ok == 0:
            continue
        for j in range(6):
            shape_num[j] += float(np.sum(pv_ref[ok, c0 + j] / end_v[ok]))
            shape_cnt[j] += n_ok
    shape = shape_num / shape_cnt                          # 每个位置的平均比值
    return shape / shape[5]                                # 末位（整点列）归一化为 1，与 probe07 同法


def decompose_shape(nodes, shape):
    """M3/M4 对照口径：整点值 × 小时内相对形状。

    输入：nodes，np.ndarray (25,)，整点结点值，kW
          shape，np.ndarray (6,)，小时内相对形状（末位=1，与整点值对齐）
    输出：np.ndarray (144,)，逐时段预报功率，kW
    """
    frame = np.asarray(nodes, dtype=float)[1:25].reshape(24, 1)   # (24,1)：每小时整点值
    return (frame * np.asarray(shape, dtype=float).reshape(1, 6)).reshape(K)


def build_pv_fc144(pv_fc, method="linear", shape=None):
    """把附件3 的整点预报批量分解为 365×4×144 的 10 分钟粒度预报（四种口径）。

    输入：pv_fc，np.ndarray (365,4,24)，原始预报，kW
          method，str，'linear'（M2 主口径）/ 'hour_const'（M1）/ 'shape'（M3/M4，需给 shape）
          shape，np.ndarray (6,) 或 None，method='shape' 时的小时内相对形状
    输出：np.ndarray (365,4,144)，分解后的逐时段预报功率，kW
    """
    n_day = pv_fc.shape[0]
    out = np.zeros((n_day, len(TAU_LIST), K), dtype=float)
    for d in range(n_day):
        for tau in TAU_LIST:
            nodes = release_nodes(pv_fc, d, tau)           # 该发布的 25 个整点结点
            if method == "linear":
                out[d, TAU_INDEX[tau]] = decompose_linear(nodes)
            elif method == "hour_const":
                out[d, TAU_INDEX[tau]] = decompose_hour_const(nodes)
            elif method == "shape":
                assert shape is not None, "method='shape' 必须提供 shape"
                out[d, TAU_INDEX[tau]] = decompose_shape(nodes, shape)
            else:
                raise ValueError("未知分解方法：%s" % method)
    # nan 兜底：任何未被任何规则覆盖的位置（理论上不存在）按 0 处理并告警
    if np.isnan(out).any():
        raise RuntimeError("分解结果出现 nan，请检查结点覆盖")
    return out


def forecast_degrade(pv_fc144, tau_index=0, bias=0.0):
    """对预报做整体乘性加偏（灵敏度 S3a 用）：P̂ ← P̂ × (1 + bias)。

    输入：pv_fc144，np.ndarray (365,4,144)，分解后的预报，kW
          tau_index，int，只对指定发布层加偏（默认 0，即所有决策都基于加偏后的 0:00 预报）
          bias，float，乘性偏差，如 0.05 表示高报 5%
    输出：np.ndarray，同形状，加偏后的预报（负值截断为 0）
    """
    out = pv_fc144.copy()
    i = int(tau_index)
    out[:, i, :] = np.maximum(out[:, i, :] * (1.0 + float(bias)), 0.0)
    return out


def resample_forecast(pv_fc144, err_pool, seed, tau_index=0):
    """按实测误差池重抽样生成"新预报"（灵敏度 S3b 用，固定种子可复现）。

    输入：pv_fc144，np.ndarray (365,4,144)，分解后的预报，kW
          err_pool，np.ndarray (m,144)，可重抽样的整日误差场池（由附件2/附件3 反演，kW）
          seed，int，随机种子
          tau_index，int，只替换指定发布层
    输出：np.ndarray，同形状，替换后的预报（负值截断为 0）
    """
    rng = np.random.default_rng(int(seed))                 # 固定种子，保证可复现
    n_day = pv_fc144.shape[0]
    idx = rng.integers(0, err_pool.shape[0], size=n_day)   # 每天抽一个整日误差场（保留日内相关）
    out = pv_fc144.copy()
    i = int(tau_index)
    out[:, i, :] = np.maximum(out[:, i, :] + err_pool[idx, :], 0.0)
    return out


def error_pool_from_data(pv_fc144, pv_actual, tau_index=0):
    r"""由"预报 − 实际"反演整日误差场池（场景构造/S3b 共用）。

    输入：pv_fc144，np.ndarray (365,4,144)，分解后的预报，kW
          pv_actual，np.ndarray (365,144)，实际光伏，kW
          tau_index，int，取自哪个发布层（默认 0:00 预报）
    输出：np.ndarray (365,144)，逐日误差场，kW（符号 $\xi$ 的绝对形式）
    """
    e = pv_fc144[:, int(tau_index), :] - pv_actual
    return e


def day_forecast_matrix(pv_fc144, day, stages=(0, 6, 12, 18)):
    """取某天各决策时刻的 144 段预报，拼成下游求解器需要的列表。

    输入：pv_fc144，np.ndarray (365,4,144)，分解后的预报，kW
          day，int，天序号（0 基）
          stages，tuple[int]，参与决策的发布时刻
    输出：dict，{tau: np.ndarray (144,)}，各决策时刻可直接使用的逐时段预报
    """
    return {tau: pv_fc144[day, TAU_INDEX[tau], :].copy() for tau in stages}


def load_forecast(method="linear", shape=None):
    """一站式读取附件3 并分解（供各脚本复用，避免重复读盘代码）。

    输入：method，str，分解口径；shape，np.ndarray(6,) 或 None
    输出：(pv_fc, pv_fc144, dates, tau_list) —— 原始 (365,4,24)、分解 (365,4,144)、日期、发布时刻
    """
    pv_fc, dates, tau_list = read_attachment3()
    pv_fc144 = build_pv_fc144(pv_fc, method=method, shape=shape)
    return pv_fc, pv_fc144, dates, tau_list
