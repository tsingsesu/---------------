"""多日滚动 LP 求解（问题 2/3/4 共用）：逐日 LP + 储电量跨日传递。

口径（`口径与假设台账.md`，不得自行更改）：
  * D-04 问题 2/3/4 终端储电量**自由** + 跨日传递 E_{144}^{d}=E_{0}^{d+1}，E_0^{1}=6000 kWh；
  * D-05 单向效率 η=0.9；
  * D-11 仿真从 2025-01-01 起，1 月为预热期，结果填报区间为 2025-02-01 至 12-31 共 334 天；
  * D-06 读法① 完全信息：计划（=实际生效购电量）覆盖净负荷，紧急购电量 r ≡ 0。

本模块只做"把单日 LP 串成多日滚动"这一件事；单日模型本体在 `lib/solve_day.py`，
储电量约束构造在 `lib/storage.py`，两处均直接复用，不作任何复制。

随机性：无（全流程为确定性 LP 求解，不需要随机种子）。
"""

import numpy as np

from lib.solve_day import solve_day
from lib.storage import E_INIT, E_MAX, E_MIN, ETA, P_MAX, U_MAX
from lib.timegrid import DT_H, K


def solve_rolling(price, load, pv_actual, e_init=E_INIT, mode="free", eta=ETA,
                  p_max=P_MAX, e_min=E_MIN, e_max=E_MAX, d_start=0, d_end=None):
    """逐日滚动 LP：按日推进，前一日 24:00 储电量作为次日 0:00 储电量。

    输入：price，np.ndarray (D,)，**当日**电价向量，元/kWh（问题 2/3 逐日相同）
          load，np.ndarray (D,K)，负载功率，kW
          pv_actual，np.ndarray (D,K)，光伏实际功率，kW
          e_init，kWh，E_0（第 d_start 天 0:00 的储电量）
          mode，str，'free'（终端自由，问题 2/3/4 主口径）或 'cyclic'（终端=当日初始，对照）
          eta / p_max / e_min / e_max，储能参数（默认取附录 1 标准值）
          d_start / d_end，int，滚动起止天序号（0 基，含 d_start、不含 d_end；None 表示到末尾）
    输出：dict，键含义——
          d_index  (n_days,)   参与滚动的天序号（0 基）
          x_plan   (n_days,K)  计划购电量，kWh
          u_chg    (n_days,K)  充电量，kWh
          v_dis    (n_days,K)  放电量，kWh
          r_emg    (n_days,K)  紧急购电量，kWh（读法① 下恒为 0，此处按定义逐时段重算）
          E_soc    (n_days,K+1) 储电量轨迹（下标 0 为当日 0:00，下标 K 为 24:00），kWh
          cost_day (n_days,)   计入紧急费用后的当日总购电费 Σp·x + 5Σp·r，元
          cost_plan_day (n_days,) 当日计划购电费 Σp·x，元
          cost_emg_day  (n_days,) 当日紧急购电费 5Σp·r，元
          status_day (n_days,) 各日求解器状态（0 为最优）
          e_init   float       滚动起点储电量（= 输入 e_init），kWh
          e_end    float       滚动终点储电量（= 最后一天 24:00），kWh

    说明：当日购电费按口径 D-12 的分项写法在 x=y（问题 2 无调整）时严格等于 Σp·x；
          紧急购电量按定义 r=[L·Δ+u−x−P·Δ−v]^+ 逐时段重算，**不由 LP 优化**，故本函数
          同时适用于"计划=实际"的读法①与"计划基于预测、实际为附件2"的读法②。
    """
    load = np.asarray(load, dtype=float)
    pv_actual = np.asarray(pv_actual, dtype=float)
    n_all = load.shape[0]                                  # 全部天数
    if d_end is None:
        d_end = n_all                                      # 默认滚到数据末尾
    d_index = np.arange(d_start, d_end)                    # 参与滚动的天序号
    n_days = d_index.size

    # 结果容器：逐日逐时段全分辨率落盘用，避免只留汇总
    x_plan = np.zeros((n_days, K))                         # 计划购电量，kWh
    u_chg = np.zeros((n_days, K))                          # 充电量，kWh
    v_dis = np.zeros((n_days, K))                          # 放电量，kWh
    r_emg = np.zeros((n_days, K))                          # 紧急购电量，kWh
    E_soc = np.zeros((n_days, K + 1))                      # 储电量轨迹，kWh
    cost_day = np.zeros(n_days)                            # 含紧急费用，元
    cost_plan_day = np.zeros(n_days)                       # 计划购电费，元
    cost_emg_day = np.zeros(n_days)                        # 紧急购电费，元
    status_day = np.zeros(n_days, dtype=int)               # 求解器状态

    e = float(e_init)                                      # 跨日传递的储电量状态，kWh
    for i, d in enumerate(d_index):
        # 单日 LP：目标 min Σp·x，供给/储能约束见 lib.solve_day；终端条件由 mode 控制
        res = solve_day(price, load[d], pv_actual[d], e_init=e, mode=mode,
                        eta=eta, p_max=p_max, e_min=e_min, e_max=e_max)
        x = res["x_plan"]; u = res["u_chg"]; v = res["v_dis"]
        # 紧急购电量按定义重算：供给缺口 = L·Δ + u − x − P·Δ − v（负值截断为 0）
        r = np.maximum(load[d] * DT_H + u - v - pv_actual[d] * DT_H - x, 0.0)
        x_plan[i] = x; u_chg[i] = u; v_dis[i] = v; r_emg[i] = r
        E_soc[i] = res["E_soc"]
        cost_plan_day[i] = float(np.dot(price, x))          # 计划购电费，元
        cost_emg_day[i] = 5.0 * float(np.dot(price, r))     # 紧急购电费（5 倍价），元
        cost_day[i] = cost_plan_day[i] + cost_emg_day[i]    # 当日总费用，元
        status_day[i] = res["status"]
        e = float(res["E_soc"][-1])                         # 传递到次日 0:00（D-04）

    return {
        "d_index": d_index,
        "x_plan": x_plan,
        "u_chg": u_chg,
        "v_dis": v_dis,
        "r_emg": r_emg,
        "E_soc": E_soc,
        "cost_day": cost_day,
        "cost_plan_day": cost_plan_day,
        "cost_emg_day": cost_emg_day,
        "status_day": status_day,
        "e_init": float(e_init),
        "e_end": float(e),
    }


def block_sums_by_k(values, k_first, k_last):
    """把一个"按真实时段序排列"的长度 144 数组按 [k_first, k_last] 区间求和。

    输入：values，np.ndarray (K,)，按真实时段序 k=1..144 排列的量（如 u_chg 或 v_dis）
          k_first / k_last，int，1 基的起止时段序号（含两端）
    输出：float，区间内求和
    """
    # 1 基闭区间 [k_first, k_last] 对应 0 基切片 [k_first-1, k_last)
    return float(np.sum(np.asarray(values, dtype=float)[k_first - 1:k_last]))
