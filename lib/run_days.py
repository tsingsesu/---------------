"""多日滚动 LP 求解（问题 2/3/4 共用）：逐日 LP + 储电量跨日传递。

本模块包含两组接口：
  1. `solve_rolling`（问题 2 主用）：逐日滚动，计划=实际（完全信息读法）；
  2. `solve_rolling_staged`（问题 3/4-3 新增）：逐日**多阶段**滚动——每天在 0:00 与若干
     调整时刻（6:00/12:00/18:00）依次决策，每个决策时刻只用**当时可得的预报**重解未来时段，
     已执行时段不可追溯（D-13）；偏差与紧急费用按 D-12 在实际光伏下结算。

口径（`口径与假设台账.md`，不得自行更改）：
  * D-04 问题 2/3/4 终端储电量**自由** + 跨日传递 E_{144}^{d}=E_{0}^{d+1}，E_0^{1}=6000 kWh；
  * D-05 单向效率 η=0.9；
  * D-11 仿真从 2025-01-01 起，1 月为预热期，结果填报区间为 2025-02-01 至 12-31 共 334 天；
  * D-06 读法① 完全信息：计划（=实际生效购电量）覆盖净负荷，紧急购电量 r ≡ 0。
  * D-07 预报分解主口径 M2（整点线性插值，见 lib/forecast.py）；
  * D-12 结算口径 J = Σ[p·min(x,y) + κ₋p(x−y)⁺ + κ₊p(y−x)⁺] + κΣp·r；
  * D-13 调整不可追溯：k∈(0:00,6:00] 用计划量；k∈(6:00,12:00] 用 6:00 决策值；余类推。

本模块只做"把单日/单阶段 LP 串成多日滚动"这一件事；单日模型本体在 `lib/solve_day.py`
（`solve_day` 与 `solve_stage`），储电量约束构造在 `lib/storage.py`，三处均直接复用。

随机性：无（全流程为确定性 LP 求解，不需要随机种子）。
"""

import numpy as np

from lib.solve_day import solve_day, solve_stage, settle_total
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


def solve_rolling_staged(price, load, pv_actual, pv_fc144, e_init=E_INIT,
                         stages=(0, 6, 12, 18), kmark=None, d_start=0, d_end=None,
                         eta=ETA, p_max=P_MAX, e_min=E_MIN, e_max=E_MAX,
                         kappa=5.0, kappa_under=0.5, kappa_over=1.5, dt_h=DT_H):
    """逐日**多阶段**滚动 LP（问题 3/4-3 主内核）：每天按决策时刻序列依次求解未来时段。

    决策流程（D-13 不可追溯，选项 C 语义）：
      τ=0   以 0:00 预报对全天 144 段求解 → 计划购电量 x^d（全时段），同时给出储能轨迹；
      τ=6   用已执行段 [1,36] 的 u,v 得 E(36)；对 [37,144] 段用 6:00 预报重解，
            **偏差结算的基准始终是 0:00 计划 x**（D-12：最终 y 与计划 x 比较）；
      τ=12/18 以此类推，各自只覆盖当时尚未执行的时段（不可追溯）。
      未列入 stages 的决策时刻不执行（用于策略对比 (a)–(d)）。

    为什么每个阶段都以 0:00 计划为偏差基准：D-12 的结算发生在"计划 x 与**最终**购电量 y"
    之间；中间阶段的改值只要被后续阶段改回，就不产生任何偏差费用。若改用"上一阶段的值"
    作基准，会把这类"改回"误记为超用/欠取，与 D-12 不符（见 `问题3/编程手汇报.md` 假设说明）。

    实际值只在评估阶段代入：先用**实际光伏**按"最终生效购电量 y + 最终充放电 u,v"重算
    紧急购电量 r = [L·Δ + u − y − P·Δ − v]⁺，再按 D-12 结算。

    输入：price，np.ndarray (K,)，当日电价，元/kWh（问题 3 逐日相同）
          load，np.ndarray (D,K)，负载功率，kW
          pv_actual，np.ndarray (D,K)，光伏**实际**功率，kW（只用于评估）
          pv_fc144，np.ndarray (D,4,K)，分解到 10 分钟粒度的光伏**预报**，kW；
                    第 2 维按发布时刻顺序 (0,6,12,18)（与 lib/forecast.TAU_LIST 一致）
          e_init，kWh，第 d_start 天 0:00 的储电量（跨日传递的起点）
          stages，tuple[int]，当天执行的决策时刻（0 必须在内；不含 0 时全天按 6:00 决策处理）
          kmark，dict 或 None，{τ: k_start}，τ 决策覆盖的时段起点（1 基时段序号-1 的 0 基下标）；
                  None 时默认 {0:0, 6:36, 12:72, 18:108}（D-13）
          d_start / d_end，int，滚动天序号（0 基，含 d_start、不含 d_end；None 到末尾）
          eta / p_max / e_min / e_max / dt_h，储能参数（默认附录 1 标准值）
          kappa / kappa_under / kappa_over，结算倍数（κ=5、κ₋=0.5、κ₊=1.5）
    输出：dict，键含义（各数组第一维均为参与滚动的天数 n_days）——
          d_index   (n_days,)   天序号（0 基）
          x_plan / y_adj / u_chg / v_dis / r_emg / g_curt  (n_days,K)  计划量/最终量/充/放/紧急/弃光，kWh
          E_soc     (n_days,K+1) 最终储电量轨迹（下标 0 为当日 0:00），kWh
          d_under / d_over  (n_days,K) 欠取量 (x−y)⁺ / 超用量 (y−x)⁺，kWh
          J_plan / J_adj / J_emg / J_day  (n_days,) 计划/调整/紧急费用与合计（主口径 D-12），元
          J_alt_day (n_days,) 对照口径（欠取按全价+违约金）合计，元
          stage_taus (n_days, n_stage) 当天实际执行的决策时刻（供审计）
          status_day (n_days,) 各日全部阶段求解器状态之和（0 表示全部最优）
          e_init / e_end float 滚动起点/终点储电量，kWh
    """
    load = np.asarray(load, dtype=float)
    pv_actual = np.asarray(pv_actual, dtype=float)
    pv_fc144 = np.asarray(pv_fc144, dtype=float)
    if kmark is None:
        kmark = {0: 0, 6: 36, 12: 72, 18: 108}                 # D-13 的覆盖起点（0 基）
    stage_list = [int(t) for t in stages]                      # 当天执行的决策时刻序列
    n_all = load.shape[0]                                      # 全部天数
    if d_end is None:
        d_end = n_all
    d_index = np.arange(d_start, d_end)                        # 参与滚动的天序号
    n_days = d_index.size
    tau_pos = {0: 0, 6: 1, 12: 2, 18: 3}                       # 发布时刻 -> pv_fc144 第 2 维下标

    # 结果容器（全分辨率，逐日逐时段，便于落盘与审计）
    x_plan = np.zeros((n_days, K))                             # 计划购电量 x，kWh
    y_adj = np.zeros((n_days, K))                              # 最终生效购电量 y，kWh
    u_chg = np.zeros((n_days, K))                              # 充电量，kWh
    v_dis = np.zeros((n_days, K))                              # 放电量，kWh
    r_emg = np.zeros((n_days, K))                              # 紧急购电量，kWh
    g_curt = np.zeros((n_days, K))                             # 弃光电量，kWh
    E_soc = np.zeros((n_days, K + 1))                          # 储电量轨迹，kWh
    d_under = np.zeros((n_days, K))                            # 欠取量，kWh
    d_over = np.zeros((n_days, K))                             # 超用量，kWh
    j_plan = np.zeros(n_days)                                  # 计划购电费 Σp·x，元
    j_adj = np.zeros(n_days)                                   # 调整相关费用（有符号），元
    j_emg = np.zeros(n_days)                                   # 紧急购电费 κΣp·r，元
    j_day = np.zeros(n_days)                                   # 主口径合计，元
    j_alt_day = np.zeros(n_days)                               # 对照口径合计，元
    stage_taus = np.zeros((n_days, len(stage_list)), dtype=int)  # 实际执行的决策时刻
    status_day = np.zeros(n_days, dtype=int)                   # 求解器状态累计

    e = float(e_init)                                          # 跨日传递的储电量状态，kWh
    assert 0 in stage_list, "多阶段滚动必须包含 0:00 决策（计划购电量的定义时刻）"
    for i, d in enumerate(d_index):
        # 状态容器：全时段先按"0:00 决策"填满，之后被各调整时刻逐步覆盖
        xx = np.zeros(K); uu = np.zeros(K); vv = np.zeros(K)
        x0 = None                                              # 0:00 计划（偏差结算的固定基准）
        e_cur = e                                              # 当前阶段的起点储电量
        st_sum = 0
        for si, tau in enumerate(stage_list):
            ks = int(kmark[tau])                               # 本阶段覆盖的 0 基起点
            n_seg = K - ks                                     # 本阶段时段数
            if n_seg <= 0:
                continue
            pv_seg = pv_fc144[d, tau_pos[tau], ks:]            # 本阶段的光伏预报（10 分钟粒度）
            # 偏差项参照 X0：0:00 计划阶段无参照（目标全价）；后续阶段一律相对 0:00 计划
            ref_seg = None if x0 is None else x0[ks:]
            res = solve_stage(price[ks:], load[d, ks:], pv_seg, e_init=e_cur,
                              x_ref=ref_seg, k_start=ks, n_period=n_seg,
                              kappa_under=kappa_under, kappa_over=kappa_over,
                              eta=eta, p_max=p_max, e_min=e_min, e_max=e_max, dt_h=dt_h)
            st_sum += res["status"]
            # 写入/覆盖本阶段各时段的值（已执行时段不在本阶段范围内，天然不可追溯）
            xx[ks:] = res["x_plan"]                            # 本阶段决策的购电量
            uu[ks:] = res["u_chg"]                             # 充电量
            vv[ks:] = res["v_dis"]                             # 放电量
            if si == 0:
                x0 = xx.copy()                                 # 0:00 计划：偏差结算基准（D-12）
                x_plan[i] = x0
            # 本阶段终点储电量传给同天下一个决策时刻（须取下一阶段起点的轨迹值，不是当日 24:00）
            if si + 1 < len(stage_list):
                ks_next = int(kmark[stage_list[si + 1]])
                e_cur = float(res["E_soc"][ks_next - ks])      # 下一阶段起点处（已锁定段末）的储电量
            else:
                e_cur = float(res["E_soc"][-1])                # 最后一个阶段：终点即当日 24:00
        # 最终生效购电量 y：本天内每个时段由"该时段之前最晚一次决策"决定（D-13）
        y_adj[i] = xx.copy()
        u_chg[i] = uu; v_dis[i] = vv
        # 最终储能轨迹（由最终充放电量独立重算，便于与仓库解核对）
        E_soc[i] = np.concatenate([[e], e + np.cumsum(eta * uu - vv / eta)])
        # 实际值评估：紧急购电量 = 供给缺口（负值截断为 0）
        r_day = np.maximum(load[d] * dt_h + uu - vv - pv_actual[d] * dt_h - xx, 0.0)
        r_emg[i] = r_day
        g_curt[i] = xx + pv_actual[d] * dt_h + vv - load[d] * dt_h - uu   # 弃光量（实际口径）
        # D-12 结算（主口径与对照口径）
        st = settle_total(price, x_plan[i], y_adj[i], r_day,
                          kappa=kappa, kappa_under=kappa_under, kappa_over=kappa_over)
        j_plan[i] = st["J_plan"]; j_adj[i] = st["J_adj"]; j_emg[i] = st["J_emg"]
        j_day[i] = st["J"]; j_alt_day[i] = st["J_alt"]
        d_under[i] = st["d_under"]; d_over[i] = st["d_over"]
        stage_taus[i, :len(stage_list)] = stage_list
        status_day[i] = st_sum
        e = float(E_soc[i, -1])                                # 传递到次日 0:00（D-04）

    return {
        "d_index": d_index,
        "x_plan": x_plan,
        "y_adj": y_adj,
        "u_chg": u_chg,
        "v_dis": v_dis,
        "r_emg": r_emg,
        "g_curt": g_curt,
        "E_soc": E_soc,
        "d_under": d_under,
        "d_over": d_over,
        "J_plan": j_plan,
        "J_adj": j_adj,
        "J_emg": j_emg,
        "J_day": j_day,
        "J_alt_day": j_alt_day,
        "stage_taus": stage_taus,
        "status_day": status_day,
        "e_init": float(e_init),
        "e_end": float(e),
    }
