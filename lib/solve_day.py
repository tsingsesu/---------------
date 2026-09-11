"""单日确定性线性规划求解器（四问共用；问题 1 为退化基线）。

模型（问题 1，符号与 `符号表.md` 一致）：

    min  Σ_{k=1..K} p_k · x_k
    s.t. x_k + P_k·Δ + v_k ≥ L_k·Δ + u_k        （供给不低于负载；等式则光伏富余时无解）
         0 ≤ u_k ≤ P̄·Δ ,  0 ≤ v_k ≤ P̄·Δ
         E_k = E_{k-1} + η·u_k − v_k/η ,  E_0 = E_INIT
         E_MIN ≤ E_k ≤ E_MAX
         mode='cyclic' 时追加 E_K = E_0    （题目"0:00 与 24:00 储电量相同"）
         x_k, u_k, v_k ≥ 0

要点说明：
  1. 供给约束写成**不等式**、x_k ≥ 0，即"微网供给不得低于负载"；富余的光伏直接丢弃
     （弃光量 g_k = x_k + P_k·Δ + v_k − L_k·Δ − u_k ≥ 0），目标函数中没有售电收益项，
     对应 `口径与假设台账.md` D-16.1（不得向电网售电）。
  2. 允许从电网购电给储能充电（D-16.2），模型里体现为 u_k 只受功率上限约束，
     不受光伏余量约束。
  3. 当 η² < 1 时存在最优解满足 u_k·v_k = 0（同时充放严格劣），故 LP 松弛即可，
     不需要引入整数变量；求解后本模块会统计同时充放时段数以便核验。
  4. 目标函数不含 u_k、v_k，二者只通过约束进入，故对偶变量可由求解器直接给出，
     用于解释"为什么某些时段购电量为 0"。

变量顺序（列）：[x_plan (K) | u_chg (K) | v_dis (K)]，共 3K 列。
"""

import numpy as np
from scipy.optimize import linprog
from scipy.sparse import csr_matrix, hstack, vstack, identity

from lib.timegrid import DT_H, K
from lib.storage import (ETA, E_INIT, E_MAX, E_MIN, P_MAX, U_MAX,
                         build_storage_blocks, power_bounds, soc_trajectory)


def solve_day(price, load, pv_plan, e_init=E_INIT, mode="cyclic", eta=ETA,
              p_max=P_MAX, e_min=E_MIN, e_max=E_MAX, dt_h=DT_H):
    """求解单日"计划购电量 + 储能充放电"线性规划。

    输入：price，np.ndarray (K,)，电价，元/kWh
          load，np.ndarray (K,)，小区负载功率，kW
          pv_plan，np.ndarray (K,)，本日采用的光伏功率（问题1 为预测值），kW
          e_init，kWh，E_0（0:00 储电量）
          mode，str，'cyclic'（锁定 E_K = E_0）或 'free'（端点自由，对照口径）
          eta，无量纲，单向充放电效率
          p_max，kW，最大充放电功率（作用于单时段充/放电量的变量上下界 u,v ≤ p_max·Δt）
          e_min / e_max，kWh，储电量允许下/上限
          dt_h，h，时段长度（Δ）
    输出：dict，键含义——
          x_plan  (K,)  计划购电量，kWh
          u_chg   (K,)  充电量，kWh
          v_dis   (K,)  放电量，kWh
          E_soc   (K+1,) 储电量轨迹，kWh（下标 0 为 0:00，下标 K 为 24:00）
          g_curt  (K,)  弃光电量，kWh（供给约束的富余量）
          cost   float  全天购电费 Σ p_k x_k，元
          pv_marginal (K,) 光伏边际价值，元/kWh（对偶信息，正数表示多用 1 kWh 光伏可省的购电费）
          status int / message str  求解器状态
    """
    # 统一转成 float 一维数组，避免调用方传入 list 或整型数组时出现类型问题
    price = np.asarray(price, dtype=float).reshape(-1)
    load = np.asarray(load, dtype=float).reshape(-1)
    pv_plan = np.asarray(pv_plan, dtype=float).reshape(-1)
    n_period = price.size                                    # 本日时段数（问题 1 为 144）
    assert load.size == n_period and pv_plan.size == n_period, "price/load/pv 长度必须一致"
    u_max = float(p_max) * float(dt_h)                       # 单时段最大充/放电量，kWh

    # ---- 目标函数：min Σ p_k x_k；充电量与放电量的系数为 0 ----
    cost_vec = np.concatenate([price, np.zeros(2 * n_period)])

    # ---- 供给约束（不等式）：−x_k + u_k − v_k ≤ Δ(P_k − L_k) ----
    a_supply = hstack([-identity(n_period, format="csr"),
                       identity(n_period, format="csr"),
                       -identity(n_period, format="csr")])
    # 右端项 = Δ(P_k − L_k)：光伏多于负载时右端为正，该时段允许少购电甚至不购电
    b_supply = dt_h * (pv_plan - load)

    # ---- 储能储电量上下界与终端条件（列顺序 [u | v]，故左侧补 K 列全零给 x） ----
    a_st, b_st, a_eq_st, b_eq_st = build_storage_blocks(
        n_period=n_period, eta=eta, e_min=e_min, e_max=e_max, e_init=e_init, mode=mode)
    a_storage = hstack([csr_matrix((2 * n_period, n_period)), a_st]).tocsr()

    # 把供给约束与储电量约束纵向拼成完整的不等式约束集合
    a_ub = vstack([a_supply, a_storage]).tocsr()
    b_ub = np.concatenate([b_supply, b_st])

    # 终端等式约束（cyclic）同样需要在左侧补 x 列的全零块
    if a_eq_st is not None:
        a_eq = hstack([csr_matrix((a_eq_st.shape[0], n_period)), a_eq_st]).tocsr()
        b_eq = np.asarray(b_eq_st, dtype=float)
    else:
        a_eq, b_eq = None, None

    # ---- 变量上下界：[x 无上界非负 | u ∈[0,u_max] | v ∈[0,u_max]] ----
    bounds = [(0.0, None)] * n_period + power_bounds(n_period=n_period, u_max=u_max)

    # ---- 求解：HiGHS 求解器，完全确定性问题，无需随机种子 ----
    res = linprog(cost_vec, A_ub=a_ub, b_ub=b_ub, A_eq=a_eq, b_eq=b_eq,
                  bounds=bounds, method="highs")

    # 求解失败时直接抛出，避免把坏结果当结果落盘
    if not res.success:
        raise RuntimeError("单日 LP 求解失败：status=%s，message=%s" % (res.status, res.message))

    # 拆分变量块：前 K 个是计划购电量，中间 K 个是充电量，最后 K 个是放电量
    x_plan = res.x[:n_period]
    u_chg = res.x[n_period:2 * n_period]
    v_dis = res.x[2 * n_period:3 * n_period]

    # 由充放电量递推储电量轨迹（不依赖求解器返回的约束值，便于独立核对）
    e_soc = soc_trajectory(u_chg, v_dis, e_init=e_init, eta=eta)

    # 弃光量 = 供给 − 需求 = x + P·Δ + v − L·Δ − u，理论上恒 ≥ 0（LP 保证）
    g_curt = x_plan + pv_plan * dt_h + v_dis - load * dt_h - u_chg

    # 光伏边际价值：供给约束右端项 b_k = Δ(P_k − L_k) 的单位本身就是 kWh，
    # 故对偶变量取负号即为"多 1 kWh 光伏可省的购电费"，单位 元/kWh
    pv_marginal = -res.ineqlin.marginals[:n_period]

    # 全天购电费按定义重算，用 4 位小数四舍五入后再取浮点，避免求解器残差带来的末位噪声
    cost = float(np.dot(price, x_plan))

    return {
        "x_plan": x_plan,
        "u_chg": u_chg,
        "v_dis": v_dis,
        "E_soc": e_soc,
        "g_curt": g_curt,
        "cost": cost,
        "pv_marginal": pv_marginal,
        "status": int(res.status),
        "message": str(res.message),
    }


def supply_residual(x_plan, load, pv_plan, u_chg, v_dis, dt_h=DT_H):
    """核对供给约束：返回 x + P·Δ + v − L·Δ − u（应恒 ≥ 0）。

    输入：x_plan/u_chg/v_dis，np.ndarray (K,)，kWh；load/pv_plan，np.ndarray (K,)，kW；dt_h，h
    输出：np.ndarray (K,)，残差，kWh；负值表示违反"供给不低于负载"
    """
    # 逐项按定义计算，不做任何化简，便于外部审计逐条对照
    return (np.asarray(x_plan, dtype=float) + np.asarray(pv_plan, dtype=float) * dt_h
            + np.asarray(v_dis, dtype=float) - np.asarray(load, dtype=float) * dt_h
            - np.asarray(u_chg, dtype=float))


def charge_discharge_summary(u_chg, v_dis, eta=ETA):
    """汇总充放电量，并给出放电量/充电量的比值（理论值等于 η²）。

    输入：u_chg / v_dis，np.ndarray (K,)，kWh；eta，无量纲
    输出：dict，含 u_total（总充电量）、v_total（总放电量）、ratio（比值）、eta_sq（η²）
    """
    u_total = float(np.sum(u_chg))                # 全天充电量合计，kWh
    v_total = float(np.sum(v_dis))                # 全天放电量合计，kWh
    # 在 E_0 = E_K 的闭环下，Σ(ηu − v/η) = 0 ⇒ Σv = η² Σu，故比值应等于 η²
    ratio = v_total / u_total if u_total > 0 else float("nan")
    return {"u_total": u_total, "v_total": v_total, "ratio": ratio, "eta_sq": eta ** 2}


def baseline_costs(price, load, pv_plan, dt_h=DT_H):
    """计算两个不储能基线方案的购电费（题目解读 §1.3 的方案 A / 方案 B）。

    方案 A：光伏不参与，全部负载由电网供给，费用 = Σ p_k·L_k·Δ
    方案 B：光伏优先自用、余电弃掉（不储能），费用 = Σ p_k·max(L_k−P_k, 0)·Δ

    输入：price (K,) 元/kWh；load (K,) kW；pv_plan (K,) kW；dt_h，h
    输出：dict，{'cost_all_grid' 元, 'cost_pv_only' 元, 'load_energy' kWh, 'pv_energy' kWh,
                'net_load_energy' kWh, 'buy_energy_pv_only' kWh}
          net_load_energy 为**有符号**净负载电量 Σ(L−P)Δ；buy_energy_pv_only 为方案 B 实际
          向电网购入的电量 Σmax(L−P,0)Δ（二者之差即被弃掉的光伏电量）
    """
    price = np.asarray(price, dtype=float)
    load = np.asarray(load, dtype=float)
    pv_plan = np.asarray(pv_plan, dtype=float)
    # 方案 A：全部电量来自电网
    cost_all_grid = float(np.sum(price * load * dt_h))
    # 方案 B：净负荷 = 负载 − 光伏（负值截断为 0，即富余光伏弃掉）
    net_load = np.maximum(load - pv_plan, 0.0)
    cost_pv_only = float(np.sum(price * net_load * dt_h))
    return {
        "cost_all_grid": cost_all_grid,
        "cost_pv_only": cost_pv_only,
        "load_energy": float(np.sum(load) * dt_h),           # 负载电量，kWh
        "pv_energy": float(np.sum(pv_plan) * dt_h),          # 光伏电量，kWh
        "net_load_energy": float(np.sum(load - pv_plan) * dt_h),   # 有符号净负载电量，kWh
        "buy_energy_pv_only": float(np.sum(net_load) * dt_h),      # 方案 B 的购电量，kWh
    }
