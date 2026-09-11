"""单日确定性线性规划求解器（四问共用；问题 1 为退化基线）。

本模块包含两组求解接口：
  1. `solve_day`（问题 1/2 主用）：单日"计划购电量 + 储能充放电"LP，目标 min Σp·x；
  2. `solve_stage`（问题 3/4-3 新增，扩展而非另起一套）：多阶段滚动决策的"单阶段"LP，
     在 solve_day 的供给/储能约束基础上增加"购电量相对参照计划的偏差分解"，
     目标含欠取（−κ₋·Σp·d⁻）与超用（+κ₊·Σp·d⁺）的偏差结算；
  3. `settle_deviation` / `settle_total`：按台账 D-12 的两种结算口径汇总逐时段偏差与分项费用。
三个函数共用同一套储能约束构造（`lib/storage.build_storage_blocks`），不存在第二套储能代码。

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

# ============================== 问题 3 结算参数（题目给定，口径见台账 D-12） ==============================

KAPPA_EMG = 5.0            # κ，紧急购电电价倍数（倍），题目"5 倍电价"
KAPPA_UNDER = 0.5          # κ₋，计划量高于调整量部分（欠取）的违约电价系数（倍）
KAPPA_OVER = 1.5           # κ₊，调整量高于计划量部分（超用）的电价系数（倍）
TOL_ZERO = 1e-9            # 偏差/残差的容差归零阈值（kWh），吸收 LP 数值噪声


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


def solve_stage(price, load, pv_fc, e_init=E_INIT, x_ref=None, k_start=0, n_period=K,
                kappa_under=KAPPA_UNDER, kappa_over=KAPPA_OVER, eta=ETA,
                p_max=P_MAX, e_min=E_MIN, e_max=E_MAX, dt_h=DT_H):
    """多阶段滚动 LP 的"单阶段"求解器（问题 3/4-3 共用；问题 1/2 的 solve_day 保持不变）。

    用途：在某个决策时刻（0:00/6:00/12:00/18:00），对从 k_start 起的 n_period 个**未来时段**
    求解"购电量 y + 储能充放电"的联合决策。目标按 D-12 结算口径写成**边际形式**：

        给定参照计划 x_ref（= 0:00 计划 x，已下达且不减损），最终购电量 y 对总费用的
        贡献为  p·min(x,y) + κ₋p(x−y)⁺ + κ₊p(y−x)⁺。
        拆成"与 y 无关的常数 + 关于 y 的分段线性项"后，优化等价于
            min Σ_k [ κ₋·p_k·y_k + (κ₊−κ₋)·p_k·d⁺_k ] ,  d⁺_k ≥ y_k − x_ref_k,  d⁺ ≥ 0
        即"基础单价 κ₋·p（少买少付、多买多付）+ 超出计划部分再加 (κ₊−κ₋)·p"。
        不做决策（y ≡ x，等价于 y 取下界 0 且 d⁺ 按式取 0）时的目标值
        κ₋Σp·x 恰好对应 J_plan−J_adj，故各阶段目标可比、可直接相加。

    模型（变量顺序 [y(n) | u(n) | v(n) | d⁺(n)]，n = n_period，本阶段首段即 k_start 段）：
        min  (1−κ₋)·Σ_k p_k·y_k + (κ₊+κ₋−1)·Σ_k p_k·d⁺_k
        s.t. y_k + P̂_k·Δ + v_k ≥ L_k·Δ + u_k                （供给不低于负载，同 solve_day）
             0 ≤ u_k, v_k ≤ P̄·Δ
             E 沿本阶段递推（起点 e_init），满足上下限；终端自由（D-04）
             d⁺_k ≥ y_k − x_ref_k,  d⁺_k ≥ 0                （超出计划的部分）
             y ≥ 0
    等价性推导：结算函数 S(y) = p·min(x₀,y) + κ₋p(x₀−y)⁺ + κ₊p(y−x₀)⁺ 对定值 x₀ 是分段线性，
    在 y ≤ x₀ 段斜率 (1−κ₋)p、截距 κ₋p·x₀；在 y ≥ x₀ 段斜率 κ₊p、截距 (1−κ₊)p·x₀。
    两段合写为 (1−κ₋)p·y + (κ₊+κ₋−1)p·(y−x₀)⁺ + κ₋p·x₀（在 y=x₀ 处两段取值相等等于 p·x₀），
    且该式对任意 κ₋ ≤ 1、κ₊ ≥ 0 成立（不限于 κ₋=0.5 的特例，故 S1 的 κ₋ 扰动可直接用本函数）。
    常数项与决策无关，不进入目标。
    x_ref=None（0:00 计划阶段）：目标退化为 min Σp·y（全价），与 solve_day 完全一致。
    d⁺ ≥ y − x_ref 为不等式约束（不是等式）：d⁺ 只出现在"要被最小化"的目标里，
    最优解自动取 d⁺ = (y − x_ref)⁺，故无需整数变量、也不会出现"d⁻ 型"变量。

    输入：price，np.ndarray (n_period,)，本阶段各时段电价，元/kWh
          load，np.ndarray (n_period,)，负载功率，kW
          pv_fc，np.ndarray (n_period,)，本阶段各时段的光伏预报（10 分钟粒度），kW
          e_init，kWh，本阶段起点储电量（= 前段已执行轨迹的当前值）
          x_ref，np.ndarray (n_period,) 或 None，参照计划（0:00 计划 x 的本阶段片段）
          k_start，int，本阶段首个真实时段的 0 基下标（仅用于返回元信息）
          n_period，int，本阶段时段数
          kappa_under / kappa_over，偏差电价系数（倍）
          eta / p_max / e_min / e_max / dt_h，储能参数（同 solve_day）
    输出：dict，键含义——
          x_plan  (n_period,) 本阶段决策的购电量 y，kWh
          u_chg   (n_period,) 充电量，kWh
          v_dis   (n_period,) 放电量，kWh
          E_soc   (n_period+1,) 本阶段储电量轨迹（下标 0 为本阶段起点），kWh
          d_under (n_period,) 相对 x_ref 的欠取量 (x_ref − y)^+，kWh（无参照时为 0）
          d_over  (n_period,) 相对 x_ref 的超用量 (y − x_ref)^+，kWh（无参照时为 0）
          g_curt  (n_period,) 弃光电量，kWh
          obj     float 本阶段 LP 目标值（边际形式，元）；可按阶段相加
          cost_plan float 本阶段全价购电费 Σ p_k·y_k，元（记录用，不参与目标）
          status int / message str 求解器状态
    """
    price = np.asarray(price, dtype=float).reshape(-1)
    load = np.asarray(load, dtype=float).reshape(-1)
    pv_fc = np.asarray(pv_fc, dtype=float).reshape(-1)
    n = int(n_period)
    assert price.size == n and load.size == n and pv_fc.size == n, "price/load/pv_fc 长度必须等于 n_period"
    if x_ref is None:
        has_ref = False
        x_ref = np.zeros(n)
    else:
        has_ref = True
        x_ref = np.asarray(x_ref, dtype=float).reshape(-1)
        assert x_ref.size == n, "x_ref 长度必须等于 n_period"
    u_max = float(p_max) * float(dt_h)                     # 单时段最大充/放电量，kWh

    # ---- 目标系数（D-12 结算的等价分段线性形式，对任意 κ₋/κ₊ 精确）----
    # 恒等式：p·min(x₀,y) + κ₋p(x₀−y)⁺ + κ₊p(y−x₀)⁺ ≡ (1−κ₋)p·y + (κ₊+κ₋−1)p·(y−x₀)⁺ + κ₋p·x₀，
    # 常数项 κ₋Σp·x₀ 不影响优化，故阶段目标取前两项；κ₋=0.5 时退化为 0.5p·y + (κ₊−0.5)p·t⁺。
    if has_ref:
        c_y = (1.0 - float(kappa_under)) * price           # 基础单价：买得越多付得越多
        c_over = (float(kappa_over) + float(kappa_under) - 1.0) * price   # 超出计划部分的追加单价
    else:
        c_y = price                                        # 0:00 计划：全价，与 solve_day 一致
        c_over = np.zeros(n)
    cost_vec = np.concatenate([c_y, np.zeros(2 * n), c_over])

    # ---- 供给约束（不等式）：−y_k + u_k − v_k ≤ Δ(P_k − L_k) ----
    a_supply = hstack([-identity(n, format="csr"), identity(n, format="csr"),
                       -identity(n, format="csr"), csr_matrix((n, n))])
    b_supply = dt_h * (pv_fc - load)

    # ---- 超用约束（不等式）：y_k − d⁺_k ≤ x_ref_k ----
    if has_ref:
        a_over = hstack([identity(n, format="csr"), csr_matrix((n, 2 * n)),
                         -identity(n, format="csr")])
        b_over = x_ref.copy()
    else:
        a_over = csr_matrix((0, 4 * n))                    # 无参照时不加该行组（列数与变量数一致）
        b_over = np.zeros(0)

    # ---- 储能储电量上下界（本阶段起点 e_init，终端自由）----
    a_st, b_st, _, _ = build_storage_blocks(
        n_period=n, eta=eta, e_min=e_min, e_max=e_max, e_init=e_init, mode="free")
    a_storage = hstack([csr_matrix((2 * n, n)), a_st, csr_matrix((2 * n, n))]).tocsr()
    # 把供给、超用、储电量约束纵向拼成完整不等式集合
    a_ub = vstack([a_supply, a_over, a_storage]).tocsr()
    b_ub = np.concatenate([b_supply, b_over, b_st])

    # ---- 变量上下界：[y ≥0 | u,v ∈[0, P̄Δ] | d⁺ ≥ 0] ----
    bounds = [(0.0, None)] * n + power_bounds(n_period=n, u_max=u_max) + [(0.0, None)] * n

    res = linprog(cost_vec, A_ub=a_ub, b_ub=b_ub, bounds=bounds, method="highs")
    if not res.success:
        raise RuntimeError("阶段 LP 求解失败：status=%s，message=%s" % (res.status, res.message))

    # 拆分变量块
    x_plan = res.x[:n]
    u_chg = res.x[n:2 * n]
    v_dis = res.x[2 * n:3 * n]
    d_over = res.x[3 * n:4 * n]
    if not has_ref:                                        # 无参照时超用量按定义恒 0
        d_over = np.zeros(n)
    d_under = np.maximum(x_ref - x_plan, 0.0)              # 欠取量按定义回算（仅供审计）

    e_soc = soc_trajectory(u_chg, v_dis, e_init=e_init, eta=eta)      # 轨迹独立重算，便于核对
    g_curt = x_plan + pv_fc * dt_h + v_dis - load * dt_h - u_chg      # 弃光量（≥0，LP 保证）
    cost_plan = float(np.dot(price, x_plan))                           # 全价购电费（记录用）
    obj = float(res.fun)                                               # 边际目标值（可加）

    return {
        "x_plan": x_plan,
        "u_chg": u_chg,
        "v_dis": v_dis,
        "E_soc": e_soc,
        "d_under": d_under,
        "d_over": d_over,
        "g_curt": g_curt,
        "obj": obj,
        "cost_plan": cost_plan,
        "k_start": int(k_start),
        "status": int(res.status),
        "message": str(res.message),
    }


def settle_deviation(price, x_plan, y_adj, kappa_under=KAPPA_UNDER, kappa_over=KAPPA_OVER):
    """按 D-12 结算口径逐时段计算调整偏差量（x 与 y 的比较），返回偏差量与分项费用。

    结算口径（台账 D-12）：
        J = Σ_k [p_k·min(x_k,y_k) + κ₋p_k(x_k−y_k)⁺ + κ₊p_k(y_k−x_k)⁺] + κΣp·r
    等价分项写法 J = Σp·x + Σ[−κ₋p(x−y)⁺ + κ₊p(y−x)⁺] + κΣp·r，故本函数返回的 dev_cost
    就是"调整购电量的相关费用"（欠取为负、超用为正）。

    输入：price，np.ndarray (n,)，电价，元/kWh
          x_plan / y_adj，np.ndarray (n,)，计划购电量 / 最终生效购电量，kWh
          kappa_under / kappa_over，偏差电价系数（倍）
    输出：dict，键含义——
          d_under (n,) 欠取量 (x−y)⁺，kWh；d_over (n,) 超用量 (y−x)⁺，kWh
          cost_plan float Σp·x，元；cost_billed float Σp·min(x,y)，元
          dev_cost float 调整相关费用（有符号），元
    """
    x_plan = np.asarray(x_plan, dtype=float).reshape(-1)
    y_adj = np.asarray(y_adj, dtype=float).reshape(-1)
    price = np.asarray(price, dtype=float).reshape(-1)
    d_under = np.maximum(x_plan - y_adj, 0.0)              # 计划高于调整的部分（违约）
    d_over = np.maximum(y_adj - x_plan, 0.0)               # 调整高于计划的部分（超用）
    # 容差归零：LP 数值噪声常在 1e-13，物理量在 1e-9 以下一律记 0
    d_under = np.where(np.abs(d_under) < TOL_ZERO, 0.0, d_under)
    d_over = np.where(np.abs(d_over) < TOL_ZERO, 0.0, d_over)
    cost_plan = float(np.dot(price, x_plan))               # 计划购电费用，元
    cost_billed = float(np.dot(price, np.minimum(x_plan, y_adj)))   # 按实际取用计价部分，元
    dev_cost = float(np.dot(kappa_over * price, d_over) - np.dot(kappa_under * price, d_under))
    return {
        "d_under": d_under,
        "d_over": d_over,
        "cost_plan": cost_plan,
        "cost_billed": cost_billed,
        "dev_cost": dev_cost,
    }


def settle_total(price, x_plan, y_adj, r_emg, kappa=KAPPA_EMG, kappa_under=KAPPA_UNDER,
                 kappa_over=KAPPA_OVER):
    """按 D-12 汇总全天总费用，并给出主口径与对照口径的分项（问题 3/4-3 共用）。

    主口径（D-12，用户已确认）：
        J = Σ[p·min(x,y) + κ₋p(x−y)⁺ + κ₊p(y−x)⁺] + κΣp·r
          = J_plan + J_adj + J_emg，其中 J_plan = Σp·x、J_adj 为调整相关费用（有符号）
    对照口径（必须报出）：欠取部分按"计划量全价照付 + 50% 违约金"处理：
        J_alt = Σp·x + Σ[κ₋p(x−y)⁺ + κ₊p(y−x)⁺] + κΣp·r
    两个口径在 y=x 时都退化为 Σp·x（与问题 2 严格一致）。

    输入：price (n,) 元/kWh；x_plan/y_adj/r_emg (n,) kWh；kappa/kappa_under/kappa_over 倍数
    输出：dict——
          J_plan / J_adj / J_emg / J（主口径分项与合计，元）
          J_alt（对照口径合计，元）；d_under/d_over (n,) 偏差量，kWh
          qty_x / qty_y / qty_r float 计划量 / 最终量 / 紧急购电量合计，kWh
    """
    st = settle_deviation(price, x_plan, y_adj, kappa_under, kappa_over)
    price = np.asarray(price, dtype=float).reshape(-1)
    r_emg = np.asarray(r_emg, dtype=float).reshape(-1)
    j_emg = float(kappa) * float(np.dot(price, r_emg))     # 紧急购电费用，元
    j_plan = st["cost_plan"]
    j_adj = st["dev_cost"]
    j_main = j_plan + j_adj + j_emg                        # 主口径总费用，元
    # 对照口径：欠取部分支付 κ₋p（违约金），但计划量本身仍按全价 p 支付
    j_alt = j_plan + float(np.dot(price, kappa_under * st["d_under"]
                                  + kappa_over * st["d_over"])) + j_emg
    return {
        "J_plan": j_plan,
        "J_adj": j_adj,
        "J_emg": j_emg,
        "J": j_main,
        "J_alt": j_alt,
        "d_under": st["d_under"],
        "d_over": st["d_over"],
        "qty_x": float(np.sum(x_plan)),
        "qty_y": float(np.sum(y_adj)),
        "qty_r": float(np.sum(r_emg)),
    }
