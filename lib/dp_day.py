"""离散化动态规划（DP）求解器：全工作区共用的独立复算工具。

用途：作为单日 LP（`lib/solve_day.py`）的**方法侧互验**——同一问题换一种算法重算，
量化网格离散化带来的费用上偏；问题 1 的 S4a 用它，问题 2–4 可直接复用。

适用性论证（为什么离散化 DP 对本题**精确**）：
  1. 目标函数 min Σ p_k x_k 与全部约束对连续变量 (x_k, u_k, v_k) 都是线性的，且
     price_k ≥ 0 ⇒ 除了负载平衡所必需的那部分，购电量取 0 最省；
     而储电量轨迹 E_0..E_K 一旦给定，各时段所需购电量立刻确定：
         x_k = max( L_k·Δ + u_k − v_k − P_k·Δ , 0 ) ,
         u_k = max( E_k − E_{k-1} , 0 ) / η ,  v_k = max( E_{k-1} − E_k , 0 ) · η 。
     故只需在**储电量网格点**上做动态规划；网格点之间的最优轨迹由线性插值取得，
     插值轨迹上的各项取值恰为上式的线性组合，可行性与目标值均可精确计算。
  2. 因此在给定网格上，DP 求得的费用就是"储电量只能取网格点"这一限制下的**全局最优**；
     随网格加密，费用单调下降并收敛到 LP 最优值（LP 的最优储电量未必落在网格点上，
     故 DP 费用 ≥ LP 费用，差值即离散化损失——这正是互验要量化的量）。
  3. 迁移的可行性判据：E_k 与 E_{k-1} 都在网格点上时，所需的 u_k、v_k 由上式唯一确定，
     只要 u_k ≤ P̄·Δ 且 v_k ≤ P̄·Δ 即可行；负载平衡由 x_k ≥ 0 的截断保证
     （x 取小到不够时，截断为 0 即用弃光/少充的富余量兜底，与 LP 的供给不等式一致）。

依赖：仅 numpy。随机性：无（完全确定性问题，不需要随机种子）。
"""

import numpy as np

from lib.storage import E_INIT, E_MAX, E_MIN, ETA, P_MAX
from lib.timegrid import DT_H


def dp_solve(price, load, pv_plan, n_grid, e_init=E_INIT, eta=ETA, p_max=P_MAX,
             e_min=E_MIN, e_max=E_MAX, dt_h=DT_H):
    """在储电量网格上做动态规划，求单日最优购电费（问题 1 端点锁定口径）。

    输入：price，np.ndarray (K,)，电价，元/kWh
          load，np.ndarray (K,)，负载功率，kW
          pv_plan，np.ndarray (K,)，光伏功率（问题1 为预测值），kW
          n_grid，int，储电量网格**区间数**（网格点数为 n_grid+1，步长 = (e_max−e_min)/n_grid）
          e_init，kWh，E_0（0:00 储电量，须落在网格点上）
          eta，无量纲，单向充放电效率
          p_max，kW，最大充放电功率
          e_min / e_max，kWh，储电量允许下/上限（网格端点）
          dt_h，h，时段长度 Δ
    输出：dict，键含义——
          cost      float，最优购电费 Σ p_k x_k，元
          E_soc     np.ndarray (K+1,)，最优网格储电量轨迹（E_0..E_K，kB 前插 E_0），kWh
          u_chg / v_dis  np.ndarray (K,)，由轨迹反推的充/放电量，kWh
          x_plan    np.ndarray (K,)，由轨迹反推的计划购电量，kWh
          n_grid    int，网格区间数（记录用）
          infeasible bool，环路不可行时为 True（正常参数下不会触发）
    """
    price = np.asarray(price, dtype=float).reshape(-1)
    load = np.asarray(load, dtype=float).reshape(-1)
    pv_plan = np.asarray(pv_plan, dtype=float).reshape(-1)
    n_period = price.size                                   # 时段数 K

    # ---- 1. 构造储电量网格，并把 E_0 对齐到最近网格点 ----
    grid = np.linspace(float(e_min), float(e_max), int(n_grid) + 1)  # 网格点数组，kWh
    step = (float(e_max) - float(e_min)) / float(n_grid)             # 网格步长，kWh
    idx0 = int(round((float(e_init) - float(e_min)) / step))         # E_0 的网格下标
    grid = grid.copy()
    grid[idx0] = float(e_init)                              # 把 E_0 精确钉在网格上（消除舍入）
    n_point = grid.size                                     # 网格点数 = n_grid+1

    # ---- 2. 预算全体状态转移（只依赖网格，与时段无关，外提以加速） ----
    e_prev = grid[:, None]                                  # 迁移起点（行），(n_point,1)
    e_next = grid[None, :]                                  # 迁移终点（列），(1,n_point)
    d_e = e_next - e_prev                                   # 迁移引起的储电量变化，kWh
    # 储电量上升 ⇒ 充电：u = ΔE/η；下降 ⇒ 放电：v = −ΔE·η；反推公式与状态方程互为逆运算
    u_step = np.where(d_e > 0.0, d_e / eta, 0.0)            # 该迁移所需充电量，kWh
    v_step = np.where(d_e < 0.0, -d_e * eta, 0.0)           # 该迁移所需放电量，kWh
    u_lim = float(p_max) * float(dt_h)                      # 单时段最大充/放电量，kWh
    # 功率可行掩码：充、放电量都不得超过 P̄·Δ（容差 1e-9 吸收网格浮点噪声）
    feasible = (u_step <= u_lim + 1e-9) & (v_step <= u_lim + 1e-9)

    # ---- 3. 逐时段前推值函数（每列是前一列所有迁移取最小） ----
    value = np.full(n_point, np.inf)                        # 值函数 f_k(E)：到达 E 的最小累计费用
    value[idx0] = 0.0                                       # 初始状态 E_0 的费用为 0
    argmin_prev = np.zeros((n_period, n_point), dtype=np.int32)   # 回溯表：每时段每状态的迁移起点
    for k in range(n_period):
        # 该时段的需求侧净电量：负载 − 光伏，kWh（可为负，表示光伏富余）
        net = (load[k] - pv_plan[k]) * dt_h
        # 给定迁移 (E_{k-1}→E_k) 后，负载平衡所需的最小购电量（不足部分由弃光兜底）
        x_min = np.maximum(net + u_step - v_step, 0.0)      # kWh，与 LP 的供给不等式 x ≥ 需要量 对应
        candidate = np.where(feasible, value[:, None] + price[k] * x_min, np.inf)  # 累计费用候选
        argmin_prev[k] = np.argmin(candidate, axis=0)       # 每列（本时段末状态）的最优前驱行号
        value = candidate.min(axis=0)                       # 本时段的值函数（沿行取最小）

    # ---- 4. 端点锁定：24:00 的储电量必须回到 E_0（题目"0:00 与 24:00 相同"） ----
    total_cost = value[idx0]
    if not np.isfinite(total_cost):
        return {"cost": float("inf"), "E_soc": None, "u_chg": None, "v_dis": None,
                "x_plan": None, "n_grid": int(n_grid), "infeasible": True}

    # ---- 5. 回溯：由 argmin_prev 反推最优网格轨迹 ----
    idx_traj = np.zeros(n_period + 1, dtype=np.int32)       # 轨迹的网格下标，下标 0 为 E_0
    idx_traj[n_period] = idx0                               # 终点回到 E_0
    for k in range(n_period - 1, -1, -1):                   # 从最后一段往前回溯
        idx_traj[k] = argmin_prev[k, idx_traj[k + 1]]       # 第 k 段起点的最优网格点
    e_soc = grid[idx_traj]                                  # 储电量轨迹，kWh

    # ---- 6. 由轨迹反推充放电量与计划购电量（与迁移约束同一组公式，便于落盘核验） ----
    d_traj = np.diff(e_soc)                                 # 逐时段 ΔE，kWh
    u_chg = np.where(d_traj > 0.0, d_traj / eta, 0.0)       # 充电量，kWh
    v_dis = np.where(d_traj < 0.0, -d_traj * eta, 0.0)      # 放电量，kWh
    net_arr = (load - pv_plan) * dt_h                       # 逐时段净负荷电量，kWh
    x_plan = np.maximum(net_arr + u_chg - v_dis, 0.0)       # 计划购电量，kWh

    return {
        "cost": float(total_cost),
        "E_soc": e_soc,
        "u_chg": u_chg,
        "v_dis": v_dis,
        "x_plan": x_plan,
        "n_grid": int(n_grid),
        "infeasible": False,
    }
