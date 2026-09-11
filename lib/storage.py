"""储能设备模型（全工作区共用）：状态方程、储电量边界与充放电功率约束的构造。

参数来源：题目附录 1（见 `题目解读.md` §1.7）：
  额定容量 C=12000 kWh；最大充放电功率 P̄=5000 kW；允许储电量区间 [1200, 10800] kWh；
  单向充放电效率 η=0.9（往返效率 0.81，口径见 `口径与假设台账.md` D-05）；
  2025-01-01 0:00 储电量 E_0 = 6000 kWh。

状态方程（本工作区统一口径）：
  E_k = E_{k-1} + η·u_k − v_k/η，其中 u_k 为充电量、v_k 为放电量（均 kWh，单时段内）。

单位约定：电量一律 kWh，功率一律 kW，时段长度 Δ=1/6 h，故单时段最大充/放电量 = P̄·Δ。
"""

import numpy as np
from scipy.sparse import csr_matrix, hstack, vstack

from lib.timegrid import DT_H, K

# ============================== 储能参数（附录 1） ==============================

ETA = 0.9                 # 单向充/放电效率，无量纲（往返效率 = ETA**2 = 0.81）
P_MAX = 5000.0            # 最大充、放电功率，kW
E_CAP = 12000.0           # 额定容量 C，kWh
E_MIN = 1200.0            # 储电量允许下限，kWh
E_MAX = 10800.0           # 储电量允许上限，kWh
E_INIT = 6000.0           # 初始储电量 E_0（2025-01-01 0:00），kWh
U_MAX = P_MAX * DT_H      # 单时段最大充电量（= 单时段最大放电量），kWh，= 833.3333


def soc_step(e_prev, u_chg, v_dis, eta=ETA):
    """单时段状态递推：由上一时段末储电量与充放电量算本时段末储电量。

    输入：e_prev，kWh，上一时段末储电量（标量或数组）
          u_chg，kWh，本时段充电量
          v_dis，kWh，本时段放电量
          eta，无量纲，单向效率
    输出：e_next，kWh，本时段末储电量（与输入同形状）
    """
    # 充电使储电量增加 eta*u，放电使储电量减少 v/eta（放电需多消耗储能以补偿损耗）
    return np.asarray(e_prev, dtype=float) + eta * np.asarray(u_chg, dtype=float) \
        - np.asarray(v_dis, dtype=float) / eta


def soc_trajectory(u_chg, v_dis, e_init=E_INIT, eta=ETA):
    """由充电量序列与放电量序列递推整条储电量轨迹。

    输入：u_chg，np.ndarray，shape (K,)，充电量，kWh
          v_dis，np.ndarray，shape (K,)，放电量，kWh
          e_init，kWh，E_0（0:00 储电量）
          eta，无量纲，单向效率
    输出：np.ndarray，shape (K+1,)，E_0..E_K（下标 0 为 0:00，下标 K 为 24:00），kWh
    """
    # 每时段的净储电量增量 = eta*u - v/eta（kWh），累加后加上初值即得轨迹
    delta_e = eta * np.asarray(u_chg, dtype=float) - np.asarray(v_dis, dtype=float) / eta
    return np.concatenate([[float(e_init)], float(e_init) + np.cumsum(delta_e)])


def build_storage_blocks(n_period=K, eta=ETA, e_min=E_MIN, e_max=E_MAX, e_init=E_INIT,
                         mode="cyclic"):
    """构造储能部分的线性约束块，供单日 LP 使用（列顺序固定为 [充电量 | 放电量]）。

    约束形式（均为 A_ub·z <= b_ub）：
      * 上界：E_k <= e_max  <=>  Σ_{j<=k}(eta·u_j − v_j/eta) <= e_max − e_init，k=1..K
      * 下界：E_k >= e_min  <=>  −Σ_{j<=k}(eta·u_j − v_j/eta) <= e_init − e_min，k=1..K
    等式的终端条件（mode='cyclic'）：
      * E_K = E_0 = e_init <=> Σ_{j<=K}(eta·u_j − v_j/eta) = 0

    输入：n_period，int，时段数（问题 1 为 144）
          eta，无量纲，单向效率
          e_min / e_max，kWh，储电量允许下/上限
          e_init，kWh，E_0
          mode，str，'cyclic' 表示锁定 E_K = E_0（问题 1 主模型）；'free' 表示端点自由（对照）
    输出：(A_ub, b_ub, A_eq, b_eq)，
          A_ub 为 (2K, 2K) 稀疏矩阵（未知量顺序 [u_1..u_K, v_1..v_K]），b_ub 为 (2K,) 数组；
          A_eq 为 (1, 2K) 稀疏矩阵或 None，b_eq 为 (1,) 数组或 None
    """
    # 下三角全 1 矩阵：第 k 行对 j<=k 的列取 1，用于表达"前缀和"型约束
    tri = csr_matrix(np.tril(np.ones((n_period, n_period), dtype=float)))

    # 上界约束：tri @ (eta*u − v/eta) <= e_max − e_init
    a_upper = hstack([tri * eta, tri * (-1.0 / eta)])
    # 下界约束：−tri @ (eta*u − v/eta) <= e_init − e_min
    a_lower = hstack([tri * (-eta), tri * (1.0 / eta)])

    a_ub = vstack([a_upper, a_lower]).tocsr()         # (2K, 2K)
    # 右端项：前 K 行是上界余量，后 K 行是下界余量（均以 e_init 为基准）
    b_ub = np.concatenate([np.full(n_period, e_max - e_init),
                           np.full(n_period, e_init - e_min)])

    # 终端条件：cyclic 时强制整日净储电量变化为 0（即 E_K = E_0）
    if mode == "cyclic":
        a_eq = hstack([csr_matrix(np.full((1, n_period), eta)),
                       csr_matrix(np.full((1, n_period), -1.0 / eta))]).tocsr()
        b_eq = np.zeros(1)
    else:
        # 端点自由：不施加终端等式，E_K 由优化自行决定（对照口径，D-03）
        a_eq, b_eq = None, None

    return a_ub, b_ub, a_eq, b_eq


def power_bounds(n_period=K, u_max=U_MAX):
    """充放电量的变量上下界（充、放电各自独立，均落在 [0, u_max]）。

    输入：n_period，int，时段数
          u_max，kWh，单时段最大充/放电量（= P̄·Δ）
    输出：list[tuple[float, float]]，长度 2K，前 K 个是充电量下上界，后 K 个是放电量下上界
    """
    # 充电量与放电量都是非负变量，且单时段不得超过功率上限折算出的电量
    return [(0.0, float(u_max))] * (2 * n_period)


def check_soc_limits(e_soc, e_min=E_MIN, e_max=E_MAX, tol=1e-6):
    """核验储电量轨迹是否满足边界约束。

    输入：e_soc，np.ndarray，shape (K+1,)，储电量轨迹，kWh
          e_min / e_max，kWh，允许下/上限
          tol，kWh，容差
    输出：(ok, 最小越限值, 最大越限值)；越限值为正表示违反
    """
    # 低于下限的越限量与高于上限的越限量分别取最大，均为非负才通过
    lower_violation = float(np.max(e_min - np.asarray(e_soc, dtype=float)))
    upper_violation = float(np.max(np.asarray(e_soc, dtype=float) - e_max))
    ok = (lower_violation <= tol) and (upper_violation <= tol)
    return ok, lower_violation, upper_violation
