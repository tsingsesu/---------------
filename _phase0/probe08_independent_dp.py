"""构思手的独立复核（与编程手互不共享代码）：

1. 用"离散化动态规划（DP）"这一与 LP 完全不同的算法复算问题 1，检验 LP 结果；
2. 独立核验 LP 解的可行性（供给、功率、储电量边界、状态递推）；
3. 独立核验"填法 Y-轮转"映射：模板行 i ↔ 当天第 (i+1)%144+1 个时段，并取出表 1 六个时段的购电量。

DP 的推导（本题特有的精确降维，也是"无同时充放"引理的构造性证明）：
  给定相邻两个时段末储电量之差 dE = E_k - E_{k-1}，要满足 dE = eta*u - v/eta（u,v>=0 且 <= U_MAX）。
  由于 u 只会抬高该时段的供电需求、v 只会降低它，固定 dE 时最优必取单方向：
     dE >= 0 : u = dE/eta, v = 0        -> 购电量 x = max(0, a_k + dE/eta)
     dE <  0 : u = 0,       v = -eta*dE -> 购电量 x = max(0, a_k + eta*dE)
  其中 a_k = (L_k - P_k)*Delta 为该时段净负荷电量。可达的 dE 区间 = [-U_MAX/eta, eta*U_MAX]。
  于是 DP 阶段代价 c_k(dE) 已知，状态只需储电量 E。
"""

import os
import numpy as np
import openpyxl
from scipy.optimize import linprog
from scipy.sparse import lil_matrix

WORK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
P = lambda *a: os.path.join(WORK, *a)
out = []
def w(s=""):
    out.append(str(s))


def load_sheet(path, sheet=None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


r1 = load_sheet(P("附件", "附件1.xlsx"))
time_lab = [r[0] for r in r1[1:]]
price = np.array([float(r[1]) for r in r1[1:]])      # 电价，元/kWh
load = np.array([float(r[2]) for r in r1[1:]])       # 负载，kW
pv = np.array([float(r[3]) for r in r1[1:]])         # 光伏预测功率，kW

K = 144
DT = 1.0 / 6.0
ETA = 0.9
P_MAX = 5000.0
U_MAX = P_MAX * DT                                    # 单时段最大充/放电量，kWh
E_CAP = 12000.0
E_MIN, E_MAX = 1200.0, 10800.0
E_INIT = 6000.0
NET = (load - pv) * DT                                # 净负荷电量，kWh（正=需外购，负=光伏富余）


# ---------------- 1. LP（与 probe06 同一模型，此处重新独立写一遍） ----------------
def solve_lp(mode="cyclic", eta=ETA):
    n = 3 * K
    c = np.concatenate([price, np.zeros(2 * K)])       # 目标：min Σ p_k x_k
    A = lil_matrix((K + 2 * K, n))
    b = np.zeros(K + 2 * K)
    for k in range(K):
        A[k, k] = -1.0                                 # -x_k
        A[k, K + k] = 1.0                              # +u_k
        A[k, 2 * K + k] = -1.0                         # -v_k
        b[k] = -NET[k]                                 # u - x - v <= -净负荷  <=>  x + pv*DT + v >= load*DT + u
    for k in range(K):
        for j in range(k + 1):
            # 上界行：E_k - E_INIT = Σ(eta*u_j - v_j/eta) <= E_MAX - E_INIT
            A[K + k, K + j] = eta
            A[K + k, 2 * K + j] = -1.0 / eta
            # 下界行：-(E_k - E_INIT) <= E_INIT - E_MIN
            A[K + K + k, K + j] = -eta
            A[K + K + k, 2 * K + j] = 1.0 / eta
        b[K + k] = E_MAX - E_INIT
        b[K + K + k] = E_INIT - E_MIN
    Aeq = beq = None
    if mode == "cyclic":
        Aeq = lil_matrix((1, n))
        for j in range(K):
            Aeq[0, K + j] = eta
            Aeq[0, 2 * K + j] = -1.0 / eta
        beq = np.array([0.0])
    bounds = [(0, None)] * K + [(0, U_MAX)] * K + [(0, U_MAX)] * K
    res = linprog(c, A_ub=A.tocsr(), b_ub=b,
                  A_eq=None if Aeq is None else Aeq.tocsr(), b_eq=beq,
                  bounds=bounds, method="highs")
    assert res.status == 0, res.message
    x = res.x[:K]; u = res.x[K:2 * K]; v = res.x[2 * K:]
    E = np.concatenate([[E_INIT], E_INIT + np.cumsum(eta * u - v / eta)])
    return dict(cost=float(res.fun), x=x, u=u, v=v, E=E)


# ---------------- 2. 离散化 DP ----------------
def solve_dp(step=10.0, mode="cyclic", eta=ETA, e_min=E_MIN, e_max=E_MAX, e_init=E_INIT):
    grid = np.arange(e_min, e_max + 1e-9, step)
    N = len(grid)
    d_min = -U_MAX / eta                                # 最大回充能力（放电使储电量下降）
    d_max = eta * U_MAX                                 # 最大充电能力
    D = grid[None, :] - grid[:, None]                   # D[i, j] = E_j - E_i（行=源，列=目标）
    ok = (D >= d_min - 1e-9) & (D <= d_max + 1e-9)
    eff = np.where(D >= 0, D / eta, eta * D)            # 净负荷的等效增量，kWh
    INF = 1e18
    i0 = int(np.argmin(np.abs(grid - e_init)))
    assert abs(grid[i0] - e_init) < 1e-9, "初始储电量必须落在网格上"
    V = np.full(N, INF)
    V[i0] = 0.0
    for k in range(K):
        cc = price[k] * np.maximum(0.0, NET[k] + eff)   # 该阶段购电费用，元
        cand = V[:, None] + cc
        cand[~ok] = INF
        V = cand.min(axis=0)
    if mode == "cyclic":
        j = int(np.argmin(np.abs(grid - e_init)))
        return float(V[j]), grid
    return float(V.min()), grid


# ---------------- 3. 独立可行性核验 ----------------
def check_feasible(R, tol=1e-6):
    x, u, v, E = R["x"], R["u"], R["v"], R["E"]
    msgs = []
    lhs = x + pv * DT + v - load * DT - u              # 供给 - 需求，应 >= -tol
    msgs.append(("供给不低于负载 min(x+pv*DT+v-load*DT-u) >= -1e-6", float(lhs.min()), bool(lhs.min() >= -tol)))
    msgs.append(("充电量 0<=u<=U_MAX", float(u.min()), bool(u.min() >= -tol and u.max() <= U_MAX + tol)))
    msgs.append(("放电量 0<=v<=U_MAX", float(v.min()), bool(v.min() >= -tol and v.max() <= U_MAX + tol)))
    msgs.append(("储电量 E_MIN<=E<=E_MAX", float(E.min()), bool(E.min() >= E_MIN - 1e-6 and E.max() <= E_MAX + 1e-6)))
    msgs.append(("E_0=E_144=6000", float(E[0]), bool(abs(E[0] - E_INIT) < 1e-9 and abs(E[-1] - E_INIT) < 1e-6)))
    dE = np.diff(E)
    rec = ETA * u - v / ETA
    msgs.append(("状态递推 max|dE-(eta*u-v/eta)| < 1e-6", float(np.abs(dE - rec).max()), bool(np.abs(dE - rec).max() < 1e-6)))
    msgs.append(("同时充放时段数 = 0", int(((u > 1e-6) & (v > 1e-6)).sum()), bool(((u > 1e-6) & (v > 1e-6)).sum() == 0)))
    msgs.append(("购电量为 0 的时段数", int((x < 1e-6).sum()), True))
    msgs.append(("全天购电量 Σx", float(x.sum()), True))
    msgs.append(("全天购电费 Σ p*x", float((price * x).sum()), True))
    msgs.append(("充电量合计 Σu", float(u.sum()), True))
    msgs.append(("放电量合计 Σv", float(v.sum()), True))
    msgs.append(("放电/充电 比值", float(v.sum() / u.sum()), True))
    return msgs


# ---------------- 4. 填法 Y-轮转 映射 ----------------
def template_rows_from_period(x):
    """输入：按真实时间顺序的 144 个时段值（0 基）；输出：模板 0 基 144 行的取值。
    规则（D-01）：模板行 i (i=0..142) 填 当天第 i+2 个时段；模板行 143 填 当天第 1 个时段。
    等价：模板行[i] = 逐时段解[(i+1) % 144]。"""
    return np.array([x[(i + 1) % K] for i in range(K)])


def period_from_template_rows(rows):
    """逆映射：模板行 -> 真实时段序。period[(i+1)%144] = rows[i]"""
    p = np.empty(K)
    for i in range(K):
        p[(i + 1) % K] = rows[i]
    return p


R_cyc = solve_lp("cyclic")
R_free = solve_lp("free")

w("### 构思手独立复核：问题 1（LP vs DP，两套独立算法）")
w()
w("--- LP 主模型（E_0=E_144=6000，eta=0.9 单向）---")
for name, val, ok in check_feasible(R_cyc):
    w("   %-46s 实测=%-16.6f 通过=%s" % (name, val, ok))
w()
w("--- LP 对照：端点自由 ---")
w("   费用 = %.4f 元 ; 购电量 = %.4f kWh ; E_144 = %.4f kWh"
  % (R_free["cost"], R_free["x"].sum(), R_free["E"][-1]))

w()
w("--- DP 复算（不同算法，用于交叉验证 LP）---")
w("   %-12s %-14s %-16s %-14s" % ("网格步长", "端点模式", "DP 费用(元)", "相对 LP 的差(元)"))
dp_result = {}
for step in (20.0, 10.0, 5.0):
    for mode, ref in (("cyclic", R_cyc["cost"]), ("free", R_free["cost"])):
        c_dp, _ = solve_dp(step=step, mode=mode)
        dp_result[(step, mode)] = c_dp
        w("   %-12s %-14s %-16.4f %-14.4f" % (step, mode, c_dp, c_dp - ref))

w()
w("--- 填法 Y-轮转：映射自检 ---")
rows = template_rows_from_period(R_cyc["x"])
back = period_from_template_rows(rows)
w("   模板行 -> 真实时段 往返映射最大误差 = %.3e" % np.abs(back - R_cyc["x"]).max())
w("   模板 144 行求和 = %.4f kWh （应等于全天购电量 %.4f）" % (rows.sum(), R_cyc["x"].sum()))
w("   模板首行(标签 0:10-0:20) = %.4f  <- 应等于 当天第2个时段 %.4f" % (rows[0], R_cyc["x"][1]))
w("   模板末行(标签 0:00+1-0:10+1) = %.4f  <- 应等于 当天第1个时段 %.4f" % (rows[-1], R_cyc["x"][0]))
w()
w("   表 1 六个时段（H:00-H:10 <-> 数据列 6H，0 基 <-> 当天第 6H+1 个时段 <-> 模板第 6H 行，1 基）:")
for hh in (10, 12, 14, 16, 18, 20):
    col = 6 * hh                                        # 数据列（0 基）
    trow = 6 * hh - 1                                    # 模板行（0 基）
    w("      %02d:00-%02d:10 : 真实时段值=%.4f ; 模板第 %d 行值=%.4f ; 该段电价=%.4f 元/kWh ; 负载=%.1f kW ; 光伏=%.1f kW"
      % (hh, hh, R_cyc["x"][col], 6 * hh, rows[trow], price[col], load[col], pv[col]))
w()
w("   表 2 六个 4 小时块（按真实时段顺序 1-144 聚合）:")
blk_names = ["0:00-4:00", "4:00-8:00", "8:00-12:00", "12:00-16:00", "16:00-20:00", "20:00-24:00"]
for i, nm in enumerate(blk_names):
    s = i * 24
    w("      %-12s 充电量=%.4f 放电量=%.4f" % (nm, R_cyc["u"][s:s + 24].sum(), R_cyc["v"][s:s + 24].sum()))
w("   0:00 储电量 = %.4f ; 24:00 储电量 = %.4f" % (R_cyc["E"][0], R_cyc["E"][-1]))

w()
w("--- 基线 ---")
base_A = float((price * load).sum() * DT)
base_B = float((price * np.maximum(load - pv, 0.0)).sum() * DT)
w("   方案A 全部向电网购电 = %.4f 元 ; 方案B 光伏自用不储能 = %.4f 元" % (base_A, base_B))
w("   主模型相对方案B 节省 = %.4f 元 (%.4f%%)" % (base_B - R_cyc["cost"], 100 * (base_B - R_cyc["cost"]) / base_B))

with open(P("_phase0", "报告13_构思手独立复核_问题1.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))

# 落盘构思手参照解（全分辨率，10 位小数），供后续验收对账 probe09 逐点比对
import csv
with open(P("_phase0", "参照解_问题1.csv"), "w", encoding="utf-8", newline="") as f:
    wr = csv.writer(f)
    wr.writerow(["时段序号", "计划购电量_kWh", "充电量_kWh", "放电量_kWh", "储电量_kWh"])
    for k in range(K):
        wr.writerow([k + 1, "%.10f" % R_cyc["x"][k], "%.10f" % R_cyc["u"][k],
                     "%.10f" % R_cyc["v"][k], "%.10f" % R_cyc["E"][k + 1]])
print("done")
