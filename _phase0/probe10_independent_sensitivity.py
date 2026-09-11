"""构思手独立参照：问题 1 的参数/数据/方法侧扰动（用于对账编程手交付的 S1–S5）。

与 probe08 同样的独立 LP/DP 实现（构思手自写，不引用 lib/）。
判据（"结论是否翻转"）定义为一个可判定的量：
  问题 1 的核心结论 = "配置储能并按峰谷价差优化购电计划，可使全天购电费显著低于不储能基线"。
  翻转判据：优化后费用 >= 方案 B（光伏自用不储能）的费用，即储能不再带来净收益。
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


def ls(path, sheet=None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


r1 = ls(P("附件", "附件1.xlsx"))
p0 = np.array([float(r[1]) for r in r1[1:]])
L0 = np.array([float(r[2]) for r in r1[1:]])
V0 = np.array([float(r[3]) for r in r1[1:]])

K = 144
DT = 1.0 / 6.0
ETA0, PMAX0 = 0.9, 5000.0
EMIN0, EMAX0, EINIT = 1200.0, 10800.0, 6000.0


def solve_lp(price, load, pv, eta=ETA0, p_max=PMAX0, e_min=EMIN0, e_max=EMAX0,
             e_init=EINIT, mode="cyclic", dt=DT, K=K):
    umax = p_max * dt
    n = 3 * K
    c = np.concatenate([price, np.zeros(2 * K)])
    A = lil_matrix((3 * K, n))
    b = np.zeros(3 * K)
    for k in range(K):
        A[k, k] = -1.0
        A[k, K + k] = 1.0
        A[k, 2 * K + k] = -1.0
        b[k] = -(load[k] - pv[k]) * dt
    for k in range(K):
        for j in range(k + 1):
            A[K + k, K + j] = eta                 # 上界行系数
            A[K + k, 2 * K + j] = -1.0 / eta
            A[K + K + k, K + j] = -eta            # 下界行系数
            A[K + K + k, 2 * K + j] = 1.0 / eta
        b[K + k] = e_max - e_init                 # 与上界行成对
        b[K + K + k] = e_init - e_min             # 与下界行成对
    Aeq = beq = None
    if mode == "cyclic":
        Aeq = lil_matrix((1, n))
        for j in range(K):
            Aeq[0, K + j] = eta
            Aeq[0, 2 * K + j] = -1.0 / eta
        beq = np.array([0.0])
    res = linprog(c, A_ub=A.tocsr(), b_ub=b,
                  A_eq=None if Aeq is None else Aeq.tocsr(), b_eq=beq,
                  bounds=[(0, None)] * K + [(0, umax)] * K + [(0, umax)] * K, method="highs")
    if res.status != 0:
        return None
    x = res.x[:K]; u = res.x[K:2 * K]; v = res.x[2 * K:]
    E = np.concatenate([[e_init], e_init + np.cumsum(eta * u - v / eta)])
    return dict(cost=float(res.fun), x=x, u=u, v=v, E=E)


def solve_dp(step, price, load, pv, eta=ETA0, p_max=PMAX0, e_min=EMIN0, e_max=EMAX0, e_init=EINIT):
    umax = p_max * DT
    grid = np.arange(e_min, e_max + 1e-9, step)
    N = len(grid)
    D = grid[None, :] - grid[:, None]
    ok = (D >= -umax / eta - 1e-9) & (D <= eta * umax + 1e-9)
    eff = np.where(D >= 0, D / eta, eta * D)
    NET = (load - pv) * DT
    INF = 1e18
    i0 = int(np.argmin(np.abs(grid - e_init)))
    V = np.full(N, INF); V[i0] = 0.0
    for k in range(K):
        cc = price[k] * np.maximum(0.0, NET[k] + eff)
        cand = V[:, None] + cc
        cand[~ok] = INF
        V = cand.min(axis=0)
    j = int(np.argmin(np.abs(grid - e_init)))
    return float(V[j])


BASE = solve_lp(p0, L0, V0)
BASE_COST = BASE["cost"]
BASE_B = float((p0 * np.maximum(L0 - V0, 0.0)).sum() * DT)

w("### 构思手独立灵敏度参照（供与编程手 S1–S5 对账）")
w()
w("基准（主模型 E_0=E_144=6000、eta=0.9 单向、P_max=5000）：")
w("   全天购电量 = %.4f kWh ; 全天购电费 = %.4f 元 ; 充电量合计 = %.4f kWh"
  % (BASE["x"].sum(), BASE_COST, BASE["u"].sum()))
w("   方案 B（光伏自用不储能）= %.4f 元 → 储能净收益 = %.4f 元" % (BASE_B, BASE_B - BASE_COST))
w()
w("判据：若某档扰动后 优化费用 >= 方案B 费用（%.4f 元），则记为【结论翻转】。" % BASE_B)


def judge(cost):
    return "翻转" if cost >= BASE_B else "未翻转"


def report_block(title, rows, ref=BASE_COST, refq=None):
    w()
    w("--- %s ---" % title)
    w("   %-22s %-16s %-14s %-12s %-12s %s" % ("扰动档", "费用(元)", "费用相对变化", "购电量(kWh)", "充电量(kWh)", "结论"))
    for name, r in rows:
        if r is None:
            w("   %-22s %-16s" % (name, "不可行"))
            continue
        w("   %-22s %-16.4f %-14s %-12.4f %-12.4f %s"
          % (name, r["cost"], "%+.4f%%" % (100 * (r["cost"] - ref) / ref),
             r["x"].sum(), r["u"].sum(), judge(r["cost"])))


# S1a 单向效率
rows = []
for eta in (0.72, 0.81, 0.855, 0.9, 0.945, 0.99, 1.08):
    rows.append(("eta=%.4f(往返%.4f)" % (eta, eta ** 2), solve_lp(p0, L0, V0, eta=eta)))
report_block("S1a 单向效率 η 扰动（D-05 主口径）", rows)

# S1b 最大充放电功率
rows = []
for f in (0.8, 0.9, 0.95, 1.0, 1.05, 1.1, 1.2):
    rows.append(("P_max=%.1f kW(%.0f%%)" % (PMAX0 * f, 100 * f), solve_lp(p0, L0, V0, p_max=PMAX0 * f)))
report_block("S1b 最大充放电功率 P_max 扰动", rows)

# S1c 储电量可用区间（围绕中点 6000 对称缩放半宽 4800）
rows = []
for f in (0.8, 0.9, 0.95, 1.0, 1.05, 1.1, 1.2):
    half = 4800.0 * f
    rows.append(("可用区间±%.0f kWh(%.0f%%)" % (half, 100 * f),
                 solve_lp(p0, L0, V0, e_min=6000 - half, e_max=6000 + half)))
report_block("S1c 储电量允许区间（对称缩放）", rows)

# S3a 电价整体缩放
rows = []
for f in (0.8, 0.9, 0.95, 1.0, 1.05, 1.1, 1.2):
    r = solve_lp(p0 * f, L0, V0)
    rows.append(("电价 x%.2f" % f, r))
report_block("S3a 电价整体缩放", rows)

# S3b 负载缩放
rows = []
for f in (0.95, 0.98, 1.0, 1.02, 1.05):
    rows.append(("负载 x%.2f" % f, solve_lp(p0, L0 * f, V0)))
report_block("S3b 负载整体缩放", rows)

# S3c 光伏缩放
rows = []
for f in (0.9, 0.95, 1.0, 1.05, 1.1):
    rows.append(("光伏 x%.2f" % f, solve_lp(p0, L0, V0 * f)))
report_block("S3c 光伏整体缩放", rows)

# S3d 小时粒度（数据粗化到 24 个 1 小时段）
def hourly(price, load, pv):
    ph = price.reshape(24, 6).mean(1)      # 小时内平均电价，元/kWh
    Lh = load.reshape(24, 6).mean(1)       # 小时内平均负载，kW
    Vh = pv.reshape(24, 6).mean(1)
    return ph, Lh, Vh

ph, Lh, Vh = hourly(p0, L0, V0)
Rh = solve_lp(ph, Lh, Vh, dt=1.0, K=24)
w()
w("--- S3d 小时粒度重算（24 段，每段 Δ=1 h）---")
w("   费用 = %.4f 元 ; 购电量 = %.4f kWh ; 充电量合计 = %.4f kWh ; 相对 10 分钟粒度 %.4f%%"
  % (Rh["cost"], Rh["x"].sum(), Rh["u"].sum(), 100 * (Rh["cost"] - BASE_COST) / BASE_COST))
basB_h = float((ph * np.maximum(Lh - Vh, 0.0)).sum())
w("   小时粒度下的方案 B = %.4f 元 ; 储能净收益 = %.4f 元" % (basB_h, basB_h - Rh["cost"]))

# S4a DP 网格
w()
w("--- S4a 离散化 DP（网格 100/200/400 步）---")
for nstep in (100, 200, 400):
    step = (EMAX0 - EMIN0) / nstep
    if abs((EINIT - EMIN0) / step - round((EINIT - EMIN0) / step)) > 1e-9:
        w("   网格 %d 步（步长 %.4f kWh）：初始储电量不在网格上，跳过" % (nstep, step))
        continue
    c = solve_dp(step, p0, L0, V0)
    w("   网格 %3d 步（步长 %8.4f kWh）：DP 费用 = %.4f 元 ; 相对 LP 高 %.4f 元（%.4f%%）"
      % (nstep, step, c, c - BASE_COST, 100 * (c - BASE_COST) / BASE_COST))

# S5 端点与效率口径对照
w()
w("--- S5 口径对照 ---")
rf = solve_lp(p0, L0, V0, mode="free")
w("   E_0=E_144=6000（主模型）= %.4f 元 ; 端点自由 = %.4f 元（再省 %.4f 元，%.4f%%），其 E_144=%.4f kWh"
  % (BASE_COST, rf["cost"], BASE_COST - rf["cost"], 100 * (BASE_COST - rf["cost"]) / BASE_COST, rf["E"][-1]))
eta_rt = 0.9 ** 0.5
rr = solve_lp(p0, L0, V0, eta=eta_rt)
w("   单向 0.9（往返 0.81）= %.4f 元 ; 往返 0.9（单向 %.4f）= %.4f 元（差 %.4f 元，%.4f%%）"
  % (BASE_COST, eta_rt, rr["cost"], BASE_COST - rr["cost"], 100 * (BASE_COST - rr["cost"]) / BASE_COST))

# 填法 X 对照（表 1 六时段取值）
def tmpl_Y(x):
    return np.array([x[(i + 1) % K] for i in range(K)])

w()
w("--- S5c 填法对照：表 1 六个时段 ---")
rowsY = tmpl_Y(BASE["x"])
w("   %-14s %-14s %-14s %s" % ("时段", "Y-轮转(采用)", "X(逐位置)", "差异"))
for hh in (10, 12, 14, 16, 18, 20):
    i = 6 * hh - 1                       # 模板行（0 基）
    vy = rowsY[i]                        # Y-轮转：模板行 i = 数据列 i+1 = 6H
    vx = BASE["x"][i]                    # X：模板行 i = 数据列 i
    w("   %02d:00-%02d:10  %-14.4f %-14.4f %+.4f" % (hh, hh, vy, vx, vx - vy))

with open(P("_phase0", "报告15_构思手独立灵敏度参照.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done")
