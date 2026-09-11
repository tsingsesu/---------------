"""构思手独立测算：问题 2 全年联合 LP 下界（对照逐日滚动策略的次优性）。

模型：变量 [x(每时段) | u | v | E(每时段末)]，共 4×52560 变量。
  supply_t:  x + v − u ≥ (L−P)Δ
  recur_t:   E_t − E_{t−1} − ηu + v/η = 0   （E_{−1} = E_INIT）
  bounds:    x ≥ 0 ; 0 ≤ u,v ≤ P̄Δ ; Ē ≥ E ≥ E̲
  min Σ p_t x_t
区间同报告17：仿真 2025-01-01 起（E=6000），统计填报区间 2.1–12.31。
"""

import os
import time
import numpy as np
import openpyxl
from scipy.optimize import linprog
from scipy.sparse import coo_matrix

WORK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
P = lambda *a: os.path.join(WORK, *a)
out = []
def w(s=""):
    out.append(str(s))


def read_sheet(path, sheet):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


r1 = read_sheet(P("附件", "附件1.xlsx"), None)
p0 = np.array([float(r[1]) for r in r1[1:145]])
rL = read_sheet(P("附件", "附件2.xlsx"), "小区负载")
rP = read_sheet(P("附件", "附件2.xlsx"), "光伏发电实际功率")
L2 = np.array([[float(v) for v in r[1:145]] for r in rL[1:]])
P2 = np.array([[float(v) for v in r[1:145]] for r in rP[1:]])
D, K = L2.shape
T = D * K

DT = 1.0 / 6.0
ETA = 0.9
UMAX = 5000.0 * DT
EMIN, EMAX, EINIT = 1200.0, 10800.0, 6000.0

price = np.tile(p0, D)                 # (T,) 电价（逐日相同）
Lt = L2.reshape(-1)                    # (T,) 负载 kW
Pt = P2.reshape(-1)                    # (T,) 光伏 kW

# 变量布局：[x(T) | u(T) | v(T) | E(T)]
IX, IU, IV, IE = 0, T, 2 * T, 3 * T
n = 4 * T
c = np.zeros(n); c[IX:IX + T] = price

t0 = time.time()
rows, cols, vals = [], [], []

# 供给约束行 r = t（3T 项）
r = np.arange(T)
rows += list(r); cols += list(IX + r); vals += [-1.0] * T
rows += list(r); cols += list(IU + r); vals += [1.0] * T
rows += list(r); cols += list(IV + r); vals += [-1.0] * T
b_ub = -(Lt - Pt) * DT
A_ub = coo_matrix((vals, (rows, cols)), shape=(T, n)).tocsr()

# 递推等式行 r = T + t：E_t − E_{t−1} − ηu_t + v_t/η = 0（E_{−1} = EINIT 在右端）
rr = np.arange(T)
rows3 = list(rr) + list(rr[1:]) + list(rr) + list(rr)
cols3 = (list(IE + rr) + list(IE + rr[:-1]) + list(IU + rr) + list(IV + rr))
vals3 = [1.0] * T + [-1.0] * (T - 1) + [-ETA] * T + [1.0 / ETA] * T
b_eq = np.zeros(T); b_eq[0] = EINIT
A_eq = coo_matrix((vals3, (rows3, cols3)), shape=(T, n)).tocsr()

bounds = [(0, None)] * T + [(0, UMAX)] * T + [(0, UMAX)] * T + [(EMIN, EMAX)] * T

w("### 问题 2 全年联合 LP（完全预见下界）——构思手独立测算")
w("变量数 = %d（4×%d），约束 = %d 行 ×2" % (n, T, T))
w("构建耗时 %.1fs，开始求解……" % (time.time() - t0))
res = linprog(c, A_ub=A_ub, b_ub=b_ub, A_eq=A_eq, b_eq=b_eq, bounds=bounds, method="highs")
w("求解状态：status=%d（0=最优），用时 %.1fs" % (res.status, time.time() - t0))
assert res.status == 0, res.message

x = res.x[IX:IX + T]
cost_total = float(res.fun)
# 填报区间 2.1–12.31 = 第 31 天起（0 基 31*144 起）
seg = slice(31 * K, T)
cost_rep = float((price[seg] * x[seg]).sum())
w()
w("全年（1.1–12.31, 365 天）联合最优总费用 = %.4f 元 ; 日均 = %.4f 元" % (cost_total, cost_total / D))
w("填报区间（2.1–12.31, 334 天）联合最优总费用 = %.4f 元 ; 日均 = %.4f 元"
  % (cost_rep, cost_rep / 334))
w()
w("对照：逐日滚动（报告17）填报区间总费用 = 12254765.7161 元；日均 36690.9153 元")
w("     滚动 − 联合下界 = %.4f 元（相对 %.4f%%）"
  % (12254765.7161 - cost_rep, 100 * (12254765.7161 - cost_rep) / cost_rep))

with open(P("_phase0", "报告18_问题2联合LP下界.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done")
