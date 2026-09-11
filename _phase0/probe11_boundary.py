"""构思手独立测算：问题 1 的适用边界（用于 适用边界.md）。

三类边界：
  B1 效率边界：单向效率 η 降到多少时，储能不再带来净收益（优化费用 = 不储能基线）
  B2 功率边界：最大充放电功率 P̄ 降到多少时，功率成为真正的瓶颈
  B3 峰谷价差边界：把电价围绕日均值压缩（p(λ) = 均值 + λ(p−均值)），λ 降到多少时储能失去价值
另附：容量区间边界的定量结果（已在 probe10 覆盖，此处只汇总）。
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
EMIN0, EMAX0, EINIT = 1200.0, 10800.0, 6000.0


def solve(price, eta=0.9, p_max=5000.0, e_min=EMIN0, e_max=EMAX0, load=None, pv=None):
    load = L0 if load is None else load
    pv = V0 if pv is None else pv
    umax = p_max * DT
    n = 3 * K
    c = np.concatenate([price, np.zeros(2 * K)])
    A = lil_matrix((3 * K, n))
    b = np.zeros(3 * K)
    for k in range(K):
        A[k, k] = -1.0; A[k, K + k] = 1.0; A[k, 2 * K + k] = -1.0
        b[k] = -(load[k] - pv[k]) * DT
    for k in range(K):
        for j in range(k + 1):
            A[K + k, K + j] = eta
            A[K + k, 2 * K + j] = -1.0 / eta
            A[K + K + k, K + j] = -eta
            A[K + K + k, 2 * K + j] = 1.0 / eta
        b[K + k] = e_max - EINIT
        b[K + K + k] = EINIT - e_min
    Aeq = lil_matrix((1, n))
    for j in range(K):
        Aeq[0, K + j] = eta
        Aeq[0, 2 * K + j] = -1.0 / eta
    res = linprog(c, A_ub=A.tocsr(), b_ub=b, A_eq=Aeq.tocsr(), b_eq=np.array([0.0]),
                  bounds=[(0, None)] * K + [(0, umax)] * K + [(0, umax)] * K, method="highs")
    if res.status != 0:
        return None
    x = res.x[:K]; u = res.x[K:2 * K]; v = res.x[2 * K:]
    return dict(cost=float(res.fun), x=x, u=u, v=v, qty=float(x.sum()), u_sum=float(u.sum()))


def baselineB(price, load=None, pv=None):
    """不储能基线：光伏优先自用，余电弃掉，缺额全额外购。"""
    load = L0 if load is None else load
    pv = V0 if pv is None else pv
    return float((price * np.maximum(load - pv, 0.0)).sum() * DT)


w("### 问题 1 适用边界测算（构思手独立执行）")
BASE = solve(p0)
BB = baselineB(p0)
w("基准：优化费用 = %.4f 元 ; 不储能基线 = %.4f 元 ; 储能净收益 = %.4f 元（%.4f%%）"
  % (BASE["cost"], BB, BB - BASE["cost"], 100 * (BB - BASE["cost"]) / BB))

# ---------------- B1 效率边界（二分求净收益 = 0 的 η*） ----------------
w()
w("--- B1 效率边界：净收益随 η 变化 ---")
w("   %-10s %-16s %-16s %-16s" % ("η 单向", "往返 η²", "优化费用(元)", "储能净收益(元)"))
tab = []
for eta in (0.50, 0.55, 0.60, 0.65, 0.70, 0.75, 0.80, 0.85, 0.90):
    R = solve(p0, eta=eta)
    tab.append((eta, R["cost"], BB - R["cost"]))
    w("   %-10.4f %-16.4f %-16.4f %-16.4f" % (eta, eta ** 2, R["cost"], BB - R["cost"]))
# 注意：净收益在 η=0.5 仍为正（2142.84 元），说明收益中有一部分来自光伏消纳而非套利，
# 搜索区间必须下探到很小的 η 才能找到真正的临界点（第一版取 [0.2, 0.9] 会误收敛到区间下界）。
lo, hi = 0.02, 0.90
for _ in range(60):
    mid = 0.5 * (lo + hi)
    R = solve(p0, eta=mid)
    if R is not None and BB - R["cost"] > 0:
        hi = mid
    else:
        lo = mid
w("   净收益 = 0 的临界单向效率 η* ≈ %.4f（往返 η*² ≈ %.4f）" % (hi, hi ** 2))
# 收益分解（三项，含交互项；基准统一为 B0 = 不储能基线）
#   总净收益   G = B0 - J(p)
#   光伏消纳项 G1 = B0_flat - J_flat          电价压平（无任何套利）时储能仍省下的钱
#   套利项     G2 = J_flat - J(p)             价格从"压平"恢复到真实形状，储能额外获得的收益
#   交互项     G3 = (B0 - J(p)) - G1 - G2     两项本就不独立（压平价格同时消掉了分时结构）
B0 = baselineB(p0)
J_flat = solve(np.full(K, p0.mean()))["cost"]
B0_flat = baselineB(np.full(K, p0.mean()))
G1 = B0_flat - J_flat
G2 = J_flat - BASE["cost"]
G3 = (B0 - BASE["cost"]) - G1 - G2
w("   收益分解（三项）：总净收益 %.4f = 光伏消纳 %.4f + 套利 %.4f + 交互 %.4f 元"
  % (B0 - BASE["cost"], G1, G2, G3))
w("   （分解口径：G1 为电价压平后储能仍省下的钱，G2 为恢复真实价格形状后额外省下的钱，")
w("     二者不独立——压平价格同时消掉分时结构，交互项即这一不可加性；避免把交互项隐式塞进任一项。）")
w("   理论套利门槛（单对充放、按全天最高/最低价）：η² > p_min/p_max = %.4f/%.4f = %.4f ⇒ η > %.4f"
  % (p0.min(), p0.max(), p0.min() / p0.max(), np.sqrt(p0.min() / p0.max())))
# D-05 对照口径：往返效率 0.9 ⇒ 单向 η = √0.9；用于核对"往返 0.9"口径下的费用数
ETA_RT = 0.9 ** 0.5
R_RT = solve(p0, eta=ETA_RT)
w("   对照（D-05 往返 0.9 = 单向 η=√0.9=%.4f）：费用 = %.4f 元；相对主口径 %.4f 元（%+.4f%%）"
  % (ETA_RT, R_RT["cost"], R_RT["cost"] - BASE["cost"],
     100 * (R_RT["cost"] - BASE["cost"]) / BASE["cost"]))

# ---------------- B2 功率边界 ----------------
w()
w("--- B2 功率边界：净收益与相对损失随 P̄ 变化 ---")
w("   %-14s %-16s %-16s %-16s" % ("P̄ (kW)", "相对基准损失(元)", "损失占比", "储能净收益(元)"))
for pm in (250, 500, 1000, 2000, 3000, 4000, 5000):
    R = solve(p0, p_max=pm)
    w("   %-14.0f %-16.4f %-16s %-16.4f"
      % (pm, R["cost"] - BASE["cost"], "%+.4f%%" % (100 * (R["cost"] - BASE["cost"]) / BASE["cost"]),
         BB - R["cost"]))

# ---------------- B3 峰谷价差边界 ----------------
w()
w("--- B3 峰谷价差边界：p(λ) = 均值 + λ(p−均值) ---")
pm_mean = p0.mean()
w("   %-10s %-14s %-16s %-16s %-16s" % ("λ", "峰谷比", "优化费用(元)", "不储能基线(元)", "储能净收益(元)"))
b3 = []
for lam in (0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.8, 1.0):
    pl = pm_mean + lam * (p0 - pm_mean)
    R = solve(pl)
    bB = baselineB(pl)
    b3.append((lam, R["cost"], bB, bB - R["cost"]))
    w("   %-10.2f %-14.4f %-16.4f %-16.4f %-16.4f"
      % (lam, pl.max() / pl.min() if pl.min() > 0 else np.inf, R["cost"], bB, bB - R["cost"]))
# λ∈[0,1] 上总净收益恒为正（λ=0 时仍有 3877.61 元，全部来自光伏消纳），不存在临界点；
# 峰谷倒挂（λ<0）时价格形状反转，仍不归零。故本节的正确结论是"净收益有下界"，而非"存在 λ*"。
w("   注：λ∈[0,1] 上净收益恒为正（下界在 λ=0 处，≈3877.61 元）；λ<0（峰谷倒挂）的数据见下：")
for lam in (-1.0, -0.5, -0.2):
    pl = pm_mean + lam * (p0 - pm_mean)
    R = solve(pl)
    bB = baselineB(pl)
    w("   λ = %+.2f（谷峰比 %.4f）：优化费用 = %.4f 元，净收益 = %.4f 元"
      % (lam, (pl.min() / pl.max() if pl.min() > 0 else float("nan")), R["cost"], bB - R["cost"]))
w("   结论：在本数据的负载/光伏结构下，储能净收益的下界由光伏消纳（避免弃光）给出，")
w("         与峰谷价差是否充分无关——即使完全没有价差或价差倒挂，储能仍为正收益。")

# ---------------- 容量区间边界 ----------------
w()
w("--- 附：可用容量区间边界（围绕中点 6000 kWh 对称缩放半宽 4800 kWh） ---")
w("   %-16s %-16s %-16s" % ("半宽(kWh)", "优化费用(元)", "储能净收益(元)"))
for half in (1200, 2400, 3600, 4800):
    R = solve(p0, e_min=6000 - half, e_max=6000 + half)
    if R is None:
        w("   %-16.0f %-16s %-16s" % (half, "不可行", "—"))
    else:
        w("   %-16.0f %-16.4f %-16.4f" % (half, R["cost"], BB - R["cost"]))

with open(P("_phase0", "报告16_适用边界测算.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done")
