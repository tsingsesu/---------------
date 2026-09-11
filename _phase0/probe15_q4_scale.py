"""构思手独立量级体检：问题 4-2（问题 2 换附件4 波动电价 + D-10 终端余值）。

口径：
  * 价格：附件4 (365,144)，逐日不同；
  * D-04：终端自由 + 跨日传递；D-11 从 2025-01-01 仿真、填报 2.1–12.31；
  * D-10 主口径：逐日 LP 目标加入终端余值 V_E·E_144，V_E = 次日最低价 / η；
    对照：终端自由（V_E=0）、终端=当日初始（cyclic）。
  * D-06 读法① 完全信息（问题 4-2 与问题 2 相同的信息结构）。
输出：三种终端口径的全年总费用 + 与问题 2（固定电价）的对照。
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


def rd(path, sheet):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


rL = rd(P("附件", "附件2.xlsx"), "小区负载")
rP = rd(P("附件", "附件2.xlsx"), "光伏发电实际功率")
L2 = np.array([[float(v) for v in r[1:145]] for r in rL[1:]])
P2 = np.array([[float(v) for v in r[1:145]] for r in rP[1:]])
r4 = rd(P("附件", "附件4.xlsx"), "Sheet1")
PR4 = np.array([[float(v) for v in r[1:145]] for r in r4[1:]])
D = 365
K = 144
DT = 1.0 / 6.0
ETA = 0.9
UMAX = 5000 * DT
EMIN, EMAX, EINIT = 1200.0, 10800.0, 6000.0


def solve_day(price, load, pv, e0, v_end=0.0):
    """单日 LP：目标 min Σp x − v_end·E_K（E_K 为终端储电量，v_end 为余值单价）。"""
    n = 3 * K + 1                                   # [x | u | v | E_K]
    c = np.concatenate([price, np.zeros(2 * K), [-v_end]])
    A = lil_matrix((3 * K, n))
    b = np.zeros(3 * K)
    for k in range(K):
        A[k, k] = -1; A[k, K + k] = 1; A[k, 2 * K + k] = -1
        b[k] = -(load[k] - pv[k]) * DT
    for k in range(K):
        for j in range(k + 1):
            A[K + k, K + j] = ETA
            A[K + k, 2 * K + j] = -1 / ETA
            A[2 * K + k, K + j] = -ETA
            A[2 * K + k, 2 * K + j] = 1 / ETA
        b[K + k] = EMAX - e0
        b[2 * K + k] = e0 - EMIN
    # E_K − Σ(ηu − v/η) = e0  ⇒  Σ(ηu − v/η) − E_K = −e0
    Aeq = lil_matrix((1, n))
    for j in range(K):
        Aeq[0, K + j] = ETA
        Aeq[0, 2 * K + j] = -1 / ETA
    Aeq[0, 3 * K] = -1.0
    beq = np.array([-e0])
    res = linprog(c, A_ub=A.tocsr(), b_ub=b, A_eq=Aeq.tocsr(), b_eq=beq,
                  bounds=[(0, None)] * K + [(0, UMAX)] * K + [(0, UMAX)] * K + [(EMIN, EMAX)],
                  method="highs")
    assert res.status == 0, res.message
    x = res.x[:K]; u = res.x[K:2 * K]; v = res.x[2 * K:3 * K]
    E = e0 + np.cumsum(ETA * u - v / ETA)
    return dict(cost=float((price * x).sum()), x=x, u=u, v=v, E=E)


def run_year(scheme):
    """scheme: 'residual'（D-10 余值）/ 'free' / 'cyclic'。

    返回逐日 (d, opcost=Σp·x（实际缴费）, E_K, Σx, x 向量, u, v)。
    说明：余值 V_E 只改变**决策**（LP 目标），不改变实际缴费；经济上正确的
    "全期真实成本" = Σ_d opcost_d − V_E·E_K^{末日}（末日电池的残值只计一次）。
    """
    e = EINIT
    rec = []
    for d in range(D):
        price = PR4[d]
        if scheme == "residual":
            nxt = PR4[d + 1] if d + 1 < D else PR4[d]
            ve = nxt.min() / ETA
        else:
            ve = 0.0
        r = solve_day(price, L2[d], P2[d], e, v_end=ve)
        rec.append((d, r["cost"], float(r["E"][-1]), float(r["x"].sum()),
                    r["x"].copy(), r["u"].copy(), r["v"].copy(), e))
        e = float(r["E"][-1])
    return rec


w("### 问题 4-2 构思手独立量级体检（附件4 波动电价；D-04/D-10/D-11）")
w()
import datetime as _dt
for scheme, name in (("residual", "D-10 终端余值 V_E=次日最低价/η（主口径）"),
                     ("free", "终端自由（对照）")):
    rec = run_year(scheme)
    rep = rec[31:]
    tot_pay = sum(r[1] for r in rep)
    e_end = rep[-1][2]
    ve_last = PR4[364].min() / ETA
    w("--- %s ---" % name)
    w("   填报区间实际缴费 Σp·x = %.4f 元 ; 日均 = %.4f 元" % (tot_pay, tot_pay / 334))
    w("   期末（12-31 24:00）储电量 = %.4f kWh（残值 %.4f 元/kWh → %.4f 元，只在期末计一次）"
      % (e_end, ve_last, e_end * ve_last))
    w("   全期真实成本（缴费 − 期末残值）= %.4f 元" % (tot_pay - e_end * ve_last))
    w("   24:00 落于下限 1200 的天数 = %d / 334" % int(sum(1 for r in rep if abs(r[2] - EMIN) < 1e-3)))
    if scheme == "residual":
        # 四指定日期锚定表
        for target in ("2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21"):
            d0 = (_dt.date.fromisoformat(target) - _dt.date(2025, 1, 1)).days
            (dd, cst, eK, qty, x, u, v, e0) = rec[d0]
            w("   --- %s ---" % target)
            vals = []
            for H in (10, 12, 14, 16, 18, 20):
                vals.append("%02d:00=%.4f" % (H, x[6 * H]))
            w("     表1: " + " ".join(vals))
            w("     全天: Σx=%.4f kWh, 费用=%.4f 元" % (qty, cst))
            blk = [(0, 24), (24, 48), (48, 72), (72, 96), (96, 120), (120, 144)]
            w("     表2充: " + " ".join("%.4f" % u[a:b2].sum() for a, b2 in blk))
            w("     表2放: " + " ".join("%.4f" % v[a:b2].sum() for a, b2 in blk))
            w("     0:00=%.4f  24:00=%.4f" % (e0, eK))
    w()
w("--- 对照：问题 2（固定电价附件1）主模型 = 12 254 765.7161 元 ---")
w("    注：跨年处理（12-31 无次日，取当日最低价/η）为构思手假设，需与编程手对齐。")

with open(P("_phase0", "报告20_问题4-2量级体检.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done")
