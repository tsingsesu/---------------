"""Phase 0 探针 6：问题 1 主模型 LP 的量级体检（只做量级与可行性验证，不产出正式交付物）。

模型：
  min  sum_k p_k * x_k
  s.t. u_k - x_k - v_k <= Delta*(P_k - L_k)              （供给不低于负载，允许弃光）
       E_MIN <= E_0 + sum_{j<=k}(eta*u_j - v_j/eta) <= E_MAX
       E_144 = E_0 = 6000 (主模型) 或自由（对照）
       0 <= x_k, 0 <= u_k,v_k <= P_MAX*Delta
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
price = np.array([float(r[1]) for r in r1[1:]])          # 元/kWh
load = np.array([float(r[2]) for r in r1[1:]])           # kW
pv = np.array([float(r[3]) for r in r1[1:]])             # kW

K = 144
DT = 1.0 / 6.0        # h
ETA = 0.9             # 单向效率
P_MAX = 5000.0        # kW
E_CAP = 12000.0
E_MIN, E_MAX = 1200.0, 10800.0
E_INIT = 6000.0
U_MAX = P_MAX * DT

def solve(mode, eta=ETA, p_max=P_MAX, e_min=E_MIN, e_max=E_MAX, e_init=E_INIT, price=price, load=load, pv=pv):
    umax = p_max * DT
    n = 3 * K
    c = np.concatenate([price, np.zeros(2 * K)])
    A = lil_matrix((K + 2 * K, n))     # K 条供给约束 + 2K 条储电量约束
    b = np.zeros(K + 2 * K)
    for k in range(K):
        A[k, k] = -1.0                 # -x_k
        A[k, K + k] = 1.0              # +u_k
        A[k, 2 * K + k] = -1.0         # -v_k
        b[k] = DT * (pv[k] - load[k])
    for k in range(K):
        for j in range(k + 1):
            A[K + k, K + j] = eta                 # +eta*u_j
            A[K + K + k, K + j] = -eta
            A[K + k, 2 * K + j] = -1.0 / eta      # -v_j/eta
            A[K + K + k, 2 * K + j] = 1.0 / eta
        b[K + k] = e_init - e_min
        b[K + K + k] = e_max - e_init
    Aeq = None; beq = None
    if mode == "cyclic":
        Aeq = lil_matrix((1, n))
        for j in range(K):
            Aeq[0, K + j] = eta
            Aeq[0, 2 * K + j] = -1.0 / eta
        beq = np.array([0.0])
    bounds = [(0, None)] * K + [(0, umax)] * K + [(0, umax)] * K
    res = linprog(c, A_ub=A.tocsr(), b_ub=b, A_eq=None if Aeq is None else Aeq.tocsr(), b_eq=beq,
                  bounds=bounds, method="highs")
    assert res.status == 0, res.message
    x = res.x[:K]; u = res.x[K:2 * K]; v = res.x[2 * K:]
    E = np.concatenate([[e_init], e_init + np.cumsum(eta * u - v / eta)])
    return dict(cost=float(res.fun), x=x, u=u, v=v, E=E, res=res)


res_cyc = solve("cyclic")
res_free = solve("free")

w("### 问题 1 量级体检（LP 实测）")
w("时段数 K=%d，单时段最大充/放电量 = %.4f kWh" % (K, U_MAX))
for name, R in [("主模型 E_0=E_144=6000", res_cyc), ("对照 端点自由", res_free)]:
    w()
    w("--- %s ---" % name)
    w("全天购电量 = %.4f kWh" % R["x"].sum())
    w("全天购电费 = %.4f 元" % R["cost"])
    w("全周期充电量合计 = %.4f kWh ; 放电量合计 = %.4f kWh" % (R["u"].sum(), R["v"].sum()))
    w("储电量 轨迹 min=%.2f max=%.2f ; 0:00=%.2f 24:00=%.2f" % (R["E"].min(), R["E"].max(), R["E"][0], R["E"][-1]))
    w("购电量为 0 的时段数 = %d ; 充放同时>1e-6 的时段数 = %d"
      % (int((R["x"] < 1e-6).sum()), int(((R["u"] > 1e-6) & (R["v"] > 1e-6)).sum())))
    w("表 1 指定时段购电量（填法 X：模板标签 'H:00-H:10' 对应 0 基索引 6H-1，即数据时间标签 H:00）:")
    for lab, hh in [("10:00-10:10", 10), ("12:00-12:10", 12), ("14:00-14:10", 14),
                    ("16:00-16:10", 16), ("18:00-18:10", 18), ("20:00-20:10", 20)]:
        k = hh * 6 - 1
        w("   %s : x = %.4f kWh （数据时间标签 %02d:00，电价 %.4f 元/kWh，负载 %.1f kW，光伏 %.1f kW）"
          % (lab, R["x"][k], hh, price[k], load[k], pv[k]))
    w("   [对照] 若取 0 基索引 6H（数据时间标签 (H+1):00）:")
    for lab, hh in [("10:00-10:10", 10), ("12:00-12:10", 12), ("14:00-14:10", 14),
                    ("16:00-16:10", 16), ("18:00-18:10", 18), ("20:00-20:10", 20)]:
        k = hh * 6
        w("      %s : x = %.4f kWh （数据时间标签 %.1f 分钟刻度）" % (lab, R["x"][k], 10 + k * 10))
    w("表 2 六个 4 小时块:")
    for i in range(6):
        s = i * 24
        w("   %d:00-%d:00 充电量=%.4f 放电量=%.4f" % (i * 4, i * 4 + 4, R["u"][s:s + 24].sum(), R["v"][s:s + 24].sum()))
    w("0:00 储电量=%.4f ; 24:00 储电量=%.4f" % (R["E"][0], R["E"][-1]))

base_A = (price * load).sum() * DT
base_B = (price * np.maximum(load - pv, 0.0)).sum() * DT
w()
w("基线：方案A 全部购电 %.2f 元 ; 方案B 光伏自用不储能 %.2f 元" % (base_A, base_B))
w("主模型相对方案B 节省 = %.2f 元 (%.2f%%)" % (base_B - res_cyc["cost"], 100 * (base_B - res_cyc["cost"]) / base_B))
w("端点自由相对主模型 再省 = %.2f 元 (%.2f%%)" % (res_cyc["cost"] - res_free["cost"],
                                              100 * (res_cyc["cost"] - res_free["cost"]) / res_cyc["cost"]))
w()
w("效率口径对照（端点自由）:")
for eta in [0.9, 0.9 ** 0.5, 0.8, 0.95]:
    R = solve("free", eta=eta)
    w("   eta=%.4f (往返 %.4f): 购电费=%.2f 元, 购电量=%.2f kWh, 充电量=%.2f kWh"
      % (eta, eta ** 2, R["cost"], R["x"].sum(), R["u"].sum()))

with open(P("_phase0", "报告11_问题1量级体检.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done")
