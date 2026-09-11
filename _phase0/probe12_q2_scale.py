"""构思手独立量级体检：问题 2 全年逐日滚动 LP（不使用编程手代码）。

口径（全部按 口径与假设台账.md 已裁决项）：
  D-01 填法 Y-轮转；D-04 终端自由 + 跨日传递，E0^1 = 6000；D-11 从 2025-01-01 仿真、填报 2.1–12.31；
  D-06 读法① 完全信息 → r ≡ 0（本脚本同时证明/验证）；D-05 单向 η=0.9。
输出：逐日费用统计 + 四指定日期表 1/表 2 锚定值 + 跨日连续性 + 与问题 1 典型日对照。
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


def read_sheet(path, sheet):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


# 附件1 电价（144 时段）
r1 = read_sheet(P("附件", "附件1.xlsx"), None)
p0 = np.array([float(r[1]) for r in r1[1:145]])
# 附件2 负载与光伏（365 天 × 144 时段）
rL = read_sheet(P("附件", "附件2.xlsx"), "小区负载")
rP = read_sheet(P("附件", "附件2.xlsx"), "光伏发电实际功率")
L2 = np.array([[float(v) for v in r[1:145]] for r in rL[1:]])
P2 = np.array([[float(v) for v in r[1:145]] for r in rP[1:]])
D = L2.shape[0]
assert L2.shape == (365, 144) and P2.shape == (365, 144), (L2.shape, P2.shape)

K = 144
DT = 1.0 / 6.0
ETA = 0.9
P_MAX = 5000.0
EMIN, EMAX, EINIT = 1200.0, 10800.0, 6000.0
UMAX = P_MAX * DT
ND = 4


def solve_day(price, load, pv, e0, e_end=None):
    """单日 LP：变量 [x(144) | u(144) | v(144)]，目标 min Σp x。e_end=None 为终端自由。"""
    n = 3 * K
    c = np.concatenate([price, np.zeros(2 * K)])
    A = lil_matrix((3 * K, n))
    b = np.zeros(3 * K)
    for k in range(K):
        A[k, k] = -1.0
        A[k, K + k] = 1.0
        A[k, 2 * K + k] = -1.0
        b[k] = -(load[k] - pv[k]) * DT
    for k in range(K):
        for j in range(k + 1):
            A[K + k, K + j] = ETA
            A[K + k, 2 * K + j] = -1.0 / ETA
            A[2 * K + k, K + j] = -ETA
            A[2 * K + k, 2 * K + j] = 1.0 / ETA
        b[K + k] = EMAX - e0
        b[2 * K + k] = e0 - EMIN
    if e_end is None:
        Aeq = None
        beq = None
    else:
        Aeq = lil_matrix((1, n))
        for j in range(K):
            Aeq[0, K + j] = ETA
            Aeq[0, 2 * K + j] = -1.0 / ETA
        beq = np.array([e_end - e0])
    res = linprog(c, A_ub=A.tocsr(), b_ub=b, A_eq=Aeq.tocsr() if Aeq is not None else None,
                  b_eq=beq, bounds=[(0, None)] * K + [(0, UMAX)] * K + [(0, UMAX)] * K,
                  method="highs")
    assert res.status == 0, res.message
    x = res.x[:K]; u = res.x[K:2 * K]; v = res.x[2 * K:]
    E = e0 + np.cumsum(ETA * u - v / ETA)
    return dict(cost=float(res.fun), x=x, u=u, v=v, E=E)


# ---- 全年逐日滚动（终端自由）----
E = EINIT
daily = []           # (费用, 购电量, 充电, 放电, E0, E144)
recs = {}
for d in range(D):
    r = solve_day(p0, L2[d], P2[d], E)
    daily.append((r["cost"], float(r["x"].sum()), float(r["u"].sum()), float(r["v"].sum()), E, float(r["E"][-1])))
    recs[d] = r
    E = float(r["E"][-1])
daily = np.array(daily)

# 填报区间 2.1–12.31 = d=31..364（0 基）
rep = daily[31:]
w("### 问题 2 构思手独立量级体检（逐日滚动 LP，终端自由，读法① 完全信息）")
w("仿真 365 天全跑；填报区间 2.1–12.31 共 334 天（d=32..365, 0 基 31..364）")
w()
w("--- 全年汇总（填报区间 334 天）---")
w("   总计划购电费 = %.4f 元" % rep[:, 0].sum())
w("   日均 = %.4f 元" % rep[:, 0].mean())
w("   总购电量 = %.4f kWh" % rep[:, 1].sum())
w("   总充电量 = %.4f kWh ; 总放电量 = %.4f kWh" % (rep[:, 2].sum(), rep[:, 3].sum()))
w("   填报区间初始储电量 E0(2.1) = %.4f kWh ; 末日 E(12.31 24:00) = %.4f kWh" % (rep[0, 4], rep[-1, 5]))
w("   储电量范围: min %.4f, max %.4f ; 触及下限天数 = %d, 触及上限天数 = %d"
  % (rep[:, 4].min(), rep[:, 5].max(),
     int((np.abs(daily[:, 5] - EMIN) < 1e-3).sum()), int((np.abs(daily[:, 5] - EMAX) < 1e-3).sum())))
w("   每日 24:00 储电量 = 下限 1200 的天数（全 365 天）= %d"
  % int((np.abs(daily[:, 5] - EMIN) < 1e-3).sum()))
w()

# ---- 与问题 1 典型日对照 ----
w("--- 与问题 1 对照 ---")
w("   问题 1 典型日（端点锁定）= 35126.9486 元 ; 端点自由 = 32909.8653 元")
w("   问题 2 日均 = %.4f 元，相对问题 1 端点锁定 %+.4f%%（凸性效应：逐日波动使日均略高属正常）"
  % (rep[:, 0].mean(), 100 * (rep[:, 0].mean() / 35126.9486 - 1)))
w()

# ---- 指定日期表 1/表 2 ----
DATES = {2025 * 10000 + 320: 31 + 31 + 28 + 19,  # 占位，下面用日期直接找
         }
import datetime as dt
for target in ("2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21"):
    d0 = (dt.date.fromisoformat(target) - dt.date(2025, 1, 1)).days
    r = recs[d0]
    E0d = daily[d0, 4]; E1d = daily[d0, 5]
    w("--- 表 1/表 2：%s（d=%d）---" % (target, d0 + 1))
    for H in (10, 12, 14, 16, 18, 20):
        k = 6 * H  # 当天第 6H+1 个时段，0 基索引 6H
        w("   %02d:00-%02d:10  x = %.4f kWh" % (H, H, r["x"][k]))
    w("   全天购电量 = %.4f kWh ; 全天购电费 = %.4f 元" % (r["x"].sum(), r["cost"]))
    # 题目表 2 规定的是六个 **4 小时块**（每块 24 个时段），不是 6 小时块
    blk = [(0, 24), (24, 48), (48, 72), (72, 96), (96, 120), (120, 144)]
    for a, b2 in blk:
        w("   %02d:00-%02d:00  充 = %.4f  放 = %.4f" % (a // 6, b2 // 6, r["u"][a:b2].sum(), r["v"][a:b2].sum()))
    w("   0:00 储电量 = %.4f ; 24:00 储电量 = %.4f" % (E0d, E1d))
    # 紧急购电（读法①：恒 0；此处按定义验证）
    residual = L2[d0] * DT + r["u"] - r["x"] - P2[d0] * DT - r["v"]
    w("   紧急购电验证: max(r) = %.6f kWh（读法① 应为 0）" % max(0.0, float(residual.max())))
    w()

with open(P("_phase0", "报告17_问题2量级体检.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done; 日均费用 = %.4f 元" % rep[:, 0].mean())
