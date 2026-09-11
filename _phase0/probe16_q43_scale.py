"""构思手独立复核：问题 4-3（附件4 波动电价 + 多阶段调整 + D-10 终端余值）。

复核策略：调用同一库（lib.forecast 分解 + lib.solve_day.solve_stage/settle_*），
但**驱动循环自写**（不经过 run_days.solve_rolling_staged 的聚合代码），独立复现
编程手 4-3 主模型总费用（其自报 J = 14 450 082.3797 元）。

口径：D-13 四阶段串联不可追溯；D-12 结算；D-10 余值 V_E=次日最低价/η（末日取当日）；
      读法① 完全信息（计划阶段用预报、评估用实际）。
"""

import os
import sys
import numpy as np

WORK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, WORK)
P = lambda *a: os.path.join(WORK, *a)
out = []
def w(s=""):
    out.append(str(s))

from lib.dataio import read_attachment2, read_attachment4
from lib.solve_day import solve_stage, settle_deviation, settle_total
from lib.forecast import load_forecast
from lib.storage import ETA, E_INIT, E_MAX, E_MIN, P_MAX
from lib.timegrid import DT_H, K

# 数据
price4, dates4 = read_attachment4()                     # (365,144)
load, pv_actual, _ = read_attachment2()                 # 各 (365,144)
pv_fc, pv_fc144, _, tau_list = load_forecast(method="linear")   # (365,4,24), (365,4,144)
D = 365
KMARKS = {0: 0, 6: 36, 12: 72, 18: 108}
STAGES = (0, 6, 12, 18)

# D-10 余值：V_E^d = 次日最低价 / η；末日取当日最低价
ve_day = np.array([price4[d + 1].min() if d + 1 < D else price4[d].min() for d in range(D)]) / ETA

e_cur = E_INIT
daily = []
for d in range(D):
    price = price4[d]                                   # 当日价格 (144,)
    x = np.zeros(K); y = np.zeros(K); u = np.zeros(K); v = np.zeros(K)
    E_full = np.zeros(K + 1); E_full[0] = e_cur
    e = e_cur
    stages_list = [(0, 0), (6, 36), (12, 72), (18, 108)]
    for si, (tau, ks) in enumerate(stages_list):
        n_seg = K - ks
        ve_seg = ve_day[d] if tau == 18 else 0.0
        x_ref = x[ks:] if tau > 0 else None
        try:
            r = solve_stage(price[ks:], load[d][ks:], pv_fc144[d, tau // 6][ks:],
                            e_init=e, x_ref=x_ref, k_start=ks, n_period=n_seg, v_end=ve_seg)
        except RuntimeError:
            print("INFEASIBLE at d=%d tau=%d e=%.4f" % (d, tau, e))
            raise
        # 写入决策
        y[ks:] = r["x_plan"]; u[ks:] = r["u_chg"]; v[ks:] = r["v_dis"]
        if tau == 0:
            x[:] = r["x_plan"]
        # 已执行段轨迹写入（用于跨阶段储电量衔接）
        E_full[ks + 1:ks + n_seg + 1] = r["E_soc"][1:]
        # 下一阶段起点储电量 = 本阶段轨迹在"下一决策时刻"处的值（而非阶段末端）
        if si + 1 < len(stages_list):
            ks_next = stages_list[si + 1][1]
            e = float(r["E_soc"][ks_next - ks])
        else:
            e = float(r["E_soc"][-1])
    E_traj = E_full
    # 评估：实际光伏
    r_emg = np.maximum(load[d] * DT_H + u - y - pv_actual[d] * DT_H - v, 0.0)
    st = settle_total(price, x, y, r_emg)
    J_plan, J_adj, J_emg = st["J_plan"], st["J_adj"], st["J_emg"]
    daily.append((d, J_plan, J_adj, J_emg, J_plan + J_adj + J_emg,
                  float(y.sum()), float(r_emg.sum()), float(r_emg.max())))
    e_cur = float(E_traj[-1])

daily = np.array(daily)
rep = daily[31:]
w("### 问题 4-3 构思手独立复核（自写驱动循环 + 库函数）")
w()
w("--- 填报区间 334 天 ---")
w("   J_plan = %.4f 元" % rep[:, 1].sum())
w("   J_adj  = %.4f 元" % rep[:, 2].sum())
w("   J_emg  = %.4f 元" % rep[:, 3].sum())
w("   总费用 J = %.4f 元（日均 %.4f）" % (rep[:, 4].sum(), rep[:, 4].mean()))
w("   Σy = %.4f kWh ; Σr = %.4f kWh ; 触发紧急天数 = %d ; max r = %.4f"
  % (rep[:, 5].sum(), rep[:, 6].sum(), int((rep[:, 7] > 1e-9).sum()), rep[:, 7].max()))
w()
w("对照编程手自报：J_plan=12 940 602.0189, J_adj=237 146.5459, J_emg=1 272 333.8149, J=14 450 082.3797")
w("对照构思手报告20（4-2 余值）= 12 815 460.2655 元")

with open(P("_phase0", "报告21_问题4-3复核.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done; J = %.4f" % rep[:, 4].sum())
