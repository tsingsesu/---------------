"""构思手独立量级体检：问题 3 主模型（多阶段确定性等价滚动，D-07/D-12/D-13 口径）。

模型（与 `问题3/交付清单.md` 的 M1 主模型一致，构思手独立实现）：
  * 0:00 计划：以 0:00 预报（→144 段，D-07 线性插值）为情景，min Σp·y（全价）；
  * tau∈{6,12,18} 调整：以该时刻预报为情景，对 (tau,24:00] 段重解，
    目标 = 结算口径下的边际成本 min Σ[0.5p·y + p·t⁺]，t⁺ ≥ y − x（x 为 0:00 计划）；
    只执行 (tau, tau+6:00] 块（不可追溯 D-13），更远段留给下一决策；
  * 评估：用附件2 实际光伏算 r=[LΔ+u−y−PΔ−v]⁺，按 D-12 结算
    J = Σ[p·min(x,y)+0.5p(x−y)⁺+1.5p(y−x)⁺] + 5Σp·r。
策略：(a) 仅 0:00；(b) +6:00；(c) +6/12；(d) +6/12/18（主模型）；(e) 完全信息（= 问题 2）。
另做退化自检：预报=实际（完全信息）时应复现问题 2 总费用 12254765.7161 元。
"""

import os
import numpy as np
import openpyxl
from scipy.optimize import linprog
from scipy.sparse import coo_matrix, vstack

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


r1 = rd(P("附件", "附件1.xlsx"), None)
p0 = np.array([float(r[1]) for r in r1[1:145]])
rL = rd(P("附件", "附件2.xlsx"), "小区负载")
rP = rd(P("附件", "附件2.xlsx"), "光伏发电实际功率")
L2 = np.array([[float(v) for v in r[1:145]] for r in rL[1:]])
P2 = np.array([[float(v) for v in r[1:145]] for r in rP[1:]])
r3 = rd(P("附件", "附件3.xlsx"), None)
D = 365
FC = np.zeros((D, 4, 24))
for i, r in enumerate(r3[1:]):
    FC[i // 4, i % 4] = [float(v) for v in r[2:26]]

K = 144
DT = 1.0 / 6.0
ETA = 0.9
UMAX = 5000 * DT
EMIN, EMAX, EINIT = 1200.0, 10800.0, 6000.0
KAPPA, KM, KP = 5.0, 0.5, 1.5

# 前缀和索引（向量化构造用，n_seg 固定取最大 144 的三角索引，按需截取）
TRIL_R, TRIL_C = np.tril_indices(K)          # 所有 (k, j<=k) 对


def build_forecast_144(d, tau_idx):
    """第 d 天 tau_idx∈{0,1,2,3} 预报 → 144 段（D-07 线性插值）。"""
    tau = tau_idx * 6
    val = np.zeros(K)
    if tau_idx == 0:
        prev = FC[d - 1, 0, 23] if d > 0 else 0.0
    else:
        prev = FC[d, tau_idx - 1, 5]             # 上一发布时刻预报的"预报6小时"值 (= tau:00)
    for mm in range(tau + 1, 25):
        B = FC[d, tau_idx, mm - tau - 1]
        s = np.arange(1, 7)
        val[(mm - 1) * 6:mm * 6] = prev + (B - prev) * s / 6.0
        prev = B
    return val


FC144 = np.zeros((D, 4, K))
for d in range(D):
    for t in range(4):
        FC144[d, t] = build_forecast_144(d, t)


def solve_seg(price, load, pv, e0, x_ref=None):
    """对一段（从某决策时刻到 24:00）求解。x_ref=None 时目标全价 Σp y；否则边际口径。"""
    n = len(price)
    marg = x_ref is not None
    # 截取三角索引：只保留 j<=k 且 k<n
    m = TRIL_R < n
    tr, tc = TRIL_R[m], TRIL_C[m]
    nv = 3 * n + (n if marg else 0)
    rows, cols, vals = [], [], []
    # 供给行 k：−y + u − v ≤ −(L−P)Δ
    ar = np.arange(n)
    rows.append(np.concatenate([ar, ar, ar]))
    cols.append(np.concatenate([ar, n + ar, 2 * n + ar]))
    vals.append(np.concatenate([-np.ones(n), np.ones(n), -np.ones(n)]))
    b_ub = -(load - pv) * DT
    # 储电量上界行 n+k：Σ_{j≤k}(ηu_j − v_j/η) ≤ EMAX − e0
    rows.append(np.concatenate([n + tr, n + tr]))
    cols.append(np.concatenate([n + tc, 2 * n + tc]))
    vals.append(np.concatenate([ETA * np.ones(len(tr)), -np.ones(len(tr)) / ETA]))
    b_ub = np.concatenate([b_ub, np.full(n, EMAX - e0)])
    # 储电量下界行 2n+k：−Σ_{j≤k}(ηu_j − v_j/η) ≤ e0 − EMIN
    rows.append(np.concatenate([2 * n + tr, 2 * n + tr]))
    cols.append(np.concatenate([n + tc, 2 * n + tc]))
    vals.append(np.concatenate([-ETA * np.ones(len(tr)), np.ones(len(tr)) / ETA]))
    b_ub = np.concatenate([b_ub, np.full(n, e0 - EMIN)])
    if marg:
        # t⁺_k − y_k ≤ −x_ref_k
        rows.append(np.concatenate([3 * n + ar, 3 * n + ar]))
        cols.append(np.concatenate([3 * n + ar, ar]))
        vals.append(np.concatenate([np.ones(n), -np.ones(n)]))
        b_ub = np.concatenate([b_ub, -x_ref])
        c = np.concatenate([0.5 * price, np.zeros(2 * n), price])
    else:
        c = np.concatenate([price, np.zeros(2 * n)])
    A = coo_matrix((np.concatenate(vals), (np.concatenate(rows), np.concatenate(cols))),
                   shape=(A_rows_n(n, marg), nv)).tocsr()
    bounds = [(0, None)] * n + [(0, UMAX)] * n + [(0, UMAX)] * n + ([(0, None)] * n if marg else [])
    res = linprog(c, A_ub=A, b_ub=b_ub, bounds=bounds, method="highs")
    assert res.status == 0, res.message
    y = res.x[:n]; u = res.x[n:2 * n]; v = res.x[2 * n:3 * n]
    E = e0 + np.cumsum(ETA * u - v / ETA)
    return dict(y=y, u=u, v=v, E=E)


def A_rows_n(n, marg):
    return (4 if marg else 3) * n


def run_day(d, e0, taus, fc_override=None):
    """taus: 含 0 的调整时刻集合。fc_override: (4,K) 或 None（用插值预报）。"""
    fcs = fc_override if fc_override is not None else FC144[d]
    x = np.zeros(K); y = np.zeros(K); u = np.zeros(K); v = np.zeros(K)
    r0 = solve_seg(p0, L2[d], fcs[0], e0, None)
    x[:] = r0["y"]; y[:] = r0["y"]; u[:] = r0["u"]; v[:] = r0["v"]
    for tau in (6, 12, 18):
        if tau not in taus:
            continue
        ks = tau * 6
        E_now = e0 + np.sum(ETA * u[:ks] - v[:ks] / ETA)
        r = solve_seg(p0[ks:], L2[d][ks:], fcs[tau // 6][ks:], E_now, x[ks:])
        y[ks:] = r["y"]; u[ks:] = r["u"]; v[ks:] = r["v"]
    return x, y, u, v


def run_year(taus, fc_override_all=None):
    e = EINIT
    rec = []
    for d in range(D):
        ov = None
        if fc_override_all is not None:
            ov = np.tile(fc_override_all[d][None, :], (4, 1))
        x, y, u, v = run_day(d, e, taus, ov)
        r_emg = np.maximum(L2[d] * DT + u - y - P2[d] * DT - v, 0.0)
        Jp = float((p0 * np.minimum(x, y)).sum())
        Ja = float((KM * p0 * np.maximum(x - y, 0) + KP * p0 * np.maximum(y - x, 0)).sum())
        Je = float((KAPPA * p0 * r_emg).sum())
        rec.append((d, Jp, Ja, Je, Jp + Ja + Je, float(y.sum()), float(r_emg.sum()),
                    float(r_emg.max())))
        e = float((e + np.sum(ETA * u - v / ETA)))
    return np.array(rec)


w("### 问题 3 构思手独立量级体检（确定性等价滚动，D-07/D-12/D-13）")
w()
# 退化自检：预报=实际（完全信息）
rec_deg = run_year({0, 6, 12, 18}, fc_override_all=P2)
deg_rep = rec_deg[31:]
w("--- 退化自检（预报=实际光伏，完全信息）---")
w("   总费用 = %.4f 元（应复现问题 2 的 12254765.7161）" % deg_rep[:, 4].sum())
w("   紧急购电总量 = %.6f（应为 0）" % deg_rep[:, 6].sum())
w()
for name, taus in (("(d) 全调整 0/6/12/18", {0, 6, 12, 18}),
                   ("(c) 0/6/12", {0, 6, 12}),
                   ("(b) 0/6", {0, 6}),
                   ("(a) 仅 0:00", {0})):
    rec = run_year(taus)
    rep = rec[31:]
    tot = rep[:, 4].sum()
    w("--- 策略 %s ---" % name)
    w("   总费用 = %.4f 元（日均 %.4f）" % (tot, tot / 334))
    w("   分项: 计划区间=%.4f  调整相关=%.4f  紧急=%.4f"
      % (rep[:, 1].sum(), rep[:, 2].sum(), rep[:, 3].sum()))
    w("   Σy = %.4f kWh ; Σr = %.4f kWh ; 触发紧急天数 = %d ; max r = %.4f"
      % (rep[:, 5].sum(), rep[:, 6].sum(), int((rep[:, 7] > 1e-9).sum()), rep[:, 7].max()))
    w()
w("--- (e) 完全信息下界 = 问题 2 主模型 = 12254765.7161 元 ---")
w()
# 四指定日期的表 1/表 2/表 3（主模型 (d) 全调整）
import datetime as _dt
rec_d = run_year({0, 6, 12, 18})
e = EINIT
# 重跑一遍并保留逐日明细便于取数（复用 run_day）
details = {}
e = EINIT
for d in range(D):
    x, y, u, v = run_day(d, e, {0, 6, 12, 18})
    r_emg = np.maximum(L2[d] * DT + u - y - P2[d] * DT - v, 0.0)
    Et = e + np.cumsum(ETA * u - v / ETA)
    details[d] = (x.copy(), y.copy(), u.copy(), v.copy(), r_emg.copy(), e, float(Et[-1]))
    e = float(Et[-1])

for target in ("2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21"):
    d0 = (_dt.date.fromisoformat(target) - _dt.date(2025, 1, 1)).days
    x, y, u, v, r_emg, E0d, E1d = details[d0]
    w("--- %s（d=%d）表 1/表 2/表 3 ---" % (target, d0 + 1))
    for H in (10, 12, 14, 16, 18, 20):
        k = 6 * H
        w("   %02d:00-%02d:10  y = %.4f kWh" % (H, H, y[k]))
    Jp = float((p0 * np.minimum(x, y)).sum())
    Ja = float((KM * p0 * np.maximum(x - y, 0) + KP * p0 * np.maximum(y - x, 0)).sum())
    Je = float((KAPPA * p0 * r_emg).sum())
    w("   全天: 购电量Σy = %.4f kWh ; 总费用 = %.4f 元（计划%.4f+调整%.4f+紧急%.4f）"
      % (y.sum(), Jp + Ja + Je, Jp, Ja, Je))
    blk = [(0, 24), (24, 48), (48, 72), (72, 96), (96, 120), (120, 144)]
    for a, b2 in blk:
        w("   %02d:00-%02d:00  充 = %.4f  放 = %.4f" % (a // 6, b2 // 6, u[a:b2].sum(), v[a:b2].sum()))
    w("   0:00 储电量 = %.4f ; 24:00 储电量 = %.4f" % (E0d, E1d))
    rk = np.where(r_emg > 1e-6)[0]
    if rk.size:
        # 聚合成连续段（rk 为 0 基段索引：段 k 覆盖 (10k min, 10(k+1) min]）
        segs = []
        s = rk[0]; p_ = rk[0]
        for kk in rk[1:]:
            if kk == p_ + 1: p_ = kk
            else: segs.append((s, p_)); s = kk; p_ = kk
        segs.append((s, p_))
        def _t(idx_end, is_end):
            """段索引 → 时刻字符串。段 k 的起=10k 分、止=10(k+1) 分。"""
            minute = 10 * (idx_end + (1 if is_end else 0))
            h, mm = divmod(minute, 60)
            return "%d:%02d" % (h, mm)
        seg_str = " ".join("%s-%s(%.4f)" % (_t(a, False), _t(b2, True), r_emg[a:b2 + 1].sum()) for a, b2 in segs)
        w("   紧急购电: Σr = %.4f kWh, 段: %s" % (r_emg.sum(), seg_str))
    else:
        w("   紧急购电: 无（Σr = 0）")
    w()

with open(P("_phase0", "报告19_问题3量级体检.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done")
