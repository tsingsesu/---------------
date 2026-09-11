"""分支答疑探针：判定"整点预报 → 10 分钟"该怎么分解，以及日内插值结构。

问题：
  Q1 附件2 的 10 分钟序列，在整点之间是不是由整点值按某种固定形状插值出来的？
     （若形状固定，则"整点预报 → 10 分钟"可以近乎精确地反演）
  Q2 用 24 个整点预报构造 144 个 10 分钟估计，哪种分解方式的误差最小？
     M1 整点值在该小时内保持不变（分段常数）
     M2 相邻整点值线性插值
     M3 整点值 × 历史平均的"小时内相对形状"
     M4 整点值 × 典型日（附件1）的"小时内相对形状"
"""

import os
import numpy as np
import openpyxl

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
A1 = np.array([float(r[3]) for r in r1[1:]])            # 附件1 光伏（典型日）
r2b = load_sheet(P("附件", "附件2.xlsx"), "光伏发电实际功率")
r2a = load_sheet(P("附件", "附件2.xlsx"), "小区负载")
r4 = load_sheet(P("附件", "附件4.xlsx"))
V = np.array([[float(v) for v in r[1:]] for r in r2b[1:]])     # (365,144) 光伏实际
L = np.array([[float(v) for v in r[1:]] for r in r2a[1:]])     # (365,144) 负载
PR = np.array([[float(v) for v in r[1:]] for r in r4[1:]])     # (365,144) 电价
r3 = load_sheet(P("附件", "附件3.xlsx"))
allrows = r3[1:]
F = np.array([[float(v) for v in r[2:]] for r in allrows])
idx0 = [i for i, r in enumerate(allrows) if r[1] == "0:00"]
F0 = F[idx0]                                                   # (365,24) 0:00 预报

w("### Q1 日内插值结构检验")
w("做法：对间隔 s（列）的每一段，取段首 A、段末 B，计算 r_m=(x[c0+m]-A)/(B-A)，")
w("      若序列是由间隔 s 的结点插值而来，r_m 应高度集中（标准差≈0）。")
w("")


def structure_test(X, name, thr, s_list=(6, 3, 2)):
    w("--- %s ---" % name)
    for s in s_list:
        pooled = [[] for _ in range(s - 1)]
        seg = 0
        for d in range(X.shape[0]):
            x = X[d]
            for c0 in range(0, 144 - s, s):
                A = x[c0]; B = x[c0 + s]
                if abs(B - A) < thr:
                    continue
                seg += 1
                for m in range(1, s):
                    pooled[m - 1].append((x[c0 + m] - A) / (B - A))
        w("   间隔 s=%d（%d 段参与）:" % (s, seg))
        for m, arr in enumerate(pooled, start=1):
            arr = np.array(arr)
            w("      r_%d: 均值=%7.4f 标准差=%7.4f 中位=%7.4f  线性基准=%.4f"
              % (m, arr.mean(), arr.std(), np.median(arr), m / s))


structure_test(V, "附件2 光伏实际（阈值 500 kW）", 500.0)
structure_test(L, "附件2 负载（阈值 300 kW）", 300.0)
structure_test(PR, "附件4 电价（阈值 0.10 元/kWh）", 0.10)

w("")
w("### Q2 整点预报 -> 144 个 10 分钟估计 的误差对比（用 0:00 预报覆盖全天 365 天）")
w("时间对齐：预报 m 小时 = 整点 m:00 的值 = 第 6m-1 列（0 基），所在小时块为第 6(m-1)..6m-1 列。")
w("评价：与附件2 实际光伏逐列比较；白天（预报>200 kW）与全样本分别统计。")

ACT = V.copy()
# 评价掩码
mask_all = np.ones_like(ACT, dtype=bool)
mask_day = np.zeros_like(ACT, dtype=bool)
for m in range(1, 25):
    mask_day[:, 6 * (m - 1):6 * m] = (F0[:, m - 1] > 200)[:, None]

# M3 的历史形状：对每个小时块位置 j=0..5，统计 V[块内第 j 列] / V[块末列(整点)]
shape_num = np.zeros(6); shape_cnt = np.zeros(6)
for m in range(1, 25):
    c0 = 6 * (m - 1)
    endv = ACT[:, c0 + 5]
    ok = endv > 500
    for j in range(6):
        shape_num[j] += (ACT[ok, c0 + j] / endv[ok]).sum()
        shape_cnt[j] += ok.sum()
S3 = shape_num / shape_cnt
S3 = S3 / S3[5]
w("   历史小时内相对形状 S3（相对该小时整点值） = %s" % np.round(S3, 4))

# M4 典型日形状
S4 = []
for j in range(6):
    num = 0.0; cnt = 0.0
    for h in range(24):
        c0 = 6 * h
        endv = A1[c0 + 5]
        if endv > 500:
            num += A1[c0 + j] / endv
            cnt += 1
    S4.append(num / cnt if cnt else np.nan)
S4 = np.array(S4); S4 = S4 / S4[5]
w("   典型日小时内相对形状 S4（附件1）        = %s" % np.round(S4, 4))

est = {}
# M1 分段常数
e1 = np.zeros_like(ACT)
for m in range(1, 25):
    e1[:, 6 * (m - 1):6 * m] = F0[:, m - 1][:, None]
est["M1 分段常数（整点值覆盖整小时）"] = e1
# M2 线性插值（整点值作为结点）
e2 = np.zeros_like(ACT)
prev = np.zeros(365)
for m in range(1, 25):
    c0 = 6 * (m - 1)
    B = F0[:, m - 1]
    for j in range(6):
        e2[:, c0 + j] = prev + (B - prev) * (j + 1) / 6.0
    prev = B
est["M2 线性插值（整点值为结点）"] = e2
# M3 / M4 形状缩放
for nm, S in [("M3 整点值 x 历史小时内形状", S3), ("M4 整点值 x 典型日小时内形状", S4)]:
    e = np.zeros_like(ACT)
    for m in range(1, 25):
        c0 = 6 * (m - 1)
        for j in range(6):
            e[:, c0 + j] = F0[:, m - 1] * S[j]
    est[nm] = e

# 参考：直接用实际值（下界 0）
w("")
w("   方法                                          MAE(白天)   RMSE(白天)   MAE(全样本)  RMSE(全样本)  RMSE(夜间)")
for nm, e in est.items():
    d_ = e - ACT
    w("   %-42s %9.1f %11.1f %11.1f %12.1f %12.1f"
      % (nm,
         np.abs(d_[mask_day]).mean(), np.sqrt((d_[mask_day] ** 2).mean()),
         np.abs(d_[mask_all]).mean(), np.sqrt((d_[mask_all] ** 2).mean()),
         np.sqrt((d_[~mask_day] ** 2).mean())))

# 偏移与偏差
w("")
w("   各方法的系统性偏差（估计 - 实际，白天）：")
for nm, e in est.items():
    d_ = (e - ACT)[mask_day]
    w("      %-42s 均值偏差=%8.1f  低估比例=%.3f  高估比例=%.3f"
      % (nm, d_.mean(), (d_ < -1e-9).mean(), (d_ > 1e-9).mean()))

# 逐小时整点处的误差（检验对齐是否正确）
w("")
w("   仅看整点列（第 6m-1 列，0 基）的误差：")
for nm, e in est.items():
    cols = [6 * m - 1 for m in range(1, 25)]
    d_ = e[:, cols] - ACT[:, cols]
    w("      %-42s MAE=%7.1f RMSE=%8.1f" % (nm, np.abs(d_).mean(), np.sqrt((d_ ** 2).mean())))
w("   仅看非整点列（每小时的 6m-6..6m-2 列）的误差：")
for nm, e in est.items():
    cols = [c for m in range(1, 25) for c in range(6 * (m - 1), 6 * m - 1)]
    d_ = e[:, cols] - ACT[:, cols]
    w("      %-42s MAE=%7.1f RMSE=%8.1f" % (nm, np.abs(d_).mean(), np.sqrt((d_ ** 2).mean())))

with open(P("_phase0", "报告12_整点预报分解判定.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done")
