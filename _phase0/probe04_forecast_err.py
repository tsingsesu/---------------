"""Phase 0 探针 4：预报误差结构、有效分辨率（插值痕迹）、价格生成规律、基线成本。"""

import os
import datetime as dt
import openpyxl
import numpy as np

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
r2a = load_sheet(P("附件", "附件2.xlsx"), "小区负载")
r2b = load_sheet(P("附件", "附件2.xlsx"), "光伏发电实际功率")
r4 = load_sheet(P("附件", "附件4.xlsx"))
r3 = load_sheet(P("附件", "附件3.xlsx"))
price1 = np.array([float(r[1]) for r in r1[1:]])
load1 = np.array([float(r[2]) for r in r1[1:]])
pv1 = np.array([float(r[3]) for r in r1[1:]])
dts = [r[0] for r in r2a[1:]]
dkey = [d.strftime("%Y-%m-%d") for d in dts]
d2idx = {k: i for i, k in enumerate(dkey)}
L = np.array([[float(v) for v in r[1:]] for r in r2a[1:]])
V = np.array([[float(v) for v in r[1:]] for r in r2b[1:]])
PR = np.array([[float(v) for v in r[1:]] for r in r4[1:]])
allrows0 = r3[1:]
F = np.array([[float(v) for v in r[2:]] for r in allrows0])
_cur = None; allrows = []
for r in allrows0:
    if r[0] not in (None, ""):
        _cur = str(r[0]).strip()
    allrows.append((_cur, r[1]))

# ---------- 1. 附件3 预报 = 该整点的 10 分钟实际值？ ----------
w("### 1. 附件3 预报与附件2 实际在“整点采样点”上的关系（读法C）")
D0 = dt.timedelta(days=1)
rec = []
for i, (ds, iss) in enumerate(allrows):
    d = dt.datetime.strptime(ds, "%Y-%m-%d")
    h0 = int(iss.split(":")[0])
    for m in range(1, 25):
        absdt = d + dt.timedelta(hours=h0 + m)
        j = d2idx.get(absdt.strftime("%Y-%m-%d"))
        if j is None:
            continue
        c = 143 if absdt.hour == 0 else absdt.hour * 6 - 1
        rec.append((i, m, F[i, m - 1], V[j, c]))
rec = np.array(rec, dtype=float)
f = rec[:, 2]; v = rec[:, 3]
diff = f - v
w("样本数 = %d" % len(rec))
w("完全相等(|差|<1e-9) 的比例 = %.4f" % float((np.abs(diff) < 1e-9).mean()))
w("|差|<1e-6 的比例 = %.4f" % float((np.abs(diff) < 1e-6).mean()))
nz = diff[np.abs(diff) >= 1e-9]
w("非零差 个数 = %d ; 均值=%.2f 中位=%.2f 标准差=%.2f 最小=%.2f 最大=%.2f"
  % (len(nz), nz.mean() if len(nz) else 0, np.median(nz) if len(nz) else 0,
     nz.std() if len(nz) else 0, nz.min() if len(nz) else 0, nz.max() if len(nz) else 0))
w("按预报时段 m(1..24) 统计 完全相等比例 与 平均绝对差:")
for m in range(1, 25):
    s = rec[rec[:, 1] == m]
    d_ = s[:, 2] - s[:, 3]
    w("   m=%2d n=%4d 相等比例=%.3f 平均绝对差=%8.1f 最大绝对差=%9.1f"
      % (m, len(s), float((np.abs(d_) < 1e-9).mean()), float(np.abs(d_).mean()), float(np.abs(d_).max())))
# 按预报时刻
w("按预报时刻统计:")
for iss in ["0:00", "6:00", "12:00", "18:00"]:
    idxs = [i for i, (ds, s) in enumerate(allrows) if s == iss]
    sel = np.isin(rec[:, 0], idxs)
    d_ = rec[sel, 2] - rec[sel, 3]
    w("   %s: n=%d 相等比例=%.3f 平均绝对差=%.1f RMSE=%.1f"
      % (iss, sel.sum(), float((np.abs(d_) < 1e-9).mean()), float(np.abs(d_).mean()), float(np.sqrt((d_ ** 2).mean()))))
# 非零差是否与真值成比例
mask = np.abs(diff) >= 1e-9
if mask.sum() > 10:
    rr = f[mask] / np.where(np.abs(v[mask]) < 1e-6, np.nan, v[mask])
    rr = rr[np.isfinite(rr)]
    w("非零点 预报/实际 比值: 中位=%.4f 25%%=%.4f 75%%=%.4f" % (np.median(rr), np.percentile(rr, 25), np.percentile(rr, 75)))
    w("非零点 实际值 分位: %s" % np.round(np.percentile(v[mask], [0, 25, 50, 75, 100]), 1))
    w("非零点 预报值 分位: %s" % np.round(np.percentile(f[mask], [0, 25, 50, 75, 100]), 1))

# ---------- 2. 有效时间分辨率：二阶差分是否在小时边界跳变 ----------
w()
w("### 2. 有效分辨率：10 分钟序列在小时内部是否为线性插值（二阶差分诊断）")


def d2diag(series, name):
    s = np.asarray(series, dtype=float)
    d2 = s[2:] - 2 * s[1:-1] + s[:-2]
    pos = np.arange(len(d2))          # d2[k] 对应 s[k], s[k+1], s[k+2]
    atbound = (pos % 6 == 4)          # (s[5],s[6],s[7]) 型：跨界
    w("%s: |二阶差分| 均值=%.4g" % (name, np.abs(d2).mean()))
    w("   跨小时位置(索引%%6==4) 的 |二阶差分| 均值=%.4g ; 小时内部 均值=%.4g"
      % (np.abs(d2[atbound]).mean(), np.abs(d2[~atbound]).mean()))
    w("   跨小时/内部 比值 = %.2f" % (np.abs(d2[atbound]).mean() / max(np.abs(d2[~atbound]).mean(), 1e-12)))


d2diag(load1, "附件1 负载")
d2diag(price1, "附件1 电价")
d2diag(pv1, "附件1 光伏")
d2diag(L[0], "附件2 负载 2025-01-01")
d2diag(V[0], "附件2 光伏 2025-01-01")
d2diag(PR[0], "附件4 电价 2025-01-01")
d2diag(L.mean(0), "附件2 负载 全年逐列均值")
d2diag(PR.mean(0), "附件4 电价 全年逐列均值")
d2diag(V.mean(0), "附件2 光伏 全年逐列均值")

# 对比：跨半小时(索引%6==1 或 4) 与 其他
w("补充：把'跨小时'放宽为 索引%3==1 的位置")
for name, s in [("附件1 电价", price1), ("附件1 负载", load1), ("附件4 电价均值", PR.mean(0))]:
    s = np.asarray(s, float)
    d2 = s[2:] - 2 * s[1:-1] + s[:-2]
    pos = np.arange(len(d2))
    a = (pos % 3 == 1)
    w("   %s: 跨半小时均值=%.4g 内部均值=%.4g 比值=%.2f" % (name, np.abs(d2[a]).mean(), np.abs(d2[~a]).mean(), np.abs(d2[a]).mean() / np.abs(d2[~a]).mean()))

# ---------- 3. 附件4 电价与附件1 电价的生成关系 ----------
w()
w("### 3. 附件4 电价与附件1 电价（全年均值型）的关系")
ratio = PR / price1[None, :]
w("比值 分位: %s" % np.round(np.percentile(ratio, [0, 1, 5, 25, 50, 75, 95, 99, 100]), 4))
w("比值 标准差(逐点) = %.4f" % ratio.std())
w("逐日 比值 的日内标准差 均值 = %.5f" % ratio.std(1).mean())
w("逐日 比值 的日均值: 均值=%.4f 标准差=%.4f 最小=%.4f 最大=%.4f" % (ratio.mean(1).mean(), ratio.mean(1).std(), ratio.mean(1).min(), ratio.mean(1).max()))
w("对数比 ln(PR/附件1) 均值=%.5f 标准差=%.5f" % (np.log(ratio).mean(), np.log(ratio).std()))
# 是否 PR = price1 * (1+a_d) + 噪声
a = PR.mean(1) / price1.mean()
resid = PR - a[:, None] * price1[None, :]
w("乘性日因子模型残差: |残差| 均值=%.5f ; 相对残差均值=%.4f" % (np.abs(resid).mean(), np.abs(resid / PR).mean()))
# 逐点比值是否只依赖于列(时间)
w("逐列 比值均值 的范围 = [%.4f, %.4f]" % (ratio.mean(0).min(), ratio.mean(0).max()))
w("逐列 比值标准差 的范围 = [%.4f, %.4f]" % (ratio.std(0).min(), ratio.std(0).max()))

# ---------- 4. 附件2 负载与附件1 的关系 ----------
w()
w("### 4. 附件2 负载/光伏与附件1（全年逐列均值）的关系")
rL = L / load1[None, :]
w("负载 逐点比值 分位: %s" % np.round(np.percentile(rL, [0, 1, 50, 99, 100]), 4))
w("负载 逐列比值均值 范围 = [%.4f, %.4f]" % (rL.mean(0).min(), rL.mean(0).max()))
rV = V / np.where(pv1[None, :] < 1e-6, np.nan, pv1[None, :])
w("光伏 逐点比值(仅白天) 分位: %s" % np.round(np.nanpercentile(rV, [0, 1, 25, 50, 75, 99, 100]), 4))

# ---------- 5. 基线成本与关键量 ----------
w()
w("### 5. 关键基线量（10 分钟区间，每个区间时长 = 1/6 h）")
H = 1.0 / 6.0
E_load = load1.sum() * H
E_pv = pv1.sum() * H
cost_all = (price1 * load1).sum() * H
cost_net = (price1 * np.maximum(load1 - pv1, 0)).sum() * H
w("附件1（典型日）: 负载电量=%.2f kWh ; 光伏电量=%.2f kWh ; 净负载=%.2f kWh" % (E_load, E_pv, E_load - E_pv))
w("  净负载/负载 = %.4f" % ((E_load - E_pv) / E_load))
w("  方案A 全部向电网购电(不储能)成本 = %.2f 元" % cost_all)
w("  方案B 光伏先自用、余电弃掉(不储能)成本 = %.2f 元" % cost_net)
w("  两者之差（光伏自用的价值）= %.2f 元" % (cost_all - cost_net))
w("  平均购电单价 = %.4f 元/kWh (典型日)" % (price1.mean()))
w("  峰谷价差: max=%.4f min=%.4f 比值=%.3f" % (price1.max(), price1.min(), price1.max() / price1.min()))
w("附件4 全年电价: max=%.4f min=%.4f 逐日极差均值=%.4f" % (PR.max(), PR.min(), (PR.max(1) - PR.min(1)).mean()))
w("储能参数: 容量 12000 kWh ; 可用区间 1200-10800 kWh (可用 9600 kWh) ; 最大充放电功率 5000 kW ; 效率 90%")
w("  满功率充电 9600 kWh 需时 = %.2f h ; 从 1200 充到 10800 需 %.2f h" % (9600 / 5000, 9600 / 5000))
w("  套利可行性判别：峰谷价比 %.3f ; 若单程效率0.9(往返0.81) 需价比>%.3f ; 若往返0.9 需价比>%.3f"
  % (price1.max() / price1.min(), 1 / 0.81, 1 / 0.9))
w("附件2 全年: 负载年电量=%.1f 万 kWh ; 光伏年电量=%.1f 万 kWh" % (L.sum() * H / 1e4, V.sum() * H / 1e4))
w("附件2 逐日 光伏/负载 比: 均值=%.4f 最小=%.4f 最大=%.4f"
  % ((V.sum(1) / L.sum(1)).mean(), (V.sum(1) / L.sum(1)).min(), (V.sum(1) / L.sum(1)).max()))
w("附件2 逐日 净负载(>0部分)电量: 均值=%.1f kWh" % ((np.maximum(L - V, 0).sum(1) * H).mean()))

with open(P("_phase0", "报告04_误差结构与基线.txt"), "w", encoding="utf-8") as fh:
    fh.write("\n".join(out))
print("done")
