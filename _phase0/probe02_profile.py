"""Phase 0 探针 2：量纲/量级/缺失/异常/对齐关系核查。输出 UTF-8 报告。"""

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


# ---------------- 附件1 ----------------
r1 = load_sheet(P("附件", "附件1.xlsx"))
h1 = r1[0]
t1 = [r[0] for r in r1[1:]]
price1 = np.array([float(r[1]) for r in r1[1:]])
load1 = np.array([float(r[2]) for r in r1[1:]])
pv1 = np.array([float(r[3]) for r in r1[1:]])
w("### 附件1")
w("表头: %s" % (h1,))
w("样本数 = %d" % len(price1))
w("时间列 前3 = %s ; 后3 = %s" % (t1[:3], t1[-3:]))
w("时间列类型 = %s" % sorted({type(x).__name__ for x in t1}))
w("电价: min=%.4f max=%.4f mean=%.4f 缺失=%d" % (price1.min(), price1.max(), price1.mean(), int(np.isnan(price1).sum())))
w("负载: min=%.4f max=%.4f mean=%.4f 缺失=%d" % (load1.min(), load1.max(), load1.mean(), int(np.isnan(load1).sum())))
w("光伏: min=%.4f max=%.4f mean=%.4f 缺失=%d 负值数=%d" % (pv1.min(), pv1.max(), pv1.mean(), int(np.isnan(pv1).sum()), int((pv1 < 0).sum())))
w("光伏>0 的时间下标范围: 首个=%s 末个=%s" % (t1[int(np.argmax(pv1 > 0))], t1[len(pv1) - 1 - int(np.argmax(pv1[::-1] > 0))]))
w("全天 负载*0.1h 之和 = %.4f kWh ; 光伏*0.1h 之和 = %.4f kWh" % (load1.sum() * 0.1, pv1.sum() * 0.1))
w("表1 指定时段在附件1中的值:")
for lab, idx in [("10:00-10:10", 0), ("12:00-12:10", 0), ("14:00-14:10", 0), ("16:00-16:10", 0), ("18:00-18:10", 0), ("20:00-20:10", 0)]:
    for i in range(len(price1)):
        if i * 10 == 0:
            pass
    pass

# ---------------- 附件2 ----------------
r2a = load_sheet(P("附件", "附件2.xlsx"), "小区负载")
r2b = load_sheet(P("附件", "附件2.xlsx"), "光伏发电实际功率")
w()
w("### 附件2")
w("小区负载表: 行=%d 列=%d" % (len(r2a), len(r2a[0])))
w("时间表头 前3=%s 后3=%s" % (r2a[0][1:4], r2a[0][-3:]))
w("日期列 前2=%s 后2=%s" % (r2a[1][0], r2a[-1][0]))
L = np.array([[float(v) for v in r[1:]] for r in r2a[1:]])
V = np.array([[float(v) for v in r[1:]] for r in r2b[1:]])
w("负载矩阵 shape=%s  NaN=%d  min=%.4f max=%.4f" % (L.shape, int(np.isnan(L).sum()), np.nanmin(L), np.nanmax(L)))
w("光伏矩阵 shape=%s  NaN=%d  min=%.4f max=%.4f 负值=%d" % (V.shape, int(np.isnan(V).sum()), np.nanmin(V), np.nanmax(V), int((V < 0).sum())))
w("负载全年日总量 min=%.1f max=%.1f mean=%.1f kWh" % tuple(np.sort([L.sum(1).min(), L.sum(1).max(), L.sum(1).mean()])))
w("光伏全年日总量 min=%.1f max=%.1f mean=%.1f kWh" % tuple(np.sort([V.sum(1).min(), V.sum(1).max(), V.sum(1).mean()])))
dts = [r[0] for r in r2a[1:]]
w("日期序列 首=%s 末=%s 个数=%d" % (dts[0], dts[-1], len(dts)))

# 小数位模式（判断哪些是原始半小时值、哪些是插值）
def dec_pattern(vec, name):
    cnt = {}
    for v in vec:
        s = ("%.6f" % v).rstrip("0")
        d = len(s.split(".")[1]) if "." in s else 0
        cnt[d] = cnt.get(d, 0) + 1
    w("%s 小数位数分布: %s" % (name, dict(sorted(cnt.items()))))
dec_pattern(L[0], "负载 2025-01-01")
dec_pattern(V[0], "光伏 2025-01-01")
# 半小时点(索引为 2,5,8...)与其他点的差异
w("负载矩阵中 索引%%3==2 的列(即XX:30整点半)、分辨率检查:")
w("  列索引样本 0..8 对应时间 = %s" % [str(x) for x in r2a[0][1:10]])

# ---------------- 附件2 vs 附件1 的匹配 ----------------
w()
w("### 附件1 与 附件2/附件4 的对齐检验")
cand = []
for i, d in enumerate(dts):
    for shift in (0, -1, 1):
        pass
for shift in (0, 1, -1):
    if shift >= 0:
        a = load1[shift:]; b = L[:, :len(a)]
    else:
        a = load1[:shift]; b = L[:, -len(a):]
    err = np.abs(b - a[None, :]).max(1)
    j = int(np.argmin(err))
    cand.append(("load shift=%d" % shift, dts[j], err[j]))
for name, d, e in cand:
    w("  %s -> 最匹配日 %s, 最大绝对差=%.4f" % (name, d, e))

r4 = load_sheet(P("附件", "附件4.xlsx"))
h4 = r4[0]
PR = np.array([[float(v) for v in r[1:]] for r in r4[1:]])
w("附件4: 行=%d 列=%d 时间表头前3=%s 后3=%s" % (len(r4), len(r4[0]), h4[1:4], h4[-3:]))
w("附件4 电价 min=%.4f max=%.4f mean=%.4f NaN=%d" % (np.nanmin(PR), np.nanmax(PR), np.nanmean(PR), int(np.isnan(PR).sum())))
for shift in (0, 1, -1):
    if shift >= 0:
        a = price1[shift:]; b = PR[:, :len(a)]
    else:
        a = price1[:shift]; b = PR[:, -len(a):]
    err = np.abs(b - a[None, :]).max(1)
    j = int(np.argmin(err))
    w("  电价 shift=%d -> 最匹配日 %s, 最大绝对差=%.4f, 全日平均绝对差=%.4f" % (shift, dts[j], err[j], np.abs(b[j] - a).mean()))

# 附件1 的价格与附件4 各日价格的日平均相关性
w("附件1 电价日形状与附件4 各日相关系数最高/最低:")
cc = []
for i in range(PR.shape[0]):
    cc.append(np.corrcoef(PR[i], price1)[0, 1])
cc = np.array(cc)
w("  最高 corr=%.4f @ %s ; 最低 corr=%.4f @ %s ; 中位=%.4f" % (cc.max(), dts[int(np.argmax(cc))], cc.min(), dts[int(np.argmin(cc))], np.median(cc)))

# 附件1 负载是否等于附件2 某日
ccL = []
for i in range(L.shape[0]):
    ccL.append(np.corrcoef(L[i], load1)[0, 1])
ccL = np.array(ccL)
w("附件1 负载与附件2 各日负载 corr 最高=%.4f @ %s ; 中位=%.4f" % (ccL.max(), dts[int(np.argmax(ccL))], np.median(ccL)))
j = int(np.argmax(ccL))
w("  该日最大绝对差=%.4f, 平均绝对差=%.4f" % (np.abs(L[j] - load1).max(), np.abs(L[j] - load1).mean()))

# 附件1 光伏(预测) 与 附件2 光伏实际 某日
ccV = []
for i in range(V.shape[0]):
    if V[i].std() > 0:
        ccV.append(np.corrcoef(V[i], pv1)[0, 1])
    else:
        ccV.append(-9)
ccV = np.array(ccV)
w("附件1 光伏(预测) 与附件2 各日光伏实际 corr 最高=%.4f @ %s" % (ccV.max(), dts[int(np.argmax(ccV))]))
j = int(np.argmax(ccV))
w("  该日绝对差 max=%.4f mean=%.4f ; 附件1 光伏日总量=%.1f kWh, 该日实际=%.1f kWh" % (np.abs(V[j] - pv1).max(), np.abs(V[j] - pv1).mean(), pv1.sum() * 0.1, V[j].sum() * 0.1))

# ---------------- 附件3 ----------------
r3 = load_sheet(P("附件", "附件3.xlsx"))
w()
w("### 附件3")
w("行=%d 列=%d 表头=%s" % (len(r3), len(r3[0]), r3[0]))
w("前5行 日期/预报时刻:")
for r in r3[1:6]:
    w("   %s | %s" % (r[0], r[1]))
w("末4行 日期/预报时刻:")
for r in r3[-4:]:
    w("   %s | %s" % (r[0], r[1]))
allrows = r3[1:]
w("总行数(数据)=%d ; 预报时刻取值集合=%s" % (len(allrows), sorted({r[1] for r in allrows})))
F = np.array([[float(v) for v in r[2:]] for r in allrows])
w("预报值矩阵 shape=%s min=%.4f max=%.4f NaN=%d 负值=%d" % (F.shape, np.nanmin(F), np.nanmax(F), int(np.isnan(F).sum()), int((F < 0).sum())))
w("单条预报24小时总量 min=%.1f max=%.1f mean=%.1f" % (F.sum(1).min(), F.sum(1).max(), F.sum(1).mean()))

# 预报时刻 -> 该日 0:00 预报 与 附件1 的光伏预测比
d0 = allrows[0]
w("2025-1-1 0:00 预报 = %s" % (list(F[0]),))
w("附件1 光伏预测(144点, 10min) 逐点 = %s" % (list(np.round(pv1, 4)),))

# 逐日 0:00 预报(24 小时) 与 附件2 该日实际 的对应关系：尝试各种小时偏移
w()
w("### 附件3 预报 -> 实际 时间对齐（用 0:00 预报 vs 附件2 当日实际小时的相关系数）")
# 附件2 的小时均值（10min 数据 -> 小时均值），小时 h 对应 列 (h-1)*6 .. h*6-1（列为 0:10..0:00+1）
H = V.reshape(V.shape[0], 24, 6).mean(2)  # 每小时 6 个 10 分钟点
w("小时均值矩阵 shape=%s" % (H.shape,))
# 只看 0:00 预报
idx0 = [i for i, r in enumerate(allrows) if r[1] == "0:00"]
F0 = F[idx0]
d0list = [allrows[i][0] for i in idx0]
w("0:00 预报条数=%d 首日=%s 末日=%s" % (len(idx0), d0list[0], d0list[-1]))
# 日期 -> 附件2 行号
d2idx = {d.strftime("%Y-%m-%d"): i for i, d in enumerate(dts)}
w("日期串格式样例: 附件3=%r 附件2=%r" % (d0list[0], dts[0].strftime("%Y-%m-%d")))
w("日期串做 %s 到 %s 的规范化示例" % (d0list[0], dt.datetime.strptime(d0list[0], "%Y-%m-%d")))
best = None
for k in range(24):
    pass
for off in range(-3, 4):
    xs, ys = [], []
    for i, ds in enumerate(d0list):
        d = dt.datetime.strptime(ds, "%Y-%m-%d")
        j = d2idx.get(d.strftime("%Y-%m-%d"))
        if j is None:
            continue
        for k in range(24):
            h = (k + 1 + off) % 24
            dd = d + dt.timedelta(days=(k + 1 + off) // 24)
            jj = d2idx.get(dd.strftime("%Y-%m-%d"))
            if jj is None:
                continue
            xs.append(F0[i, k]); ys.append(H[jj, h])
    xs = np.array(xs); ys = np.array(ys)
    w("  偏移 off=%+d (预报k小时 -> 时刻 k+1+off 点): corr=%.4f  样本=%d  RMSE=%.1f  MAE=%.1f"
      % (off, np.corrcoef(xs, ys)[0, 1], len(xs), float(np.sqrt(((xs - ys) ** 2).mean())), float(np.abs(xs - ys).mean())))

with open(P("_phase0", "报告02_数据体检.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done")
