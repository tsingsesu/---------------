"""Phase 0 探针 3：时间对齐约定、有效分辨率、异常值、结果模板标签逐条抄录。"""

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


r2a = load_sheet(P("附件", "附件2.xlsx"), "小区负载")
r2b = load_sheet(P("附件", "附件2.xlsx"), "光伏发电实际功率")
r4 = load_sheet(P("附件", "附件4.xlsx"))
r3 = load_sheet(P("附件", "附件3.xlsx"))
dts = [r[0] for r in r2a[1:]]
dkey = [d.strftime("%Y-%m-%d") for d in dts]
d2idx = {k: i for i, k in enumerate(dkey)}
L = np.array([[float(v) for v in r[1:]] for r in r2a[1:]])
V = np.array([[float(v) for v in r[1:]] for r in r2b[1:]])
PR = np.array([[float(v) for v in r[1:]] for r in r4[1:]])
allrows = r3[1:]
F = np.array([[float(v) for v in r[2:]] for r in allrows])
# 附件3 的“日期”列只在每天第一行填写（合并单元格），向下填充
_filled = []
_cur = None
for r in allrows:
    d = r[0]
    if d not in (None, ""):
        _cur = str(d).strip()
    _filled.append((_cur, r[1]))
allrows = _filled
w("附件3 日期向下填充后 前5 = %s ; 末4 = %s" % (allrows[:5], allrows[-4:]))

# ---------- 1. 预报->实际 对齐的精细判定 ----------
w("### 1. 附件3 预报与附件2 实际的时间对齐（分预报时刻、分辨 '区间平均' 与 '整点瞬时'）")
w("约定说明：附件2 第 c 列(c=0..143) 的时间标签 = 0:10 + c*10min，最后为 0:00+1。")
w("  区间平均读法 A_m: 预报m小时 = 标签为 m:00 的 10 分钟点所属的那一小时块 [m-1,m)，列 (m-1)*6 .. m*6-1")
w("  区间平均读法 B_m: 预报m小时 = 小时块 [m,m+1)，列 m*6 .. m*6+5")
w("  整点瞬时读法 C_m: 预报m小时 = 标签为 m:00 的单个 10 分钟点，列 m*6-1")
w("  整点瞬时读法 D_m: 预报m小时 = 标签为 (m+1):00 的单个 10 分钟点，列 m*6+5")
issues = ["0:00", "6:00", "12:00", "18:00"]
for iss in issues:
    rows_iss = [(i, r) for i, r in enumerate(allrows) if r[1] == iss]
    # 该预报时刻对应的当日小时偏移
    h0 = int(iss.split(":")[0])
    stats = {"A": [], "B": [], "C": [], "D": []}
    D0 = dt.timedelta(days=1)
    for i, r in rows_iss:
        d = dt.datetime.strptime(r[0], "%Y-%m-%d")
        for m in range(1, 25):
            # absdt = 预报 m 小时所指的那个整点时刻
            absdt = d + dt.timedelta(hours=h0 + m)
            jm = d2idx.get(absdt.strftime("%Y-%m-%d"))
            jp = d2idx.get((absdt - D0).strftime("%Y-%m-%d"))
            jn = d2idx.get((absdt + D0).strftime("%Y-%m-%d"))
            f = F[i, m - 1]
            # 读法 A: 结束于 absdt 的一小时 (absdt-1h, absdt]
            j = jp if absdt.hour == 0 else jm
            if j is not None:
                c = 138 if absdt.hour == 0 else (absdt.hour - 1) * 6
                stats["A"].append(abs(f - V[j, c:c + 6].mean()))
            # 读法 B: 开始于 absdt 的一小时 [absdt, absdt+1h)
            if jm is not None:
                c = absdt.hour * 6
                stats["B"].append(abs(f - V[jm, c:c + 6].mean()))
            # 读法 C: 与 absdt 同一标签的单个 10 分钟采样点
            j = jp if absdt.hour == 0 else jm
            if j is not None:
                c = 143 if absdt.hour == 0 else absdt.hour * 6 - 1
                stats["C"].append(abs(f - V[j, c]))
            # 读法 D: absdt 之后 10 分钟的采样点
            j = jn if absdt.hour == 23 else jm
            if j is not None:
                c = 5 if absdt.hour == 23 else absdt.hour * 6 + 5
                stats["D"].append(abs(f - V[j, c]))
    w("  预报时刻 %s (样本 %d 条):" % (iss, len(rows_iss)))
    for k in ["A", "B", "C", "D"]:
        a = np.array(stats[k])
        w("     读法%s: MAE=%.1f  中位AE=%.1f  RMSE=%.1f  n=%d" % (k, a.mean(), np.median(a), np.sqrt((a ** 2).mean()), len(a)))

# ---------- 2. 附件1 是否为全年平均日 ----------
r1 = load_sheet(P("附件", "附件1.xlsx"))
price1 = np.array([float(r[1]) for r in r1[1:]])
load1 = np.array([float(r[2]) for r in r1[1:]])
pv1 = np.array([float(r[3]) for r in r1[1:]])
w()
w("### 2. 附件1 与 全年数据的均值关系")
w("附件1 电价 与 附件4 各列均值 的最大绝对差 = %.6f" % np.abs(PR.mean(0) - price1).max())
w("附件1 负载 与 附件2 负载各列均值 的最大绝对差 = %.6f" % np.abs(L.mean(0) - load1).max())
w("附件1 光伏 与 附件2 光伏各列均值 的最大绝对差 = %.6f" % np.abs(V.mean(0) - pv1).max())
w("附件4 电价 全样本均值=%.6f ; 附件1 电价均值=%.6f" % (PR.mean(), price1.mean()))
w("附件2 负载 全样本均值=%.6f ; 附件1 负载均值=%.6f" % (L.mean(), load1.mean()))
w("附件2 光伏 全样本均值=%.6f ; 附件1 光伏均值=%.6f" % (V.mean(), pv1.mean()))
# 附件3 的 0:00 预报小时均值 与 附件1 的小时均值
idx0 = [i for i, r in enumerate(allrows) if r[1] == "0:00"]
F0 = F[idx0]
w("附件3 0:00 预报(小时) 全年均值 = %s" % np.round(F0.mean(0), 2))
w("附件1 光伏 小时块均值(列 (h)*6..h*6+5, h=0..23) = %s" % np.round(pv1.reshape(24, 6).mean(1), 2))
w("附件1 光伏 小时块均值(列 (h-1)*6..h*6-1) = %s" % np.round(np.array([pv1[max(0, h * 6 - 6):h * 6].mean() for h in range(1, 25)]), 2))

# ---------- 3. 有效时间分辨率（小数位按列位置统计） ----------
w()
w("### 3. 有效时间分辨率：按列位置(0..143)统计“恰好2位小数”的出现次数")
def dec2count(mat):
    cnt = np.zeros(mat.shape[1], dtype=int)
    for c in range(mat.shape[1]):
        col = mat[:, c]
        n = 0
        for v in col:
            s = ("%.6f" % float(v)).rstrip("0")
            d = len(s.split(".")[1]) if "." in s else 0
            if d == 2:
                n += 1
        cnt[c] = n
    return cnt

for name, mat in [("附件1 电价(单日)", price1.reshape(1, -1)), ("附件1 负载(单日)", load1.reshape(1, -1)),
                  ("附件1 光伏(单日)", pv1.reshape(1, -1)), ("附件4 电价(365日)", PR),
                  ("附件2 负载(365日)", L), ("附件2 光伏(365日)", V)]:
    c2 = dec2count(mat)
    frac = c2 / mat.shape[0]
    w("%s: 每列恰2位小数的比例 前18列 = %s" % (name, np.round(frac[:18], 2)))
    w("    全144列 比例 = 均值%.2f 最小%.2f 最大%.2f ; 恰2位比例>0.9的列号 = %s"
      % (frac.mean(), frac.min(), frac.max(), np.where(frac > 0.9)[0][:60]))

# ---------- 4. 异常值 ----------
w()
w("### 4. 异常值排查")
w("附件4 电价 分位数: " + str(np.round(np.percentile(PR, [0, 0.1, 1, 5, 25, 50, 75, 95, 99, 99.9, 100]), 4)))
lo = PR < 0.05
w("附件4 电价 < 0.05 元/kWh 的点数 = %d ; 涉及日期 = %s" % (int(lo.sum()), sorted({dkey[i] for i in range(365) if lo[i].any()})))
for i in range(365):
    if lo[i].any():
        c = np.where(lo[i])[0]
        w("    %s 列号 %s 值 %s" % (dkey[i], list(c), list(np.round(PR[i, c], 4))))
w("附件4 电价 <=0 的点数 = %d" % int((PR <= 0).sum()))
w("附件2 负载 分位数: " + str(np.round(np.percentile(L, [0, 0.1, 1, 5, 25, 50, 75, 95, 99, 99.9, 100]), 1)))
w("附件2 光伏 分位数: " + str(np.round(np.percentile(V, [0, 1, 25, 50, 75, 95, 99, 99.9, 100]), 1)))
w("附件2 光伏 夜间的非零小值(<10kW) 个数 = %d ; 占全样本 %.3f%%" % (int(((V > 0) & (V < 10)).sum()), 100.0 * ((V > 0) & (V < 10)).sum() / V.size))
w("附件2 光伏 白天时段列(6:00~18:00)出现 0 的个数 = %d" % int((V[:, 35:107] == 0).sum()))
# 每日光伏为零的行（全阴天）
zero_days = [dkey[i] for i in range(365) if V[i].sum() < 1]
w("光伏日总量 < 1 kWh 的日期数 = %d ; 样例 = %s" % (len(zero_days), zero_days[:10]))

# ---------- 5. 结果模板逐条抄录 ----------
w()
w("### 5. 附件5 结果模板的精确规定")
for fn in ["result1.xlsx", "result2.xlsx", "result3.xlsx", "result4-2.xlsx", "result4-3.xlsx"]:
    wb = openpyxl.load_workbook(P("附件", "附件5", fn), read_only=True, data_only=True)
    w("--- %s : 工作表 = %s" % (fn, wb.sheetnames))
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        w("  [%s] 行数=%d 列数=%d" % (ws.title, len(rows), len(rows[0]) if rows else 0))
        hdr = rows[0]
        w("    表头第1个 = %r ; 表头个数 = %d" % (hdr[0] if hdr else None, len(hdr)))
        if ws.title == "计划购电量":
            w("    表头 前4 = %s" % (hdr[1:5],))
            w("    表头 后4 = %s" % (hdr[-4:],))
            w("    数据行列数 = %d ; 首行标签=%r 末行标签=%r" % (len(rows[0]), rows[1][0], rows[-1][0]))
            if len(rows[0]) > 5:
                w("    时间列总数 = %d" % (len(hdr) - 1))
                w("    第2列标签=%r 第3列标签=%r 倒数第3列=%r 倒数第2列=%r 倒数第1列=%r" % (hdr[1], hdr[2], hdr[-3], hdr[-2], hdr[-1]))
        elif ws.title in ("充放电量",):
            for r in rows[:8]:
                w("      %s" % (r,))
        elif ws.title == "紧急购电量":
            for r in rows[:3]:
                w("      %s" % (r,))
        elif ws.title == "调整购电量":
            w("    表头 前4 = %s ; 后4 = %s" % (hdr[1:5], hdr[-4:]))
    wb.close()

with open(P("_phase0", "报告03_对齐与异常.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done")
