"""构思手的交付物验收工具：读取编程手交付的 result1.xlsx 与逐时段 CSV，
与构思手自己的独立 LP/DP 参照解逐条对账。只读 问题1/ 下的文件，不写入。

用法：python _phase0/probe09_verify_deliverable.py
"""

import os
import csv
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


# ---------- 参照：附件1 与构思手自己的 LP ----------
r1 = load_sheet(P("附件", "附件1.xlsx"))
price = np.array([float(r[1]) for r in r1[1:]])
load = np.array([float(r[2]) for r in r1[1:]])
pv = np.array([float(r[3]) for r in r1[1:]])
K, DT, ETA = 144, 1.0 / 6.0, 0.9
U_MAX, E_MIN, E_MAX, E_INIT = 5000.0 * DT, 1200.0, 10800.0, 6000.0
REF_COST = 35126.948589
REF_QTY = 59482.698998
REF_TABLE1 = {10: 0.0000, 12: 480.4124, 14: 0.0000, 16: 445.4317, 18: 531.8940, 20: 0.0000}
REF_BLOCK = {"0:00-4:00": (4500.0000, 0.0000), "4:00-8:00": (833.3333, 6365.8412),
             "8:00-12:00": (4787.9643, 1702.9970), "12:00-16:00": (5286.0352, 91.1014),
             "16:00-20:00": (0.0000, 5780.1319), "20:00-24:00": (5333.3333, 2859.8681)}

res_path = P("问题1", "result1.xlsx")
csv_path = P("问题1", "问题1_逐时段结果.csv")

w("### 交付物验收对账（构思手独立执行）")
w("待检文件：")
w("   %s  存在=%s" % (res_path, os.path.exists(res_path)))
w("   %s  存在=%s" % (csv_path, os.path.exists(csv_path)))
if not os.path.exists(res_path):
    w("")
    w("**结论：result1.xlsx 尚未落盘，无法验收。**")
elif not os.path.exists(csv_path):
    # 退化为只检查 xlsx
    w("")
    w("CSV 未落盘，仅对 result1.xlsx 做格式与总量检查。")
    wb = openpyxl.load_workbook(res_path, read_only=True, data_only=True)
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        w("--- 工作表 [%s] 行=%d 列=%d" % (ws.title, len(rows), len(rows[0])))
        for i, r in enumerate(rows[:3]):
            w("    [%d] %s" % (i, r))
        for i in range(max(3, len(rows) - 2), len(rows)):
            w("    [%d] %s" % (i, rows[i]))
    wb.close()
else:
    # ---------- 1. 模板格式比对 ----------
    tmpl = load_sheet(P("附件", "附件5", "result1.xlsx"))
    from openpyxl import load_workbook
    wb_t = load_workbook(P("附件", "附件5", "result1.xlsx"), read_only=True, data_only=True)
    wb_d = load_workbook(res_path, read_only=True, data_only=True)
    w("--- 1. 模板结构比对 ---")
    w("   模板工作表 = %s ; 交付工作表 = %s" % (wb_t.sheetnames, wb_d.sheetnames))
    w("   工作表名一致 = %s" % (wb_t.sheetnames == wb_d.sheetnames))
    for sn in wb_t.sheetnames:
        rt = [list(r) for r in wb_t[sn].iter_rows(values_only=True)]
        rd = [list(r) for r in wb_d[sn].iter_rows(values_only=True)]
        w("   [%s] 行数 模板=%d 交付=%d ; 列数 模板=%d 交付=%d" % (sn, len(rt), len(rd), len(rt[0]), len(rd[0])))
        same_labels = all(
            (rt[i][0] == rd[i][0]) and all((rt[0][j] == rd[0][j]) for j in range(len(rt[0])))
            for i in range(len(rt)))
        w("        行标签与表头逐格一致 = %s" % same_labels)
    wb_t.close(); wb_d.close()

    # ---------- 2. 逐时段 CSV ----------
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        rd_ = list(csv.DictReader(f))
    w("")
    w("--- 2. 逐时段结果 CSV ---")
    w("   行数 = %d ; 列名 = %s" % (len(rd_), list(rd_[0].keys())))
    x = np.array([float(r["计划购电量_kWh"]) for r in rd_])
    u = np.array([float(r["充电量_kWh"]) for r in rd_])
    v = np.array([float(r["放电量_kWh"]) for r in rd_])
    E = np.array([float(r["储电量_kWh"]) for r in rd_])
    # 注意：交付 CSV 的“储电量_kWh”列定义为【时段末】，E[0] 是第 1 时段末，不是 0:00 的值
    p_csv = np.array([float(r["电价_元每kWh"]) for r in rd_])
    L_csv = np.array([float(r["负载_kW"]) for r in rd_])
    Pv_csv = np.array([float(r["光伏_kW"]) for r in rd_])

    # 交付文件统一保留 4 位小数：单个表达式最多叠加 3 个舍入项，误差上限约 3e-4，
    # 因此逐点核验容差取 1e-3（若取 1e-6 会把正常的十进制舍入误判成违规）
    TOL = 1e-3

    # 0:00 / 24:00 储电量以交付 xlsx 的 `充放电量` 表为准（CSV 只含 144 个时段末值）
    wb_tmp = load_workbook(res_path, read_only=True, data_only=True)
    blk = [list(r) for r in wb_tmp["充放电量"].iter_rows(values_only=True)]
    wb_tmp.close()
    soc_0000 = float(blk[1][4])
    soc_2400 = float(blk[2][4])

    w("   %-40s %-18s %-18s %s" % ("检查项", "交付实测", "构思手参照", "判定"))
    def row(name, got, ref, tol=TOL):
        ok = abs(got - ref) <= tol
        w("   %-40s %-18.4f %-18.4f %s" % (name, got, ref, "通过" if ok else "**不符**"))
        return ok

    row("全天购电量 Σx (kWh)", x.sum(), REF_QTY, 0.01)
    row("全天购电费 Σp·x (元)", float((p_csv * x).sum()), REF_COST, 0.01)
    row("充电量合计 Σu (kWh)", u.sum(), 20740.6661, 0.01)
    row("放电量合计 Σv (kWh)", v.sum(), 16799.9396, 0.01)
    row("放电/充电 比值", v.sum() / u.sum(), 0.8100, 1e-6)
    row("0:00 储电量 (kWh) [xlsx 充放电量]", soc_0000, 6000.0, TOL)
    row("24:00 储电量 (kWh) [xlsx 充放电量]", soc_2400, 6000.0, TOL)
    row("CSV 首行 = 第1时段末储电量 (kWh)", E[0], 6750.0, TOL)
    row("CSV 末行 = 第144时段末储电量 (kWh)", E[-1], 6000.0, TOL)
    row("储电量轨迹最小值 (kWh)", E.min(), 1200.0, TOL)
    row("储电量轨迹最大值 (kWh)", E.max(), 10800.0, TOL)
    w("   %-40s %-18d %-18d %s" % ("购电量为 0 的时段数", int((x < 1e-6).sum()), 58,
                                  "通过" if int((x < 1e-6).sum()) == 58 else "**不符**"))
    w("   %-40s %-18d %-18d %s" % ("同时充放时段数", int(((u > 1e-6) & (v > 1e-6)).sum()), 0,
                                  "通过" if ((u > 1e-6) & (v > 1e-6)).sum() == 0 else "**不符**"))

    # ---------- 3. 逐点约束核验（用交付的 x,u,v 独立重算） ----------
    w("")
    w("--- 3. 逐点约束核验（用交付的 x,u,v 重算，不信任其自检） ---")
    w("   容差说明：交付值保留 4 位小数，单个表达式最多叠加 3 个舍入项，误差上限约 3e-4，故取容差 %.0e" % TOL)
    lhs = x + Pv_csv * DT + v - L_csv * DT - u
    w("   供给约束 min(x+Pv·Δ+v-L·Δ-u) = %.6f （要求 >= -%.0e）%s"
      % (lhs.min(), TOL, "通过" if lhs.min() >= -TOL else "**不符**"))
    w("   max|u| 越界量 = %.6f ; max|v| 越界量 = %.6f"
      % (max(0.0, u.max() - U_MAX), max(0.0, v.max() - U_MAX)))
    chain = np.concatenate([[E_INIT], E])
    rec = ETA * u - v / ETA
    w("   状态递推 max|ΔE-(ηu-v/η)| = %.6f （容差 %.0e）%s"
      % (np.abs(np.diff(chain) - rec).max(), TOL,
         "通过" if np.abs(np.diff(chain) - rec).max() < TOL else "**不符**"))
    chk = E_INIT + np.concatenate([[0.0], np.cumsum(rec)])
    w("   储电量链式重算 max|E-重算| = %.6f （容差 %.0e）%s"
      % (np.abs(chk - chain).max(), TOL,
         "通过" if np.abs(chk - chain).max() < TOL else "**不符**"))
    dE = np.diff(np.concatenate([[E_INIT], E]))
    w("   CSV 电价/负载/光伏与附件1 一致性: max|Δ|=%.6f / %.6f / %.6f"
      % (np.abs(p_csv - price).max(), np.abs(L_csv - load).max(), np.abs(Pv_csv - pv).max()))
    w("   CSV 是否按真实时间顺序（时段序号 1..144 递增）: %s"
      % ([int(r["时段序号"]) for r in rd_][:3] == [1, 2, 3] and
         [int(r["时段序号"]) for r in rd_][-3:] == [142, 143, 144]))

    # ---------- 4. 表 1 / 表 2 与 xlsx 对账 ----------
    w("")
    w("--- 4. 表 1 六个时段（H:00-H:10 ↔ CSV 第 6H+1 个时段 ↔ 模板第 6H 行） ---")
    wb_d = load_workbook(res_path, read_only=True, data_only=True)
    trows = [list(r) for r in wb_d["计划购电量"].iter_rows(values_only=True)]
    vals = [float(r[1]) for r in trows[1:] if r[1] is not None]
    w("   模板 144 行取值个数 = %d ; 求和 = %.4f kWh" % (len(vals), sum(vals)))
    for hh in (10, 12, 14, 16, 18, 20):
        csv_val = x[6 * hh]
        xlsx_val = float(trows[6 * hh][1])
        ref = REF_TABLE1[hh]
        ok = abs(csv_val - ref) <= 0.001 and abs(xlsx_val - ref) <= 0.001
        w("   %02d:00-%02d:10  CSV=%.4f  模板第%d行=%.4f  参照=%.4f  %s"
          % (hh, hh, csv_val, 6 * hh, xlsx_val, ref, "通过" if ok else "**不符**"))

    w("")
    w("--- 5. 表 2 六个 4 小时块（真实时段 1-144 顺序聚合） ---")
    brows = [list(r) for r in wb_d["充放电量"].iter_rows(values_only=True)]
    names = ["0:00-4:00", "4:00-8:00", "8:00-12:00", "12:00-16:00", "16:00-20:00", "20:00-24:00"]
    for i, nm in enumerate(names):
        s = i * 24
        got_c = u[s:s + 24].sum(); got_d = v[s:s + 24].sum()
        xl_c = brows[i + 1][1]; xl_d = brows[i + 1][2]
        rc, rd_ = REF_BLOCK[nm]
        w("   %-12s CSV(充/放)=%.4f/%.4f  xlsx=%.4f/%.4f  参照=%.4f/%.4f  %s"
          % (nm, got_c, got_d, float(xl_c), float(xl_d), rc, rd_,
             "通过" if (abs(got_c - rc) < 0.01 and abs(got_d - rd_) < 0.01 and
                        abs(float(xl_c) - rc) < 0.01 and abs(float(xl_d) - rd_) < 0.01) else "**不符**"))
    w("   充放电量表 时刻列 = %r" % [brows[1][3], brows[2][3]])
    w("   充放电量表 储电量列 = %r" % [brows[1][4], brows[2][4]])
    wb_d.close()

    # ---------- 6. 与构思手参照解逐点比较 ----------
    w("")
    w("--- 6. 与构思手参照解逐点比较（参照解由 probe08 独立 LP 生成） ---")
    ref_path = P("_phase0", "参照解_问题1.csv")
    if os.path.exists(ref_path):
        with open(ref_path, "r", encoding="utf-8-sig", newline="") as f:
            rr = list(csv.DictReader(f))
        xr = np.array([float(r["计划购电量_kWh"]) for r in rr])
        ur = np.array([float(r["充电量_kWh"]) for r in rr])
        vr = np.array([float(r["放电量_kWh"]) for r in rr])
        Er = np.array([float(r["储电量_kWh"]) for r in rr])
        w("   max|x-x_ref| = %.6f kWh ; max|u-u_ref| = %.6f ; max|v-v_ref| = %.6f ; max|E-E_ref| = %.6f"
          % (np.abs(x - xr).max(), np.abs(u - ur).max(), np.abs(v - vr).max(), np.abs(E - Er).max()))
        w("   费用差 = %.6f 元（%.2e 相对）" % (abs(float((p_csv * x).sum()) - REF_COST),
                                              abs(float((p_csv * x).sum()) - REF_COST) / REF_COST))
        w("   说明：LP 最优解可能不唯一（价格分段常数导致退化），因此逐点差可以不严格为 0；")
        w("         判定以费用与全部总量指标为准。若逐点差大于 1 kWh，需编程手说明是否有另一组最优解。")
    else:
        w("   参照解文件缺失（先运行 probe08_independent_dp.py）")

w("")
w("### 验收结论")
w("以上任一项标 **不符** 即视为未通过，需编程手给出差异来源或重做。")

with open(P("_phase0", "报告14_交付物验收对账.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done")
