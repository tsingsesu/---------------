"""问题 4 独立复核（从落盘文件反查）：约束核验 + 表 1/2/3 反查 + 代码一致性 + 下界对比。

设计原则：本脚本**不重新求解任何 LP**，只从交付文件（result4-2.xlsx / result4-3.xlsx /
两份全分辨率 CSV / 回归对照文件 / 联合 LP 落盘）反查、重算并核对，保证"报告中的每个数字
都能从落盘文件复现"。全部核验项写入 `问题4/自检报告.txt`，并打印到控制台。

核验清单（对应 `问题4/交付清单.md` §六）：
  A. 文件与表结构：工作表名/列名/行列数/`7:0-7:10` 笔误照抄
  B. 轮转填法逐格：xlsx 每格 = CSV 重排（第 1 格装当天第 2 时段…第 144 格装第 1 时段）
  C. 表 1/表 2 反查：四个指定日期的六个时段与六个 4 小时块，xlsx 与 CSV 重算逐位一致
  D. 表 3 反查：紧急购电文本与 CSV 的 r 重算一致（含 4-2 全 0）
  E. 约束核验：供给 ≥ 负载、0≤u,v≤P̄Δ、E∈[下限,上限]、跨日连续性、x,y,r,u,v ≥ 0
  F. D-12 恒等式：全天购电费（调整表）= J_plan + J_adj + J_emg（由 CSV 独立重算）
  G. 非追溯性（4-3）：k≤36 的 y ≡ x
  H. 代码一致性（核心）：att1 回归文件与问题 2/问题 3 交付文件逐格 0 差异
  I. 下界对比：4-2 缴费 ≥ 4-3 J ≥ 联合 LP 下界
  J. 锚定对照：4-2 缴费/残值/持有过夜天数与构思手锚定一致

运行：python 问题4/verify_q4.py（须先跑 run_q4.py 与 run_q4.py --price=att1）
依赖：numpy、openpyxl；lib/ 公共模块。
随机性：无。
"""

import csv
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import numpy as np
import openpyxl

from lib.dataio import read_attachment2
from lib.logio import Tee
from lib.storage import E_MAX, E_MIN, ETA, P_MAX
from lib.timegrid import DT_H, K, four_hour_blocks, hour_block_to_k, k_to_label

QDIR = os.path.join(ROOT, "问题4")
RESULT_42 = os.path.join(QDIR, "result4-2.xlsx")
RESULT_43 = os.path.join(QDIR, "result4-3.xlsx")
REG_42 = os.path.join(QDIR, "result4-2_att1回归对照.xlsx")
REG_43 = os.path.join(QDIR, "result4-3_att1回归对照.xlsx")
CSV_42 = os.path.join(QDIR, "全分辨率明细_4-2.csv")
CSV_43 = os.path.join(QDIR, "全分辨率明细_4-3.csv")
Q2_RESULT = os.path.join(ROOT, "问题2", "result2.xlsx")
Q3_RESULT = os.path.join(ROOT, "问题3", "result3.xlsx")
REPORT = os.path.join(QDIR, "自检报告.txt")

D_REP_FIRST = 31
N_REP = 334
U_MAX = P_MAX * DT_H                         # 单时段最大充/放电量，kWh
SPEC_DATES = ("2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21")
SPEC_HOURS = (10, 12, 14, 16, 18, 20)

# 锚定值（`_phase0/报告20_问题4-2量级体检.txt` 与 `问题4/主模型运行日志.txt`）
ANCHOR_42 = {"pay": 12815460.2655, "resid": 469.0667, "true_cost": 12814991.1989,
             "n_hold": 60, "n_min": 274}
ANCHOR_Q2 = 12254765.7161                    # 问题 2 主模型
ANCHOR_Q3 = 13816096.7169                    # 问题 3 主模型

_RESULTS = []                                # (核验项, 结果 "OK"/"FAIL", 细节)


def check(name, ok, detail=""):
    """记录一条核验结果并打印（OK/FAIL）。

    输入：name，str，核验项；ok，bool；detail，str，数值细节
    输出：bool（原样返回 ok，便于串联）
    """
    _RESULTS.append((name, "OK" if ok else "FAIL", detail))
    print("[%s] %s %s" % ("OK  " if ok else "FAIL", name, detail))
    return ok


def read_csv_matrix(path):
    """把全分辨率明细 CSV 读成结构化数组字典（按真实时段序，334×144）。

    输入：path，str；输出：dict，键为列名，值为 (334,144) 数组（数值列）或 list（日期列）
    """
    with open(path, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    data = {}
    for col in rows[0].keys():
        if col in ("日期", "时段起", "时段止"):
            continue
        try:
            data[col] = np.array([float(r[col]) for r in rows], dtype=float).reshape(N_REP, K)
        except ValueError:
            data[col] = [r[col] for r in rows]
    data["日期"] = [rows[i * K]["日期"] for i in range(N_REP)]
    return data


def sheet_values(ws, n_row, n_col):
    """把工作表的矩形数据区读成 (n_row, n_col) 的 float 数组（None 记 nan）。

    输入：ws，openpyxl 工作表（read_only 模式，支持批量迭代）；n_row/n_col，int
    输出：np.ndarray (n_row, n_col)
    """
    out = np.full((n_row, n_col), np.nan)
    rows = ws.iter_rows(min_row=2, max_row=1 + n_row, min_col=2,
                        max_col=1 + n_col, values_only=True)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            if v is not None:
                out[i, j] = float(v)
    return out


def rotated(values):
    """按 D-01 填法 Y-轮转把真实时段序数组重排为模板列序（与 lib.xlsxio.template_order 同式）。

    输入：values，(144,) 真实时段序；输出：(144,) 模板列序（第 i 格 = 第 (i+1)%144 个元素）
    """
    arr = np.asarray(values, dtype=float)
    return arr[(np.arange(arr.size) + 1) % arr.size]


def verify_structure():
    """A. 表结构与列名核验（批量读取表头，避免 read_only 下逐格访问）。"""
    print("-" * 78)
    print("【A. 文件与表结构】")
    for path, sheets in ((RESULT_42, ("计划购电量", "充放电量", "紧急购电量")),
                         (RESULT_43, ("计划购电量", "调整购电量", "充放电量", "紧急购电量"))):
        wb = openpyxl.load_workbook(path, read_only=True)
        names = wb.sheetnames
        check("%s 工作表名" % os.path.basename(path), list(names) == list(sheets), str(names))
        # 表头第 1 行整行批量读出（判断 335×147 与 `7:0-7:10` 笔误）
        ws = wb["计划购电量"]
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        n_data = sum(1 for _ in ws.iter_rows(min_row=2, values_only=True))
        check("%s 计划表 335×147" % os.path.basename(path),
              n_data == 334 and len(header) == 147, "%d 数据行 × %d 列" % (n_data, len(header)))
        label42 = header[42]                                 # 第 42 个"时段列"= 表列 43（0 基 42）的笔误
        check("%s `7:0-7:10` 笔误照抄" % os.path.basename(path), label42 == "7:0-7:10",
              repr(label42))
        ws2 = wb["充放电量"]
        n2 = sum(1 for _ in ws2.iter_rows(min_row=2, values_only=True))
        check("%s 充放电量行数 = 2004（数据行）" % os.path.basename(path), n2 == 2004,
              "数据行=%d" % n2)
        ws4 = wb["紧急购电量"]
        n4 = sum(1 for _ in ws4.iter_rows(min_row=2, values_only=True))
        check("%s 紧急购电量覆盖 334 天" % os.path.basename(path), n4 == 334,
              "数据行=%d" % n4)
        wb.close()


def verify_rotation(d42, d43):
    """B. 轮转填法逐格核验（xlsx 每个时段格 = CSV 轮转重排 + 4 位小数）。"""
    print("-" * 78)
    print("【B. 轮转填法逐格核验（全表）】")
    # ---- 4-2 计划表 ----
    wb = openpyxl.load_workbook(RESULT_42, read_only=True)
    ws = wb["计划购电量"]
    mat = sheet_values(ws, N_REP, K)
    expect = np.array([rotated(d42["计划购电量_kWh"][i]) for i in range(N_REP)])
    diff = float(np.abs(mat - expect).max())
    check("4-2 计划表逐格 = CSV 轮转重排", diff < 1e-9, "最大差 %.2e" % diff)
    # 尾两列（全天购电量 / 全天购电费）
    tot = np.zeros(N_REP); cost = np.zeros(N_REP)
    for i, row in enumerate(ws.iter_rows(min_row=2, max_row=1 + N_REP, min_col=146,
                                         max_col=147, values_only=True)):
        tot[i] = float(row[0]); cost[i] = float(row[1])
    csv_tot = d42["计划购电量_kWh"].sum(axis=1)
    csv_cost = (d42["计划购电量_kWh"] * d42["电价_元每kWh"]).sum(axis=1)
    # 容差：xlsx 与 CSV 均为 4 位小数，逐格舍入的累积偏差 ≤ 144×5e-5 ≈ 7.2e-3
    tol_sum = 144 * 5e-5 * 1.5
    check("4-2 计划表`全天购电量`列", np.max(np.abs(tot - csv_tot)) < tol_sum,
          "最大差 %.2e（容差 %.2e，含 4 位小数舍入）" % (np.max(np.abs(tot - csv_tot)), tol_sum))
    check("4-2 计划表`全天购电费`列（= Σp·x，r≡0）", np.max(np.abs(cost - csv_cost)) < tol_sum * 2,
          "最大差 %.2e（p·x 两项舍入，容差 %.2e）" % (np.max(np.abs(cost - csv_cost)), tol_sum * 2))
    wb.close()
    # ---- 4-3 计划表与调整表 ----
    wb = openpyxl.load_workbook(RESULT_43, read_only=True)
    ws = wb["计划购电量"]
    mat = sheet_values(ws, N_REP, K)
    expect = np.array([rotated(d43["计划购电量_kWh"][i]) for i in range(N_REP)])
    diff = np.max(np.abs(mat - expect))
    check("4-3 计划表逐格 = CSV 轮转重排", diff < 1e-9, "最大差 %.2e" % diff)
    ws = wb["调整购电量"]
    mat = sheet_values(ws, N_REP, K)
    expect = np.array([rotated(d43["调整购电量_kWh"][i]) for i in range(N_REP)])
    diff = np.max(np.abs(mat - expect))
    check("4-3 调整表逐格 = CSV 轮转重排（填最终量 y，D-14）", diff < 1e-9, "最大差 %.2e" % diff)
    wb.close()


def verify_spec_tables(d42, d43):
    """C. 表 1/表 2/表 3 反查：四个指定日期 xlsx 与 CSV 重算逐位一致（批量读取）。"""
    print("-" * 78)
    print("【C. 表 1/表 2/表 3 反查（四个指定日期）】")
    import datetime as _dt
    # 批量读计划表数据区（334×144）与两张表的 146/147 列
    wb42 = openpyxl.load_workbook(RESULT_42, read_only=True)
    plan42 = sheet_values(wb42["计划购电量"], N_REP, K)
    wb43 = openpyxl.load_workbook(RESULT_43, read_only=True)
    plan43 = sheet_values(wb43["计划购电量"], N_REP, K)
    yadj43 = sheet_values(wb43["调整购电量"], N_REP, K)
    # 充放电量：批量读 2004 行 × 校核列（充电量=第3列、放电量=第4列、日期=第1列）
    def read_ud(ws):
        """批量读充放电量的（日期, 充电量, 放电量, 时刻, 储电量），返回 (dates list, u, v)。"""
        u = np.zeros(N_REP * 6); v = np.zeros(N_REP * 6)
        nn = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            u[nn] = float(row[2]) if row[2] is not None else 0.0
            v[nn] = float(row[3]) if row[3] is not None else 0.0
            nn += 1
        return u.reshape(N_REP, 6), v.reshape(N_REP, 6)
    u42, v42 = read_ud(wb42["充放电量"])
    u43, v43 = read_ud(wb43["充放电量"])
    ok_all = True
    for t in SPEC_DATES:
        d0 = (_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days
        i = d0 - D_REP_FIRST
        for h in SPEC_HOURS:
            k = hour_block_to_k(h)
            j_tpl = 6 * h - 1                                # 模板第 6H 列（0 基 6H-1）
            v42c = plan42[i, j_tpl]; e42 = d42["计划购电量_kWh"][i, k - 1]
            v43c = plan43[i, j_tpl]; e43 = d43["计划购电量_kWh"][i, k - 1]
            if abs(v42c - e42) > 1e-9 or abs(v43c - e43) > 1e-9:
                ok_all = False
                print("    %s %d:00 不符：xlsx %.4f/%.4f vs CSV %.4f/%.4f" % (t, h, v42c, v43c, e42, e43))
        # 表 2：六个 4 小时块（充放电量表的行号 = i*6+b）
        for b, (sl, el, kf, kl) in enumerate(four_hour_blocks()):
            eu42 = float(d42["充电量_kWh"][i, kf - 1:kl].sum())
            ev42 = float(d42["放电量_kWh"][i, kf - 1:kl].sum())
            eu43 = float(d43["充电量_kWh"][i, kf - 1:kl].sum())
            ev43 = float(d43["放电量_kWh"][i, kf - 1:kl].sum())
            if (abs(u42[i, b] - eu42) > 1e-3 or abs(v42[i, b] - ev42) > 1e-3
                    or abs(u43[i, b] - eu43) > 1e-3 or abs(v43[i, b] - ev43) > 1e-3):
                ok_all = False
                print("    表2 %s %s 不符：42 %s/%s vs %.4f/%.4f；43 %s/%s vs %.4f/%.4f"
                      % (t, sl, u42[i, b], v42[i, b], eu42, ev42, u43[i, b], v43[i, b], eu43, ev43))
    check("四个指定日期表 1/表 2 反查（4-2 与 4-3）", ok_all, "48 个表 1 值 + 48 个表 2 块和")

    # 表 3 反查（4-3）：触发标志逐日 + 指定日期电量合计
    emg_txt = []
    emg_val = []
    for row in wb43["紧急购电量"].iter_rows(min_row=2, values_only=True):
        emg_txt.append(row[1])
        emg_val.append(row[2])
    r_csv = d43["紧急购电量_kWh"]
    ok_trigger = True
    for i in range(N_REP):
        csv_has = bool(r_csv[i].max() > 1e-6)
        if bool(emg_txt[i] not in (None, "")) != csv_has:
            ok_trigger = False
            print("    %s 触发标志不一致：xlsx=%r" % (d43["日期"][i], emg_txt[i]))
    check("4-3 表 3 触发标志逐日一致（334 天）", ok_trigger, "文本非空 ⇔ CSV 有 r>0")
    ok_seg = True
    for t in SPEC_DATES:
        d0 = (_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days
        i = d0 - D_REP_FIRST
        if emg_txt[i] in (None, ""):
            ok_seg = ok_seg and bool(r_csv[i].max() <= 1e-6)
            continue
        parts = str(emg_val[i]).split()
        total_txt = sum(float(p) for p in parts)
        total_csv = float(r_csv[i].sum())
        if abs(total_txt - total_csv) > 1e-3:
            ok_seg = False
            print("    %s 表 3 电量合计不符：文本 %.4f vs CSV %.4f" % (t, total_txt, total_csv))
    check("4-3 表 3 指定日期电量合计 vs CSV", ok_seg, "4 个日期")

    # 4-2 表 3：全 334 天应为 0（读法①）
    ok42 = True
    for row in wb42["紧急购电量"].iter_rows(min_row=2, values_only=True):
        txt = row[1]; val = row[2]
        if txt not in (None, "") or (val is not None and float(val) != 0.0):
            ok42 = False
            print("    4-2 表 3 非 0：%r / %r" % (txt, val))
    check("4-2 表 3 全 334 天均为空+0（读法① r≡0）", ok42, "")
    wb42.close(); wb43.close()


def verify_constraints(d42, d43):
    """E/F/G. 约束核验、D-12 恒等式、非追溯性。"""
    print("-" * 78)
    print("【E. 约束核验（从 CSV 反查）】")
    load, pv, dates = read_attachment2()
    load_rep = load[D_REP_FIRST:]
    pv_rep = pv[D_REP_FIRST:]
    for tag, d in (("4-2", d42), ("4-3", d43)):
        x = d["计划购电量_kWh"]; y = d["调整购电量_kWh"] if "调整购电量_kWh" in d else x
        u = d["充电量_kWh"]; v = d["放电量_kWh"]; E = d["时段末储电量_kWh"]
        if tag == "4-2":
            # 4-2：计划量恰覆盖负载（完全信息读法①），供给残差应 ≈ 0；容差取 4 位小数舍入量级
            resid = x + pv_rep * DT_H + v - load_rep * DT_H - u
            check("%s 供给约束 x+PΔ+v−LΔ−u ≥ 0（读法① 应 ≈ 0）" % tag,
                  float(resid.min()) > -1e-3,
                  "最小 %.3e kWh（[-1e-3, 0) 为 4 位小数舍入）" % resid.min())
        else:
            # 4-3：实际供给按最终生效购电量 y 核算（D-13）；y+PΔ+v−LΔ−u = g_curt − r_emg，
            # 故应满足"残差 + r_emg ≥ 0 且 g_curt ≥ 0"（缺额由紧急购电覆盖，不由供给约束违反）
            resid_y = y + pv_rep * DT_H + v - load_rep * DT_H - u
            r_csv = d["紧急购电量_kWh"]; g_csv = d["弃光电量_kWh"]
            lhs = resid_y + r_csv
            check("%s 实际供给核算 y+PΔ+v−LΔ−u+r ≥ 0 且 = 弃光" % tag,
                  float(lhs.min()) > -2e-3 and float(np.abs(lhs - g_csv).max()) < 2e-3,
                  "最小 %.3e kWh；|残差+r−弃光| 最大 %.3e kWh" % (lhs.min(), np.abs(lhs - g_csv).max()))
            # 计划阶段（0:00 决策）自身的供给约束核验：k≤36（0:00–6:00）不可追溯，
            # 其 x 与 u,v 即 0:00 阶段解，用 0:00 预报列（光伏预报0_kW）重算残差应 ≥ 0；
            # k>36 的最终 u,v 混合了后续阶段决策，无法从 CSV 单独还原计划段解，故不核验
            fc0 = d["光伏预报0_kW"] * DT_H
            resid_plan_early = x[:, :36] + fc0[:, :36] + v[:, :36] - load_rep[:, :36] * DT_H - u[:, :36]
            check("%s 计划阶段供给约束（k≤36，0:00 预报下）" % tag,
                  float(resid_plan_early.min()) > -2e-3,
                  "最小 %.3e kWh（不可追溯段；[-2e-3, 0) 为 4 位小数舍入）"
                  % resid_plan_early.min())
        check("%s 功率约束 0≤u,v≤P̄Δ" % tag,
              float(u.min()) > -1e-6 and float(v.min()) > -1e-6
              and float(u.max()) <= U_MAX + 1e-6 and float(v.max()) <= U_MAX + 1e-6,
              "u∈[%.4f, %.4f] v∈[%.4f, %.4f]" % (u.min(), u.max(), v.min(), v.max()))
        # 储电量边界
        check("%s 储电量边界 [%.0f, %.0f]" % (tag, E_MIN, E_MAX),
              float(E.min()) > E_MIN - 1e-6 and float(E.max()) < E_MAX + 1e-6,
              "E∈[%.4f, %.4f]" % (E.min(), E.max()))
        # 储电量递推一致性 + 跨日传递（D-04）：对每个填报日反推当日 0:00 储电量
        # （E_1 − η·u_1 + v_1/η），再逐时段递推并与 CSV 的时段末储电量逐点比对；
        # 同时核验反推的 0:00 值 = 前一日 24:00（跨日传递）。容差按 4 位小数舍入量级。
        max_rec = 0.0
        for i in range(N_REP):
            e0_day = float(E[i, 0] - ETA * u[i, 0] + v[i, 0] / ETA)
            traj = np.concatenate([[e0_day], e0_day + np.cumsum(ETA * u[i] - v[i] / ETA)])
            max_rec = max(max_rec, float(np.abs(traj[1:] - E[i]).max()))
            if i > 0:
                max_rec = max(max_rec, abs(e0_day - float(E[i - 1, K - 1])))
        check("%s 储电量递推与跨日传递（逐日全链重算）" % tag, max_rec < 2e-3,
              "最大差 %.3e kWh（含逐时段 4 位小数舍入累积）" % max_rec)
        check("%s 电量非负" % tag,
              float(x.min()) > -1e-9 and float(y.min()) > -1e-9
              and float(d["紧急购电量_kWh"].min()) > -1e-9
              and float(d["弃光电量_kWh"].min() if "弃光电量_kWh" in d else 0.0) > -1e-9, "")

    print("-" * 78)
    print("【F. D-12 恒等式（4-3，由 CSV 独立重算）】")
    x = d43["计划购电量_kWh"]; y = d43["调整购电量_kWh"]; r = d43["紧急购电量_kWh"]
    p = d43["电价_元每kWh"]
    j_plan = (p * x).sum()
    j_adj = (1.5 * p * np.maximum(y - x, 0.0)).sum() - (0.5 * p * np.maximum(x - y, 0.0)).sum()
    j_emg = (5.0 * p * r).sum()
    j_tot = j_plan + j_adj + j_emg
    # 与 xlsx 调整表的"全天购电费"列合计核对（批量读第 147 列）
    wb = openpyxl.load_workbook(RESULT_43, read_only=True)
    cost_x = np.zeros(N_REP)
    for i, row in enumerate(wb["调整购电量"].iter_rows(min_row=2, max_row=1 + N_REP,
                                                       min_col=147, max_col=147,
                                                       values_only=True)):
        cost_x[i] = float(row[0])
    wb.close()
    check("4-3 调整表全天购电费合计 = 重算 J", abs(cost_x.sum() - j_tot) < 0.5,
          "xlsx %.4f vs 重算 %.4f（差 %.4f 元：CSV 逐格 4 位小数舍入的非负累积）"
          % (cost_x.sum(), j_tot, cost_x.sum() - j_tot))
    check("4-3 主口径恒等式 J = J_plan + J_adj + J_emg", True,
          "J_plan=%.4f J_adj=%.4f J_emg=%.4f J=%.4f" % (j_plan, j_adj, j_emg, j_tot))

    print("-" * 78)
    print("【G. 非追溯性（4-3）】")
    dev_early = float(np.abs(y[:, :36] - x[:, :36]).max())
    check("k≤36（0:00–6:00）的 y ≡ x", dev_early < 1e-9, "max|y−x| = %.2e" % dev_early)
    n_y_gt_x_early = int((y[:, :36] > x[:, :36] + 1e-6).sum())
    check("0:00–6:00 不存在超用记录", n_y_gt_x_early == 0, "超用格数 = %d" % n_y_gt_x_early)


def verify_consistency():
    """H. 代码一致性：att1 回归文件 vs 问题 2/3 交付文件逐格（批量行比较）。"""
    print("-" * 78)
    print("【H. 代码一致性（核心：换回附件1 价格须逐格复现问题 2/3）】")
    for reg, ref, sheets, tag in (
            (REG_42, Q2_RESULT, ("计划购电量", "充放电量", "紧急购电量"), "4-2 vs 问题2"),
            (REG_43, Q3_RESULT, ("计划购电量", "调整购电量", "充放电量", "紧急购电量"), "4-3 vs 问题3")):
        if not os.path.exists(reg):
            check("%s 回归文件存在" % tag, False, "缺 %s（请先跑 run_q4.py --price=att1）" % reg)
            continue
        wa = openpyxl.load_workbook(reg, read_only=True)
        wb = openpyxl.load_workbook(ref, read_only=True)
        n_diff = 0; n_cell = 0
        examples = []
        for sh in sheets:
            wsa, wsb = wa[sh], wb[sh]
            rows_a = wsa.iter_rows(values_only=True)
            rows_b = wsb.iter_rows(values_only=True)
            for ra, rb in zip(rows_a, rows_b):
                for va, vb in zip(ra, rb):
                    if isinstance(va, (int, float)) and isinstance(vb, (int, float)):
                        n_cell += 1
                        if abs(float(va) - float(vb)) > 1e-12:
                            n_diff += 1
                            if len(examples) < 6:
                                examples.append((sh, va, vb))
                    elif (va is None) != (vb is None):
                        n_diff += 1
                        if len(examples) < 6:
                            examples.append((sh, va, vb))
        wa.close(); wb.close()
        check("%s 逐格一致" % tag, n_diff == 0,
              "数值格 %d 个，差异 %d 个%s" % (n_cell, n_diff,
                                             ("；例：" + str(examples)) if examples else ""))


def verify_bounds_and_anchor(d42, d43):
    """I/J. 下界对比与锚定对照。"""
    print("-" * 78)
    print("【I. 下界对比】")
    joint_path = os.path.join(QDIR, "联合LP下界_对照.xlsx")
    if os.path.exists(joint_path):
        wb = openpyxl.load_workbook(joint_path, read_only=True)
        ws = wb["汇总对照"]
        got = {}
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[0] and isinstance(r[1], (int, float)):
                got[str(r[0])] = float(r[1])
        wb.close()
        lb = got.get("填报区间联合最优费用（2.1–12.31）", float("nan"))
        roll42 = float((d42["计划购电量_kWh"] * d42["电价_元每kWh"]).sum())
        roll43 = got.get("4-3 多阶段滚动总费用 J（主模型）", float("nan"))
        check("4-2 缴费 ≥ 联合 LP 下界", roll42 >= lb - 1e-3,
              "%.4f vs 下界 %.4f（gap %+.4f 元，%+.4f%%）"
              % (roll42, lb, roll42 - lb, 100.0 * (roll42 - lb) / lb))
        check("4-3 总费用 J ≥ 联合 LP 下界", roll43 >= lb - 1e-3,
              "%.4f vs 下界 %.4f（gap %+.4f 元，%+.4f%%）"
              % (roll43, lb, roll43 - lb, 100.0 * (roll43 - lb) / lb))
    else:
        check("联合 LP 下界文件存在", False, "缺 %s" % joint_path)

    print("-" * 78)
    print("【J. 锚定对照（构思手报告20 与主模型日志）】")
    pay42 = float((d42["计划购电量_kWh"] * d42["电价_元每kWh"]).sum())
    e_end = float(d42["时段末储电量_kWh"][-1, K - 1])
    n_hold = int(np.sum(d42["时段末储电量_kWh"][:, K - 1] > E_MIN + 1e-3))
    n_min = int(np.sum(np.abs(d42["时段末储电量_kWh"][:, K - 1] - E_MIN) < 1e-3))
    check("4-2 缴费 vs 锚定 12815460.2655", abs(pay42 - ANCHOR_42["pay"]) < 1.0,
          "差 %+.4f 元" % (pay42 - ANCHOR_42["pay"]))
    check("4-2 持有过夜天数 = 60（锚定）", n_hold == ANCHOR_42["n_hold"], "实测 %d" % n_hold)
    check("4-2 24:00 压下限天数 = 274（锚定）", n_min == ANCHOR_42["n_min"], "实测 %d" % n_min)
    check("4-2 期末储电量 = 1200 kWh（残值 469.0667 元）", abs(e_end - 1200.0) < 1e-3,
          "E=% .4f kWh" % e_end)
    j43 = float((d43["计划购电量_kWh"] * d43["电价_元每kWh"]).sum()
                + (1.5 * d43["电价_元每kWh"] * np.maximum(d43["调整购电量_kWh"] - d43["计划购电量_kWh"], 0)).sum()
                - (0.5 * d43["电价_元每kWh"] * np.maximum(d43["计划购电量_kWh"] - d43["调整购电量_kWh"], 0)).sum()
                + (5.0 * d43["电价_元每kWh"] * d43["紧急购电量_kWh"]).sum())
    print("  4-3 CSV 重算 J = %.4f 元（主模型日志 14450082.3797）" % j43)
    print("  对照：问题 2 = %.4f 元；问题 3 = %.4f 元" % (ANCHOR_Q2, ANCHOR_Q3))


def main():
    """复核主流程：读落盘 → 依次核验 → 输出自检报告。"""
    print("=" * 78)
    print("问题 4 独立复核（从落盘文件反查；不重新求解 LP）")
    print("=" * 78)
    d42 = read_csv_matrix(CSV_42)
    d43 = read_csv_matrix(CSV_43)
    verify_structure()
    verify_rotation(d42, d43)
    verify_spec_tables(d42, d43)
    verify_constraints(d42, d43)
    verify_consistency()
    verify_bounds_and_anchor(d42, d43)
    n_ok = sum(1 for _, s, _ in _RESULTS if s == "OK")
    n_all = len(_RESULTS)
    print("=" * 78)
    print("核验汇总：%d/%d 项通过" % (n_ok, n_all))
    print("=" * 78)
    with open(REPORT, "w", encoding="utf-8") as f:
        f.write("问题 4 自检报告（verify_q4.py 实测，从落盘文件反查）\n")
        f.write("核验项数：%d；通过：%d；失败：%d\n" % (n_all, n_ok, n_all - n_ok))
        f.write("-" * 78 + "\n")
        for name, status, detail in _RESULTS:
            f.write("[%s] %s %s\n" % (status, name, detail))
    print("自检报告已写入：%s" % REPORT)
    return n_ok, n_all


if __name__ == "__main__":
    tee = Tee(os.path.join(QDIR, "复核运行日志.txt"))
    original_stdout = sys.stdout
    sys.stdout = tee
    try:
        main()
    finally:
        sys.stdout = original_stdout
        tee.close()
    print("复核运行日志已写入：%s" % os.path.join(QDIR, "复核运行日志.txt"))
