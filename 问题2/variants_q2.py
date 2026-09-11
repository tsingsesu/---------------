"""问题 2 备选读法与口径对照：读法②（不完全信息）+ 每日终端=当日初始。

按 `问题2/交付清单.md` §三 执行，全部结果落盘；口径以 `口径与假设台账.md` 为准。

两套对照：
  A. 读法②（交付清单 §三 "备选读法②"）：0:00 制定计划时**不知道**当天的实际负载与光伏，
     只能依据"当天可得的典型日预测"（附件1 的负载/光伏列）制定计划；实际执行为附件2。
     计划量与实际需求的缺口按期前约束导出：r_k=[L_k^d·Δ+u_k−x_k−P_k^d·Δ−v_k]^+，
     费用含紧急购电项 5·p·r。**计划的储能轨迹跨日传递**（计划侧的 E 序列自洽）。
     另报"计划=前一日实际曲线"的变体（读法②b）作稳健性对照。
  B. 每日终端=当日初始（交付清单 §三）：滚动时每天加约束 E_144 = E_0（'cyclic'），
     用于量化问题 2 主模型"终端自由"的代价（D-04 要求的对照）。

输出文件（均在 问题2/ 下）：
  备选读法②_汇总.xlsx     读法②/②b 的汇总、逐日明细与四指定日期表 3
  备选读法②_逐日紧急购电.csv  读法② 334 天 × 144 时段紧急购电全分辨率明细
  口径对照_终端与读法.xlsx  终端自由 vs 终端=初始；读法① vs 读法② vs 读法②b

运行：python 问题2/variants_q2.py（建议先跑 run_q2.py 生成主模型产物）
依赖：numpy、scipy、openpyxl；lib/ 公共模块。
随机性：无（全流程确定性 LP，不需要随机种子）。
"""

import os
import sys

# 把工作区根目录加入模块路径，保证任意工作目录下可导入 lib
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import numpy as np
import openpyxl

from lib.dataio import read_attachment1, read_attachment2
from lib.solve_day import solve_day
from lib.storage import E_INIT, E_MAX, E_MIN, ETA, P_MAX
from lib.timegrid import DT_H, K

QDIR = os.path.join(ROOT, "问题2")
FIGDIR = os.path.join(ROOT, "图片", "问题2")                        # 图片统一目录（2026-09-11 起，用户要求）
os.makedirs(FIGDIR, exist_ok=True)                                # 确保目录存在（重跑时自动建）
XLSX_READING2 = os.path.join(QDIR, "备选读法②_汇总.xlsx")
CSV_EMG2 = os.path.join(QDIR, "备选读法②_逐日紧急购电.csv")
XLSX_VARIANT = os.path.join(QDIR, "口径对照_终端与读法.xlsx")
LOG_PATH = os.path.join(QDIR, "备选读法运行日志.txt")

ND = 4                                   # 小数位数（D-15）
D_REP_FIRST = 31                         # 填报区间首日 2025-02-01（0 基 31）
D_REP_LAST = 365                         # 填报区间末日 2025-12-31（0 基 364，不含于切片）
N_REP = D_REP_LAST - D_REP_FIRST         # 填报天数 = 334
SPEC_DATES = ("2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21")

# 构思手快速测算锚定（`进度台账.md` Phase 2 读法② 行）——只用于日志对照打印
ANCHOR_R2 = {
    "plan": 11757539.8955,       # 计划购电费，元
    "emg": 13393308.2086,        # 紧急购电费（5 倍价），元
    "day_mean": 75301.9404,      # 日均总费用，元
    "trigger_days": 293,         # 触发紧急购电的天数
    "max_r": 625.2438,           # 单时段最大紧急购电量，kWh
}


class Tee:
    """把控制台输出同时写入日志文件（与 run_q2.py 同型）。

    输入：path，str，日志路径；输出：无
    """

    def __init__(self, path):
        self.stdout = sys.stdout
        try:
            self.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
        self.file = open(path, "w", encoding="utf-8")

    def write(self, text):
        # 双通道写入，日志与回显一致
        self.stdout.write(text)
        self.file.write(text)

    def flush(self):
        self.stdout.flush()
        self.file.flush()

    def close(self):
        self.file.close()


def r4(values, tol=1e-6):
    """容差归零后四舍五入到 4 位小数（同 run_q2.py，D-15）。

    输入：values，array_like；tol，归零阈值
    输出：np.ndarray 或 float，4 位小数
    """
    arr = np.asarray(values, dtype=float)
    arr = np.where(np.abs(arr) < tol, 0.0, arr)
    return np.round(arr, ND)


def merge_intervals(r_day):
    """把一天 144 个时段的紧急购电量聚合成连续时段（表 4 写法）。

    输入：r_day，np.ndarray (K,)，kWh
    输出：(time_text, energy_text, n_seg)，无紧急购电时 ('', '0', 0)
    """
    from lib.timegrid import k_to_label
    active = np.asarray(r_day, dtype=float) > 1e-6
    if not active.any():
        return "", "0", 0
    idx = np.where(active)[0]
    segments, seg_start = [], idx[0]
    for j in range(1, idx.size + 1):
        # 遇不连续或末尾则封段
        if j == idx.size or idx[j] != idx[j - 1] + 1:
            segments.append((seg_start, idx[j - 1]))
            if j < idx.size:
                seg_start = idx[j]
    t_parts, e_parts = [], []
    for a, b in segments:
        t_parts.append("%s-%s" % (k_to_label(a + 1)[0], k_to_label(b + 1)[1]))
        seg_energy = r4(float(np.sum(r_day[a:b + 1])))
        e_parts.append(("%.4f" % seg_energy).rstrip("0").rstrip("."))
    return " ".join(t_parts), " ".join(e_parts), len(segments)


def roll_with_plan(price, load_plan, pv_plan, load_actual, pv_actual, e_init=E_INIT,
                   mode="free"):
    """"计划—实际"两阶段滚动：计划按预测曲线优化，实际缺口记紧急购电。

    输入：price，np.ndarray (K,)，电价，元/kWh
          load_plan / pv_plan，np.ndarray (D,K)，制定计划所用的负载/光伏预测，kW
          load_actual / pv_actual，np.ndarray (D,K)，实际执行的负载/光伏，kW
          e_init，kWh，滚动起点储电量（第 0 天 0:00）
          mode，str，单日 LP 的终端条件（'free' 为主口径）
    输出：dict，键含义——
          x_plan (D,K) 计划购电量；u_chg/v_dis (D,K) 计划充放电量；
          r_emg (D,K) 紧急购电量（按定义重算）；E_soc (D,K+1) 计划侧储电量轨迹；
          cost_plan_day (D,) 计划购电费；cost_emg_day (D,) 紧急购电费（5 倍价）；
          cost_day (D,) 合计；status_day (D,)
    """
    n_day = load_actual.shape[0]
    x_plan = np.zeros((n_day, K)); u_chg = np.zeros((n_day, K)); v_dis = np.zeros((n_day, K))
    r_emg = np.zeros((n_day, K)); E_soc = np.zeros((n_day, K + 1))
    cost_plan_day = np.zeros(n_day); cost_emg_day = np.zeros(n_day)
    cost_day = np.zeros(n_day); status_day = np.zeros(n_day, dtype=int)
    e = float(e_init)
    for d in range(n_day):
        # 计划阶段：0:00 用预测曲线求解当天计划（含储能充放电计划）
        res = solve_day(price, load_plan[d], pv_plan[d], e_init=e, mode=mode)
        x = res["x_plan"]; u = res["u_chg"]; v = res["v_dis"]
        # 实际执行：供给缺口按定义导出紧急购电量（5 倍价）
        r = np.maximum(load_actual[d] * DT_H + u - v - pv_actual[d] * DT_H - x, 0.0)
        x_plan[d] = x; u_chg[d] = u; v_dis[d] = v; r_emg[d] = r; E_soc[d] = res["E_soc"]
        cost_plan_day[d] = float(np.dot(price, x))           # 计划购电费，元
        cost_emg_day[d] = 5.0 * float(np.dot(price, r))      # 紧急购电费，元
        cost_day[d] = cost_plan_day[d] + cost_emg_day[d]
        status_day[d] = res["status"]
        e = float(res["E_soc"][-1])                          # 计划侧储电量跨日传递
    return {"x_plan": x_plan, "u_chg": u_chg, "v_dis": v_dis, "r_emg": r_emg,
            "E_soc": E_soc, "cost_plan_day": cost_plan_day, "cost_emg_day": cost_emg_day,
            "cost_day": cost_day, "status_day": status_day}


def write_xlsx(path, sheets):
    """把 {工作表名: 行列表} 写成 xlsx（每行是 list，元素可为 str/float/None）。

    输入：path，str；sheets，dict[str, list[list]]
    输出：str，写入路径
    """
    wb = openpyxl.Workbook()
    first = True
    for name, rows in sheets.items():
        ws = wb.active if first else wb.create_sheet()
        ws.title = name
        first = False
        for row in rows:
            ws.append(row)
        # 数值列统一 4 位小数格式（对 General 之外的字符串单元格无影响）
        for r in ws.iter_rows(min_row=2):
            for c in r:
                if isinstance(c.value, float):
                    c.number_format = "0.0000"
    wb.save(path)
    return path


def compute_readings(price, load1, pv1, load2, pv_actual, dates):
    """计算读法②（典型日计划）与读法②b（前一日实际计划）的全部结果。

    输入：price (K,)；load1/pv1 (K,) 典型日；load2/pv_actual (D,K)；dates list
    输出：dict，含两个变体的滚动结果与统计
    """
    # ---- 读法②：计划 = 典型日（附件1 的负载/光伏列广播到每天） ----
    load_typ = np.tile(load1, (365, 1))                  # (D,K) 典型日负载
    pv_typ = np.tile(pv1, (365, 1))                      # (D,K) 典型日光伏
    res_r2 = roll_with_plan(price, load_typ, pv_typ, load2, pv_actual, e_init=E_INIT)

    # ---- 读法②b：计划 = 前一日实际曲线（第 1 天无前一日，用典型日代替） ----
    load_prev = np.vstack([load1[None, :], load2[:-1]])  # 第 d 天的计划曲线 = 第 d−1 天实际
    pv_prev = np.vstack([pv1[None, :], pv_actual[:-1]])
    res_r2b = roll_with_plan(price, load_prev, pv_prev, load2, pv_actual, e_init=E_INIT)
    return {"r2": res_r2, "r2b": res_r2b}


def summarize(res, price):
    """对一段时间序列结果做填报区间统计。

    输入：res，roll_with_plan 返回值；price (K,)
    输出：dict，汇总统计（含逐日起止）
    """
    sl = slice(D_REP_FIRST, D_REP_LAST)
    cost_day = res["cost_day"][sl]; plan = res["cost_plan_day"][sl]; emg = res["cost_emg_day"][sl]
    r = res["r_emg"][sl]
    emg_kwh = float(r.sum())
    trigger = int(np.sum(r.max(axis=1) > 1e-6))          # 触发天数：当天存在 r>0 的时段
    return {
        "total": float(cost_day.sum()), "plan": float(plan.sum()), "emg": float(emg.sum()),
        "day_mean": float(cost_day.mean()), "emg_kwh": emg_kwh,
        "trigger_days": trigger, "max_r": float(r.max()),
        "max_day_emg": float(r.sum(axis=1).max()),
        "r2_trigger_rate": trigger / N_REP,
        "E0_rep": float(res["E_soc"][D_REP_FIRST, 0]),
        "E144_end": float(res["E_soc"][D_REP_LAST - 1, K]),
    }


def draw_emergency_figure(res_r2, dates):
    """绘制"紧急购电量分布图"（读法① 下作为机制说明图：假设按典型日计划会发生的紧急购电）。

    输入：res_r2，roll_with_plan 的返回值（读法②：计划=典型日、实际=附件2）
          dates，list[datetime.date]，365 天日期
    输出：str，PNG 落盘路径
    """
    import matplotlib.pyplot as plt
    from lib.plotstyle import (COLOR_CHG, COLOR_DIS, COLOR_PRICE, COLOR_REF,
                               FIGSIZE_TALL, apply_chinese_style, save_figure)
    apply_chinese_style()
    r = res_r2["r_emg"][D_REP_FIRST:D_REP_LAST]          # 填报区间 (334,144)
    daily_kwh = r.sum(axis=1)                            # 逐日紧急购电量，kWh
    # 各小时的平均紧急购电量：144 个 10 分钟时段按小时合并（24×6），再除以 334 天
    hour_kwh = r.sum(axis=0).reshape(24, 6).sum(axis=1) / 6.0 / 334.0
    fig, axes = plt.subplots(1, 2, figsize=FIGSIZE_TALL)
    # 左：逐日紧急购电量分布直方图
    axes[0].hist(daily_kwh, bins=24, color=COLOR_DIS, edgecolor="white", alpha=0.9)
    axes[0].axvline(daily_kwh.mean(), color=COLOR_PRICE, linestyle="--", linewidth=1.6,
                    label="日均 %.1f kWh" % daily_kwh.mean())
    axes[0].set_xlabel("单日紧急购电量（kWh）")
    axes[0].set_ylabel("天数（天）")
    axes[0].set_title("读法② 日紧急购电量分布\n（293/334 天触发，均值 %.0f、最大 %.0f kWh）"
                      % (daily_kwh.mean(), daily_kwh.max()))
    axes[0].legend()
    # 右：小时级时段分布（说明紧急购电出现在哪些时刻）
    hours = [("%d:00" % h) for h in range(24)]
    axes[1].bar(range(24), hour_kwh, width=0.8, color=COLOR_CHG, alpha=0.9)
    axes[1].set_xticks(range(0, 24, 2), ["%d:00" % h for h in range(0, 24, 2)], fontsize=9)
    axes[1].set_xlabel("时刻（h）")
    axes[1].set_ylabel("平均小时紧急购电量（kWh/h）")
    axes[1].set_title("紧急购电的时段分布\n（高发在 18:00–22:00 晚峰与清晨启动时段）")
    fig.suptitle("问题 2 紧急购电机制说明：若按典型日制定计划（读法②），实际负载/伏较典型日的\n"
                 "偏离将产生紧急购电（电价 5 倍）；主模型（读法① 完全信息）下该机制不被触发（r≡0）",
                 fontsize=14)
    fig.tight_layout(rect=(0, 0, 1, 0.93))
    return save_figure(fig, os.path.join(FIGDIR, "紧急购电量分布图.png"))


def main():
    """备选读法与终端对照的主流程：计算、落盘、打印。"""
    import datetime as _dt
    price, load1, pv1, _ = read_attachment1()
    load2, pv_actual, dates = read_attachment2()
    print("=" * 78)
    print("问题 2 备选读法②（不完全信息）与终端条件对照")
    print("=" * 78)

    # ---------------- 1. 读法② 两个变体 ----------------
    out = compute_readings(price, load1, pv1, load2, pv_actual, dates)
    s2 = summarize(out["r2"], price)
    s2b = summarize(out["r2b"], price)

    print("-" * 78)
    print("【读法② 汇总（填报区间 334 天）】计划 = 典型日（附件1 负载/光伏）；实际 = 附件2")
    print("计划购电费 = %.4f 元（锚定 %.4f，差 %+.4f）"
          % (s2["plan"], ANCHOR_R2["plan"], s2["plan"] - ANCHOR_R2["plan"]))
    print("紧急购电费 = %.4f 元（锚定 %.4f，差 %+.4f）"
          % (s2["emg"], ANCHOR_R2["emg"], s2["emg"] - ANCHOR_R2["emg"]))
    print("总费用 = %.4f 元；日均 = %.4f 元（锚定 %.4f）"
          % (s2["total"], s2["day_mean"], ANCHOR_R2["day_mean"]))
    print("紧急购电量 = %.4f kWh；触发天数 = %d / 334（锚定 %d）；单时段最大 r = %.4f kWh（锚定 %.4f）"
          % (s2["emg_kwh"], s2["trigger_days"], ANCHOR_R2["trigger_days"],
             s2["max_r"], ANCHOR_R2["max_r"]))
    print("最大单日紧急购电量 = %.4f kWh；填报期初 E_0(2.1) = %.4f kWh；期末 = %.4f kWh"
          % (s2["max_day_emg"], s2["E0_rep"], s2["E144_end"]))
    print("-" * 78)
    print("【读法②b 汇总】计划 = 前一日实际曲线（第 1 天用典型日）")
    print("计划购电费 = %.4f 元；紧急购电费 = %.4f 元；总费用 = %.4f 元；日均 = %.4f 元"
          % (s2b["plan"], s2b["emg"], s2b["total"], s2b["day_mean"]))
    print("紧急购电量 = %.4f kWh；触发天数 = %d / 334；单时段最大 r = %.4f kWh"
          % (s2b["emg_kwh"], s2b["trigger_days"], s2b["max_r"]))
    print("对照：读法①（主模型）总费用 12254765.7161 元，日均 36690.9153 元 —— "
          "信息不足使日均升到 %.4f 倍（读法②）/ %.4f 倍（读法②b）"
          % (s2["day_mean"] / 36690.9153, s2b["day_mean"] / 36690.9153))

    # ---------------- 2. 终端条件对照：终端自由 vs 终端=初始 ----------------
    # 终端=初始（cyclic）：每天加约束 E_144 = E_0，E_0^1 = 6000（跨日仍连续但恒定 6000）
    from lib.run_days import solve_rolling
    roll_cyclic = solve_rolling(price, load2, pv_actual, e_init=E_INIT, mode="cyclic")
    sl = slice(D_REP_FIRST, D_REP_LAST)
    cyc_total = float(roll_cyclic["cost_day"][sl].sum())
    cyc_mean = cyc_total / N_REP
    cyc_E0 = float(roll_cyclic["E_soc"][D_REP_FIRST, 0])
    print("-" * 78)
    print("【终端条件对照】")
    print("终端自由（主模型）总费用 = 12254765.7161 元；日均 = 36690.9153 元；每日 24:00 全部压到 1200 kWh")
    print("终端=当日初始（cyclic）总费用 = %.4f 元；日均 = %.4f 元；每日 0:00/24:00 恒为 %.4f kWh"
          % (cyc_total, cyc_mean, cyc_E0))
    cyc_diff = cyc_total - 12254765.7161                 # cyclic − 终端自由（负值 = cyclic 更省）
    cyc_pct = 100.0 * cyc_diff / 12254765.7161
    if cyc_diff >= 0:
        print("终端锁定使总费用上升 %.4f 元（+%.4f%%）——即'自由终端让每日末压到下限'带来的节省上限"
              % (cyc_diff, cyc_pct))
    else:
        print("终端锁定使总费用下降 %.4f 元（%.4f%%，即 cyclic 比终端自由更省）——"
              "方向来自 1 月初值存量被循环链逐步变现的承诺效应；两者均远高于联合下界，结论不翻转"
              % (abs(cyc_diff), cyc_pct))

    # ---------------- 3. 落盘 ----------------
    # 3.1 读法② 汇总 + 逐日明细 + 四个指定日期表 3
    detail_rows = [["日期", "计划购电费_元", "紧急购电费_元", "总费用_元", "紧急购电量_kWh",
                    "紧急购电区间数", "单时段最大紧急购电量_kWh", "0:00储电量_kWh", "24:00储电量_kWh"]]
    for d in range(D_REP_FIRST, D_REP_LAST):
        t_text, e_text, n_seg = merge_intervals(out["r2"]["r_emg"][d])
        detail_rows.append([
            dates[d].isoformat(), r4(out["r2"]["cost_plan_day"][d]), r4(out["r2"]["cost_emg_day"][d]),
            r4(out["r2"]["cost_day"][d]), r4(float(out["r2"]["r_emg"][d].sum())), n_seg,
            r4(float(out["r2"]["r_emg"][d].max())),
            r4(out["r2"]["E_soc"][d, 0]), r4(out["r2"]["E_soc"][d, K]),
        ])
    # 四个指定日期表 3（按题目表 3 的"4 日期并排"精神，转成逐日期两列纵向排列）
    spec_rows = [["日期", "时间段", "购电量_kWh"]]
    for t in SPEC_DATES:
        d0 = (_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days
        t_text, e_text, _ = merge_intervals(out["r2"]["r_emg"][d0])
        spec_rows.append([t, t_text if t_text else "无（0）", e_text])
    summary_rows = [
        ["指标", "读法②（计划=典型日）", "读法②b（计划=前一日实际）", "读法①（主模型，对照）"],
        ["计划购电费_元", r4(s2["plan"]), r4(s2b["plan"]), 12254765.7161],
        ["紧急购电费_元", r4(s2["emg"]), r4(s2b["emg"]), 0.0],
        ["总费用_元", r4(s2["total"]), r4(s2b["total"]), 12254765.7161],
        ["日均费用_元", r4(s2["day_mean"]), r4(s2b["day_mean"]), 36690.9153],
        ["紧急购电量_kWh", r4(s2["emg_kwh"]), r4(s2b["emg_kwh"]), 0.0],
        ["触发紧急购电天数（占 334 天）", s2["trigger_days"], s2b["trigger_days"], 0],
        ["单时段最大紧急购电量_kWh", r4(s2["max_r"]), r4(s2b["max_r"]), 0.0],
        ["最大单日紧急购电量_kWh", r4(s2["max_day_emg"]), r4(s2b["max_day_emg"]), 0.0],
        ["填报期初储电量_kWh", r4(s2["E0_rep"]), r4(s2b["E0_rep"]), 1200.0],
        ["填报期末储电量_kWh", r4(s2["E144_end"]), r4(s2b["E144_end"]), 1200.0],
        ["口径说明", "计划侧储能轨迹跨日传递；实际缺口按 5 倍电价紧急购电", "同左，计划曲线换为前一日实际",
         "完全信息：计划恰好覆盖净负荷，r≡0；该列 计划购电费=总费用、紧急费用=0"],
    ]
    write_xlsx(XLSX_READING2, {
        "汇总": summary_rows,
        "逐日明细": detail_rows,
        "表3_四指定日期": spec_rows,
    })
    # 3.2 读法② 逐时段紧急购电 CSV（全分辨率审计）
    lines = ["日期,时段序号,时段起,时段止,紧急购电量_kWh"]
    from lib.timegrid import k_to_label
    for d in range(D_REP_FIRST, D_REP_LAST):
        for k in range(1, K + 1):
            lines.append("%s,%d,%s,%s,%.4f" % (
                dates[d].isoformat(), k, k_to_label(k)[0], k_to_label(k)[1],
                r4(out["r2"]["r_emg"][d, k - 1])))
    with open(CSV_EMG2, "w", encoding="utf-8", newline="") as f:
        f.write("\n".join(lines) + "\n")
    # 3.3 终端与读法总对照
    write_xlsx(XLSX_VARIANT, {
        "终端条件对照": [
            ["口径", "总费用_元", "日均_元", "每日 24:00 储电量", "与主模型差_元", "与主模型差_%"],
            ["终端自由（主模型，D-04）", 12254765.7161, 36690.9153, "全部 1200 kWh", 0.0, 0.0],
            ["终端=当日初始（cyclic）", r4(cyc_total), r4(cyc_mean), "恒为 %.4f kWh" % cyc_E0,
             r4(cyc_total - 12254765.7161), r4(100.0 * (cyc_total - 12254765.7161) / 12254765.7161)],
        ],
        "读法对照": [
            ["口径", "总费用_元", "日均_元", "紧急购电费_元", "触发天数", "单时段最大_r_kWh"],
            ["读法① 完全信息（主模型）", 12254765.7161, 36690.9153, 0.0, 0, 0.0],
            ["读法② 计划=典型日", r4(s2["total"]), r4(s2["day_mean"]), r4(s2["emg"]),
             s2["trigger_days"], r4(s2["max_r"])],
            ["读法②b 计划=前一日实际", r4(s2b["total"]), r4(s2b["day_mean"]), r4(s2b["emg"]),
             s2b["trigger_days"], r4(s2b["max_r"])],
        ],
    })
    print("已落盘：%s" % XLSX_READING2)
    print("已落盘：%s" % CSV_EMG2)
    print("已落盘：%s" % XLSX_VARIANT)
    png_emg = draw_emergency_figure(out["r2"], dates)
    print("已落盘：%s" % png_emg)


if __name__ == "__main__":
    tee = Tee(LOG_PATH)
    original_stdout = sys.stdout
    sys.stdout = tee
    try:
        main()
    finally:
        sys.stdout = original_stdout
        tee.close()
        print("运行日志已写入：%s" % LOG_PATH)
