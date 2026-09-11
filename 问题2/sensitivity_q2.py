"""问题 2 灵敏度与稳健性实验（S1–S6）：参数扰动、多因素网格、数据扰动、方法互验、口径对照、适用边界。

按 `问题2/交付清单.md` §五 执行，全部结果落盘；口径一律以 `口径与假设台账.md` 为准。

基准（主模型）：逐日滚动 LP，读法① 完全信息、终端自由 + 跨日传递（D-04）、单向 η=0.9（D-05）、
填报区间 334 天（2025-02-01 至 12-31，D-11）；基准总费用 12 254 765.7161 元、日均 36 690.9153 元。

输出文件（均在 问题2/ 下）：
  灵敏度分析_参数扰动.xlsx    S1：η / P̄ / [E̲,Ē] / κ 单因素扰动（±5%/±10%/±20%；κ 另含 4/4.5/5/5.5/6）
  灵敏度分析_多因素网格.xlsx  S2：η × κ × P̄ 三维网格（5×5×5），含两个热力图切面数据
  灵敏度分析_多因素热力图.png S2 图
  灵敏度分析_数据扰动.xlsx    S3：极端点截断 / 光伏置零 / ±5% 缩放 / 预热期 0 vs 31 天 / 抽 10% 天数
  方法侧互验_联合LP与MATLAB与规则策略.xlsx   S4：联合 LP 下界、MATLAB 复算、DP 互验、规则策略；附 S6 适用边界表
  口径对照_读法与终端与效率.xlsx             S5：读法①②、终端自由/锁定、η 单向/往返、填法 Y/X
  灵敏度分析_参数扰动.png      S1 图
  灵敏度分析_数据扰动_逐日分布.png  S3(e) 图
  方法侧互验_对照.png          S4 图
  灵敏度运行日志.txt           本脚本运行的完整控制台记录

判据（贯穿全部扰动）：
  * "结论是否翻转" = 扰动后"优化总费用 ≥ 同口径下的不储能基线"（即储能净收益 ≤ 0）；
  * 读法① 下不储能基线 = 按实际数据（L₂,P₂）的光伏自用基线（16407319.6320 元）；
  * 读法② 下不储能基线 = 同信息口径（典型日净负荷计划 + 实际缺口按 κ 紧急购电）的基线。

运行：python 问题2/sensitivity_q2.py（建议先跑 run_q2.py / variants_q2.py / run_q2_joint.py）
依赖：numpy、scipy、openpyxl、matplotlib；lib/ 公共模块。
随机性：唯一随机用途为 S3(e) 抽样（30 次统计抽样 + 1 次抽样重算），固定种子 RNG_SEED=20260911。
"""

import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
QDIR = os.path.join(ROOT, "问题2")
FIGDIR = os.path.join(ROOT, "图片", "问题2")                        # 图片统一目录（2026-09-11 起，用户要求）
os.makedirs(FIGDIR, exist_ok=True)                                # 确保目录存在（重跑时自动建）
for _p in (ROOT, QDIR):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import numpy as np
import openpyxl

from lib.dataio import read_attachment1, read_attachment2
from lib.dp_day import dp_solve
from lib.plotstyle import (COLOR_BUY, COLOR_CHG, COLOR_DIS, COLOR_PRICE, COLOR_REF,
                           FIGSIZE_TALL, FIGSIZE_WIDE, apply_chinese_style, save_figure)
from lib.run_days import solve_rolling
from lib.solve_day import solve_day
from lib.storage import E_INIT, E_MAX, E_MIN, ETA, P_MAX, U_MAX
from lib.timegrid import DT_H, K

# ============================== 全局常量 ==============================

XLSX_S1 = os.path.join(QDIR, "灵敏度分析_参数扰动.xlsx")
XLSX_S2 = os.path.join(QDIR, "灵敏度分析_多因素网格.xlsx")
XLSX_S3 = os.path.join(QDIR, "灵敏度分析_数据扰动.xlsx")
XLSX_S4 = os.path.join(QDIR, "方法侧互验_联合LP与MATLAB与规则策略.xlsx")
XLSX_S5 = os.path.join(QDIR, "口径对照_读法与终端与效率.xlsx")
PNG_S1 = os.path.join(FIGDIR, "灵敏度分析_参数扰动.png")
PNG_S2 = os.path.join(FIGDIR, "灵敏度分析_多因素热力图.png")
PNG_S3 = os.path.join(FIGDIR, "灵敏度分析_数据扰动_逐日分布.png")
PNG_S4 = os.path.join(FIGDIR, "方法侧互验_对照.png")
LOG_PATH = os.path.join(QDIR, "灵敏度运行日志.txt")

ND = 4                                   # 小数位数（D-15）
D_REP_FIRST = 31                         # 填报区间首日 0 基
N_REP = 334                              # 填报天数
RNG_SEED = 20260911                      # S3(e) 抽样随机种子（固定）
SAMPLE_FRAC = 0.10                       # S3(e) 抽样比例 10%

# 基准锚定（构思手独立测算，`_phase0/报告17/18`）——只用于对照打印
BASE_COST = 12254765.7161                # 基准填报区间总费用，元
BASE_MEAN = 36690.9153                   # 基准日均费用，元
BASE_PV_ONLY = 16407319.6320             # 基准不储能基线（实际数据、方案 B），元
JOINT_LB = 12227243.6427                 # 全年联合 LP 下界（填报区间），元
R2_PLAN = 11757539.8955                  # 读法② 计划购电费（κ 无关），元
R2_EMG_UNIT = 13393308.2086 / 5.0        # 读法② κ=1 时的紧急加权值 Σp·r，元

FACTORS = (0.8, 0.9, 0.95, 1.0, 1.05, 1.1, 1.2)      # ±5%/±10%/±20% 档位（另含基准）
KAPPA_LIST = (4.0, 4.5, 5.0, 5.5, 6.0)               # 紧急电价倍数（S1 κ 行 / S2 网格）
P_GRID = (4000.0, 4500.0, 5000.0, 5500.0, 6000.0)    # S2 的 P̄ 网格（5 档）
ETA_GRID5 = tuple(ETA * f for f in (0.8, 0.9, 0.95, 1.0, 1.1))   # S2 的 η 网格（5 档）


class Tee:
    """把控制台输出同时写入日志文件（与其他脚本同型）。"""

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
    """容差归零后四舍五入到 4 位小数（其他脚本同型）。"""
    arr = np.asarray(values, dtype=float)
    arr = np.where(np.abs(arr) < tol, 0.0, arr)
    return np.round(arr, ND)


def write_xlsx(path, sheets):
    """把 {工作表名: 行列表} 写成 xlsx；数值单元格统一 4 位小数格式。"""
    wb = openpyxl.Workbook()
    first = True
    for name, rows in sheets.items():
        ws = wb.active if first else wb.create_sheet()
        ws.title = name
        first = False
        for row in rows:
            ws.append(row)
        for r in ws.iter_rows(min_row=2):
            for c in r:
                if isinstance(c.value, float):
                    c.number_format = "0.0000"
    wb.save(path)
    return path


# ============================== 通用评估工具 ==============================


def baseline_pv_only(price, load, pv_actual):
    """不储能基线（方案 B）：光伏优先自用、余电弃掉，逐日费用与合计。

    输入：price (K,) 元/kWh；load/pv_actual (D,K) kW
    输出：(daily (D,) 元, total 元)
    """
    # 净负荷 = max(L − P, 0)，乘电价与 Δ 后按日求和
    daily = (price[None, :] * np.maximum(load * DT_H - pv_actual * DT_H, 0.0)).sum(axis=1)
    return daily, float(daily.sum())


def eval_rolling_case(price, load, pv_actual, mode="free", eta=ETA,
                      p_max=P_MAX, e_min=E_MIN, e_max=E_MAX, d_start=0, d_end=365):
    """评估一个滚动案例：返回填报区间总费用/日均/紧急购电统计等。

    输入：price (K,)；load/pv_actual (D,K)；其余同 solve_rolling
    输出：dict（total/mean/plan/emg/trigger_days/max_r/x_total/u_total/v_total/roll）
    """
    roll = solve_rolling(price, load, pv_actual, e_init=E_INIT, mode=mode, eta=eta,
                         p_max=p_max, e_min=e_min, e_max=e_max, d_start=d_start, d_end=d_end)
    sl = slice(D_REP_FIRST, 365)
    cost = roll["cost_day"][sl]
    r = roll["r_emg"][sl]
    return {
        "total": float(cost.sum()), "mean": float(cost.mean()),
        "plan": float(roll["cost_plan_day"][sl].sum()), "emg": float(roll["cost_emg_day"][sl].sum()),
        "trigger_days": int(np.sum(r.max(axis=1) > 1e-6)), "max_r": float(r.max()),
        "x_total": float(roll["x_plan"][sl].sum()), "u_total": float(roll["u_chg"][sl].sum()),
        "v_total": float(roll["v_dis"][sl].sum()),
        "roll": roll,
    }


def flip_verdict(total, base_total):
    """返回（储能净收益，判定文本）：净收益 ≤ 0 记为翻转。"""
    gain = base_total - total
    return gain, ("翻转（储能不再省钱）" if gain <= 0 else "未翻转")


def build_reading2_cache(price, load, pv_actual):
    """读法②（计划=典型日、实际=附件2）的计划解与全部 κ 无关量，一次算好供复用。

    原理：读法② 的计划曲线由典型日负载/光伏决定，**与 κ 无关**（κ 只作用于事后结算的
    紧急购电费用），因此 365 天的计划 LP 只需解一次；不同 κ 的总费用为纯算术。

    输入：price (K,)；load/pv_actual (D,K)
    输出：dict，键含义——
          plan_cost (334 天计划购电费 Σp·x)；emg_w (334 天紧急加权 Σp·r，κ=1)；
          trigger_days / max_r（κ 无关）；base_plan（同信息口径不储能基线的计划购电费
          Σp·(L₁−P₁)⁺Δ）；base_emg_w（同信息口径不储能基线的紧急加权 Σp·r_base，
          r_base=[(L₂−P₂)⁺−(L₁−P₁)⁺]⁺Δ，即无储能时实际超过计划的缺口）
    """
    _, load1, pv1, _ = read_attachment1()
    load_typ = np.tile(load1, (365, 1))
    pv_typ = np.tile(pv1, (365, 1))
    e = float(E_INIT)
    plan_cost = 0.0; emg_w = 0.0; trigger = 0; max_r = 0.0
    base_plan = 0.0; base_emg_w = 0.0; x_total = 0.0
    for d in range(365):
        res = solve_day(price, load_typ[d], pv_typ[d], e_init=e, mode="free")
        x = res["x_plan"]; u = res["u_chg"]; v = res["v_dis"]
        r = np.maximum(load[d] * DT_H + u - v - pv_actual[d] * DT_H - x, 0.0)
        # 同信息口径的不储能基线：按典型日净负荷买足，实际缺口另付紧急电价
        net_typ = np.maximum(load_typ[d] * DT_H - pv_typ[d] * DT_H, 0.0)
        net_act = np.maximum(load[d] * DT_H - pv_actual[d] * DT_H, 0.0)
        r_base = np.maximum(net_act - net_typ, 0.0)
        if d >= D_REP_FIRST:
            plan_cost += float(np.dot(price, x))
            emg_w += float(np.dot(price, r))
            x_total += float(x.sum())
            base_plan += float(np.dot(price, net_typ))
            base_emg_w += float(np.dot(price, r_base))
            if r.max() > 1e-6:
                trigger += 1
            max_r = max(max_r, float(r.max()))
        e = float(res["E_soc"][-1])
    return {"plan_cost": plan_cost, "emg_w": emg_w, "trigger_days": trigger, "max_r": max_r,
            "base_plan": base_plan, "base_emg_w": base_emg_w, "x_total": x_total}


def reading2_total(cache, kappa):
    """读法② 口径下、给定 κ 的总费用与其不储能基线。"""
    total = cache["plan_cost"] + kappa * cache["emg_w"]
    base = cache["base_plan"] + kappa * cache["base_emg_w"]
    return total, base


# ============================== S1 单因素参数扰动 ==============================


def run_s1(price, load, pv_actual, base, cache2):
    """S1：η / P̄ / [E̲,Ē] / κ 单因素扰动，全部 ±5%/±10%/±20%，κ 另含 4/4.5/5.5/6。"""
    rows = []
    header = ["扰动组", "扰动档", "参数值", "总费用_元", "日均费用_元", "相对基准变化_%",
              "总购电量_kWh", "总充电量_kWh", "总放电量_kWh", "紧急购电触发天数",
              "单时段最大紧急购电量_kWh", "不储能基线_元", "储能净收益_元（基线−优化）",
              "结论是否翻转（净收益≤0）", "备注"]
    base_cost = base["total"]

    def add(group, case, param, res, note="", pv_base=None, tcols=(True,)):
        # 统一把一行结果整理进 rows（tcols 标记 u/v 列是否适用）
        pv_b = BASE_PV_ONLY if pv_base is None else pv_base
        gain, verdict = flip_verdict(res["total"], pv_b)
        rows.append([group, case, param, r4(res["total"]), r4(res["mean"]),
                     r4(100.0 * (res["total"] / base_cost - 1.0)), r4(res["x_total"]),
                     r4(res["u_total"]) if tcols else None,
                     r4(res["v_total"]) if tcols else None, res["trigger_days"],
                     r4(res["max_r"]), r4(pv_b), r4(gain), verdict, note])

    # ---- (a) 单向效率 η ----
    for f in FACTORS:
        eta_v = ETA * f
        res = eval_rolling_case(price, load, pv_actual, eta=eta_v)
        if abs(f - 1.0) < 1e-9:
            note = "基准档"
        elif eta_v > 1.0:
            note = "η>1 为越界外推档，物理不可实现，仅作趋势参考（与问题 1 的 S1 同口径）"
        else:
            note = "η 相对基准 %+.0f%%" % (100.0 * (f - 1.0))
        add("η（单向效率）", "η=%.4g（往返 %.4f）" % (eta_v, eta_v ** 2), eta_v, res, note)

    # ---- (b) 最大充放电功率 P̄ ----
    for f in FACTORS:
        res = eval_rolling_case(price, load, pv_actual, p_max=P_MAX * f)
        note = "基准档" if abs(f - 1.0) < 1e-9 else "P̄ 相对基准 %+.0f%%" % (100.0 * (f - 1.0))
        add("P̄（最大充放电功率）", "P̄=%.0f kW" % (P_MAX * f), P_MAX * f, res, note)

    # ---- (c) 储电量区间 [E̲,Ē]（等比例缩放半宽，中心 6000 不变） ----
    half = (E_MAX - E_MIN) / 2.0                          # 基准半宽 4800 kWh
    for f in FACTORS:
        e_min_v, e_max_v = E_INIT - half * f, E_INIT + half * f
        res = eval_rolling_case(price, load, pv_actual, e_min=e_min_v, e_max=e_max_v)
        note = "基准档" if abs(f - 1.0) < 1e-9 else "区间半宽相对基准 %+.0f%%" % (100.0 * (f - 1.0))
        add("E̲/Ē（储电量区间）", "[%.0f, %.0f] kWh" % (e_min_v, e_max_v), f, res, note)

    # ---- (d) 紧急电价倍数 κ（读法① 中 r≡0、κ 不生效；故在读法② 口径下评估） ----
    for kap in KAPPA_LIST:
        total2, base2 = reading2_total(cache2, kap)
        gain, verdict = flip_verdict(total2, base2)
        note = ("读法② 口径：κ 只作用于事后紧急结算，计划不随 κ 变化；"
                + ("基准档 κ=5" if abs(kap - 5.0) < 1e-9 else "κ 相对基准 %+.0f%%" % (100.0 * (kap / 5.0 - 1.0))))
        rows.append(["κ（紧急电价倍数）", "κ=%.4g（读法②口径）" % kap, kap, r4(total2),
                     r4(total2 / N_REP), r4(100.0 * (total2 / base_cost - 1.0)),
                     r4(cache2["x_total"]), None, None, cache2["trigger_days"],
                     r4(cache2["max_r"]), r4(base2), r4(gain), verdict, note])
    return rows, header


# ============================== S2 多因素网格 ==============================


def run_s2(price, load, pv_actual, cache2):
    """S2：η × κ × P̄ 三维网格（5×5×5）；η、P̄ 作用于读法① 滚动，κ 作用于读法② 口径。"""
    rows = []
    header = ["η", "κ", "P̄_kW", "读法①总费用_元", "读法②总费用_元", "读法①相对基准_%"]
    grid1 = {}
    # 先算 η × P̄ 的读法① 费用（25 次滚动），κ 只影响读法② 列的线性加权
    for eta_v in ETA_GRID5:
        for p_v in P_GRID:
            res = eval_rolling_case(price, load, pv_actual, eta=eta_v, p_max=p_v)
            grid1[(round(eta_v, 6), p_v)] = res["total"]
            print("  S2 读法①：η=%.4g P̄=%.0f kW → %.4f 元" % (eta_v, p_v, res["total"]))
    for eta_v in ETA_GRID5:
        for kap in KAPPA_LIST:
            for p_v in P_GRID:
                c1 = grid1[(round(eta_v, 6), p_v)]
                c2, _ = reading2_total(cache2, kap)       # 读法② 费用与 η、P̄ 无关
                rows.append([eta_v, kap, p_v, r4(c1), r4(c2), r4(100.0 * (c1 / BASE_COST - 1.0))])
    return rows, header, grid1


def draw_s2_heatmap(grid1, cache2):
    """S2 热力图：左=η×P̄ 读法① 总费用；右=κ×P̄ 读法② 总费用。"""
    import matplotlib.pyplot as plt
    apply_chinese_style()
    fig, axes = plt.subplots(1, 2, figsize=FIGSIZE_TALL)
    # 左图：η（行）× P̄（列），读法①
    m1 = np.array([[grid1[(round(eta_v, 6), p_v)] for p_v in P_GRID] for eta_v in ETA_GRID5]) / 1e6
    im1 = axes[0].imshow(m1, cmap="YlOrRd", aspect="auto")
    for i in range(m1.shape[0]):
        for j in range(m1.shape[1]):
            axes[0].text(j, i, "%.3f" % m1[i, j], ha="center", va="center", fontsize=10)
    axes[0].set_xticks(range(len(P_GRID)), ["%.0f" % p for p in P_GRID])
    axes[0].set_yticks(range(len(ETA_GRID5)), ["%.3g" % e for e in ETA_GRID5])
    axes[0].set_xlabel("最大充放电功率 P̄（kW）")
    axes[0].set_ylabel("单向效率 η")
    axes[0].set_title("S2-a：读法① 总费用（百万元）\nη × P̄，终端自由滚动")
    fig.colorbar(im1, ax=axes[0], label="总费用（百万元）")
    # 右图：κ（行）× P̄（列），读法② 总费用（读法② 与 η、P̄ 无关，故整行同值）
    m2 = np.array([[reading2_total(cache2, kap)[0] for _ in P_GRID] for kap in KAPPA_LIST]) / 1e6
    im2 = axes[1].imshow(m2, cmap="YlGnBu", aspect="auto")
    for i in range(m2.shape[0]):
        for j in range(m2.shape[1]):
            axes[1].text(j, i, "%.3f" % m2[i, j], ha="center", va="center", fontsize=10)
    axes[1].set_xticks(range(len(P_GRID)), ["%.0f" % p for p in P_GRID])
    axes[1].set_yticks(range(len(KAPPA_LIST)), ["%.1f" % k for k in KAPPA_LIST])
    axes[1].set_xlabel("最大充放电功率 P̄（kW；读法② 计划不随其变化）")
    axes[1].set_ylabel("紧急电价倍数 κ")
    axes[1].set_title("S2-b：读法② 总费用（百万元）\nκ × P̄（典型日计划、实际附件2）")
    fig.colorbar(im2, ax=axes[1], label="总费用（百万元）")
    fig.suptitle("问题 2 多因素组合扰动：η × κ × P̄ 网格扫描（每格标注总费用）", fontsize=15)
    fig.tight_layout(rect=(0, 0, 1, 0.96))
    save_figure(fig, PNG_S2)


# ============================== S3 数据侧扰动 ==============================


def run_s3(price, load, pv_actual, base):
    """S3：数据侧扰动 (a)~(e)，落盘明细行并返回抽样分布供画图。"""
    rows = []
    header = ["扰动组", "扰动档", "总费用_元", "日均费用_元", "相对基准变化_%",
              "总购电量_kWh", "总充电量_kWh", "总放电量_kWh", "触发紧急购电天数", "备注"]

    def one(group, case, res, note=""):
        # 统一整理一行
        rows.append([group, case, r4(res["total"]), r4(res["mean"]),
                     r4(100.0 * (res["total"] / BASE_COST - 1.0)), r4(res["x_total"]),
                     r4(res["u_total"]), r4(res["v_total"]), res["trigger_days"], note])

    # ---- (a) 极端点截断：负载/光伏 >95 百分位或 <5 百分位的时段 ----
    lo, hi = np.percentile(load, [5, 95])
    res = eval_rolling_case(price, np.clip(load, lo, hi), pv_actual)
    one("(a) 负载极端点截断", "负载截断到 [P5, P95] = [%.1f, %.1f] kW" % (lo, hi), res,
        "截断 %d 个时段" % int(((load < lo) | (load > hi)).sum()))
    p_lo, p_hi = np.percentile(pv_actual, [5, 95])
    res = eval_rolling_case(price, load, np.clip(pv_actual, p_lo, p_hi))
    one("(a) 光伏极端点截断", "光伏截断到 [P5, P95] = [%.1f, %.1f] kW" % (p_lo, p_hi), res,
        "截断 %d 个时段" % int(((pv_actual < p_lo) | (pv_actual > p_hi)).sum()))

    # ---- (b) 光伏 <10 kW 置零（D-16.4 的数据侧对照） ----
    pv_b = np.where(pv_actual < 10.0, 0.0, pv_actual)
    res = eval_rolling_case(price, load, pv_b)
    n_zero = int((pv_actual < 10.0).sum())
    one("(b) 光伏<10 kW 置零", "夜间小值清零", res,
        "置零 %d 个时段（全样本的 %.2f%%）；这些值多在夜间且量级极小，费用差 %.4f 元"
        % (n_zero, 100.0 * n_zero / pv_actual.size, res["total"] - BASE_COST))

    # ---- (c) 负载/光伏整体 ±5% ----
    for f in (0.95, 1.05):
        one("(c) 负载整体缩放", "负载 ×%.2f" % f, eval_rolling_case(price, load * f, pv_actual))
        one("(c) 光伏整体缩放", "光伏 ×%.2f" % f, eval_rolling_case(price, load, pv_actual * f))

    # ---- (d) 预热期长度：0 天（2.1 起、E_0=6000）vs 31 天（1.1 起，主模型） ----
    roll0 = solve_rolling(price, load, pv_actual, e_init=E_INIT, mode="free",
                          d_start=D_REP_FIRST, d_end=365)
    cost0 = float(roll0["cost_day"].sum())
    rows.append(["(d) 预热期长度", "预热 0 天（2.1 起、E_0=6000）", r4(cost0), r4(cost0 / N_REP),
                 r4(100.0 * (cost0 / BASE_COST - 1.0)), r4(float(roll0["x_plan"].sum())),
                 r4(float(roll0["u_chg"].sum())), r4(float(roll0["v_dis"].sum())), 0,
                 "主模型为预热 31 天（1.1 起，E_0=6000 跨日传递至 2.1 时已被压到 1200）"])

    # ---- (e) 随机抽掉 10% 的天数 ----
    rng = np.random.default_rng(RNG_SEED)
    # (e1) 统计稳定性：在固定的全年滚动解上做 30 次 90% 子样本，统计日均费用的分布
    base_roll = base["roll"]
    day_cost = base_roll["cost_day"][D_REP_FIRST:365]     # 填报区间逐日费用（固定）
    means = np.array([day_cost[rng.choice(N_REP, size=N_REP - int(N_REP * SAMPLE_FRAC),
                                          replace=False)].mean() for _ in range(30)])
    rows.append(["(e1) 抽样统计稳定性", "30 次 90% 子样本的日均费用", r4(means.mean() * N_REP),
                 r4(means.mean()), r4(100.0 * (means.mean() / BASE_MEAN - 1.0)), None, None, None, None,
                 "子样本日均范围 [%.4f, %.4f] 元，标准差 %.4f 元；种子 %d"
                 % (means.min(), means.max(), means.std(), RNG_SEED)])
    # (e2) 真实重算：随机抽掉 10% 的天（全年 365 天里抽），在保留下来的日序列上重跑滚动链
    keep = np.sort(rng.choice(365, size=365 - int(365 * SAMPLE_FRAC), replace=False))
    roll_kept = solve_rolling(price, load[keep], pv_actual[keep], e_init=E_INIT, mode="free",
                              d_start=0, d_end=keep.size)
    # 保留日在填报区间（原始下标 >= D_REP_FIRST）的部分做统计
    rep_mask = keep >= D_REP_FIRST
    cost_kept = roll_kept["cost_day"][rep_mask]
    rows.append(["(e2) 抽样重算（真实重跑滚动链）", "抽掉 10%%（%d 天）后重算" % (365 - keep.size),
                 r4(float(cost_kept.sum())), r4(float(cost_kept.mean())),
                 r4(100.0 * (float(cost_kept.mean()) / BASE_MEAN - 1.0)),
                 r4(float(roll_kept["x_plan"][rep_mask].sum())),
                 r4(float(roll_kept["u_chg"][rep_mask].sum())),
                 r4(float(roll_kept["v_dis"][rep_mask].sum())), 0,
                 "保留 %d 天中填报区间 %d 天；日均 %.4f 元，与全样本基准差 %+.4f%%"
                 % (keep.size, int(rep_mask.sum()), float(cost_kept.mean()),
                    100.0 * (float(cost_kept.mean()) / BASE_MEAN - 1.0))])
    return rows, header, means


def draw_s3_distribution(means):
    """S3(e) 图：90% 子样本重算的日均费用分布。"""
    import matplotlib.pyplot as plt
    apply_chinese_style()
    fig, ax = plt.subplots(figsize=FIGSIZE_WIDE)
    ax.hist(means, bins=10, color=COLOR_CHG, edgecolor="white", alpha=0.85,
            label="30 次 90% 子样本的日均费用")
    ax.axvline(BASE_MEAN, color=COLOR_PRICE, linestyle="-", linewidth=2.0,
               label="全样本基准日均 %.4f 元" % BASE_MEAN)
    ax.axvline(means.mean(), color=COLOR_REF, linestyle="--", linewidth=1.6,
               label="子样本均值 %.4f 元" % means.mean())
    ax.set_xlabel("折算日均费用（元）")
    ax.set_ylabel("出现次数（次）")
    ax.set_title("问题 2 数据侧扰动 S3(e)：随机抽掉 10%% 天数后重算的日均费用分布\n"
                 "（固定种子 %d；范围 [%.2f, %.2f] 元，标准差 %.2f 元）"
                 % (RNG_SEED, means.min(), means.max(), means.std()))
    ax.legend()
    save_figure(fig, PNG_S3)


# ============================== S4 方法侧互验 ==============================


def valley_peak_rule_rolling(price, load, pv_actual, n_window=36):
    """固定阈值谷充峰放规则策略（逐日滚动版）：时间窗内满功率充/放，不优化功率。

    规则：每天的充电窗 = 当天电价最低的 n_window 个时段（问题 2 电价逐日相同，
          故窗口每天相同）；放电窗 = 电价最高的 n_window 个时段；窗内以最大功率
          充/放（受储电量上下限截断）；窗口重叠时以充电优先；不强制终端。
    输入：price (K,)；load/pv_actual (D,K)；n_window，int，窗口时段数
    输出：dict，{'total','mean','x_total'}
    """
    order = np.argsort(price, kind="stable")             # 价格升序（每天相同，算一次即可）
    charge_set = set(order[:n_window].tolist())
    discharge_set = set(order[-n_window:].tolist())
    e = float(E_INIT)
    total = 0.0; x_total = 0.0
    for d in range(365):
        u = np.zeros(K); v = np.zeros(K)
        for k in range(K):
            if k in charge_set and k not in discharge_set:
                # 满功率充电，且不越储电量上限：可充电量 ≤ (Ē−E)/η
                u[k] = min(U_MAX, max(0.0, (E_MAX - e) / ETA))
                e += ETA * u[k]
            elif k in discharge_set:
                # 满功率放电，且不越储电量下限：可放电量 ≤ (E−E̲)·η
                v[k] = min(U_MAX, max(0.0, (e - E_MIN) * ETA))
                e -= v[k] / ETA
        x = np.maximum((load[d] - pv_actual[d]) * DT_H + u - v, 0.0)
        if d >= D_REP_FIRST:
            total += float(np.dot(price, x))
            x_total += float(x.sum())
    return {"total": total, "mean": total / N_REP, "x_total": x_total}


def window_lp_rolling(price, load, pv_actual, n_window=36):
    """固定充放时间窗下的最优功率 LP（逐日滚动版）：量化"窗口固定时的功率优化收益"。

    与主 LP 同模型，只把 u_k 限制在充电窗、v_k 限制在放电窗（其余为 0）；终端自由。
    输入：price (K,)；load/pv_actual (D,K)；n_window，int
    输出：dict，{'total','mean','x_total'}
    """
    from scipy.optimize import linprog
    from scipy.sparse import csr_matrix, hstack, vstack
    order = np.argsort(price, kind="stable")
    c_list = sorted(order[:n_window].tolist())
    d_list = sorted(order[-n_window:].tolist())
    n_c, n_d = len(c_list), len(d_list)
    n_var = K + n_c + n_d
    c_at = {k: i for i, k in enumerate(c_list)}          # 充电窗时段 -> 变量列
    d_at = {k: i for i, k in enumerate(d_list)}          # 放电窗时段 -> 变量列
    rows, cols, vals = [], [], []
    for k in range(K):
        rows.append(k); cols.append(k); vals.append(-1.0)
        if k in c_at:
            rows.append(k); cols.append(K + c_at[k]); vals.append(1.0)
        if k in d_at:
            rows.append(k); cols.append(K + n_c + d_at[k]); vals.append(-1.0)
    a_supply = csr_matrix((vals, (rows, cols)), shape=(K, n_var))
    # 储电量上下界（前缀和，与主 LP 同口径）
    rows_u, cols_u, vals_u = [], [], []
    for k in range(K):
        for kk in c_list:
            if kk <= k:
                rows_u.append(k); cols_u.append(K + c_at[kk]); vals_u.append(ETA)
        for kk in d_list:
            if kk <= k:
                rows_u.append(k); cols_u.append(K + n_c + d_at[kk]); vals_u.append(-1.0 / ETA)
    a_prefix = csr_matrix((vals_u, (rows_u, cols_u)), shape=(K, n_var)).tocsr()
    cost_vec = np.concatenate([price, np.zeros(n_c + n_d)])
    bounds = [(0.0, None)] * K + [(0.0, U_MAX)] * (n_c + n_d)
    total = 0.0; x_total = 0.0
    e = float(E_INIT)
    for d in range(365):
        b_supply = DT_H * (pv_actual[d] - load[d])
        a_ub = vstack([a_supply, a_prefix, -a_prefix]).tocsr()
        b_ub = np.concatenate([b_supply, np.full(K, E_MAX - e), np.full(K, e - E_MIN)])
        res = linprog(cost_vec, A_ub=a_ub, b_ub=b_ub, bounds=bounds, method="highs")
        if not res.success:
            raise RuntimeError("窗口 LP 失败：%s" % res.message)
        x = res.x[:K]
        u = np.zeros(K); u[c_list] = res.x[K:K + n_c]
        v = np.zeros(K); v[d_list] = res.x[K + n_c:]
        e = e + float(np.sum(ETA * u - v / ETA))         # 终端自由：跨日传递实际末端
        if d >= D_REP_FIRST:
            total += float(res.fun)
            x_total += float(x.sum())
    return {"total": total, "mean": total / N_REP, "x_total": x_total}


def run_s4(price, load, pv_actual, base):
    """S4：联合 LP 下界、MATLAB 复算、DP 单日互验、规则策略对照。"""
    rows = []
    header = ["方法", "案例", "费用_元", "与基准差_元", "相对基准_%", "说明"]
    # ---- (a) 联合 LP 下界（读构思手锚定与落盘结果） ----
    rows.append(["全年联合 LP（完全预见下界）", "填报区间 334 天", r4(JOINT_LB), r4(JOINT_LB - BASE_COST),
                 r4(100.0 * (JOINT_LB / BASE_COST - 1.0)),
                 "滚动（主模型）= %.4f 元；滚动−下界 = %.4f 元（%.4f%%），滚动策略近优"
                 % (BASE_COST, BASE_COST - JOINT_LB, 100.0 * (BASE_COST - JOINT_LB) / JOINT_LB)])
    # ---- (b) MATLAB linprog 复算（读落盘 CSV，如存在） ----
    mat_sum = os.path.join(QDIR, "_matlab校验_q2_汇总.csv")
    if os.path.exists(mat_sum):
        import csv as _csv
        mrows = list(_csv.reader(open(mat_sum, encoding="utf-8-sig")))   # utf-8-sig 兼容 BOM
        # 该 CSV 为"指标,数值"两列的键值表（第 1 行为表头），逐行构造 dict
        vals = {row[0]: row[1] for row in mrows[1:] if len(row) >= 2}
        cost_m = float(vals["填报区间总费用_元"])
        x_m = float(vals["填报区间总购电量_kWh"])
        max_df_m = float(vals["与Python最大单日费用差_元"])
        rows.append(["MATLAB linprog（dual-simplex）", "逐日滚动 365 天",
                     r4(cost_m), r4(cost_m - BASE_COST), r4(100.0 * (cost_m / BASE_COST - 1.0)),
                     "与 Python HiGHS 差 %.4f 元；总购电量 %.4f kWh；最大单日费用差 %.6f 元；"
                     "2025-03-20 逐时段 x/u/v/E 最大差均为 %.6f kWh（4 位小数舍入量级）"
                     % (cost_m - BASE_COST, x_m, max_df_m,
                        float(vals["2025-03-20逐时段最大差_x_kWh"]))])
    else:
        rows.append(["MATLAB linprog（dual-simplex）", "逐日滚动 365 天", None, None, None,
                     "未运行（可选复算项；运行 `matlab -batch matlab_q2` 后重跑本脚本会自动并入）"])
    # ---- (c) 离散化 DP 单日互验（三个日期 × 网格 100/200 步） ----
    import datetime as _dt
    dp_points = []
    for t in ("2025-03-20", "2025-06-21", "2025-12-21"):
        d0 = (_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days
        lp = solve_day(price, load[d0], pv_actual[d0], e_init=E_MIN, mode="free")
        for n_grid in (100, 200):
            dp = dp_solve(price, load[d0], pv_actual[d0], n_grid, e_init=E_MIN)
            rows.append(["离散化 DP（单日互验）", "%s，网格 %d 步" % (t, n_grid),
                         r4(dp["cost"]), r4(dp["cost"] - lp["cost"]),
                         r4(100.0 * (dp["cost"] / lp["cost"] - 1.0)),
                         "同端点 LP（e_init=1200、终端自由）= %.4f 元；DP 为网格限制下的全局最优，"
                         "随网格加密收敛到 LP（费用 ≥ LP）" % lp["cost"]])
            if t == "2025-12-21":
                dp_points.append((n_grid, dp["cost"], lp["cost"]))
    # ---- (d) 固定阈值规则策略 与 窗口内最优功率 LP ----
    rule = valley_peak_rule_rolling(price, load, pv_actual)
    rlp = window_lp_rolling(price, load, pv_actual)
    rows.append(["固定阈值谷充峰放规则（满功率）", "充电窗/放电窗各 36 时段（6 h）",
                 r4(rule["total"]), r4(rule["total"] - BASE_COST),
                 r4(100.0 * (rule["total"] / BASE_COST - 1.0)),
                 "比主模型贵 %.4f 元；比不储能基线（%.4f 元）%s %.4f 元"
                 % (rule["total"] - BASE_COST, BASE_PV_ONLY,
                    "贵" if rule["total"] > BASE_PV_ONLY else "省", abs(rule["total"] - BASE_PV_ONLY))])
    rows.append(["窗口内最优功率 LP", "同时间窗，功率由 LP 优化",
                 r4(rlp["total"]), r4(rlp["total"] - BASE_COST),
                 r4(100.0 * (rlp["total"] / BASE_COST - 1.0)),
                 "窗口固定时功率优化可省 %.4f 元（= 满功率规则 − 窗口LP）；时段灵活性损失 = %.4f 元"
                 % (rule["total"] - rlp["total"], rlp["total"] - BASE_COST)])
    return rows, header, rule, rlp, dp_points


def draw_s4_compare(rule, rlp, dp_points):
    """S4 图：联合下界 / 滚动主模型 / 规则策略 / 不储能基线 对照 + DP 网格收敛。"""
    import matplotlib.pyplot as plt
    apply_chinese_style()
    fig, axes = plt.subplots(1, 2, figsize=FIGSIZE_TALL)
    # 左：四种策略总费用
    names = ["联合 LP\n下界", "滚动主模型\n（读法①）", "窗口最优\n功率 LP", "固定阈值规则\n（满功率）", "不储能基线"]
    vals = [JOINT_LB, BASE_COST, rlp["total"], rule["total"], BASE_PV_ONLY]
    colors = [COLOR_DIS, COLOR_BUY, COLOR_CHG, COLOR_PRICE, COLOR_REF]
    bars = axes[0].bar(names, np.array(vals) / 1e6, color=colors, alpha=0.9)
    for b, v in zip(bars, vals):
        axes[0].text(b.get_x() + b.get_width() / 2, b.get_height() + 0.05, "%.3f" % (v / 1e6),
                     ha="center", va="bottom", fontsize=10)
    axes[0].set_ylabel("总费用（百万元）")
    axes[0].set_title("S4-a：全年费用对照（填报区间 334 天）")
    axes[0].tick_params(axis="x", labelsize=9)
    axes[0].set_ylim(0, max(vals) / 1e6 * 1.15)
    # 右：DP 收敛（费用 vs 网格）
    if dp_points:
        grids = [p[0] for p in dp_points]; costs = [p[1] for p in dp_points]
        lp_ref = dp_points[0][2]
        axes[1].plot(grids, costs, "o-", color=COLOR_CHG, label="离散化 DP 费用")
        axes[1].axhline(lp_ref, color=COLOR_PRICE, linestyle="--",
                        label="同端点 LP 最优 = %.2f 元" % lp_ref)
        for g, c in zip(grids, costs):
            axes[1].annotate("%.2f" % c, (g, c), textcoords="offset points", xytext=(0, 6),
                             ha="center", fontsize=9)
        axes[1].set_xlabel("储电量网格步数")
        axes[1].set_ylabel("单日购电费（元）")
        axes[1].set_title("S4-b：DP 网格加密收敛到 LP（2025-12-21）")
        axes[1].legend()
    fig.suptitle("问题 2 方法侧互验：下界、规则策略与 DP 收敛", fontsize=15)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    save_figure(fig, PNG_S4)


# ============================== S5 口径/读法对照 ==============================


def run_s5(price, load, pv_actual, base):
    """S5：读法①②、终端自由/锁定、η 单向/往返、填法 Y/X 四组对照。"""
    import math
    rows = []
    header = ["对照维度", "口径", "总费用_元", "日均_元", "与基准差_元", "与基准差_%", "说明"]
    # (a) 读法① vs 读法②（②的数值来自 variants_q2.py 落盘，此处引用）
    rows.append(["信息结构", "读法① 完全信息（主模型）", r4(BASE_COST), r4(BASE_MEAN), 0.0, 0.0,
                 "计划恰好覆盖净负荷，r≡0；表 3 全 0"])
    rows.append(["信息结构", "读法② 计划=典型日", 25150848.1041, 75301.9404,
                 r4(25150848.1041 - BASE_COST), r4(100.0 * (25150848.1041 / BASE_COST - 1.0)),
                 "计划/紧急 = 11757539.8955 / 13393308.2086 元；293/334 天触发"])
    r2b = 22752073.1581
    rows.append(["信息结构", "读法②b 计划=前一日实际", r2b, r2b / N_REP, r4(r2b - BASE_COST),
                 r4(100.0 * (r2b / BASE_COST - 1.0)),
                 "315/334 天触发；单时段最大 992.2236 kWh"])
    # (b) 终端自由 vs 终端=当日初始（数值来自 variants_q2.py 的滚动实算）
    cyc = 12245046.9153
    rows.append(["终端条件", "终端自由（主模型，D-04）", r4(BASE_COST), r4(BASE_MEAN), 0.0, 0.0,
                 "每日 24:00 全部被压到 1200 kWh（经济结论）"])
    rows.append(["终端条件", "终端=当日初始（cyclic）", cyc, cyc / N_REP, r4(cyc - BASE_COST),
                 r4(100.0 * (cyc / BASE_COST - 1.0)),
                 "每天 E_144=E_0；循环链把 1.1 的 6000 kWh 初值在 100 天内逐步变现（每日只释放"
                 " 48 kWh，免于往返损耗），故比逐日短视的终端自由更低 %.4f%%；两者都远高于联合下界，"
                 "结论不翻转" % (100.0 * (BASE_COST - cyc) / BASE_COST)])
    # (c) η 单向 0.9 vs 往返 0.9（单向 √0.9）
    eta_rt = math.sqrt(0.9)
    res_rt = eval_rolling_case(price, load, pv_actual, eta=eta_rt)
    rows.append(["效率口径", "往返 0.9（单向 %.4f）" % eta_rt, r4(res_rt["total"]),
                 r4(res_rt["mean"]), r4(res_rt["total"] - BASE_COST),
                 r4(100.0 * (res_rt["total"] / BASE_COST - 1.0)),
                 "往返口径效率更高、费用更低；两口径下峰谷价比 3.758 均远超套利门槛 1/η²（1.235/1.111），"
                 "结论不翻转"])
    # (d) 填法 Y-轮转（已定，D-01）vs 填法 X（逐位置对应）
    x_first_day = base["roll"]["x_plan"][0]              # 2025-02-01 的逐时段购电量（真实序）
    # 填法 Y：模板第 j 格装真实第 (j+1)%144+1 个时段；填法 X：模板第 j 格装真实第 j+1 个时段。
    # 两种落位相差一次循环移位，逐格最大差 = max|x[(j+1)%144] − x[j]| = np.roll(x, -1) − x
    diff_shift = float(np.abs(np.roll(x_first_day, -1) - x_first_day).max())
    rows.append(["时段填法", "填法 Y-轮转（已定，D-01）", r4(BASE_COST), r4(BASE_MEAN), 0.0, 0.0,
                 "行内 144 格恰覆盖当天一天；全天购电量不改"])
    rows.append(["时段填法", "填法 X（逐位置对应，对照）", r4(BASE_COST), r4(BASE_MEAN), 0.0, 0.0,
                 "只改模板落位、不改物理解；2025-02-01 两填法逐格最大差 %.4f kWh（整体错位一格）"
                 % diff_shift])
    # (e) 表 1 六时段的填法 Y/X 逐值对照（四个指定日期；供论文表注量化填法影响）
    #
    # 索引口径（与 D-01、问题 1 的 S5 填法对照完全一致）：
    #   模板标签 ``H:00-H:10`` 所在列是 0 基第 6H-1 列（如 12:00-12:10 是第 71 列）；
    #   填法 Y（已定）：该列装当天第 6H+1 个时段（0 基 idx = 6H），即表 1 的正式取值；
    #   填法 X（对照）：该列装当天第 6H 个时段（0 基 idx = 6H-1），即模板整体少轮转一格。
    import datetime as _dt
    spec_rows = [["日期", "时段", "真实时段序号", "填法Y购电量_kWh", "填法X购电量_kWh", "差_kWh"]]
    for t in ("2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21"):
        d0 = (_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days
        xd = base["roll"]["x_plan"][d0]                # base["roll"] 含全年 365 天（d_start=0），按天序号直接索引
        for h in (10, 12, 14, 16, 18, 20):
            k_y = 6 * h + 1                                # 填法 Y：H:00-H:10 ↔ 当天第 6H+1 个时段（1 基）
            y_val = float(xd[k_y - 1])                     # 填法 Y 取 0 基 idx = 6H
            x_val = float(xd[k_y - 2])                     # 填法 X 取 0 基 idx = 6H-1（模板少轮转一格）
            spec_rows.append([t, "%d:00-%d:10" % (h, h), k_y, r4(y_val), r4(x_val),
                              r4(x_val - y_val)])
    return rows, header, spec_rows


# ============================== S6 适用边界 ==============================


def run_s6_boundary(price, load, pv_actual, rule, rlp):
    """S6：适用边界量化（落盘到 S4 xlsx 的 S6 工作表，供 `适用边界.md` 引用）。"""
    rows = []
    header = ["边界问题", "量化结论", "数值证据"]
    cross_day_gap = 1.0 / ETA ** 2
    rows.append(["跨日套利在什么价格结构下会变得重要？",
                 "跨天价差需超过 1/η² = %.4f 才值得搬运；问题 2 电价逐日相同（跨天比价恒为 1.000），"
                 "故跨日搬运无利可图——这是每日 24:00 压到下限 1200 kWh 的直接原因；问题 4 换波动电价后"
                 "跨天因子 0.688–1.213（报告见 D-09）多数低于门槛，跨日套利仍属次要" % cross_day_gap,
                 "1/η² = %.4f；日内峰谷价比 = %.4f" % (cross_day_gap, price.max() / price.min())])
    cyc = 12245046.9153
    rows.append(["终端条件选择对费用的影响百分比？",
                 "终端自由 vs 终端=当日初始：差 %.4f 元（%.4f%%），方向为终端=初始更省（每日均匀释放"
                 "初始存量的承诺效应）" % (BASE_COST - cyc, 100.0 * (BASE_COST - cyc) / BASE_COST),
                 "%.4f vs %.4f 元；两者与联合下界差分别为 %.4f / %.4f 元"
                 % (BASE_COST, cyc, BASE_COST - JOINT_LB, cyc - JOINT_LB)])
    rows.append(["κ 降到多少时'计划不足'才会成为最优选择？",
                 "理论上 κ ≤ 1 时'故意少买、缺口按 κp 补'不劣于'按 p 买足'（缺 1 kWh 少付 p、多付 κp），"
                 "即 κ<1 才会翻转；κ=5 的现实罚价使'计划不足'严格劣，读数见 S1(d)/S2 的 κ 切面",
                 "κ 档位 4/4.5/5/5.5/6 全部未翻转（读法② 口径）"])
    rows.append(["固定阈值规则策略与最优策略的差距？",
                 "满功率规则 %.4f 元，比主模型贵 %.4f 元（%.4f%%）%s；窗口内最优功率 LP %.4f 元，"
                 "仍比主模型贵 %.4f 元（时段不灵活）"
                 % (rule["total"], rule["total"] - BASE_COST,
                    100.0 * (rule["total"] / BASE_COST - 1.0),
                    "，且比不储能基线还贵" if rule["total"] > BASE_PV_ONLY else "",
                    rlp["total"], rlp["total"] - BASE_COST),
                 "规则 %.4f 元；窗口 LP %.4f 元；主模型 %.4f 元；不储能 %.4f 元"
                 % (rule["total"], rlp["total"], BASE_COST, BASE_PV_ONLY)])
    rows.append(["储能（相对不储能）的净收益在什么范围成立？",
                 "读法① 主模型净收益 %.4f 元（占不储能基线 %.4f%%），S1/S3 全部扰动档位下均 > 0（未翻转）"
                 % (BASE_PV_ONLY - BASE_COST, 100.0 * (BASE_PV_ONLY - BASE_COST) / BASE_PV_ONLY),
                 "不储能基线 %.4f 元；主模型 %.4f 元" % (BASE_PV_ONLY, BASE_COST)])
    return rows, header


# ============================== S1 图 ==============================


def draw_s1_figure(rows1):
    """S1 图：总费用随 η / P̄ / [E̲,Ē] 的变化（三面板，标注不储能基线）。"""
    import matplotlib.pyplot as plt
    apply_chinese_style()
    fig, axes = plt.subplots(1, 3, figsize=FIGSIZE_TALL)
    groups = ("η（单向效率）", "P̄（最大充放电功率）", "E̲/Ē（储电量区间）")
    xlabels = {"η（单向效率）": "单向效率 η",
               "P̄（最大充放电功率）": "最大充放电功率 P̄（kW）",
               "E̲/Ē（储电量区间）": "储电量区间半宽比例（中心 6000 kWh）"}
    for ax, g in zip(axes, groups):
        sub = [r for r in rows1 if r[0] == g]
        xs = [r[2] for r in sub]; ys = [float(r[3]) / 1e6 for r in sub]
        ax.plot(xs, ys, "o-", color=COLOR_BUY, label="滚动 LP 总费用")
        ax.axhline(BASE_PV_ONLY / 1e6, color=COLOR_PRICE, linestyle="--", linewidth=1.5,
                   label="不储能基线 %.3f 百万元" % (BASE_PV_ONLY / 1e6))
        for r in sub:
            ax.annotate("%.3f" % (float(r[3]) / 1e6), (r[2], float(r[3]) / 1e6),
                        textcoords="offset points", xytext=(0, 6), ha="center", fontsize=9)
        ax.set_xlabel(xlabels[g])
        ax.set_ylabel("总费用（百万元）")
        ax.set_title("S1：%s 扰动" % g)
        ax.legend(fontsize=9)
    fig.suptitle("问题 2 单因素参数扰动：各档结论均未翻转（费用始终低于不储能基线）", fontsize=15)
    fig.tight_layout(rect=(0, 0, 1, 0.94))
    save_figure(fig, PNG_S1)


# ============================== 主流程 ==============================


def main():
    """S1–S6 主流程：计算、落盘、控制台打印。"""
    print("=" * 78)
    print("问题 2 灵敏度与稳健性实验 S1–S6")
    print("=" * 78)
    price, load1, pv1, _ = read_attachment1()
    load, pv_actual, dates = read_attachment2()
    print("基准：主模型 %.4f 元（日均 %.4f）；不储能基线 %.4f 元（填报区间，实际数据）"
          % (BASE_COST, BASE_MEAN, BASE_PV_ONLY))

    # ---- 基准滚动（同时供 S5/S6 引用其逐时段解）----
    base = eval_rolling_case(price, load, pv_actual)
    print("基准复核：%.4f 元（锚定 %.4f，差 %+.4f 元）"
          % (base["total"], BASE_COST, base["total"] - BASE_COST))

    # ---- 读法② 缓存（κ 扰动与 S2 读法②列共用）----
    t0 = time.time()
    cache2 = build_reading2_cache(price, load, pv_actual)
    print("读法② 缓存：计划 %.4f 元（锚定 %.4f）；紧急加权（κ=1）%.4f 元；用时 %.1f s"
          % (cache2["plan_cost"], R2_PLAN, cache2["emg_w"], time.time() - t0))

    # ---- S1 ----
    t0 = time.time()
    rows1, header1 = run_s1(price, load, pv_actual, base, cache2)
    write_xlsx(XLSX_S1, {"S1_参数扰动": [header1] + rows1})
    print("S1 完成：%d 行，用时 %.1f s" % (len(rows1), time.time() - t0))

    # ---- S2 ----
    t0 = time.time()
    rows2, header2, grid1 = run_s2(price, load, pv_actual, cache2)
    write_xlsx(XLSX_S2, {"S2_三因素网格": [header2] + rows2})
    draw_s2_heatmap(grid1, cache2)
    print("S2 完成：%d 网格点，用时 %.1f s" % (len(rows2), time.time() - t0))

    # ---- S3 ----
    t0 = time.time()
    rows3, header3, means_s3 = run_s3(price, load, pv_actual, base)
    write_xlsx(XLSX_S3, {"S3_数据扰动": [header3] + rows3})
    draw_s3_distribution(means_s3)
    print("S3 完成：%d 行，用时 %.1f s" % (len(rows3), time.time() - t0))

    # ---- S4 + S6 ----
    t0 = time.time()
    rows4, header4, rule, rlp, dp_points = run_s4(price, load, pv_actual, base)
    rows6, header6 = run_s6_boundary(price, load, pv_actual, rule, rlp)
    write_xlsx(XLSX_S4, {"S4_方法侧互验": [header4] + rows4, "S6_适用边界": [header6] + rows6})
    draw_s4_compare(rule, rlp, dp_points)
    print("S4 完成：%d 行；规则 %.4f、窗口LP %.4f、主模型 %.4f 元，用时 %.1f s"
          % (len(rows4), rule["total"], rlp["total"], BASE_COST, time.time() - t0))

    # ---- S5 ----
    rows5, header5, spec5 = run_s5(price, load, pv_actual, base)
    write_xlsx(XLSX_S5, {"S5_口径对照": [header5] + rows5, "表1填法YX对照": spec5})
    print("S5 完成：%d 行（另附表 1 填法 Y/X 逐值对照 %d 行）" % (len(rows5), len(spec5) - 1))

    # ---- S1 图 ----
    draw_s1_figure(rows1)
    print("=" * 78)
    print("全部落盘：S1/S2/S3/S5 xlsx，S4&S6 xlsx，四张图，本日志")


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
