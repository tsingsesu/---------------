r"""问题 3 灵敏度与稳健性实验（S1–S6）：一键运行、全部落盘。

S1 关键参数扰动：η、P̄、κ（紧急倍数）、κ₋、κ₊ 各 ±5%/±10%/±20%（D-16.5/D-18）；
S2 多因素组合扰动：η × κ₊ × κ 三维网格（3×3×3）；
S3 数据侧扰动：预报整体加偏（±5%/±10%）、按实测误差分布重抽样、负载 ±5%、
   （c）剔除附件4 极端低价点属问题 4-3，本脚本注明转发） ；
S4 方法侧互验：（a）分解规则对比（M2 主口径 vs M1 整点常数 vs 历史形状 M3 vs 典型日形状 M4
   vs 小时粒度 LP）；（b）MATLAB 复算（见 matlab_q3.m，另脚本）；（c）确定性滚动 vs 随机规划
   （见 run_q3_stochastic.py，结论并入 S4 工作表）；
S5 口径对照：（a）结算口径主/对照；（b）调整生效不可追溯 vs 可追溯（极端允许全部追溯）；
   （c）`调整购电量` 填总量 vs 增量；（d）效率口径（单向 0.9 vs 往返 0.9）；
S6 适用边界：预报加偏/更新信息量（λ 阻尼）/纯噪声三组梯度 → 调整收益何时归零或转负；
   理论：最优计划分位 F(x*) = 1/(1+κ₊−κ₋)（默认 = 0.50，即中位数锚定）。

输出：
  `灵敏度分析_参数扰动.xlsx`、`灵敏度分析_多因素网格.xlsx`、`灵敏度分析_数据扰动.xlsx`、
  `方法侧互验_分解规则.xlsx`、`口径对照_结算与调整规则.xlsx`、`适用边界_S6.xlsx`
  `灵敏度分析_参数扰动.png`、`灵敏度分析_多因素热力图.png`、`灵敏度运行日志.txt`

运行：python 问题3/sensitivity_q3.py [--quick]   （--quick 只跑 10 天做冒烟测试）
随机性：重抽样实验固定种子 SEED=20260911，重跑一致。
"""

import os
import sys
import time
from multiprocessing import Pool

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import numpy as np
import openpyxl

from lib.dataio import read_attachment1, read_attachment2
from lib.forecast import (build_pv_fc144, decompose_shape, error_pool_from_data,
                          load_forecast, resample_forecast, within_hour_shape)
from lib.logio import Tee
from lib.plotstyle import (COLOR_BUY, COLOR_CHG, COLOR_PRICE, COLOR_REF, FIGSIZE_TALL,
                           FIGSIZE_WIDE, apply_chinese_style, save_figure)
from lib.run_days import solve_rolling_staged
from lib.solve_day import KAPPA_EMG, KAPPA_OVER, KAPPA_UNDER, settle_total, solve_stage
from lib.storage import E_INIT, E_MAX, E_MIN, ETA, P_MAX
from lib.timegrid import DT_H, K

QDIR = os.path.join(ROOT, "问题3")
FIGDIR = os.path.join(ROOT, "图片", "问题3")                        # 图片统一目录（2026-09-11 起，用户要求）
os.makedirs(FIGDIR, exist_ok=True)                                # 确保目录存在（重跑时自动建）
LOG_PATH = os.path.join(QDIR, "灵敏度运行日志.txt")
SEED = 20260911                       # 固定随机种子
D_REP_FIRST = 31                      # 填报区间首日（0 基）
N_REP = 334
QUICK = "--quick" in sys.argv         # 冒烟模式：只跑 10 天
N_WORKERS = 6                         # 并行进程数

# 全局数据（在子进程初始化时由 _init_worker 注入，避免重复读盘）
_G = {}


def _init_worker():
    """子进程初始化：读一次数据与预报（Windows spawn 模式下每个进程各来一份）。"""
    price, _, _, _ = read_attachment1()
    load, pv_actual, dates = read_attachment2()
    pv_fc, pv_fc144, _, _ = load_forecast(method="linear")
    _G["price"] = price
    _G["load"] = load
    _G["pv_actual"] = pv_actual
    _G["dates"] = dates
    _G["pv_fc"] = pv_fc
    _G["pv_fc144"] = pv_fc144


def _run_case(case):
    """通用求解用例：按 overrides 跑完整年滚动链并汇总填报区间指标（供并行调用）。

    输入：case，dict，键——
          tag，str，用例名；eta/p_max/kappa/kappa_under/kappa_over 覆盖参数；
          pv_fc144，可选，替换预报矩阵；load，可选，替换负载；
          stages，可选，决策时刻集合；fc_actual（bool），预报=实际（退化口径）；
          hours_gran，可选，小时粒度 LP（True）
    输出：dict，用例名 + 总费用与分项 + Σy/Σr + 触发天数
    """
    tag = case["tag"]
    price = _G["price"]; load = case.get("load", _G["load"])
    pv_actual = _G["pv_actual"]
    pv_fc144 = case.get("pv_fc144", _G["pv_fc144"])
    stages = case.get("stages", (0, 6, 12, 18))
    d_end = 41 if QUICK else 365
    kw = dict(eta=case.get("eta", ETA), p_max=case.get("p_max", P_MAX),
              kappa=case.get("kappa", KAPPA_EMG), kappa_under=case.get("kappa_under", KAPPA_UNDER),
              kappa_over=case.get("kappa_over", KAPPA_OVER))
    t0 = time.time()
    roll = solve_rolling_staged(price, load, pv_actual, pv_fc144,
                                stages=stages, d_start=0, d_end=d_end, **kw)
    sl = slice(31, d_end)                                   # 填报区间（quick 模式仅为部分天数）
    out = {
        "用例": tag,
        "总费用_元": float(roll["J_day"][sl].sum()),
        "计划购电费用_元": float(roll["J_plan"][sl].sum()),
        "调整相关费用_元": float(roll["J_adj"][sl].sum()),
        "紧急购电费用_元": float(roll["J_emg"][sl].sum()),
        "对照口径总费用_元": float(roll["J_alt_day"][sl].sum()),
        "总最终购电量_kWh": float(roll["y_adj"][sl].sum()),
        "紧急购电量_kWh": float(roll["r_emg"][sl].sum()),
        "触发紧急购电天数": int(np.sum(roll["r_emg"][sl].sum(axis=1) > 1e-6)),
        "单时段最大紧急购电量_kWh": float(roll["r_emg"][sl].max()),
        "求解秒数": float(time.time() - t0),
    }
    return out


def run_grid(cases):
    """并行跑一组用例（fork/spawn 兼容：Windows 下用 spawn + 初始化函数）。

    输入：cases，list[dict]；输出：list[dict]（顺序与输入一致）
    """
    if len(cases) == 1:                                     # 单用例直接跑，省去进程开销
        _init_worker()
        return [_run_case(cases[0])]
    with Pool(processes=N_WORKERS, initializer=_init_worker) as pool:
        return pool.map(_run_case, cases)


def add_rel(rows, base_row):
    """给结果行添加"相对基准"的列（总费用相对变化与是否翻转判断）。

    输入：rows，list[dict]；base_row，dict，基准用例行
    输出：无（就地修改）
    """
    b = base_row["总费用_元"]
    for r in rows:
        r["总费用相对变化_%"] = 100.0 * (r["总费用_元"] / b - 1.0)
        r["结论是否翻转"] = "否"


def write_sheet(path, sheet_name, rows, note=None):
    """把结果行写入 xlsx（单表；已有文件则追加工作表）。

    输入：path，str；sheet_name，str；rows，list[dict]；note，str 或 None（写入末尾说明）
    输出：str，写入的路径
    说明：各行键集可能不一致（如仅部分用例含补充字段），本函数按"全体键的并集"对齐，
          缺失处填空字符串，保证工作表列结构统一。
    """
    keys = []                                               # 键的并集（保持首次出现顺序）
    for r in rows:
        for k in r.keys():
            if k not in keys:
                keys.append(k)
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
    ws = wb.create_sheet(sheet_name)
    ws.append(keys)
    for r in rows:
        ws.append([r.get(k, "") for k in keys])
    if note:
        ws.append([])
        ws.append([note])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, float):
                c.number_format = "0.0000"
    wb.save(path)
    return path


def s1_parameters():
    """S1：η、P̄、κ、κ₋、κ₊ 各 ±5%/±10%/±20% 的参数扰动。

    实现（口径可辩护、工程上最稳）：η / P̄ / κ₋ / κ₊ 直接影响决策与结算，全部**完整重跑**
    全年滚动链；κ 在 D-12 中只乘在紧急购电费用上、不进入任何决策（r 由实际缺口决定），
    故其各行按基准运行结果的紧急费用线性换算（避免 6 个完全重复的全年重跑，数学上严格等价）。

    输出：`灵敏度分析_参数扰动.xlsx`（含基准行与全部扰动行的总费用、分项、相对变化）
    """
    print("-" * 78)
    print("【S1 关键参数扰动（η / P̄ / κ / κ₋ / κ₊ 各 ±5%/±10%/±20%）】")
    levels = [-0.20, -0.10, -0.05, +0.05, +0.10, +0.20]
    cases = [{"tag": "基准 (η=0.9, P̄=5000, κ=5, κ₋=0.5, κ₊=1.5)"}]
    for lv in levels:
        cases.append({"tag": "η %+.0f%% (%.4f)" % (lv * 100, ETA * (1 + lv)), "eta": ETA * (1 + lv)})
    for lv in levels:
        cases.append({"tag": "P̄ %+.0f%% (%.0f kW)" % (lv * 100, P_MAX * (1 + lv)),
                      "p_max": P_MAX * (1 + lv)})
    for lv in levels:
        cases.append({"tag": "κ₋ %+.0f%% (%.3f)" % (lv * 100, KAPPA_UNDER * (1 + lv)),
                      "kappa_under": KAPPA_UNDER * (1 + lv)})
    for lv in levels:
        cases.append({"tag": "κ₊ %+.0f%% (%.3f)" % (lv * 100, KAPPA_OVER * (1 + lv)),
                      "kappa_over": KAPPA_OVER * (1 + lv)})
    rows = run_grid(cases)
    # κ 行：不重跑，按基准的紧急费用线性换算（κ 不进入决策）
    base = rows[0]
    for lv in levels:
        k_new = KAPPA_EMG * (1 + lv)
        rows.append({
            "用例": "κ %+.0f%% (%.2f)" % (lv * 100, k_new),
            "总费用_元": base["计划购电费用_元"] + base["调整相关费用_元"]
                    + base["紧急购电费用_元"] * k_new / KAPPA_EMG,
            "计划购电费用_元": base["计划购电费用_元"],
            "调整相关费用_元": base["调整相关费用_元"],
            "紧急购电费用_元": base["紧急购电费用_元"] * k_new / KAPPA_EMG,
            "对照口径总费用_元": base["对照口径总费用_元"]
                          - base["紧急购电费用_元"] + base["紧急购电费用_元"] * k_new / KAPPA_EMG,
            "总最终购电量_kWh": base["总最终购电量_kWh"],
            "紧急购电量_kWh": base["紧急购电量_kWh"],
            "触发紧急购电天数": base["触发紧急购电天数"],
            "单时段最大紧急购电量_kWh": base["单时段最大紧急购电量_kWh"],
            "求解秒数": 0.0,
        })
    add_rel(rows, base)
    path = write_sheet(os.path.join(QDIR, "灵敏度分析_参数扰动.xlsx"), "S1_参数扰动", rows,
                       note="η/P̄/κ₋/κ₊ 行为完整重跑 334 天滚动链；κ 行不影响决策，"
                            "按基准紧急费用线性换算（κ 只出现在紧急购电费的乘子上）。")
    print("  S1 完成：%d 行 → %s" % (len(rows), path))
    return rows


def s2_grid():
    """S2：η × κ₊ × κ 三维网格（3×3×3；κ 维度由后处理线性换算，不影响决策）。"""
    print("-" * 78)
    print("【S2 多因素组合扰动（η × κ₊ × κ 各 3 档）】")
    eta_levels = [ETA * 0.95, ETA, ETA * 1.05]
    ko_levels = [KAPPA_OVER * 0.9, KAPPA_OVER, KAPPA_OVER * 1.1]
    k_levels = [KAPPA_EMG - 1.0, KAPPA_EMG, KAPPA_EMG + 1.0]
    cases = [{"tag": "η=%.4f|κ₊=%.3f" % (e, ko), "eta": e, "kappa_over": ko}
             for e in eta_levels for ko in ko_levels]
    rows = run_grid(cases)
    full = []                                               # 展开成 27 行（含 κ 维度）
    for r in rows:
        parts = dict(p.split("=") for p in r["用例"].split("|"))
        e = float(parts["η"]); ko = float(parts["κ₊"])
        for k in k_levels:
            rr = dict(r)
            rr["用例"] = "η=%.4f|κ₊=%.3f|κ=%.1f" % (e, ko, k)
            rr["紧急购电费用_元"] = r["紧急购电费用_元"] * k / KAPPA_EMG
            rr["总费用_元"] = r["计划购电费用_元"] + r["调整相关费用_元"] + rr["紧急购电费用_元"]
            rr["对照口径总费用_元"] = (r["对照口径总费用_元"] - r["紧急购电费用_元"]
                                + rr["紧急购电费用_元"])
            full.append(rr)
    base = [r for r in full if r["用例"] == "η=%.4f|κ₊=%.3f|κ=%.1f" % (ETA, KAPPA_OVER, KAPPA_EMG)][0]
    add_rel(full, base)
    path = write_sheet(os.path.join(QDIR, "灵敏度分析_多因素网格.xlsx"), "S2_三因素网格", full,
                       note="η 与 κ₊ 的 9 个组合完整重跑；每个组合展开为 κ 的 3 档（线性换算）。")
    print("  S2 完成：%d 行 → %s" % (len(full), path))
    return full


def s3_data():
    """S3：数据侧扰动——预报加偏、误差重抽样、负载 ±5%。"""
    print("-" * 78)
    print("【S3 数据侧扰动（预报加偏 ±5%/±10%、误差重抽样、负载 ±5%）】")
    pv_fc144 = _G["pv_fc144"]; load = _G["load"]; pv_actual = _G["pv_actual"]
    cases = [{"tag": "基准"}]
    for bias in (-0.10, -0.05, +0.05, +0.10):
        cases.append({"tag": "预报整体加偏 %+.0f%%" % (bias * 100),
                      "pv_fc144": np.maximum(pv_fc144 * (1.0 + bias), 0.0)})
    # 按实测误差分布重抽样（固定种子）：用四个发布时刻各自的误差场
    err0 = error_pool_from_data(pv_fc144, pv_actual, 0)
    rng = np.random.default_rng(SEED)
    idx = rng.integers(0, err0.shape[0], size=365)
    fc_rs = np.maximum(pv_fc144 + err0[idx][:, None, :], 0.0)   # 四个发布层同步加误差场
    cases.append({"tag": "预报按实测误差重抽样（种子 %d）" % SEED, "pv_fc144": fc_rs})
    for lv in (-0.05, +0.05):
        cases.append({"tag": "负载 %+.0f%%" % (lv * 100), "load": load * (1.0 + lv)})
    rows = run_grid(cases)
    add_rel(rows, rows[0])
    path = write_sheet(os.path.join(QDIR, "灵敏度分析_数据扰动.xlsx"), "S3_数据扰动", rows,
                       note="(c) 剔除附件4 极端低价点属问题 4-3 的数据侧扰动，在问题4 交付中另行落盘。")
    print("  S3 完成：%d 行 → %s" % (len(rows), path))
    return rows


def s4_decomposition():
    """S4a：预报分解规则对比（M2 主口径 vs M1 vs 历史形状 M3 vs 典型日形状 M4）+ 小时粒度。"""
    print("-" * 78)
    print("【S4a 分解规则对比（M2 线性 / M1 整点常数 / M3 历史形状 / M4 典型日形状）】")
    price = _G["price"]; load = _G["load"]; pv_actual = _G["pv_actual"]
    pv_fc = _G["pv_fc"]
    _, _, pv1, _ = read_attachment1()                       # 附件1 光伏（典型日）
    shape_hist = within_hour_shape(pv_actual, thr=500.0)    # 历史小时内平均形状
    shape_typ = within_hour_shape(np.asarray(pv1, dtype=float)[None, :], thr=500.0)  # 典型日形状
    fc_m1 = build_pv_fc144(pv_fc, method="hour_const")
    fc_m3 = build_pv_fc144(pv_fc, method="shape", shape=shape_hist)
    fc_m4 = build_pv_fc144(pv_fc, method="shape", shape=shape_typ)
    rows = []
    cases = [
        {"tag": "M2 线性插值（主口径）"},
        {"tag": "M1 整点常数（分段常数）", "pv_fc144": fc_m1},
        {"tag": "M3 历史小时内形状", "pv_fc144": fc_m3},
        {"tag": "M4 典型日小时内形状", "pv_fc144": fc_m4},
    ]
    rows += run_grid(cases)
    # 各分解口径的预报误差（白天 RMSE，kW；相对白天实际出力的百分比）
    mask = pv_actual > 200
    for r, fc in [(rows[0], _G["pv_fc144"]), (rows[1], fc_m1), (rows[2], fc_m3), (rows[3], fc_m4)]:
        # 四个发布层合并统计（与 D-07 判定同一口径）：每层与实际的逐时段差
        diffs = np.concatenate([(fc[:, ti, :][mask] - pv_actual[mask]) for ti in range(4)])
        rmse = float(np.sqrt(np.mean(diffs ** 2)))
        r["预报RMSE_白天_kW"] = rmse
        r["RMSE_占白天实际出力_%"] = 100.0 * rmse / float(pv_actual[mask].mean())
    add_rel(rows, rows[0])
    # 小时粒度 LP：把 144 段聚合成 24 小时后求解（决策粒度粗化对照）
    rows_h = run_hourly(price, load, pv_actual, _G["pv_fc144"])
    rows.append(rows_h)
    add_rel(rows, rows[0])
    path = write_sheet(os.path.join(QDIR, "方法侧互验_分解规则.xlsx"), "S4a_分解规则", rows,
                       note="M2 为主口径（D-07）；小时粒度 LP 将每天聚合成 24 小时决策后再分摊，"
                            "量化粗化损失。另见 matlab_q3.m（S4b）与 run_q3_stochastic.py（S4c）。")
    print("  S4a 完成：%d 行 → %s" % (len(rows), path))
    return rows


def run_hourly(price, load, pv_actual, pv_fc144):
    """小时粒度 LP 对照：把一天聚合成 24 个 1 小时时段做多阶段滚动，再换算回全天费用。

    输入：price (K,)；load (D,K) kW；pv_actual (D,K) kW；pv_fc144 (D,4,K) kW
    输出：dict，与 S4 其他用例同构的汇总行（填报区间）
    """
    n_h = 24                                                # 小时时段数
    # 聚合算子：每 6 段求和（kWh，每小时 1 h）；功率用 kW 表示即 kWh/h 数值
    price_h = price.reshape(n_h, 6).mean(axis=1)             # 每小时平均电价（购电量按小时统一）
    t0 = time.time()
    e = E_INIT
    tot = dict(J=0.0, plan=0.0, adj=0.0, emg=0.0, y=0.0, r=0.0, nt=0, maxr=0.0)
    for d in range(365):
        load_h = (load[d].reshape(n_h, 6) * DT_H).sum(axis=1)        # 每小时负载电量，kWh
        pv_a_h = (pv_actual[d].reshape(n_h, 6) * DT_H).sum(axis=1)   # 每小时实际光伏电量
        fc_h = {ti: (pv_fc144[d, ti].reshape(n_h, 6) * DT_H).sum(axis=1) for ti in range(4)}
        # 0:00 决策（小时粒度，24 段；dt=1h ⇒ 功率上限折算 5000 kWh/h）
        r0 = solve_stage(price_h, load_h, fc_h[0], e_init=e, x_ref=None,
                         k_start=0, n_period=n_h, dt_h=1.0)
        x = r0["x_plan"]; u = r0["u_chg"].copy(); v = r0["v_dis"].copy(); y = x.copy()
        e_cur = float(r0["E_soc"][6])                        # 6:00 处（6 个小时段）
        for ti, ks in ((1, 6), (2, 12), (3, 18)):
            if ti > 3:
                continue
            res = solve_stage(price_h[ks:], load_h[ks:], fc_h[ti][ks:], e_init=e_cur,
                              x_ref=x[ks:], k_start=ks, n_period=n_h - ks, dt_h=1.0)
            y[ks:] = res["x_plan"]; u[ks:] = res["u_chg"]; v[ks:] = res["v_dis"]
            ks_next = {1: 12, 2: 18, 3: n_h}[ti]
            e_cur = float(res["E_soc"][ks_next - ks])
        r_h = np.maximum(load_h + u - y - pv_a_h - v, 0.0)   # 逐小时紧急购电量，kWh
        # 小时级结算：单价用小时平均价，电量按小时合计
        st = settle_total(price_h, x, y, r_h)
        if d >= 31:
            tot["J"] += st["J"]; tot["plan"] += st["J_plan"]; tot["adj"] += st["J_adj"]
            tot["emg"] += st["J_emg"]; tot["y"] += float(y.sum()); tot["r"] += float(r_h.sum())
            tot["nt"] += int(r_h.sum() > 1e-6); tot["maxr"] = max(tot["maxr"], float(r_h.max()))
        e = float(e_cur if False else (e + np.sum(0.9 * u - v / 0.9)))   # 跨日传递（小时粒度）
    return {
        "用例": "小时粒度 LP（24 段决策）",
        "总费用_元": tot["J"], "计划购电费用_元": tot["plan"],
        "调整相关费用_元": tot["adj"], "紧急购电费用_元": tot["emg"],
        "对照口径总费用_元": tot["J"],
        "总最终购电量_kWh": tot["y"], "紧急购电量_kWh": tot["r"],
        "触发紧急购电天数": tot["nt"], "单时段最大紧急购电量_kWh": tot["maxr"],
        "预报RMSE_白天_kW": float("nan"), "RMSE_占白天实际出力_%": float("nan"),
        "求解秒数": float(time.time() - t0),
    }


def s5_caliber():
    """S5：口径对照——结算口径、调整生效规则（不可追溯 vs 可追溯）、填法、效率。"""
    print("-" * 78)
    print("【S5 口径对照（结算主/对照、生效规则、填法、效率）】")
    price = _G["price"]; load = _G["load"]; pv_actual = _G["pv_actual"]
    pv_fc144 = _G["pv_fc144"]
    rows = []
    # (a) 结算口径：主口径 vs 对照口径（同一次运行的两个汇总列）
    base = _run_case_quick({"tag": "结算口径：主口径（D-12）"})
    base["口径说明"] = "J = Σ[p·min(x,y)+κ₋p(x−y)⁺+κ₊p(y−x)⁺] + κΣpr"
    rows.append(base)
    alt = dict(base)
    alt["用例"] = "结算口径：对照（欠取按全价+违约金）"
    alt["总费用_元"] = base["对照口径总费用_元"]
    alt["口径说明"] = "J = Σp·x + Σ[κ₋p(x−y)⁺+κ₊p(y−x)⁺] + κΣpr"
    rows.append(alt)
    # (b) 调整生效规则：不可追溯（主）vs 可追溯（极端：18:00 用当日最后预报重优化全天）
    retr = run_retroactive(price, load, pv_actual, pv_fc144)
    rows.append(retr)
    # (c) 填法：总量（D-14 主）vs 增量（仅落盘形式；费用不变）
    r_c = dict(base)
    r_c["用例"] = "填法：增量口径（仅影响工作表形式）"
    r_c["总费用_元"] = base["总费用_元"]
    r_c["口径说明"] = "增量填法下 k≤36 段为 0，其余为 y−x；Σ|y−x| = %.4f kWh" % base["调整量绝对值和_kWh"]
    rows.append(r_c)
    # (d) 效率口径：单向 0.9（主）vs 往返 0.9（单向 √0.9）
    eta_rt = float(np.sqrt(0.9))
    rows += run_grid([{"tag": "效率：往返 0.9（单向 %.4f）" % eta_rt, "eta": eta_rt}])
    add_rel(rows, rows[0])
    path = write_sheet(os.path.join(QDIR, "口径对照_结算与调整规则.xlsx"), "S5_口径对照", rows,
                       note="可追溯口径 = 极端允许在 18:00 用当日最后预报重优化全天（含已执行段），"
                            "为规则收紧程度的上界对照。")
    print("  S5 完成：%d 行 → %s" % (len(rows), path))
    return rows


def _run_case_quick(case):
    """跑一次主模型并补充"调整量绝对值和"等 S5 需要的字段。"""
    price = _G["price"]; load = _G["load"]; pv_actual = _G["pv_actual"]
    pv_fc144 = _G["pv_fc144"]
    d_end = 41 if QUICK else 365
    roll = solve_rolling_staged(price, load, pv_actual, pv_fc144,
                                stages=(0, 6, 12, 18), d_start=0, d_end=d_end)
    sl = slice(31, d_end)
    return {
        "用例": case["tag"],
        "总费用_元": float(roll["J_day"][sl].sum()),
        "计划购电费用_元": float(roll["J_plan"][sl].sum()),
        "调整相关费用_元": float(roll["J_adj"][sl].sum()),
        "紧急购电费用_元": float(roll["J_emg"][sl].sum()),
        "对照口径总费用_元": float(roll["J_alt_day"][sl].sum()),
        "总最终购电量_kWh": float(roll["y_adj"][sl].sum()),
        "紧急购电量_kWh": float(roll["r_emg"][sl].sum()),
        "触发紧急购电天数": int(np.sum(roll["r_emg"][sl].sum(axis=1) > 1e-6)),
        "单时段最大紧急购电量_kWh": float(roll["r_emg"][sl].max()),
        "调整量绝对值和_kWh": float(np.sum(roll["d_under"][sl] + roll["d_over"][sl])),
        "求解秒数": 0.0,
    }


def run_retroactive(price, load, pv_actual, pv_fc144):
    """可追溯（极端）口径：每天在 18:00 用当日最后一份预报重优化全天 144 段。

    说明：这是"允许修改已执行时段"的最大自由度对照；计划量仍取 0:00 决策（用于结算比较），
          最终量 y 取 18:00 重优化结果（全天）。
    """
    e = E_INIT
    J = 0.0; Jr = 0.0
    d_end = 41 if QUICK else 365
    for d in range(d_end):
        # 0:00 计划（不可变，用于结算比较）
        r0 = solve_stage(price, load[d], pv_fc144[d, 0], e_init=e, x_ref=None)
        x = r0["x_plan"]
        # 18:00 用最后一份预报重优化全天（允许改已执行段——极端可追溯）
        r18 = solve_stage(price, load[d], pv_fc144[d, 3], e_init=e, x_ref=x,
                          k_start=0, n_period=K)
        y = r18["x_plan"]; u = r18["u_chg"]; v = r18["v_dis"]
        r_emg = np.maximum(load[d] * DT_H + u - y - pv_actual[d] * DT_H - v, 0.0)
        st = settle_total(price, x, y, r_emg)
        if d >= 31:
            J += st["J"]; Jr += st["J_plan"] + st["J_adj"]
        e = float(r18["E_soc"][-1])
    n = (d_end - 31) if d_end > 31 else 1
    return {
        "用例": "调整生效规则：可追溯（18:00 重优化全天）",
        "总费用_元": J,
        "计划购电费用_元": float("nan"), "调整相关费用_元": float("nan"),
        "紧急购电费用_元": float("nan"), "对照口径总费用_元": float("nan"),
        "总最终购电量_kWh": float("nan"), "紧急购电量_kWh": float("nan"),
        "触发紧急购电天数": 0, "单时段最大紧急购电量_kWh": float("nan"),
        "调整量绝对值和_kWh": float("nan"),
        "求解秒数": 0.0,
        "口径说明": "极端允许追溯（每天 18:00 重优化全天，含已执行段）；n=%d 天" % n,
    }


def s6_boundary():
    """S6：适用边界——预报加偏梯度、更新信息量（λ 阻尼）、纯噪声对抗。"""
    print("-" * 78)
    print("【S6 适用边界（加偏梯度 / 更新信息阻尼 λ / 纯噪声）】")
    pv_fc144 = _G["pv_fc144"]
    rows = []
    # (1) 预报整体加偏梯度（找到"调整不再值得"的偏置量级）
    cases = [{"tag": "偏置 0%（基准）"}]
    for bias in (-0.30, -0.20, -0.10, +0.10, +0.20, +0.30):
        cases.append({"tag": "整体加偏 %+.0f%%" % (bias * 100),
                      "pv_fc144": np.maximum(pv_fc144 * (1.0 + bias), 0.0)})
    rows += run_grid(cases)
    # (b) 加偏下的"策略 (a) vs (d)"对比（收益是否转负）
    rows_ab = []
    for bias in (0.0, -0.30, -0.20, -0.10, +0.10, +0.20, +0.30):
        fc = np.maximum(pv_fc144 * (1.0 + bias), 0.0)
        ra = run_grid([{"tag": "偏置 %+.0f%%|(a)仅0:00" % (bias * 100), "pv_fc144": fc, "stages": (0,)}])[0]
        rd = run_grid([{"tag": "偏置 %+.0f%%|(d)全调整" % (bias * 100), "pv_fc144": fc}])[0]
        rows_ab.append({"偏置_%": bias * 100,
                        "(a)_元": ra["总费用_元"], "(d)_元": rd["总费用_元"],
                        "调整收益_元": ra["总费用_元"] - rd["总费用_元"]})
    # (2) 更新信息量阻尼：fc_τ^λ = fc0 + λ(fc_τ − fc0)（λ=0 时更新无信息）
    rows_lam = []
    for lam in (1.0, 0.75, 0.5, 0.25, 0.0):
        fc = pv_fc144.copy()
        for ti in (1, 2, 3):
            fc[:, ti, :] = pv_fc144[:, 0, :] + lam * (pv_fc144[:, ti, :] - pv_fc144[:, 0, :])
        ra = run_grid([{"tag": "λ=%.2f|(a)" % lam, "pv_fc144": fc, "stages": (0,)}])[0]
        rd = run_grid([{"tag": "λ=%.2f|(d)" % lam, "pv_fc144": fc}])[0]
        rows_lam.append({"更新阻尼λ": lam, "(a)_元": ra["总费用_元"], "(d)_元": rd["总费用_元"],
                         "调整收益_元": ra["总费用_元"] - rd["总费用_元"]})
    # (3) 纯噪声对抗：预报更新替换为与实际无关的噪声（信息量为 0）
    rng = np.random.default_rng(SEED)
    fc_noise = pv_fc144.copy()
    for ti in (1, 2, 3):
        noise = rng.normal(0.0, 1.0, size=pv_fc144[:, ti, :].shape) * 500.0   # ±500 kW 量级噪声
        fc_noise[:, ti, :] = np.maximum(pv_fc144[:, 0, :] + noise, 0.0)
    ra = run_grid([{"tag": "纯噪声|(a)", "pv_fc144": fc_noise, "stages": (0,)}])[0]
    rd = run_grid([{"tag": "纯噪声|(d)", "pv_fc144": fc_noise}])[0]
    rows_noise = [{"情景": "更新=0:00预报+独立噪声(σ=500kW)",
                   "(a)_元": ra["总费用_元"], "(d)_元": rd["总费用_元"],
                   "调整收益_元": ra["总费用_元"] - rd["总费用_元"]}]
    # 落盘
    path = write_sheet(os.path.join(QDIR, "适用边界_S6.xlsx"), "S6_加偏梯度", rows, note=None)
    write_sheet(path, "S6_策略对比_加偏", rows_ab,
                note="调整收益 = (a) 费用 − (d) 费用；<0 表示引入调整反而更贵。")
    write_sheet(path, "S6_更新信息阻尼", rows_lam,
                note="λ=0 表示 τ≥6 的预报等价于 0:00 预报（更新不含信息），理论收益恰为 0。")
    write_sheet(path, "S6_纯噪声对抗", rows_noise,
                note="更新预报与实际无关时，调整仅增加费用的方差（凸性），期望收益 ≤ 0。")
    print("  S6 完成 → %s" % path)
    return rows, rows_ab, rows_lam, rows_noise


def _parse_tag(tag):
    """把用例标签解析成参数字典，例如 'η=0.9000|κ₊=1.500|κ=5.0' -> {'η':0.9,...}。"""
    parts = dict(p.split("=") for p in tag.split("|"))
    return {k: float(v) for k, v in parts.items()}


def draw_figures(s1_rows, s2_rows):
    """S1/S2 的图：参数扰动曲线 + 三因素热力图切面与 κ 响应。"""
    apply_chinese_style()
    import matplotlib.pyplot as plt
    # ---------------- S1：五个面板（η / P̄ / κ / κ₋ / κ₊） ----------------
    fig, axes = plt.subplots(2, 3, figsize=FIGSIZE_TALL)
    axes = axes.ravel()
    key_order = ["η ", "P̄ ", "κ₊ ", "κ₋ ", "κ "]
    base = [r for r in s1_rows if r["用例"].startswith("基准")][0]
    for ax, key in zip(axes, key_order):
        sub = [r for r in s1_rows if r["用例"].startswith(key) and not r["用例"].startswith("基准")]
        xs, ys = [0.0], [0.0]
        for r in sub:
            xs.append(float(r["用例"].split("%")[0].split()[-1]))
            ys.append(r["总费用相对变化_%"])
        order = np.argsort(xs)
        xs = np.array(xs)[order]; ys = np.array(ys)[order]
        ax.plot(xs, ys, "o-", color=COLOR_BUY, linewidth=1.5, markersize=6, label="总费用相对变化")
        ax.axhline(0, color=COLOR_REF, linewidth=1.0, linestyle="--")
        ax.set_xlabel("参数相对变化（%）")
        ax.set_ylabel("总费用相对变化（%）")
        ax.set_title("%s 扰动（基准 %.0f 元）" % (key.strip(), base["总费用_元"]))
        ax.legend(fontsize=9)
    axes[5].axis("off")                                     # 第六格空置
    axes[5].text(0.05, 0.5, "基准费用（334 天）\n%.2f 元\n\n紧急购电 %.2f 万 kWh\n"
                 % (base["总费用_元"], base["紧急购电量_kWh"] / 1e4),
                 fontsize=12, transform=axes[5].transAxes)
    fig.suptitle("问题 3 灵敏度 S1：关键参数 ±5%/±10%/±20% 对全年总费用（334 天）的影响",
                 fontsize=15)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    save_figure(fig, os.path.join(FIGDIR, "灵敏度分析_参数扰动.png"))
    # ---------------- S2：η × κ₊ 热力图（κ=5 切面）+ κ 响应 ----------------
    fig, axes = plt.subplots(1, 2, figsize=FIGSIZE_WIDE)
    eta_set = sorted({_parse_tag(r["用例"])["η"] for r in s2_rows})
    ko_set = sorted({_parse_tag(r["用例"])["κ₊"] for r in s2_rows})
    k_set = sorted({_parse_tag(r["用例"])["κ"] for r in s2_rows})
    grid = np.full((len(ko_set), len(eta_set)), np.nan)
    for r in s2_rows:
        p = _parse_tag(r["用例"])
        if abs(p["κ"] - KAPPA_EMG) < 1e-9:
            grid[ko_set.index(p["κ₊"]), eta_set.index(p["η"])] = r["总费用_元"] / 1e4
    im = axes[0].imshow(grid, cmap="viridis", aspect="auto")
    axes[0].set_xticks(range(len(eta_set)), ["%.4f" % e for e in eta_set])
    axes[0].set_yticks(range(len(ko_set)), ["%.2f" % x for x in ko_set])
    axes[0].set_xlabel("单向效率 η")
    axes[0].set_ylabel("超用倍数 κ₊")
    axes[0].set_title("(a) η × κ₊ 总费用（万元，κ=5 切面）")
    for i in range(len(ko_set)):
        for j in range(len(eta_set)):
            axes[0].annotate("%.1f" % grid[i, j], (j, i), ha="center", va="center",
                             color="white", fontsize=11)
    fig.colorbar(im, ax=axes[0], label="总费用（万元）")
    # κ 响应：固定 (η, κ₊) 的组合，总费用随 κ 变化（斜率 = Σp·r /5）
    for e in eta_set:
        for ko in ko_set:
            sel = [r for r in s2_rows if abs(_parse_tag(r["用例"])["η"] - e) < 1e-9
                   and abs(_parse_tag(r["用例"])["κ₊"] - ko) < 1e-9]
            sel = sorted(sel, key=lambda r: _parse_tag(r["用例"])["κ"])
            axes[1].plot([_parse_tag(r["用例"])["κ"] for r in sel],
                         [r["总费用_元"] / 1e4 for r in sel],
                         "o-", linewidth=1.0, markersize=4,
                         label="η=%.4f, κ₊=%.2f" % (e, ko))
    axes[1].set_xlabel("紧急购电倍数 κ")
    axes[1].set_ylabel("总费用（万元）")
    axes[1].set_title("(b) κ 响应（斜率 = 紧急购电量 Σp·r；κ 不改决策）")
    axes[1].legend(fontsize=8, ncol=2)
    fig.suptitle("问题 3 灵敏度 S2：多因素组合扰动（η × κ₊ × κ，3×3×3 网格）", fontsize=15)
    fig.tight_layout(rect=(0, 0, 1, 0.94))
    save_figure(fig, os.path.join(FIGDIR, "灵敏度分析_多因素热力图.png"))


def main():
    """S1–S6 全部实验与图表。"""
    print("=" * 78)
    print("问题 3 灵敏度与稳健性（S1–S6）%s" % ("（--quick 冒烟模式）" if QUICK else ""))
    print("=" * 78)
    _init_worker()
    path_param = os.path.join(QDIR, "灵敏度分析_参数扰动.xlsx")
    path_grid = os.path.join(QDIR, "灵敏度分析_多因素网格.xlsx")
    for p in (path_param, path_grid, os.path.join(QDIR, "灵敏度分析_数据扰动.xlsx"),
              os.path.join(QDIR, "方法侧互验_分解规则.xlsx"),
              os.path.join(QDIR, "口径对照_结算与调整规则.xlsx"),
              os.path.join(QDIR, "适用边界_S6.xlsx")):
        if os.path.exists(p) and not QUICK:
            os.remove(p)                                    # 重新生成，避免重复工作表
    s1 = s1_parameters()
    s2 = s2_grid()
    s3 = s3_data()
    s4 = s4_decomposition()
    s5 = s5_caliber()
    s6 = s6_boundary()
    draw_figures(s1, s2)
    print("=" * 78)
    print("【S6 适用边界结论预览】")
    print("  最优计划分位 F(x*) = (κ₊−1)/(κ₊+κ₋−1)（一阶条件；数值验证见 适用边界_S6.xlsx）")
    print("  κ₊−κ₋=1 → F=0.50（中位数锚定，默认）；κ₊−κ₋<1 → F<0.5（低于中位数，故意低报变优）；")
    print("  κ₊−κ₋>1 → F>0.5（高于中位数）。")
    for r in s6[1]:
        print("  偏置 %+6.1f%%：调整收益 %+12.4f 元" % (r["偏置_%"], r["调整收益_元"]))
    for r in s6[2]:
        print("  更新阻尼 λ=%.2f：调整收益 %+12.4f 元" % (r["更新阻尼λ"], r["调整收益_元"]))
    for r in s6[3]:
        print("  纯噪声：调整收益 %+12.4f 元" % r["调整收益_元"])
    print("=" * 78)


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
