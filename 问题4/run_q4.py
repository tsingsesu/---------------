"""问题 4 主脚本：把问题 2 / 问题 3 的求解链原样套上附件4 的逐日波动电价。

设计纪律（`问题4/交付清单.md` §一）：本问**不产生任何新的储能模型、平衡约束或模板写出代码**。
允许的唯一输入差异是价格矩阵：问题 2/3 用附件1 电价广播到 365 天，本问用附件4 (365,144)。
具体地：
  * 4-2 = 问题 2 的代码链（`lib/run_days.solve_rolling`）+ 附件4 价格 + D-10 终端余值；
  * 4-3 = 问题 3 的代码链（`lib/run_days.solve_rolling_staged`，多阶段调整）+ 附件4 价格 + D-10 终端余值；
  * 结果文件写出直接复用 `问题2/run_q2.py` 与 `问题3/run_q3.py` 的写表函数
    （把模板路径临时指向附件5/result4-2.xlsx、result4-3.xlsx 后调用），本文件不复制写出逻辑。

口径（全部来自 `口径与假设台账.md`，不得自行更改）：
  * D-01 填法 Y-轮转；D-04 终端自由 + 跨日传递；D-05 单向 η=0.9；
  * D-06 读法① 完全信息（4-2 与问题 2 相同的信息结构，r ≡ 0）；
  * D-07 预报分解主口径 M2（问题 3 链，`lib/forecast.py`）；
  * D-10 终端余值：V_E = 次日最低价 / η（末日无次日取当日最低价 / η，与构思手探针一致）；
        终端余值只改变**决策**，结果文件中报告的费用一律是实际缴费 Σp·x（或 D-12 总费用 J）；
  * D-11 从 2025-01-01 仿真（1 月预热），结果文件只填 2025-02-01 至 12-31 共 334 天；
  * D-12/D-13 问题 3 的调整结算与不可追溯规则（4-3 与问题 3 完全一致）；
  * D-15 统一 4 位小数；D-16.8/9/10 模板扩表、紧急购电写法、`7:0-7:10` 笔误照抄。

关键交叉检查（交付清单 §三）：`--price=att1` 时价格换回附件1 广播、终端余值清零，
应逐格复现问题 2 / 问题 3 的结果——这是"只换一个输入"设计纪律的实测证据。

运行：
  python 问题4/run_q4.py                 # 主模型：附件4 波动电价 + D-10 终端余值 → result4-2/4-3.xlsx
  python 问题4/run_q4.py --price=att1    # 回归对照：附件1 广播 + 终端自由 → 独立的 *_att1回归对照.xlsx
  python 问题4/run_q4.py --quick         # 冒烟：只跑前 41 天（不覆盖交付文件，输出 *_quick.xlsx）
依赖：numpy、scipy、openpyxl（均本机已装）；lib/ 公共模块 + 问题2/问题3 的写表函数。
随机性：无（全流程确定性 LP，不需要随机种子）。
"""

import argparse
import os
import sys

# 把工作区根目录加入模块搜索路径，保证在任意工作目录下都能 from lib... import ...
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import numpy as np

from lib.dataio import (attach_path, read_attachment1, read_attachment2,
                        read_attachment4)
from lib.forecast import load_forecast
from lib.logio import Tee
from lib.run_days import solve_rolling, solve_rolling_staged
from lib.solve_day import terminal_value_definitions
from lib.storage import E_INIT, E_MAX, E_MIN
from lib.timegrid import (DT_H, K, four_hour_blocks, hour_block_to_k,
                          hour_block_to_template_row, k_to_label)
from lib.xlsxio import r4

# 问题 2/问题 3 的写表函数（复用，不复制）——把它们的模板路径临时指向 result4-2/4-3 即可
sys.path.insert(0, os.path.join(ROOT, "问题2"))
sys.path.insert(0, os.path.join(ROOT, "问题3"))
import run_q2                                  # noqa: E402  仅取 fill_result_xlsx（问题 2 的写出代码）
import run_q3                                  # noqa: E402  仅取 fill_result_xlsx（问题 3 的写出代码）

QDIR = os.path.join(ROOT, "问题4")                                   # 问题 4 交付目录
TEMPLATE_42 = attach_path("附件5", "result4-2.xlsx")                 # 模板（只读）
TEMPLATE_43 = attach_path("附件5", "result4-3.xlsx")
RESULT_42 = os.path.join(QDIR, "result4-2.xlsx")                     # 交付结果文件
RESULT_43 = os.path.join(QDIR, "result4-3.xlsx")
AUDIT_CSV_42 = os.path.join(QDIR, "全分辨率明细_4-2.csv")             # 4-2 全分辨率明细
AUDIT_CSV_43 = os.path.join(QDIR, "全分辨率明细_4-3.csv")             # 4-3 全分辨率明细
LOG_PATH = os.path.join(QDIR, "主模型运行日志.txt")                   # 运行日志

D_REP_FIRST = 31                                                     # 填报区间首日 2025-02-01（0 基 31）
N_REP = 334                                                          # 填报区间天数（2025-02-01 至 12-31）
SPEC_HOURS = (10, 12, 14, 16, 18, 20)                                # 题目表 1 的六个指定时段
SPEC_DATES = ("2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21")  # 表 1/2/3 的四个指定日期
STAGES_MAIN = (0, 6, 12, 18)                                         # 4-3 主模型：全调整（同问题 3）

# 构思手独立锚定值（`_phase0/报告20_问题4-2量级体检.txt`，主口径 = D-10 终端余值）——只作对照打印
ANCHOR_42 = {
    "pay": 12815460.2655,         # 填报区间实际缴费 Σp·x，元
    "day_mean": 38369.6415,       # 日均缴费，元
    "residual": 469.0667,         # 期末残值 V_E(末)·E(12.31 24:00)，元
    "true_cost": 12814991.1989,   # 全期真实成本 = 缴费 − 期末残值，元
    "n_hold": 274,                # 24:00 高于下限（持有过夜）的天数 / 334
    "free": 12841054.3634,        # 对照：终端自由（V_E=0）缴费，元
    "q2": 12254765.7161,          # 对照：问题 2（固定电价）主模型，元
}

# 4-2 全分辨率明细 CSV 的列定义（在问题 2 的 13 列基础上追加终端余值单价，供 D-10 审计）
CSV_HEADER_42 = ["日期", "天序号", "时段序号", "时段起", "时段止", "电价_元每kWh",
                 "负载_kW", "光伏_kW", "计划购电量_kWh", "充电量_kWh", "放电量_kWh",
                 "紧急购电量_kWh", "时段末储电量_kWh", "终端余值单价_元每kWh"]

# 4-3 全分辨率明细 CSV 的列定义（在问题 3 的 21 列基础上追加终端余值单价，供 D-10 审计）
CSV_HEADER_43 = ["日期", "天序号", "时段序号", "时段起", "时段止", "电价_元每kWh",
                 "负载_kW", "光伏实际_kW", "光伏预报0_kW", "光伏预报6_kW", "光伏预报12_kW",
                 "光伏预报18_kW", "计划购电量_kWh", "调整购电量_kWh", "欠取量_kWh", "超用量_kWh",
                 "充电量_kWh", "放电量_kWh", "紧急购电量_kWh", "弃光电量_kWh", "时段末储电量_kWh",
                 "终端余值单价_元每kWh"]


def solve_q4_2(price_mat, load, pv_actual, ve_arr):
    """4-2 = 问题 2 代码链 + 附件4 价格 + D-10 终端余值（唯一输入差异是价格矩阵）。

    输入：price_mat，np.ndarray (D,K)，逐日逐时段电价，元/kWh
          load / pv_actual，np.ndarray (D,K)，kW
          ve_arr，np.ndarray (D,)，逐日终端余值单价 V_E，元/kWh（回归对照时传全 0）
    输出：dict，solve_rolling 的返回值（365 天全链）
    """
    # 直接调用问题 2 的滚动求解器；逐日价格矩阵与 v_end 数组是本问唯一新增的输入
    return solve_rolling(price_mat, load, pv_actual, e_init=E_INIT, mode="free",
                         v_end=ve_arr)


def solve_q4_3(price_mat, load, pv_actual, pv_fc144, ve_arr_or_none):
    """4-3 = 问题 3 代码链（多阶段调整）+ 附件4 价格 + D-10 终端余值。

    输入：price_mat，np.ndarray (D,K)，元/kWh；load/pv_actual，np.ndarray (D,K)，kW
          pv_fc144，np.ndarray (D,4,K)，M2 分解后的光伏预报，kW
          ve_arr_or_none，np.ndarray (D,) 或 None，逐日终端余值（回归对照时传 None）
    输出：dict，solve_rolling_staged 的返回值（365 天全链）
    """
    # 多阶段滚动（0:00 计划 + 6/12/18 调整）与问题 3 完全同一条链；终端余值加在 18:00 阶段
    return solve_rolling_staged(price_mat, load, pv_actual, pv_fc144,
                                stages=STAGES_MAIN, d_start=0, d_end=365,
                                v_end_day=ve_arr_or_none)


def write_audit_csv_42(path, dates, price_mat, load, pv_actual, ve_arr, roll):
    """落盘 4-2 全分辨率明细 CSV（334 天 × 144 时段 × 14 列）。

    输入：path，str；dates，list[date]；price_mat (D,K) 元/kWh；load/pv_actual (D,K) kW
          ve_arr (D,) 元/kWh；roll，4-2 链返回值（含 1 月预热期的 365 天）
    输出：str，写入路径
    """
    lines = [",".join(CSV_HEADER_42)]
    for i, d in enumerate(roll["d_index"]):
        for k in range(1, K + 1):
            start_label, end_label = k_to_label(k)          # 真实时段序的（起, 止）钟点
            row = [
                dates[d].isoformat(), str(d + 1), str(k), start_label, end_label,
                "%.4f" % price_mat[d, k - 1],                # 当日该时段电价（附件4 或广播价）
                "%.4f" % load[d, k - 1], "%.4f" % pv_actual[d, k - 1],
                "%.4f" % r4(roll["x_plan"][i, k - 1]),       # 计划购电量，kWh
                "%.4f" % r4(roll["u_chg"][i, k - 1]),        # 充电量，kWh
                "%.4f" % r4(roll["v_dis"][i, k - 1]),        # 放电量，kWh
                "%.4f" % r4(roll["r_emg"][i, k - 1]),        # 紧急购电量，kWh（读法① 恒 0）
                "%.4f" % r4(roll["E_soc"][i, k]),            # 时段 k 末储电量，kWh
                "%.4f" % ve_arr[d],                          # 当日终端余值单价 V_E^d，元/kWh
            ]
            lines.append(",".join(row))
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("\n".join(lines) + "\n")
    return path


def write_audit_csv_43(path, dates, price_mat, load, pv_actual, pv_fc144, ve_arr, roll):
    """落盘 4-3 全分辨率明细 CSV（334 天 × 144 时段 × 22 列）。

    输入：path，str；dates，list[date]；price_mat (D,K)；load/pv_actual (D,K) kW
          pv_fc144 (D,4,K) kW；ve_arr (D,) 或 None；roll，4-3 链返回值（365 天）
    输出：str，写入路径
    """
    lines = [",".join(CSV_HEADER_43)]
    for i, d in enumerate(roll["d_index"]):
        for k in range(1, K + 1):
            start_label, end_label = k_to_label(k)
            ve_d = 0.0 if ve_arr is None else float(ve_arr[d])   # 回归对照模式无终端余值
            row = [
                dates[d].isoformat(), str(d + 1), str(k), start_label, end_label,
                "%.4f" % price_mat[d, k - 1],
                "%.4f" % load[d, k - 1], "%.4f" % pv_actual[d, k - 1],
                "%.4f" % pv_fc144[d, 0, k - 1], "%.4f" % pv_fc144[d, 1, k - 1],
                "%.4f" % pv_fc144[d, 2, k - 1], "%.4f" % pv_fc144[d, 3, k - 1],
                "%.4f" % r4(roll["x_plan"][i, k - 1]),       # 计划购电量 x，kWh
                "%.4f" % r4(roll["y_adj"][i, k - 1]),        # 调整后的最终购电量 y，kWh
                "%.4f" % r4(roll["d_under"][i, k - 1]),      # 欠取量，kWh
                "%.4f" % r4(roll["d_over"][i, k - 1]),       # 超用量，kWh
                "%.4f" % r4(roll["u_chg"][i, k - 1]),        # 充电量，kWh
                "%.4f" % r4(roll["v_dis"][i, k - 1]),        # 放电量，kWh
                "%.4f" % r4(roll["r_emg"][i, k - 1]),        # 紧急购电量，kWh
                "%.4f" % r4(roll["g_curt"][i, k - 1]),       # 弃光电量，kWh
                "%.4f" % r4(roll["E_soc"][i, k]),            # 时段 k 末储电量，kWh
                "%.4f" % ve_d,                               # 当日终端余值单价 V_E^d，元/kWh
            ]
            lines.append(",".join(row))
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("\n".join(lines) + "\n")
    return path


def print_spec_42(dates, roll_rep, ve_arr):
    """打印四个指定日期的表 1 / 表 2（4-2，从链结果直接取值，与落盘文件同源）。

    输入：dates，list[date]；roll_rep，填报区间滚动结果；ve_arr (D,) 元/kWh
    输出：无（仅打印）
    """
    import datetime as _dt
    for t in SPEC_DATES:
        d0 = (_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days
        i = d0 - D_REP_FIRST
        print("  --- %s（天序号 %d）---" % (t, d0 + 1))
        for h in SPEC_HOURS:
            k = hour_block_to_k(h)                       # H:00-H:10 ↔ 第 6H+1 个时段
            j0 = hour_block_to_template_row(h)           # ↔ 模板第 6H 个数据行（0 基）
            print("    %2d:00-%2d:10  真实时段 %3d  模板列 %3d   x = %10.4f kWh"
                  % (h, h, k, j0 + 1, roll_rep["x_plan"][i, k - 1]))
        print("    全天购电量 = %.4f kWh；全天购电费 = %.4f 元"
              % (float(roll_rep["x_plan"][i].sum()), float(roll_rep["cost_day"][i])))
        for start_label, end_label, k_first, k_last in four_hour_blocks():
            print("    %s-%s  充 = %10.4f  放 = %10.4f"
                  % (start_label, end_label,
                     float(np.sum(roll_rep["u_chg"][i, k_first - 1:k_last])),
                     float(np.sum(roll_rep["v_dis"][i, k_first - 1:k_last]))))
        print("    0:00 储电量 = %.4f kWh；24:00 储电量 = %.4f kWh；当日 V_E = %.4f 元/kWh"
              % (roll_rep["E_soc"][i, 0], roll_rep["E_soc"][i, K], ve_arr[d0]))
        print("    表 3（紧急购电）：时间段='%s'；r 总量 = %.6f kWh"
              % ("", float(roll_rep["r_emg"][i].sum())))


def print_spec_43(dates, roll_rep, ve_arr):
    """打印四个指定日期的表 1 / 表 2 / 表 3（4-3）。

    输入：dates，list[date]；roll_rep，填报区间多阶段结果；ve_arr (D,) 或 None
    输出：无（仅打印）
    """
    import datetime as _dt
    for t in SPEC_DATES:
        d0 = (_dt.date.fromisoformat(t) - _dt.date(2025, 1, 1)).days
        i = d0 - D_REP_FIRST
        print("  --- %s（天序号 %d）---" % (t, d0 + 1))
        for h in SPEC_HOURS:
            k = hour_block_to_k(h)
            j0 = hour_block_to_template_row(h)
            print("    %2d:00-%2d:10  真实时段 %3d  模板列 %3d   y = %10.4f kWh"
                  % (h, h, k, j0 + 1, roll_rep["y_adj"][i, k - 1]))
        print("    全天购电量(y) = %.4f kWh；计划购电费 = %.4f 元；全天总费用 J = %.4f 元"
              % (float(roll_rep["y_adj"][i].sum()), float(roll_rep["J_plan"][i]),
                 float(roll_rep["J_day"][i])))
        print("    调整相关费用 = %.4f 元；紧急购电费 = %.4f 元；Σr = %.4f kWh"
              % (float(roll_rep["J_adj"][i]), float(roll_rep["J_emg"][i]),
                 float(roll_rep["r_emg"][i].sum())))
        for start_label, end_label, k_first, k_last in four_hour_blocks():
            print("    %s-%s  充 = %10.4f  放 = %10.4f"
                  % (start_label, end_label,
                     float(np.sum(roll_rep["u_chg"][i, k_first - 1:k_last])),
                     float(np.sum(roll_rep["v_dis"][i, k_first - 1:k_last]))))
        ve_d = 0.0 if ve_arr is None else float(ve_arr[d0])
        print("    0:00 储电量 = %.4f kWh；24:00 储电量 = %.4f kWh；当日 V_E = %.4f 元/kWh"
              % (roll_rep["E_soc"][i, 0], roll_rep["E_soc"][i, K], ve_d))
        # 表 3：紧急购电的时间段与电量文本（用问题 3 的 emergency_text 生成，保证与表格一致）
        from lib.xlsxio import emergency_text
        time_text, energy_text = emergency_text(roll_rep["r_emg"][i], k_to_label)
        print("    表 3（紧急购电）：时间段='%s'，购电量='%s'"
              % (time_text, energy_text if time_text != "" else "0"))


def main():
    """问题 4 主流程：解析参数 → 读数据 → 4-2 链 → 4-3 链 → 落盘 → 打印关键数值与锚定对照。"""
    parser = argparse.ArgumentParser(description="问题 4：附件4 波动电价下的 4-2 与 4-3")
    parser.add_argument("--price", choices=("att4", "att1"), default="att4",
                        help="价格矩阵：att4 = 附件4 逐日波动电价（主口径）；"
                             "att1 = 附件1 广播（回归对照，自动关闭终端余值）")
    parser.add_argument("--quick", action="store_true",
                        help="冒烟模式：仍跑全链，但输出写 *_quick.xlsx，不覆盖交付文件")
    args = parser.parse_args()
    price_mode = args.price

    # 输出路径：主口径写交付名；回归对照与冒烟模式写独立后缀名，绝不动交付文件
    # （CSV 明细同样按模式分离，避免回归运行覆盖交付的附件4 明细）
    if args.quick:
        result_42 = os.path.join(QDIR, "result4-2_quick.xlsx")
        result_43 = os.path.join(QDIR, "result4-3_quick.xlsx")
        csv_42 = os.path.join(QDIR, "全分辨率明细_4-2_quick.csv")
        csv_43 = os.path.join(QDIR, "全分辨率明细_4-3_quick.csv")
    elif price_mode == "att1":
        result_42 = os.path.join(QDIR, "result4-2_att1回归对照.xlsx")
        result_43 = os.path.join(QDIR, "result4-3_att1回归对照.xlsx")
        csv_42 = os.path.join(QDIR, "全分辨率明细_4-2_att1回归对照.csv")
        csv_43 = os.path.join(QDIR, "全分辨率明细_4-3_att1回归对照.csv")
    else:
        result_42, result_43 = RESULT_42, RESULT_43
        csv_42, csv_43 = AUDIT_CSV_42, AUDIT_CSV_43

    print("=" * 78)
    print("问题 4：附件4 逐日波动电价 + D-10 终端余值（4-2 与 4-3 各用问题 2/3 同一条代码链）")
    print("价格模式 = %s；%s" % (price_mode,
          "主口径：V_E^d = 次日最低价/η（D-10）" if price_mode == "att4" else "回归对照：V_E=0（终端自由）"))
    print("=" * 78)

    # ---------------- 1. 读数据（附件1 电价备用 / 附件2 负载与光伏 / 附件3 预报 / 附件4 电价） ----------------
    price1, load1, pv1, _ = read_attachment1()           # 附件1 典型日（回归对照用其电价列）
    load, pv_actual, dates = read_attachment2()          # 附件2 负载（视为已知）与光伏实际（仅评估）
    pv_fc, pv_fc144, dates3, tau_list = load_forecast(method="linear")   # D-07 M2 分解
    pr4, dates4 = read_attachment4()                     # 附件4 逐日逐时段电价 (365,144)
    assert load.shape == (365, K) and pv_actual.shape == (365, K)
    assert pr4.shape == (365, K) and pv_fc144.shape == (365, 4, K)

    # ---------------- 2. 组装唯一输入差异：价格矩阵 + 终端余值 ----------------
    if price_mode == "att4":
        price_mat = pr4                                   # 唯一输入差异：价格矩阵换成附件4
        ve_arr = terminal_value_definitions(pr4)          # D-10：V_E^d = 次日最低价/η（末日取当日）
        ve_for_43 = ve_arr                                # 4-3：终端余值加在 18:00 阶段
        print("附件4 电价：%.4f–%.4f 元/kWh（逐日波动）；V_E 范围 = %.4f–%.4f 元/kWh"
              % (pr4.min(), pr4.max(), ve_arr.min(), ve_arr.max()))
    else:
        price_mat = np.tile(price1, (365, 1))             # 附件1 电价广播到 365 天（问题 2/3 口径）
        ve_arr = np.zeros(365)                            # 回归对照：关闭终端余值
        ve_for_43 = None
        print("回归对照模式：价格 = 附件1 广播（%.4f–%.4f 元/kWh），终端自由（V_E=0）"
              % (price1.min(), price1.max()))
        print("预期：4-2 逐格复现 问题2/result2.xlsx；4-3 逐格复现 问题3/result3.xlsx")

    if args.quick:
        print("冒烟模式：全链仍跑 365 天，输出写独立的 *_quick.xlsx（不覆盖交付文件）")

    # ---------------- 3. 4-2 链：问题 2 的滚动 + 附件4 价格 + D-10 终端余值 ----------------
    roll42_all = solve_q4_2(price_mat, load, pv_actual, ve_arr)
    roll42 = {key: roll42_all[key][D_REP_FIRST:] for key in
              ("d_index", "x_plan", "u_chg", "v_dis", "r_emg", "E_soc",
               "cost_day", "cost_plan_day", "cost_emg_day", "status_day")}
    print("4-2 滚动完成：365 天全部 status=0（最优）=%s" % bool(np.all(roll42_all["status_day"] == 0)))

    # ---------------- 4. 4-3 链：问题 3 的多阶段滚动 + 附件4 价格 + D-10 终端余值 ----------------
    roll43_all = solve_q4_3(price_mat, load, pv_actual, pv_fc144, ve_for_43)
    keys43 = ("d_index", "x_plan", "y_adj", "u_chg", "v_dis", "r_emg", "g_curt", "E_soc",
              "d_under", "d_over", "J_plan", "J_adj", "J_emg", "J_day", "J_alt_day")
    roll43 = {key: roll43_all[key][D_REP_FIRST:] for key in keys43}
    print("4-3 多阶段滚动完成：365 天全部阶段 status 之和为 0 = %s"
          % bool(np.all(roll43_all["status_day"] == 0)))

    if not args.quick:
        # ---------------- 5. 落盘（CSV + 两个 result 文件，写出函数复用问题 2/3 的） ----------------
        write_audit_csv_42(csv_42, dates, price_mat, load, pv_actual, ve_arr, roll42)
        write_audit_csv_43(csv_43, dates, price_mat, load, pv_actual, pv_fc144, ve_arr, roll43)
        print("已落盘：%s（334 天 × 144 时段 × %d 列）" % (csv_42, len(CSV_HEADER_42)))
        print("已落盘：%s（334 天 × 144 时段 × %d 列）" % (csv_43, len(CSV_HEADER_43)))

        # 模板写出复用：把问题 2/3 写表函数使用的模板路径临时指向 result4-2/4-3 的模板
        # （附件5/result4-2.xlsx 与 result2.xlsx 结构逐格相同；两问的"轮转填法"实现同一套）
        run_q2.TEMPLATE_XLSX = TEMPLATE_42
        run_q2.fill_result_xlsx(result_42, dates, roll42, price_mat[0])
        run_q3.TEMPLATE_XLSX = TEMPLATE_43
        run_q3.fill_result_xlsx(result_43, dates, roll43, price_mat[0])
        print("已落盘：%s（由附件5/result4-2.xlsx 模板复制后填写三张表，附件未改动）" % result_42)
        print("已落盘：%s（由附件5/result4-3.xlsx 模板复制后填写四张表，附件未改动）" % result_43)

    # ---------------- 6. 关键数值打印 ----------------
    pay42 = float(roll42["cost_day"].sum())               # 填报区间实际缴费 Σp·x（r≡0），元
    x42 = float(roll42["x_plan"].sum())
    e_end = float(roll42["E_soc"][-1, K])                 # 12-31 24:00 储电量，kWh
    residual = e_end * float(ve_arr[-1])                  # 期末残值（只在期末计一次），元
    n_hold = int(np.sum(roll42["E_soc"][:, K] > E_MIN + 1e-3))   # 24:00 高于下限的天数
    n_min = int(np.sum(np.abs(roll42["E_soc"][:, K] - E_MIN) < 1e-3))  # 24:00 落于下限的天数
    print("-" * 78)
    print("【4-2 全年汇总（填报区间 334 天，2025-02-01 至 12-31）】")
    print("实际缴费 Σp·x = %.4f 元；日均 = %.4f 元" % (pay42, pay42 / N_REP))
    if price_mode == "att4":
        print("锚定对照：缴费 %.4f（差 %+.4f）；日均 %.4f（差 %+.4f）"
              % (ANCHOR_42["pay"], pay42 - ANCHOR_42["pay"],
                 ANCHOR_42["day_mean"], pay42 / N_REP - ANCHOR_42["day_mean"]))
    print("期末（12.31 24:00）储电量 = %.4f kWh；期末残值 = %.4f 元；全期真实成本 = %.4f 元"
          % (e_end, residual, pay42 - residual))
    if price_mode == "att4":
        print("锚定对照：期末残值 %.4f（差 %+.4f）；全期真实成本 %.4f（差 %+.4f）"
              % (ANCHOR_42["residual"], residual - ANCHOR_42["residual"],
                 ANCHOR_42["true_cost"], (pay42 - residual) - ANCHOR_42["true_cost"]))
    print("24:00 落于下限 1200 kWh 的天数 = %d / 334；持有过夜（>下限）的天数 = %d / 334"
          % (n_min, n_hold))
    print("总购电量 Σx = %.4f kWh；总充电量 = %.4f kWh；总放电量 = %.4f kWh；Σr = %.6f kWh"
          % (x42, float(roll42["u_chg"].sum()), float(roll42["v_dis"].sum()),
             float(roll42["r_emg"].sum())))
    if price_mode == "att4":
        print("与问题 2 对照：问题 2（固定电价）= %.4f 元；波动电价缴费增量 = %+.4f 元（%+.4f%%）"
              % (ANCHOR_42["q2"], pay42 - ANCHOR_42["q2"],
                 100.0 * (pay42 - ANCHOR_42["q2"]) / ANCHOR_42["q2"]))

    j43 = float(roll43["J_day"].sum())
    print("-" * 78)
    print("【4-3 全年汇总（填报区间 334 天）】")
    print("计划购电费用 J_plan = %.4f 元；调整相关费用 J_adj = %.4f 元（欠取为负、超用为正）"
          % (float(roll43["J_plan"].sum()), float(roll43["J_adj"].sum())))
    print("紧急购电费用 J_emg = %.4f 元；总费用 J = %.4f 元；日均 = %.4f 元"
          % (float(roll43["J_emg"].sum()), j43, j43 / N_REP))
    e_end43 = float(roll43["E_soc"][-1, K])
    print("对照口径总费用（欠取按全价+违约金）= %.4f 元（差 %+.4f 元）"
          % (float(roll43["J_alt_day"].sum()), float(roll43["J_alt_day"].sum()) - j43))
    print("期末（12.31 24:00）储电量 = %.4f kWh；全期真实成本（J − 期末残值）= %.4f 元"
          % (e_end43, j43 - e_end43 * float(ve_arr[-1])))
    print("总计划量 Σx = %.4f kWh；总最终量 Σy = %.4f kWh；调整量 Σ|y−x| = %.4f kWh"
          % (float(roll43["x_plan"].sum()), float(roll43["y_adj"].sum()),
             float(np.sum(roll43["d_under"] + roll43["d_over"]))))
    print("紧急购电量 Σr = %.4f kWh；触发天数 = %d / 334；单时段最大 r = %.4f kWh"
          % (float(roll43["r_emg"].sum()),
             int(np.sum(roll43["r_emg"].max(axis=1) > 1e-6)), float(roll43["r_emg"].max())))

    print("-" * 78)
    print("【题目表 1/表 2：四个指定日期（4-2；真实时段序 k=1..144；H:00-H:10 ↔ 第 6H+1 个时段）】")
    print_spec_42(dates, roll42, ve_arr)
    print("-" * 78)
    print("【题目表 1/表 2/表 3：四个指定日期（4-3）】")
    print_spec_43(dates, roll43, ve_arr)
    print("=" * 78)
    return roll42, roll43


if __name__ == "__main__":
    # 把标准输出同时写入日志，保证报告中出现的每个数字都能在文件里找到出处
    tee = Tee(LOG_PATH)
    original_stdout = sys.stdout
    sys.stdout = tee
    try:
        main()
    finally:
        sys.stdout = original_stdout
        tee.close()
    print("运行日志已写入：%s" % LOG_PATH)
