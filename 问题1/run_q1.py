"""问题 1 主脚本：典型日单日确定性线性规划（LP）。

流程：读附件1 → 构造 LP → 求解 → 落盘逐时段 CSV 与结果文件 result1.xlsx → 打印关键数值。

口径（全部来自 `口径与假设台账.md`，不得自行更改）：
  * D-01 时段口径与填法 Y-轮转：模板第 i 行（0 基）装当天第 (i+1)%144+1 个时段；
         第 144 行在时间上是当天最早的一段；一切聚合按真实时段序 k=1..144。
  * D-03 主模型取 E_0 = E_144 = 6000 kWh，并同时给出"端点自由"对照。
  * D-05 主模型取单向效率 η=0.9（往返 0.81），并同时给出"往返 0.9"对照。
  * D-15 电量（kWh）与费用（元）统一保留 4 位小数。
  * D-16.1/16.2 不得向电网售电；允许从电网购电给储能充电。

运行：python 问题1/run_q1.py
依赖：numpy、scipy、openpyxl（均本机已装）。
"""

import math
import os
import shutil
import sys

# 把工作区根目录加入模块搜索路径，保证在任意工作目录下都能 from lib... import ...
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import numpy as np
import openpyxl

from lib.dataio import attach_path, read_attachment1, template_rows_for_plan
from lib.solve_day import (baseline_costs, charge_discharge_summary, solve_day,
                           supply_residual)
from lib.storage import E_INIT, E_MAX, E_MIN, ETA, P_MAX, U_MAX
from lib.timegrid import (DT_H, K, four_hour_blocks, hour_block_to_k,
                          hour_block_to_template_row, k_to_label)

QDIR = os.path.join(ROOT, "问题1")                                  # 问题 1 交付目录
TEMPLATE_XLSX = attach_path("附件5", "result1.xlsx")                # 题目模板（只读）
RESULT_XLSX = os.path.join(QDIR, "result1.xlsx")                    # 交付结果文件
CSV_PATH = os.path.join(QDIR, "问题1_逐时段结果.csv")                # 全分辨率中间结果
LOG_PATH = os.path.join(QDIR, "主模型运行日志.txt")                  # 运行日志（数值证据）
ND = 4                                                              # 统一小数位数（D-15）
SPEC_HOURS = (10, 12, 14, 16, 18, 20)                               # 表 1 的六个指定时段
CSV_HEADER = ["时段序号", "时段起", "时段止", "电价_元每kWh", "负载_kW", "光伏_kW",
              "计划购电量_kWh", "充电量_kWh", "放电量_kWh", "储电量_kWh"]


def snap_zero(value, tol=1e-6):
    """把求解器容差量级的极小值归零，仅用于打印，不参与任何计算。

    输入：value，float，待显示的数值
          tol，float，归零阈值
    输出：float，|value| < tol 时返回 0.0，否则原值
    """
    # LP 的残差常在 1e-13 量级，直接打印会得到 "-0.0000" 这种容易被误读的显示
    return 0.0 if abs(value) < tol else float(value)


def r4(values):
    """把数组/标量四舍五入到 4 位小数（D-15 的统一精度口径）。

    输入：values，array_like，任意数值
    输出：np.ndarray 或 float，保留 4 位小数
    """
    # 用 np.round 而非格式化字符串，保证落盘的就是 4 位小数本身而非显示格式
    return np.round(np.asarray(values, dtype=float), ND)


def set_decimal_format(cell):
    """仅当单元格原格式为 General 时才设为 4 位小数，避免覆盖模板自带的数字格式。

    输入：cell，openpyxl 单元格对象
    输出：无（就地修改 cell.number_format）
    """
    # 模板 计划购电量 的 购电量 列自带 '0.0000_ '，不应改动；其余空单元格为 General
    if cell.number_format in ("General", "general"):
        cell.number_format = "0.0000"


class Tee:
    """把控制台输出同时写入日志文件，保证"落盘"不依赖人工复制。

    输入：path，str，日志文件路径
    输出：无（作为上下文管理器使用）
    """

    def __init__(self, path):
        # 控制台默认可能是 GBK 代码页，先切到 utf-8，避免中文与数学符号报编码错
        try:
            self.stdout = sys.stdout
            self.stdout.reconfigure(encoding="utf-8")
        except Exception:
            self.stdout = sys.stdout
        # 以 utf-8 打开日志文件（python 文件本身不写编码注释，编码在 open 时指定）
        self.file = open(path, "w", encoding="utf-8")

    def write(self, text):
        # 同时写日志与回显，二者内容完全一致
        self.stdout.write(text)
        self.file.write(text)

    def flush(self):
        # 刷新两个通道，避免异常退出时日志残缺
        self.stdout.flush()
        self.file.flush()

    def close(self):
        # 关闭日志文件句柄
        self.file.close()


def write_period_csv(path, price, load, pv_plan, res):
    """落盘 10 列的逐时段结果 CSV（真实时间顺序 k=1..144，外部审计的唯一依据）。

    输入：path，str，目标 csv 路径
          price/load/pv_plan，np.ndarray (144,)，电价 元/kWh、负载 kW、光伏 kW
          res，dict，solve_day 的返回值
    输出：str，写入的路径
    """
    # 逐时段按真实时间顺序拼行：时段起止标签由 timegrid 统一生成
    lines = [",".join(CSV_HEADER)]
    for k in range(1, K + 1):
        start_label, end_label = k_to_label(k)          # 该时段的（起, 止）钟点
        row = [
            str(k),                                     # 时段序号，1..144
            start_label,                                # 时段起
            end_label,                                  # 时段止
            "%.4f" % price[k - 1],                      # 电价，元/kWh
            "%.4f" % load[k - 1],                       # 负载功率，kW
            "%.4f" % pv_plan[k - 1],                    # 光伏功率，kW
            "%.4f" % r4(res["x_plan"][k - 1]),          # 计划购电量，kWh
            "%.4f" % r4(res["u_chg"][k - 1]),           # 充电量，kWh
            "%.4f" % r4(res["v_dis"][k - 1]),           # 放电量，kWh
            "%.4f" % r4(res["E_soc"][k]),               # 时段末储电量，kWh
        ]
        lines.append(",".join(row))
    # 写文件时显式指定 utf-8 与换行，保证 Excel 与 pandas 都能正确读取
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("\n".join(lines) + "\n")
    return path


def fill_result_xlsx(path, res):
    """把 LP 解按填法 Y-轮转填入 result1.xlsx（先复制模板，绝不改写附件）。

    输入：path，str，目标 result1.xlsx 路径
          res，dict，solve_day 的返回值
    输出：dict，写入的摘要（表 1 六个时段值、表 2 六块值、0:00/24:00 储电量）
    """
    # 先复制模板：附件目录只读，一切填写都在问题1/ 下的副本上完成（D-16.6）
    shutil.copyfile(TEMPLATE_XLSX, path)
    wb = openpyxl.load_workbook(path)                   # 载入副本（保留模板全部标签）

    # ---------------- 工作表 1：计划购电量（144 个数据行，2 列） ----------------
    ws = wb["计划购电量"]
    # 模板行序的值：第 i 行装当天第 (i+1)%144+1 个时段（dataio 已实现该重排）
    plan_in_template_order = template_rows_for_plan(r4(res["x_plan"]))
    for i in range(K):
        cell = ws.cell(row=2 + i, column=2)             # 第 1 行是表头，数据从第 2 行起
        cell.value = float(plan_in_template_order[i])   # 只写入数值，标签与数字格式原样保留
        # 模板该列数字格式本就是 0.0000，故此处不改写格式，保证结构与模板逐格一致

    # ---------------- 工作表 2：充放电量（6 个数据行，5 列） ----------------
    ws2 = wb["充放电量"]
    blocks = four_hour_blocks()                         # 六个 4 小时块（真实时间顺序）
    assert len(blocks) == 6, "表 2 必须是 6 个 4 小时块"
    for b, (start_label, end_label, k_first, k_last) in enumerate(blocks):
        row = 2 + b                                     # 数据从第 2 行起
        # 块内 36 个时段的充电量之和（k 为 1 基，故切片下标用 k_first-1 : k_last）
        u_block = float(np.sum(res["u_chg"][k_first - 1:k_last]))
        v_block = float(np.sum(res["v_dis"][k_first - 1:k_last]))
        ws2.cell(row=row, column=2).value = float(r4(u_block))   # 充电量，kWh
        ws2.cell(row=row, column=3).value = float(r4(v_block))   # 放电量，kWh
        # 模板该表数字格式为 General；按 D-15 的 4 位小数口径，仅对填入数值的单元格设置格式
        set_decimal_format(ws2.cell(row=row, column=2))
        set_decimal_format(ws2.cell(row=row, column=3))
    # 时刻列（第 4 列）模板已填 '0:00' / '24:00'，此处只做核对、不改写
    hour_col = [ws2.cell(row=2, column=4).value, ws2.cell(row=3, column=4).value]
    assert hour_col == ["0:00", "24:00"], "模板 时刻 列应为 0:00 / 24:00，实际为 %s" % hour_col
    # 储电量列（第 5 列）：第 1 个数据行填 0:00 储电量 E_0，第 2 个数据行填 24:00 储电量 E_K
    ws2.cell(row=2, column=5).value = float(r4(res["E_soc"][0]))          # E_0，kWh
    ws2.cell(row=3, column=5).value = float(r4(res["E_soc"][K]))          # E_K，kWh
    set_decimal_format(ws2.cell(row=2, column=5))
    set_decimal_format(ws2.cell(row=3, column=5))

    wb.save(path)                                       # 保存副本，附件模板不受影响

    # 返回写入摘要，供控制台打印与自检脚本核对
    summary = {
        "plan_template_order": plan_in_template_order,
        "spec_hours": {},
        "blocks": [(b[0], b[1], float(r4(np.sum(res["u_chg"][b[2] - 1:b[3]]))),
                    float(r4(np.sum(res["v_dis"][b[2] - 1:b[3]])))) for b in blocks],
    }
    for h in SPEC_HOURS:
        k = hour_block_to_k(h)                          # 表 1 时段 H:00-H:10 对应的真实时段
        summary["spec_hours"][h] = float(r4(res["x_plan"][k - 1]))
    return summary


def main():
    """问题 1 主流程：求解、落盘、打印。"""
    print("=" * 78)
    print("问题 1：典型日单日确定性线性规划（主模型 E_0 = E_144 = 6000 kWh，单向 η = 0.9）")
    print("=" * 78)
    print("储能参数：额定容量=%.0f kWh，最大充放电功率=%.0f kW，[E_MIN, E_MAX]=[%.0f, %.0f] kWh，"
          "η=%.4f，E_0=%.0f kWh，Δ=%.6f h，单时段最大充/放电量=%.4f kWh"
          % (12000.0, P_MAX, E_MIN, E_MAX, ETA, E_INIT, DT_H, U_MAX))

    # ---------------- 1. 读附件1（典型日） ----------------
    price, load, pv_plan, time_labels = read_attachment1()
    print("附件1 读入：%d 个时段；电价 %.4f–%.4f 元/kWh；负载 %.2f–%.2f kW；光伏 %.2f–%.2f kW"
          % (K, price.min(), price.max(), load.min(), load.max(), pv_plan.min(), pv_plan.max()))
    print("附件1 首末时间标签：%s ... %s（标签为时段右端点，见口径 D-01）"
          % (time_labels[0], time_labels[-1]))

    # ---------------- 2. 求解主模型与三个对照口径 ----------------
    res_main = solve_day(price, load, pv_plan, mode="cyclic")                       # 主模型
    res_free = solve_day(price, load, pv_plan, mode="free")                         # 端点自由对照
    # 往返效率 0.9 口径：单向效率 = sqrt(0.9) = 0.948683（D-05 要求的对照）
    eta_roundtrip = math.sqrt(0.9)
    res_rt_free = solve_day(price, load, pv_plan, mode="free", eta=eta_roundtrip)
    res_rt_cyclic = solve_day(price, load, pv_plan, mode="cyclic", eta=eta_roundtrip)

    # ---------------- 3. 落盘逐时段 CSV 与 result1.xlsx ----------------
    write_period_csv(CSV_PATH, price, load, pv_plan, res_main)
    summary = fill_result_xlsx(RESULT_XLSX, res_main)
    print("已落盘：%s" % CSV_PATH)
    print("已落盘：%s（由附件5 模板复制后填写，附件未改动）" % RESULT_XLSX)

    # ---------------- 4. 关键数值打印 ----------------
    tot_x = float(np.sum(res_main["x_plan"]))                       # 全天购电量，kWh
    cd = charge_discharge_summary(res_main["u_chg"], res_main["v_dis"])   # 充放电汇总
    residual = supply_residual(res_main["x_plan"], load, pv_plan,
                               res_main["u_chg"], res_main["v_dis"])      # 供给约束残差
    base = baseline_costs(price, load, pv_plan)                     # 不储能基线

    print("-" * 78)
    print("【主模型关键数值】")
    print("全天购电量 = %.4f kWh" % tot_x)
    print("全天购电费 = %.4f 元" % res_main["cost"])
    print("充电量合计 = %.4f kWh；放电量合计 = %.4f kWh；比值 = %.6f（理论 η² = %.6f）"
          % (cd["u_total"], cd["v_total"], cd["ratio"], cd["eta_sq"]))
    print("储电量轨迹：最小 %.4f kWh、最大 %.4f kWh；0:00 = %.4f kWh、24:00 = %.4f kWh"
          % (res_main["E_soc"].min(), res_main["E_soc"].max(),
             res_main["E_soc"][0], res_main["E_soc"][K]))
    print("储电量触及上限的时段数 = %d；触及下限的时段数 = %d"
          % (int(np.sum(np.abs(res_main["E_soc"] - E_MAX) < 1e-6)),
             int(np.sum(np.abs(res_main["E_soc"] - E_MIN) < 1e-6))))
    print("购电量为 0 的时段数 = %d / %d；同时充放（u>0 且 v>0）的时段数 = %d"
          % (int(np.sum(res_main["x_plan"] < 1e-6)), K,
             int(np.sum((res_main["u_chg"] > 1e-6) & (res_main["v_dis"] > 1e-6)))))
    print("供给约束最小残差（x + PΔ + v − LΔ − u）= %.6e kWh（≥ 0 表示满足）"
          % float(residual.min()))
    print("弃光电量合计 = %.4f kWh（供给约束的富余量之和）"
          % snap_zero(float(np.sum(res_main["g_curt"]))))
    print("光伏边际价值（对偶）：最小 %.4f 元/kWh、最大 %.4f 元/kWh、均值 %.4f 元/kWh"
          % (res_main["pv_marginal"].min(), res_main["pv_marginal"].max(),
             res_main["pv_marginal"].mean()))

    print("-" * 78)
    print("【题目表 1：指定时段购电量（kWh）】（H:00-H:10 ↔ 当天第 6H+1 个时段 ↔ 模板第 6H 行）")
    for h in SPEC_HOURS:
        k = hour_block_to_k(h)                                      # 真实时段序号
        row0 = hour_block_to_template_row(h)                         # 0 基模板行号
        print("  %2d:00-%2d:10  真实时段 %3d  模板行 %3d  x = %10.4f kWh   "
              "（电价 %.4f 元/kWh，负载 %.1f kW，光伏 %.1f kW）"
              % (h, h, k, row0 + 1, res_main["x_plan"][k - 1],
                 price[k - 1], load[k - 1], pv_plan[k - 1]))
    print("  全天购电量 = %.4f kWh；全天购电费 = %.4f 元" % (tot_x, res_main["cost"]))

    print("-" * 78)
    print("【题目表 2：六个 4 小时块充放电量与首末储电量（kWh）】（按真实时间顺序）")
    for start_label, end_label, u_block, v_block in summary["blocks"]:
        print("  %s-%s  充电量 = %10.4f  放电量 = %10.4f"
              % (start_label, end_label, u_block, v_block))
    print("  0:00 储电量 = %.4f kWh；24:00 储电量 = %.4f kWh"
          % (res_main["E_soc"][0], res_main["E_soc"][K]))

    print("-" * 78)
    print("【基线对照】")
    print("方案 A（全部向电网购电，不储能不弃光）= %.4f 元，购电量 %.4f kWh"
          % (base["cost_all_grid"], base["load_energy"]))
    print("方案 B（光伏优先自用、余电弃掉，不储能）= %.4f 元，购电量 %.4f kWh"
          % (base["cost_pv_only"], base["buy_energy_pv_only"]))
    print("本问优化解 = %.4f 元；相对方案 B 节省 %.4f 元（%.4f%%）"
          % (res_main["cost"], base["cost_pv_only"] - res_main["cost"],
             100.0 * (base["cost_pv_only"] - res_main["cost"]) / base["cost_pv_only"]))
    print("负载电量 = %.4f kWh；光伏电量 = %.4f kWh（占负载 %.2f%%）；有符号净负载电量 = %.4f kWh"
          % (base["load_energy"], base["pv_energy"],
             100.0 * base["pv_energy"] / base["load_energy"], base["net_load_energy"]))
    print("峰谷价比 = %.4f；往返效率 %0.4f ⇒ 套利门槛 1/η² = %.4f ⇒ 套利空间充足"
          % (price.max() / price.min(), ETA ** 2, 1.0 / ETA ** 2))

    print("-" * 78)
    print("【口径对照（D-03 / D-05，必须报出）】")
    print("(a) 端点自由（周期稳态最优）：购电量 = %.4f kWh，购电费 = %.4f 元，24:00 储电量 = %.4f kWh"
          % (float(np.sum(res_free["x_plan"])), res_free["cost"], res_free["E_soc"][K]))
    print("    端点自由相对主模型再省 %.4f 元（%.4f%%）"
          % (res_main["cost"] - res_free["cost"],
             100.0 * (res_main["cost"] - res_free["cost"]) / res_main["cost"]))
    print("(b) 往返效率 0.9（单向 %.4f）端点自由口径：购电量 = %.4f kWh，购电费 = %.4f 元"
          % (eta_roundtrip, float(np.sum(res_rt_free["x_plan"])), res_rt_free["cost"]))
    print("    往返效率 0.9（单向 %.4f）端点锁定口径：购电量 = %.4f kWh，购电费 = %.4f 元"
          % (eta_roundtrip, float(np.sum(res_rt_cyclic["x_plan"])), res_rt_cyclic["cost"]))
    print("    往返口径相对单向口径（同为端点自由）差 %.4f 元（%.4f%%）"
          % (res_free["cost"] - res_rt_free["cost"],
             100.0 * (res_free["cost"] - res_rt_free["cost"]) / res_free["cost"]))
    print("=" * 78)
    print("求解器状态：status=%d，message=%s" % (res_main["status"], res_main["message"]))


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
