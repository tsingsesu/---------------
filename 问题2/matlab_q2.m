% matlab_q2.m —— 问题 2 的 MATLAB 独立复算（方法侧互验 S4b）
%
% 目的：用 MATLAB 的 linprog（对偶单纯形）独立重算"全年 365 天逐日滚动 LP"，与
%       `问题2/run_q2.py` 的 HiGHS 解比对填报区间总费用/购电量与逐时段购电量，
%       验证"换工具结果一致"（口径与 `lib/solve_day.py`、`lib/run_days.py` 完全一致）。
%
% 模型（与 lib/solve_day.py 一致，符号见 `符号表.md`）：
%   min  Σ_k p_k·x_k
%   s.t. -x_k + u_k - v_k ≤ Δ(P_k − L_k)                      （供给不低于负载）
%        E_k = E_0 + Σ_{j≤k}(η·u_j − v_k/η)，E_MIN ≤ E_k ≤ E_MAX（储能状态与边界）
%        终端自由（D-04：不加 E_K = E_0 的等式约束）
%        0 ≤ u_k, v_k ≤ P̄·Δ，x_k ≥ 0
%
% 输入（只读）：附件/附件1.xlsx（电价）、附件/附件2.xlsx（负载与光伏实际）
% 输出（落盘，均在 问题2/ 下）：
%   _matlab校验_q2_汇总.csv    —— 指标,数值（填报区间总费用/总购电量/与落盘CSV的最大单日差）
%   _matlab校验_q2_逐时段.csv  —— 2025-03-20 的 时段序号,计划购电量_kWh,充电量_kWh,放电量_kWh,储电量_kWh
%
% 运行：在 问题2/ 目录下执行  matlab -batch matlab_q2
% 说明：变量名沿用 `符号表.md` §五.3 的对照（x_plan/u_chg/v_dis/E_soc/price/load）。

% ---------- 0. 定位路径（脚本在 问题2/ 下，附件在工作区根的 附件/ 下） ----------
script_dir = fileparts(mfilename('fullpath'));                 % 本脚本所在目录，即 问题2/
root_dir   = fileparts(script_dir);                            % 工作区根目录
attach1    = fullfile(root_dir, '附件', '附件1.xlsx');          % 附件1（电价）
attach2    = fullfile(root_dir, '附件', '附件2.xlsx');          % 附件2（负载/光伏实际）
csv_py     = fullfile(script_dir, '逐日结果.csv');              % Python 落盘的全分辨率明细
csv_sum    = fullfile(script_dir, '_matlab校验_q2_汇总.csv');   % 输出：汇总指标
csv_k      = fullfile(script_dir, '_matlab校验_q2_逐时段.csv'); % 输出：3.20 逐时段解

% ---------- 1. 读入电价（附件1 第 2 列）与负载/光伏（附件2 两张表） ----------
C1 = readcell(attach1, 'Sheet', 1);       % 附件1 第 1 张表
price = cell2mat(C1(2:145, 2));                                % 电价，元/kWh，144 个时段

CL = readcell(attach2, 'Sheet', '小区负载');       % 负载表（366 行 × 145 列）
CP = readcell(attach2, 'Sheet', '光伏发电实际功率'); % 光伏表
load_all = zeros(365, 144);                                    % 负载矩阵 (365,144)，kW
pv_all   = zeros(365, 144);                                    % 光伏矩阵 (365,144)，kW
for d = 1:365
    load_all(d, :) = cell2mat(CL(d + 1, 2:145));               % 第 2..145 列是 144 个时段值
    pv_all(d, :)   = cell2mat(CP(d + 1, 2:145));
end

% ---------- 2. 模型参数（与 lib/storage.py 一致，附录 1 给定） ----------
K      = 144;                                                  % 一天 10 分钟时段数
DT_H   = 1/6;                                                  % 单时段长度 Δ，h
ETA    = 0.9;                                                  % 单向充放电效率 η，无量纲
P_MAX  = 5000;                                                 % 最大充放电功率 P̄，kW
E_MIN  = 1200;                                                 % 储电量下限，kWh
E_MAX  = 10800;                                                % 储电量上限，kWh
E_INIT = 6000;                                                 % 初始储电量 E_0（2025-01-01 0:00），kWh
U_MAX  = P_MAX * DT_H;                                         % 单时段最大充/放电量 = 833.3333，kWh
N_REP  = 334;                                                  % 填报区间天数（2.1–12.31）

% ---------- 3. 预构造与储电量无关的约束块（三天复用同一模式，只是右端随 E 更新） ----------
% 3.1 供给不等式：-x_k + u_k - v_k ≤ Δ(P_k − L_k)，共 K 行
A_sup_row = [1:K, 1:K, 1:K];                                   % 行索引：每个时段一行
A_sup_col = [1:K, K + (1:K), 2*K + (1:K)];                     % 列索引：[x | u | v]
A_sup_val = [-ones(1, K), ones(1, K), -ones(1, K)];            % 系数：-1, +1, -1
A_sup = sparse(A_sup_row, A_sup_col, A_sup_val, K, 3 * K);     % K×3K 稀疏矩阵

% 3.2 储能上下界（前缀和形式）：Σ_{j≤k}(η·u_j − v_j/η) ≤ E_MAX−E_0 与 ≥ E_MIN−E_0
tri = tril(ones(K));                                           % K×K 下三角全 1 矩阵（前缀和算子）
A_st_ub = sparse([zeros(K, K), tri * ETA, tri * (-1 / ETA)]);  % 上界行：x 块补零，+η·u − v/η 的逐前缀和
A_st_lb = sparse([zeros(K, K), tri * (-ETA), tri * (1 / ETA)]);% 下界行：x 块补零，−η·u + v/η
A_fixed = [A_sup; A_st_ub; A_st_lb];                           % 3K×3K 常数部分（不含 E_0）

% 3.3 目标与变量界：只对购电量 x 计电价；u/v 单位时段电量不超过 P̄·Δ
f  = [price; zeros(2 * K, 1)];                                 % 目标系数向量 3K×1
lb = [zeros(K, 1); zeros(K, 1); zeros(K, 1)];                  % 下界：全部非负
ub = [inf(K, 1); U_MAX * ones(K, 1); U_MAX * ones(K, 1)];      % 上界：x 无上界，u/v ≤ 833.3333

options = optimoptions('linprog', 'Algorithm', 'dual-simplex', 'Display', 'off');

% ---------- 4. 全年 365 天逐日滚动（终端自由 + 跨日传递，D-04） ----------
daily_cost = zeros(365, 1);                                    % 各日购电费（元）
daily_x    = zeros(365, 1);                                    % 各日购电量（kWh）
E_state = E_INIT;                                              % 当日 0:00 储电量，kWh
z_0320 = [];                                                   % 保存 2025-03-20 的完整解（供逐时段对照）
for d = 1:365
    b_ub = [DT_H * (pv_all(d, :)' - load_all(d, :)');          % 供给右端：Δ(P−L)
            (E_MAX - E_state) * ones(K, 1);                    % 储电量上界右端（随 E_state 更新）
            (E_state - E_MIN) * ones(K, 1)];                   % 储电量下界右端
    [z, J, flag] = linprog(f, A_fixed, b_ub, [], [], lb, ub, options);
    if flag ~= 1
        error('第 %d 天 LP 未收敛：flag=%d', d, flag);
    end
    x = z(1:K); u = z(K + 1:2 * K); v = z(2 * K + 1:3 * K);   % 拆解变量块
    E_state = E_state + sum(ETA * u - v / ETA);                % 终端自由：动态递推到次日 0:00
    daily_cost(d) = J;                                         % 记录当日费用
    daily_x(d) = sum(x);                                       % 记录当日购电量
    if d == 79                                                 % 2025-03-20 是第 79 天
        z_0320 = z;                                            % 保存该日完整解（供逐时段对照）
    end
end

% ---------- 5. 与 Python 落盘 CSV 对照（填报区间 2.1–12.31 = 第 32..365 天） ----------
Cpy = readcell(csv_py, 'Encoding', 'UTF-8', 'Delimiter', ','); % 读 Python 全分辨率明细
py_cost = zeros(365, 1);                                       % 由 CSV 重算的逐日费用
for d = 32:365
    idx = 2 + (d - 32) * K;                                    % 该日第一行在 CSV 中的行号（含表头偏移）
    blk = cell2mat(Cpy(idx:idx + K - 1, [6, 9]));              % 第 6 列电价、第 9 列计划购电量
    py_cost(d) = sum(blk(:, 1) .* blk(:, 2));                  % 逐时段 p·x 求和 = 当日费用
end
idx_rep = 32:365;                                              % 填报区间天序号
tot_cost_m = sum(daily_cost(idx_rep));                         % MATLAB 填报区间总费用（元）
tot_x_m    = sum(daily_x(idx_rep));                            % MATLAB 填报区间总购电量（kWh）
max_day_df = max(abs(daily_cost(idx_rep) - py_cost(idx_rep))); % 与 Python 的最大单日费用差（元）

% 3.20 逐时段解对比：MATLAB 与 CSV 的 x/u/v/E 最大差
idx_320 = 2 + (79 - 32) * K;                                   % 2025-03-20 在 CSV 中的首行
blk320 = cell2mat(Cpy(idx_320:idx_320 + K - 1, [9, 10, 11, 13]));  % x, u, v, E（时段末）
x_m = z_0320(1:K); u_m = z_0320(K + 1:2 * K); v_m = z_0320(2 * K + 1:3 * K);
E_m = E_MIN + cumsum(ETA * u_m - v_m / ETA);                   % 该日 0:00 储电量 = 1200（跨日传递）
dx320 = max(abs(x_m - blk320(:, 1)));                          % 购电量逐时段最大差（kWh）
du320 = max(abs(u_m - blk320(:, 2)));                          % 充电量逐时段最大差（kWh）
dv320 = max(abs(v_m - blk320(:, 3)));                          % 放电量逐时段最大差（kWh）
dE320 = max(abs(E_m - blk320(:, 4)));                          % 储电量逐时段最大差（kWh）

% ---------- 6. 落盘汇总与逐时段 CSV ----------
fid = fopen(csv_sum, 'w', 'n', 'UTF-8');                % 汇总文件（UTF-8）
fprintf(fid, '指标,数值\n');
fprintf(fid, '填报区间总费用_元,%.4f\n', tot_cost_m);
fprintf(fid, '填报区间总购电量_kWh,%.4f\n', tot_x_m);
fprintf(fid, '与Python最大单日费用差_元,%.6f\n', max_day_df);
fprintf(fid, '2025-03-20逐时段最大差_x_kWh,%.6f\n', dx320);
fprintf(fid, '2025-03-20逐时段最大差_u_kWh,%.6f\n', du320);
fprintf(fid, '2025-03-20逐时段最大差_v_kWh,%.6f\n', dv320);
fprintf(fid, '2025-03-20逐时段最大差_E_kWh,%.6f\n', dE320);
fprintf(fid, '备注,MATLAB linprog dual-simplex 全年 365 天滚动；口径 D-04 终端自由、η=0.9 单向\n');
fclose(fid);

fid = fopen(csv_k, 'w', 'n', 'UTF-8');                  % 3.20 逐时段解文件
fprintf(fid, '时段序号,计划购电量_kWh,充电量_kWh,放电量_kWh,储电量_kWh\n');
for k = 1:K
    fprintf(fid, '%d,%.4f,%.4f,%.4f,%.4f\n', k, x_m(k), u_m(k), v_m(k), E_m(k));
end
fclose(fid);

fprintf('MATLAB 复算完成：填报区间总费用 = %.4f 元；总购电量 = %.4f kWh\n', tot_cost_m, tot_x_m);
fprintf('与 Python 最大单日费用差 = %.6f 元；3.20 逐时段最大差 x/u/v/E = %.6f/%.6f/%.6f/%.6f\n', ...
        max_day_df, dx320, du320, dv320, dE320);
