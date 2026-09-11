% matlab_q1.m —— 问题 1 的 MATLAB 独立复算（方法侧互验 S4b）
%
% 目的：用 MATLAB 的 linprog（对偶单纯形/内点）独立求解与 Python 完全相同的一个 LP，
%       与 `问题1/run_q1.py` 的 HiGHS 解比对总费用与逐时段购电量，验证"换工具结果一致"。
%
% 模型（与 lib/solve_day.py 完全一致，符号见 `符号表.md`）：
%   min  Σ_k p_k·x_k
%   s.t. -x_k + u_k - v_k ≤ Δ(P_k − L_k)                      （供给不低于负载）
%        E_k = E_0 + Σ_{j≤k}(η·u_j − v_k/η)，E_MIN ≤ E_k ≤ E_MAX（储能状态与边界）
%        Σ_{j≤K}(η·u_j − v_j/η) = 0                           （端点锁定 E_K = E_0）
%        0 ≤ u_k, v_k ≤ P̄·Δ，x_k ≥ 0
%
% 输入（只读）：问题1/问题1_逐时段结果.csv（列：时段序号,时段起,时段止,电价_元每kWh,负载_kW,光伏_kW,...）
% 输出（落盘）：问题1/_matlab校验_q1_汇总.csv      —— 指标,数值（费用/购电量/充放电量/末储电量）
%               问题1/_matlab校验_q1_逐时段.csv    —— 时段序号,计划购电量_kWh,充电量_kWh,放电量_kWh,储电量_kWh
%
% 运行：在 问题1/ 目录下执行  matlab -batch matlab_q1
% 说明：变量名沿用 `符号表.md` §五.3 的对照（x_plan/u_chg/v_dis/E_soc/price/load），
%       数据全部读自落盘 CSV（与 Python 端同源），不依赖任何内存状态。

% ---------- 0. 定位脚本所在目录（保证在任意工作目录下启动都能找到输入输出） ----------
script_dir = fileparts(mfilename('fullpath'));            % 本脚本所在目录，即 问题1/
csv_in  = fullfile(script_dir, '问题1_逐时段结果.csv');    % 输入：逐时段结果（utf-8）
csv_sum = fullfile(script_dir, '_matlab校验_q1_汇总.csv'); % 输出：汇总指标
csv_k   = fullfile(script_dir, '_matlab校验_q1_逐时段.csv');% 输出：逐时段解

% ---------- 1. 读入电价/负载/光伏（与 Python 端同一份落盘数据，只读） ----------
C = readcell(csv_in, 'Encoding', 'UTF-8');                % 读成单元格数组（含中文表头与钟点标签）
price   = cell2mat(C(2:145, 4));                          % 第 4 列：电价，元/kWh，144 个时段
load_kW = cell2mat(C(2:145, 5));                          % 第 5 列：负载功率，kW
pv_kW   = cell2mat(C(2:145, 6));                          % 第 6 列：光伏功率，kW

% ---------- 2. 模型参数（与 lib/storage.py 一致，附录 1 给定） ----------
K      = 144;                                             % 一天 10 分钟时段数
DT_H   = 1/6;                                             % 单时段长度 Δ，h
ETA    = 0.9;                                             % 单向充放电效率 η，无量纲
P_MAX  = 5000;                                            % 最大充放电功率 P̄，kW
E_MIN  = 1200;                                            % 储电量下限，kWh
E_MAX  = 10800;                                           % 储电量上限，kWh
E_INIT = 6000;                                            % 初始储电量 E_0，kWh
U_MAX  = P_MAX * DT_H;                                    % 单时段最大充/放电量 = 833.3333，kWh

% ---------- 3. 构造 LP：变量 z = [x_plan(144); u_chg(144); v_dis(144)]，共 3K 列 ----------
n_var = 3 * K;                                            % 决策变量个数 432
f = [price; zeros(2*K, 1)];                               % 目标系数：只对购电量 x 计电价，u/v 系数为 0

% 3.1 供给不等式：-x_k + u_k - v_k ≤ Δ(P_k − L_k)，共 K 行
%     取值向量必须与行/列索引等长（3K 个），故用 ones 生成而非标量列表
A_sup = sparse([1:K, 1:K, 1:K], [1:K, K+(1:K), 2*K+(1:K)], ...
               [-ones(1,K), ones(1,K), -ones(1,K)], K, n_var);
b_sup = DT_H * (pv_kW - load_kW);                         % 右端：光伏富余为负负载，为正时少购电

% 3.2 储能上下界：前缀和形式，Σ_{j≤k}(η·u_j − v_j/η) 分别 ≤ E_MAX−E_0 与 ≥ 由 −(·) 表达
tri = tril(ones(K));                                      % K×K 下三角全 1 矩阵（前缀和算子）
A_ub_st = sparse([tri * ETA,  tri * (-1/ETA)]);           % 上界行：+η·u − v/η 的逐前缀和
A_lb_st = sparse([tri * (-ETA), tri * (1/ETA)]);          % 下界行：−η·u + v/η，配右端 E_0−E_MIN
b_ub_st = (E_MAX - E_INIT) * ones(K, 1);                  % 上界右端：Ē − E_0（配对陷阱：不可写反）
b_lb_st = (E_INIT - E_MIN) * ones(K, 1);                  % 下界右端：E_0 − E̲
A_st = [A_ub_st; A_lb_st];                                % 2K 行储能不等式（列序 [u | v]）
b_st = [b_ub_st; b_lb_st];                                % 对应右端向量
A_ub = [A_sup; [sparse(2*K, K), A_st]];                   % 左侧给 2K 行储能约束补 x 的 K 列零块
b_ub = [b_sup; b_st];                                     % 完整不等式右端

% 3.3 端点锁定等式：Σ_{j≤K}(η·u_j − v_j/η) = 0，即 E_K = E_0
A_eq = [sparse(1, K), [ETA * ones(1, K), -1/ETA * ones(1, K)]];  % 左侧补 x 的零块，成 1×3K
b_eq = 0;                                                 % 右端为 0

% 3.4 变量上下界：x ≥ 0 无上界；u, v ∈ [0, P̄·Δ]
lb = [zeros(K, 1); zeros(2*K, 1)];                        % 下界全 0
ub = [inf(K, 1); U_MAX * ones(2*K, 1)];                   % 仅 u/v 有上界 U_MAX

% ---------- 4. 求解（dual-simplex，输出尽量少） ----------
options = optimoptions('linprog', 'Algorithm', 'dual-simplex', 'Display', 'off');
[z, J, exitflag] = linprog(f, A_ub, b_ub, A_eq, b_eq, lb, ub, options);
if exitflag ~= 1
    error('matlab_q1: 求解未收敛，exitflag=%d', exitflag);
end

% ---------- 5. 拆分变量并重算储电量轨迹（不信任求解器返回值，独立递推） ----------
x_plan = z(1:K);                                          % 计划购电量，kWh
u_chg  = z(K+1:2*K);                                      % 充电量，kWh
v_dis  = z(2*K+1:3*K);                                    % 放电量，kWh
dE     = ETA * u_chg - v_dis / ETA;                       % 每时段储电量增量 η·u − v/η，kWh
E_soc  = [E_INIT; E_INIT + cumsum(dE)];                   % 储电量轨迹 E_0..E_K，kWh

% ---------- 6. 落盘汇总指标（utf-8，供 Python 端对账） ----------
fid = fopen(csv_sum, 'w', 'n', 'UTF-8');                  % 新建 utf-8 文本文件
fprintf(fid, '指标,数值\n');                               % 表头
fprintf(fid, '费用_元,%.4f\n', J);                         % 全天购电费，元
fprintf(fid, '购电量_kWh,%.4f\n', sum(x_plan));            % 全天购电量，kWh
fprintf(fid, '充电量_kWh,%.4f\n', sum(u_chg));             % 充电量合计，kWh
fprintf(fid, '放电量_kWh,%.4f\n', sum(v_dis));             % 放电量合计，kWh
fprintf(fid, '末储电量_kWh,%.4f\n', E_soc(end));           % 24:00 储电量，kWh
fprintf(fid, '求解状态_flag,%d\n', exitflag);              % linprog 退出标志（1=收敛）
fclose(fid);                                              % 关闭文件

% ---------- 7. 落盘逐时段解（供 Python 端逐点比对） ----------
fid = fopen(csv_k, 'w', 'n', 'UTF-8');                    % 新建 utf-8 文本文件
fprintf(fid, '时段序号,计划购电量_kWh,充电量_kWh,放电量_kWh,储电量_kWh\n');  % 表头
for k = 1:K                                               % 逐时段写 144 行
    fprintf(fid, '%d,%.4f,%.4f,%.4f,%.4f\n', ...
            k, x_plan(k), u_chg(k), v_dis(k), E_soc(k+1));  % E_soc(k+1) 为该时段末储电量
end
fclose(fid);                                              % 关闭文件

% ---------- 8. 控制台打印关键数值（便于直接核对） ----------
fprintf('MATLAB 复算：费用 = %.4f 元；购电量 = %.4f kWh；充电 = %.4f kWh；放电 = %.4f kWh；末储电量 = %.4f kWh\n', ...
        J, sum(x_plan), sum(u_chg), sum(v_dis), E_soc(end));
