% matlab_q3.m —— 问题 3 的 MATLAB 独立复算（方法侧互验 S4b）
%
% 目的：用 MATLAB 的 linprog（对偶单纯形）独立重算"多阶段滚动 LP"的关键数值：
%       1) 整点预报 → 144 段的线性插值分解（D-07 主口径 M2）与 Python 一致；
%       2) 2025-03-20 当天的 0:00 计划 + 6/12/18 调整四阶段解与结算费用，
%          与 `问题3/主模型运行日志.txt` / `策略对比.xlsx` 的锚定值对照；
%       3) 退化自检：把预报换成实际光伏，y≡x、r≡0、费用 = 问题 2 的当日费用。
%
% 模型（与 lib/solve_day.solve_stage / lib/run_days.solve_rolling_staged 完全一致）：
%   阶段 0（0:00，全价）：min Σ p·y
%     s.t. y_k + P̂_k·Δ + v_k ≥ L_k·Δ + u_k；0 ≤ u,v ≤ P̄Δ；E 递推（起点 e0）、上下限、终端自由
%   阶段 τ（6/12/18，偏差结算）：min Σ (1−κ₋)p·y + (κ₊+κ₋−1)p·t
%     s.t. 同上，再增 t_k ≥ y_k − x_k（x 为 0:00 计划，D-12 结算基准）
%   结算（实际值代入）：J = Σ[p·min(x,y) + κ₋p(x−y)⁺ + κ₊p(y−x)⁺] + κΣp·r
%
% 输入（只读）：附件/附件1.xlsx、附件2.xlsx、附件3.xlsx
% 输出（落盘，均在 问题3/ 下）：
%   _matlab校验_q3_汇总.csv    —— 指标,数值（分解一致性 / 2025-03-20 费用与分项 / 退化检验）
%   _matlab校验_q3_逐时段.csv  —— 2025-03-20 的 时段序号,计划x,最终y,充电u,放电v
%
% 运行：matlab -batch matlab_q3（在 问题3/ 目录下，或给出脚本绝对路径）

% ---------- 0. 定位路径 ----------
script_dir = fileparts(mfilename('fullpath'));                 % 本脚本所在目录，即 问题3/
root_dir   = fileparts(script_dir);                            % 工作区根目录
attach1    = fullfile(root_dir, '附件', '附件1.xlsx');
attach2    = fullfile(root_dir, '附件', '附件2.xlsx');
attach3    = fullfile(root_dir, '附件', '附件3.xlsx');
csv_sum    = fullfile(script_dir, '_matlab校验_q3_汇总.csv');
csv_k      = fullfile(script_dir, '_matlab校验_q3_逐时段.csv');

% ---------- 1. 读入电价 / 负载 / 光伏实际 ----------
C1 = readcell(attach1, 'Sheet', 1);
price = cell2mat(C1(2:145, 2));                                % 电价（144，1），元/kWh

CL = readcell(attach2, 'Sheet', '小区负载');
CP = readcell(attach2, 'Sheet', '光伏发电实际功率');
load_all = zeros(365, 144); pv_all = zeros(365, 144);
for d = 1:365
    load_all(d, :) = cell2mat(CL(d + 1, 2:145));
    pv_all(d, :)   = cell2mat(CP(d + 1, 2:145));
end

% ---------- 2. 读入预报并做 M2 线性插值分解 ----------
C3 = readcell(attach3, 'Sheet', 1);                            % 第 1 行表头，其后 4×365 行
FC = zeros(365, 4, 24);                                        % (天, 发布时刻, 提前期)
for i = 1:size(C3, 1) - 1
    d = floor((i - 1) / 4) + 1;                                % 天序号（1 基）
    t = mod(i - 1, 4) + 1;                                     % 发布时刻序号（0/6/12/18 -> 1..4）
    FC(d, t, :) = cell2mat(C3(i + 1, 3:26));                   % 第 3..26 列 = 预报 1..24 小时
end
tau_list = [0 6 12 18];
K = 144; DT = 1/6; ETA = 0.9; PMAX = 5000; UMAX = PMAX * DT;
EMIN = 1200; EMAX = 10800; EINIT = 6000;
KU = 0.5; KO = 1.5; KE = 5;                                    % κ₋, κ₊, κ
D320 = 79;                                                     % 2025-03-20（1 基天序号）
E320 = 1200;                                                   % 当日 0:00 储电量（滚动链稳态值）

% M2 分解（与 lib/forecast.release_nodes + decompose_linear 同规则）
FC144 = zeros(365, 4, 144);
for d = 1:365
    for ti = 1:4
        tau = tau_list(ti);
        nodes = nan(1, 25);                                    % 0:00..24:00 结点
        if d > 1
            nodes(1) = FC(d - 1, 1, 24);                       % 0:00 结点取前一日 0:00 预报第 24 值
        else
            nodes(1) = 0;                                      % 2025-01-01 取 0（D-07）
        end
        for h = 1:24
            i_star = 0;                                        % 覆盖 h 的最新发布（τ'<h 且 τ'≤τ）
            for i2 = 1:4
                if h > tau_list(i2) && tau_list(i2) <= tau
                    i_star = i2;
                end
            end
            nodes(h + 1) = FC(d, i_star, h - tau_list(i_star));
        end
        for k = 1:144
            tt = k / 6;                                        % 段右端点钟点
            m = ceil(tt - 1e-12);                              % 所在结点区间右端号（1..24）
            if m < 1, m = 1; end
            frac = m - tt;                                     % 距右端结点的比例
            FC144(d, ti, k) = nodes(m) * frac + nodes(m + 1) * (1 - frac);
        end
    end
end

% ---------- 3. 单阶段 LP 求解器（与 solve_stage 同构） ----------
% 变量顺序 [y(K) | u(K) | v(K) | t(K)]；x_ref 为 [] 时目标全价、无 t 变量
function [y, u, v, Esoc] = solve_stage_m(p, L, pv, e0, x_ref, K, DT, ETA, UMAX, EMIN, EMAX, KU, KO)
    n = K;
    has_ref = ~isempty(x_ref);
    nv = 3 * n + (n * has_ref);
    % 目标
    if has_ref
        c = [ (1 - KU) * p(:); zeros(2 * n, 1); (KO + KU - 1) * p(:) ];
    else
        c = [ p(:); zeros(2 * n, 1) ];
    end
    % 约束行：供给 K + 储能 2K + （超用 K）
    nrow = 3 * n + (n * has_ref);
    rows = []; cols = []; vals = []; b = zeros(nrow, 1);
    ar = (1:n)';
    % 供给：−y + u − v ≤ Δ(P−L)
    rows = [rows; ar; ar; ar];
    cols = [cols; ar; n + ar; 2 * n + ar];
    vals = [vals; -ones(n,1); ones(n,1); -ones(n,1)];
    b(1:n) = (pv(:) - L(:)) * DT;
    % 储能前缀和（下三角），行 n+1..3n
    [tr, tc] = find(tril(ones(n)));
    rows = [rows; n + tr; n + tr; 2 * n + tr; 2 * n + tr];
    cols = [cols; n + tc; 2 * n + tc; n + tc; 2 * n + tc];
    vals = [vals; ETA * ones(length(tr),1); -ones(length(tr),1)/ETA; ...
            -ETA * ones(length(tr),1); ones(length(tr),1)/ETA];
    b(n + 1:2 * n) = EMAX - e0;
    b(2 * n + 1:3 * n) = e0 - EMIN;
    % 超用约束：y − t ≤ x_ref（第 3n+1..4n 行）
    if has_ref
        rows = [rows; 3 * n + ar; 3 * n + ar];
        cols = [cols; ar; 3 * n + ar];
        vals = [vals; ones(n,1); -ones(n,1)];
        b(3 * n + 1:4 * n) = x_ref(:);
    end
    A = sparse(rows, cols, vals, nrow, nv);
    lb = zeros(nv, 1); ub = inf(nv, 1);
    ub(n + 1:3 * n) = UMAX;
    options = optimoptions('linprog', 'Algorithm', 'dual-simplex', 'Display', 'off');
    [z, ~, flag] = linprog(c, A, b, [], [], lb, ub, options);
    if flag ~= 1
        error('阶段 LP 未收敛：flag=%d', flag);
    end
    y = z(1:n); u = z(n + 1:2 * n); v = z(2 * n + 1:3 * n);
    Esoc = e0 + cumsum(ETA * u - v / ETA);
end

% ---------- 4. 2025-03-20 的四阶段多阶段解 ----------
% 阶段 0：0:00 计划（全价）
[x0, u0, v0, E0] = solve_stage_m(price, load_all(D320,:), squeeze(FC144(D320,1,:)), ...
                                 E320, [], K, DT, ETA, UMAX, EMIN, EMAX, KU, KO);
x = x0; y = x0; u = u0; v = v0;
ks = [36 72 108];                                              % 6/12/18 覆盖的 0 基起点
e_cur = E0(ks(1));                                             % 6:00 处储电量
for ti = 2:4
    k0 = ks(ti - 1);
    [yy, uu, vv, Ee] = solve_stage_m(price(k0 + 1:end), load_all(D320, k0 + 1:end), ...
        squeeze(FC144(D320, ti, k0 + 1:end)), e_cur, x(k0 + 1:end), K - k0, DT, ETA, ...
        UMAX, EMIN, EMAX, KU, KO);
    y(k0 + 1:end) = yy; u(k0 + 1:end) = uu; v(k0 + 1:end) = vv;
    if ti < 4
        e_cur = Ee(ks(ti) - k0);                               % 下一阶段起点（已锁定段末）
    end
end
% 结算（实际值）
r_emg = max(load_all(D320,:)' * DT + u - y - pv_all(D320,:)' * DT - v, 0);
J_plan = sum(price .* min(x, y));
J_adj = sum(KU * price .* max(x - y, 0) + KO * price .* max(y - x, 0));
J_emg = KE * sum(price .* r_emg);
J_320 = J_plan + J_adj + J_emg;

% ---------- 5. 退化自检：预报=实际（同一天） ----------
[xd, ud, vd, ~] = solve_stage_m(price, load_all(D320,:), pv_all(D320,:)', ...
                                E320, [], K, DT, ETA, UMAX, EMIN, EMAX, KU, KO);
rd = max(load_all(D320,:)' * DT + ud - xd - pv_all(D320,:)' * DT - vd, 0);
J_deg = sum(price .* xd) + KE * sum(price .* rd);              % 完全信息：J = Σp·x、r=0

% ---------- 6. 与 Python 落盘对照 ----------
% Python 锚定（问题3/主模型运行日志.txt）：3.20 全天 J = 50206.2807 元；
% 退化（预报=实际）全年 = 12 254 765.7161 元，故当日退化费用应与问题 2 的 3.20 一致。
csv_q2 = fullfile(root_dir, '问题2', '逐日结果.csv');           % 问题 2 全分辨率明细
J_q2_320 = NaN;
if exist(csv_q2, 'file')
    T = readcell(csv_q2, 'Encoding', 'UTF-8', 'Delimiter', ',');
    % 3.20 的行：第 2 + (79-32)*144 行起 144 行（当日逐时段 p·x 求和）
    i0 = 2 + (79 - 32) * K;
    blk = cell2mat(T(i0:i0 + K - 1, [6, 9]));                  % 第 6 列电价、第 9 列计划购电量
    J_q2_320 = sum(blk(:, 1) .* blk(:, 2));
end

% ---------- 7. 落盘 ----------
fid = fopen(csv_sum, 'w', 'n', 'UTF-8');
fprintf(fid, '指标,数值\n');
fprintf(fid, 'MATLAB_2025-03-20_总费用_元,%.4f\n', J_320);
fprintf(fid, 'MATLAB_2025-03-20_计划购电费_元,%.4f\n', J_plan);
fprintf(fid, 'MATLAB_2025-03-20_调整相关费_元,%.4f\n', J_adj);
fprintf(fid, 'MATLAB_2025-03-20_紧急购电费_元,%.4f\n', J_emg);
fprintf(fid, 'MATLAB_2025-03-20_紧急购电量_kWh,%.6f\n', sum(r_emg));
fprintf(fid, 'Python锚定_2025-03-20_总费用_元,50206.2807\n');
fprintf(fid, 'MATLAB_退化自检_2025-03-20_费用_元,%.4f\n', J_deg);
fprintf(fid, '问题2_2025-03-20_费用_元,%.4f\n', J_q2_320);
fprintf(fid, '备注,口径与 lib/solve_day.solve_stage 一致（D-12 结算、D-13 不可追溯、D-04 终端自由）\n');
fclose(fid);

fid = fopen(csv_k, 'w', 'n', 'UTF-8');
fprintf(fid, '时段序号,计划购电量x_kWh,最终购电量y_kWh,充电量_kWh,放电量_kWh\n');
for k = 1:K
    fprintf(fid, '%d,%.6f,%.6f,%.6f,%.6f\n', k, x(k), y(k), u(k), v(k));
end
fclose(fid);

fprintf('MATLAB 复算完成：3.20 总费用 = %.4f 元（Python 锚定 50206.2807）\n', J_320);
fprintf('退化自检：3.20 费用 = %.4f 元（问题 2 当日 %.4f 元）\n', J_deg, J_q2_320);
fprintf('分解一致性抽查：D320 的 0:00 预报第 61 段 = %.4f（应与 Python 一致）\n', FC144(D320,1,61));
