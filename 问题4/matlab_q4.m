% matlab_q4.m —— 问题 4 的 MATLAB 独立复算（方法侧互验 S4c）
%
% 目的：用 MATLAB 的 linprog（对偶单纯形）独立重算附件4 波动电价下的关键数值：
%       1) 整点预报 → 144 段的线性插值分解（D-07 主口径 M2）与 Python 一致；
%       2) 4-2（逐日滚动）三个指定日期的单日 LP：2025-03-20、2025-06-21、2025-12-21，
%          终端余值 V_E = 次日最低价/η（D-10 主口径），与 问题4/全分辨率明细_4-2.csv 对照；
%       3) 4-3（多阶段）2025-03-20 的 0:00 计划 + 6/12/18 调整四阶段解与 D-12 结算，
%          与 问题4/全分辨率明细_4-3.csv 对照。
%
% 模型（与 lib/solve_day.solve_day / solve_stage 完全一致）：
%   单日 4-2：min Σ p·x − V_E·E_K
%     s.t. x_k + P_k·Δ + v_k ≥ L_k·Δ + u_k；0 ≤ u,v ≤ P̄Δ
%          E 前缀和 ∈ [E_MIN, E_MAX]；Σ(η·u − v/η) − E_K = −e0（显式 E_K 变量，与 Python 同构）
%   阶段 4-3（τ>0）：min Σ (1−κ₋)p·y + (κ₊+κ₋−1)p·t [− V_E·E_K（仅末阶段）]
%     s.t. 同上，再增 t_k ≥ y_k − x_k（x 为 0:00 计划，D-12 结算基准）
%   结算（实际值代入）：J = Σ[p·min(x,y) + κ₋p(x−y)⁺ + κ₊p(y−x)⁺] + κΣp·r
%
% 输入（只读）：附件/附件1~4.xlsx 与 问题4/全分辨率明细_4-2.csv、全分辨率明细_4-3.csv
% 输出（落盘，均在 问题4/ 下）：
%   _matlab校验_q4_汇总.csv   指标,数值,相对差 —— MATLAB vs Python 对照
%   _matlab校验_q4_逐时段.csv 3.20 单日 LP 逐时段 x/u/v（供逐点抽查）
%
% 运行：matlab -batch matlab_q4（在 问题4/ 目录下，或给出脚本绝对路径）

% ---------- 0. 定位路径 ----------
script_dir = fileparts(mfilename('fullpath'));                 % 本脚本所在目录，即 问题4/
root_dir   = fileparts(script_dir);                            % 工作区根目录
attach1    = fullfile(root_dir, '附件', '附件1.xlsx');
attach2    = fullfile(root_dir, '附件', '附件2.xlsx');
attach3    = fullfile(root_dir, '附件', '附件3.xlsx');
attach4    = fullfile(root_dir, '附件', '附件4.xlsx');
csv42      = fullfile(script_dir, '全分辨率明细_4-2.csv');
csv43      = fullfile(script_dir, '全分辨率明细_4-3.csv');
csv_sum    = fullfile(script_dir, '_matlab校验_q4_汇总.csv');
csv_k      = fullfile(script_dir, '_matlab校验_q4_逐时段.csv');

% ---------- 1. 读入电价 / 负载 / 光伏实际 / 附件4 电价 ----------
C1 = readcell(attach1, 'Sheet', 1);
price1 = cell2mat(C1(2:145, 2));                               % 附件1 电价（144，1），元/kWh
CL = readcell(attach2, 'Sheet', '小区负载');
CP = readcell(attach2, 'Sheet', '光伏发电实际功率');
load_all = zeros(365, 144); pv_all = zeros(365, 144);
for d = 1:365
    load_all(d, :) = cell2mat(CL(d + 1, 2:145));
    pv_all(d, :)   = cell2mat(CP(d + 1, 2:145));
end
C4 = readcell(attach4, 'Sheet', 'Sheet1');
price4 = zeros(365, 144);
for d = 1:365
    price4(d, :) = cell2mat(C4(d + 1, 2:145));                 % 附件4 逐日逐时段电价，元/kWh
end

K = 144; DT = 1/6; ETA = 0.9; PMAX = 5000; UMAX = PMAX * DT;
EMIN = 1200; EMAX = 10800; EINIT = 6000;
KU = 0.5; KO = 1.5; KE = 5;                                    % κ₋, κ₊, κ

% ---------- 2. 读入预报并做 M2 线性插值分解（与 matlab_q3.m 同一实现） ----------
C3 = readcell(attach3, 'Sheet', 1);
FC = zeros(365, 4, 24);
for i = 1:size(C3, 1) - 1
    d = floor((i - 1) / 4) + 1;
    t = mod(i - 1, 4) + 1;
    FC(d, t, :) = cell2mat(C3(i + 1, 3:26));
end
tau_list = [0 6 12 18];
FC144 = zeros(365, 4, 144);
for d = 1:365
    for ti = 1:4
        tau = tau_list(ti);
        nodes = nan(1, 25);
        if d > 1
            nodes(1) = FC(d - 1, 1, 24);                       % 0:00 结点取前一日 0:00 预报第 24 值
        else
            nodes(1) = 0;
        end
        for h = 1:24
            i_star = 0;
            for i2 = 1:4
                if h > tau_list(i2) && tau_list(i2) <= tau
                    i_star = i2;
                end
            end
            nodes(h + 1) = FC(d, i_star, h - tau_list(i_star));
        end
        for k = 1:144
            tt = k / 6;
            m = ceil(tt - 1e-12);
            if m < 1, m = 1; end
            frac = m - tt;
            FC144(d, ti, k) = nodes(m) * frac + nodes(m + 1) * (1 - frac);
        end
    end
end

% ---------- 3. 单日 LP（4-2：含显式终端变量 E_K 与终端余值） ----------
% 语法说明：MATLAB 脚本中函数必须放在文件末尾（R2016b 起），本文件全部函数在文末定义。

% ---------- 4. 从 Python 落盘 CSV 读取对照数据 ----------
T42 = readcell(csv42, 'Encoding', 'UTF-8', 'Delimiter', ',');
T43 = readcell(csv43, 'Encoding', 'UTF-8', 'Delimiter', ',');

% 4-3 CSV 读入前先做表头列核对（列序号：6=电价 13=x 14=y 17=u 18=v 19=r 21=时段末储电量）
assert(strcmp(char(string(T43{1, 13})), '计划购电量_kWh'), '4-3 CSV 第 13 列应为计划购电量');
assert(strcmp(char(string(T43{1, 21})), '时段末储电量_kWh'), '4-3 CSV 第 21 列应为时段末储电量');

% ---------- 5. 4-2 三日复算（3.20 / 6.21 / 12.21） ----------
out_lines = {};                                                % 汇总行收集（先初始化）
days_42 = [79 172 355];                                        % 1 基天序号
cost_py_all = zeros(1, 3); cost_ml_all = zeros(1, 3);
for ii = 1:3
    d0 = days_42(ii);
    day_str = datestr(datetime(2025, 1, 1) + days(d0 - 1), 'yyyy-mm-dd');
    prev_str = datestr(datetime(2025, 1, 1) + days(d0 - 2), 'yyyy-mm-dd');
    % 0:00 储电量 = 前一日 CSV 末段的时段末储电量（跨日连续性，来自 Python 链）
    blk_prev = day_block(T42, prev_str, K);
    e0 = cell2mat(blk_prev(K, 13));
    % 终端余值 V_E = 次日最低价 / η（D-10；末日退回当日最低价，与 Python 一致）
    if d0 < 365
        ve = min(price4(d0 + 1, :)) / ETA;
    else
        ve = min(price4(d0, :)) / ETA;
    end
    [x, u, v, cost] = solve_day_m(price4(d0, :)', load_all(d0, :)', pv_all(d0, :)', ...
                                         e0, ve, K, DT, ETA, UMAX, EMIN, EMAX);
    % Python 对照：同日 CSV 的 Σx 与 Σp·x
    blk = day_block(T42, day_str, K);
    x_py = cell2mat(blk(:, 9)); p_py = cell2mat(blk(:, 6));
    cost_py = sum(p_py .* x_py); xsum_py = sum(x_py);
    cost_ml_all(ii) = cost; cost_py_all(ii) = cost_py;
    out_lines{end + 1} = sprintf('%s_4-2_缴费_元,%.4f,%.4f', day_str, cost, cost - cost_py);      %#ok<SAGROW>
    out_lines{end + 1} = sprintf('%s_4-2_Σx_kWh,%.4f,%.4f', day_str, sum(x), sum(x) - xsum_py);   %#ok<SAGROW>
    out_lines{end + 1} = sprintf('%s_4-2_V_E_元每kWh,%.6f,0', day_str, ve);                       %#ok<SAGROW>
    if ii == 1
        % 3.20 单日 LP 逐时段落盘（供逐点抽查）
        fid = fopen(csv_k, 'w', 'n', 'UTF-8');
        fprintf(fid, '时段序号,MATLAB_x_kWh,Python_x_kWh,MATLAB_u_kWh,MATLAB_v_kWh\n');
        for k = 1:K
            fprintf(fid, '%d,%.6f,%.6f,%.6f,%.6f\n', k, x(k), x_py(k), u(k), v(k));
        end
        fclose(fid);
    end
end

% ---------- 6. 4-3 复算 2025-03-20（四阶段 + D-12 结算） ----------
% 阶段起点储电量取 Python 落盘 CSV 的对应时段末值（6:00→第 36 段末、12:00→72、18:00→108），
% 并核对 MATLAB 阶段 0 轨迹在 6:00 处与 CSV 的一致性（跨日/跨阶段连续性核验）
blk43 = day_block(T43, '2025-03-20', K);
blk43_prev = day_block(T43, '2025-03-19', K);
e0_43 = cell2mat(blk43_prev(K, 21));
E36 = cell2mat(blk43(36, 21)); E72 = cell2mat(blk43(72, 21)); E108 = cell2mat(blk43(108, 21));
ve_320 = min(price4(80, :)) / ETA;
% 阶段 0：0:00 计划（全价，无终端余值）
[x0, u0, v0, E0] = solve_stage_m(price4(79, :)', load_all(79, :)', squeeze(FC144(79, 1, :))', ...
                                 e0_43, [], 0, K, DT, ETA, UMAX, EMIN, EMAX, KU, KO);
cont_check = E0(36) - E36;                                     % 应为 0（跨阶段连续性）
x = x0; y = x0; u = u0; v = v0;
ks = [36 72 108];
e_cur = E36;
for ti = 2:4
    k0 = ks(ti - 1);
    if ti == 4
        ve_seg = ve_320;                                       % 末阶段（覆盖到 24:00）含终端余值
    else
        ve_seg = 0;
    end
    [yy, uu, vv, ~] = solve_stage_m(price4(79, k0 + 1:end)', load_all(79, k0 + 1:end)', ...
        squeeze(FC144(79, ti, k0 + 1:end))', e_cur, x(k0 + 1:end), ve_seg, K - k0, DT, ETA, ...
        UMAX, EMIN, EMAX, KU, KO);
    y(k0 + 1:end) = yy; u(k0 + 1:end) = uu; v(k0 + 1:end) = vv;
    if ti == 2, e_cur = E72; end                               % 12:00 起点（CSV 值）
    if ti == 3, e_cur = E108; end                              % 18:00 起点（CSV 值）
end
% 结算（实际值）
r_emg = max(load_all(79,:)' * DT + u - y - pv_all(79,:)' * DT - v, 0);
J_plan = sum(price4(79, :)' .* min(x, y));
J_adj = sum(KU * price4(79, :)' .* max(x - y, 0) + KO * price4(79, :)' .* max(y - x, 0));
J_emg = KE * sum(price4(79, :)' .* r_emg);
J_320 = J_plan + J_adj + J_emg;
% Python 对照：4-3 CSV 同日逐时段重算
x_p = cell2mat(blk43(:, 13)); y_p = cell2mat(blk43(:, 14));
pp = cell2mat(blk43(:, 6)); r_p = cell2mat(blk43(:, 19));
J_py = sum(pp .* min(x_p, y_p) + KU * pp .* max(x_p - y_p, 0) + KO * pp .* max(y_p - x_p, 0)) ...
       + KE * sum(pp .* r_p);
max_dx = max(abs(x - x_p)); max_dy = max(abs(y - y_p));
out_lines{end + 1} = sprintf('2025-03-20_4-3_总费用J_元,%.4f,%.4f', J_320, J_320 - J_py);         %#ok<SAGROW>
out_lines{end + 1} = sprintf('2025-03-20_4-3_计划购电费_元,%.4f,0', J_plan);                      %#ok<SAGROW>
out_lines{end + 1} = sprintf('2025-03-20_4-3_调整相关费_元,%.4f,0', J_adj);                       %#ok<SAGROW>
out_lines{end + 1} = sprintf('2025-03-20_4-3_紧急购电费_元,%.4f,0', J_emg);                       %#ok<SAGROW>
out_lines{end + 1} = sprintf('2025-03-20_4-3_Σr_kWh,%.4f,0', sum(r_emg));                         %#ok<SAGROW>
out_lines{end + 1} = sprintf('2025-03-20_4-3_逐时段max|Δx|_kWh,%.6f,0', max_dx);                  %#ok<SAGROW>
out_lines{end + 1} = sprintf('2025-03-20_4-3_逐时段max|Δy|_kWh,%.6f,0', max_dy);                  %#ok<SAGROW>
out_lines{end + 1} = sprintf('2025-03-20_4-3_阶段0在6点处连续性_ΔE_kWh,%.6f,0', cont_check);      %#ok<SAGROW>

% ---------- 7. 汇总结论落盘 ----------
fid = fopen(csv_sum, 'w', 'n', 'UTF-8');
fprintf(fid, '指标,数值,相对差\n');
for i = 1:numel(out_lines)
    fprintf(fid, '%s\n', out_lines{i});
end
fprintf(fid, '备注,MATLAB linprog(dual-simplex) 独立复算；口径与 lib/solve_day.py 一致（D-10 终端余值、D-12 结算、M2 分解）\n');
fclose(fid);

fprintf('MATLAB 复算完成：4-2 缴费 3.20 = %.4f、6.21 = %.4f、12.21 = %.4f 元\n', ...
        cost_ml_all(1), cost_ml_all(2), cost_ml_all(3));
fprintf('对照差异（MATLAB − Python）：%.4f / %.4f / %.4f 元\n', ...
        cost_ml_all - cost_py_all);
fprintf('4-3 3.20 总费用 J = %.4f 元（Python 逐时段重算 %.4f 元）\n', J_320, J_py);
fprintf('汇总已落盘：%s\n', csv_sum);

% ---------- 文末函数定义 ----------
function blk = day_block(T, day_str, K)
    % 从读入的 CSV 单元数组中取出指定日期的 144 行（日期列 = 第 1 列，格式 'yyyy-mm-dd'）
    rows_idx = [];
    for r = 2:size(T, 1)
        v = T{r, 1};
        if isa(v, 'datetime')
            s = datestr(v, 'yyyy-mm-dd');
        else
            s = char(string(v));
        end
        if strcmp(s, day_str)
            rows_idx = [rows_idx; r];                          %#ok<AGROW>
        end
    end
    assert(numel(rows_idx) == K, '日期 %s 的行数 %d 不等于 144', day_str, numel(rows_idx));
    blk = T(rows_idx(1):rows_idx(1) + K - 1, :);
end

function [x, u, v, c54] = solve_day_m(p, L, pv, e0, v_end, K, DT, ETA, UMAX, EMIN, EMAX)
    % 单日 4-2 LP（显式 E_K 变量与终端余值 −V_E·E_K，与 Python solve_day 的 v_end>0 分支同构）
    % 输入：p/L/pv (K,1) 元/kWh、kW、kW；e0 kWh；v_end 元/kWh；其余为参数
    % 输出：x/u/v (K,1) kWh；c54 实际缴费 Σp·x 元
    nv = 3 * K + 1;                                            % 变量 [x|u|v|E_K]
    c = [p(:); zeros(2 * K, 1); -v_end];                       % 目标 min Σp·x − V_E·E_K
    nrow = 3 * K;                                              % 供给 K + 储能上下界 2K
    rows = []; cols = []; vals = []; b = zeros(nrow, 1);
    ar = (1:K)';
    rows = [rows; ar; ar; ar];
    cols = [cols; ar; K + ar; 2 * K + ar];
    vals = [vals; -ones(K,1); ones(K,1); -ones(K,1)];
    b(1:K) = (pv(:) - L(:)) * DT;                              % 右端 Δ(P−L)，kWh
    [tr, tc] = find(tril(ones(K)));
    rows = [rows; K + tr; K + tr; 2 * K + tr; 2 * K + tr];
    cols = [cols; K + tc; 2 * K + tc; K + tc; 2 * K + tc];
    vals = [vals; ETA * ones(length(tr),1); -ones(length(tr),1)/ETA; ...
            -ETA * ones(length(tr),1); ones(length(tr),1)/ETA];
    b(K + 1:2 * K) = EMAX - e0;
    b(2 * K + 1:3 * K) = e0 - EMIN;
    A = sparse(rows, cols, vals, nrow, nv);                    % 第 3K+1 列（E_K）自动补零
    Aeq = sparse(1, nv);
    Aeq(1, K + 1:2 * K) = ETA;                                 % Σ(η·u − v/η) − E_K = −e0
    Aeq(1, 2 * K + 1:3 * K) = -1 / ETA;
    Aeq(1, 3 * K + 1) = -1;
    beq = -e0;
    lb = zeros(nv, 1); ub = inf(nv, 1);
    ub(K + 1:3 * K) = UMAX;                                    % u,v 功率上限，kWh
    lb(3 * K + 1) = EMIN; ub(3 * K + 1) = EMAX;                % E_K 储电量边界，kWh
    options = optimoptions('linprog', 'Algorithm', 'dual-simplex', 'Display', 'off');
    [z, ~, flag] = linprog(c, A, b, Aeq, beq, lb, ub, options);
    if flag ~= 1
        error('单日 LP 未收敛：flag=%d', flag);
    end
    x = z(1:K); u = z(K + 1:2 * K); v = z(2 * K + 1:3 * K);
    c54 = sum(p(:) .* x);                                      % 实际缴费 Σp·x，元
end

function [y, u, v, Esoc] = solve_stage_m(p, L, pv, e0, x_ref, v_end, K, DT, ETA, UMAX, EMIN, EMAX, KU, KO)
    % 阶段 LP（4-3：偏差结算等价边际形式 + 可选终端余值），与 Python solve_stage 同构
    % 输入：p/L/pv (K,1)；e0 kWh；x_ref (K,1) 或 []；v_end 元/kWh；KU/KO 倍数
    % 输出：y/u/v (K,1) kWh；Esoc (K+1,1) 阶段储电量轨迹，kWh
    has_ref = ~isempty(x_ref);
    nv = 3 * K + K * has_ref + (v_end ~= 0);                   % 变量 [y|u|v|t|E_K]
    if has_ref
        c = [ (1 - KU) * p(:); zeros(2 * K, 1); (KO + KU - 1) * p(:) ];
    else
        c = [ p(:); zeros(2 * K, 1) ];
    end
    if v_end ~= 0
        c = [c; -v_end];                                       % 末阶段追加 −V_E·E_K
    end
    nrow = 3 * K + K * has_ref;
    rows = []; cols = []; vals = []; b = zeros(nrow, 1);
    ar = (1:K)';
    rows = [rows; ar; ar; ar];
    cols = [cols; ar; K + ar; 2 * K + ar];
    vals = [vals; -ones(K,1); ones(K,1); -ones(K,1)];
    b(1:K) = (pv(:) - L(:)) * DT;
    [tr, tc] = find(tril(ones(K)));
    rows = [rows; K + tr; K + tr; 2 * K + tr; 2 * K + tr];
    cols = [cols; K + tc; 2 * K + tc; K + tc; 2 * K + tc];
    vals = [vals; ETA * ones(length(tr),1); -ones(length(tr),1)/ETA; ...
            -ETA * ones(length(tr),1); ones(length(tr),1)/ETA];
    b(K + 1:2 * K) = EMAX - e0;
    b(2 * K + 1:3 * K) = e0 - EMIN;
    if has_ref
        rows = [rows; 3 * K + ar; 3 * K + ar];
        cols = [cols; ar; 3 * K + ar];
        vals = [vals; ones(K,1); -ones(K,1)];
        b(3 * K + 1:4 * K) = x_ref(:);
    end
    A = sparse(rows, cols, vals, nrow, nv);
    lb = zeros(nv, 1); ub = inf(nv, 1);
    ub(K + 1:3 * K) = UMAX;
    Aeq = []; beq = [];
    if v_end ~= 0
        Aeq = sparse(1, nv);
        Aeq(1, K + 1:2 * K) = ETA;                             % Σ(η·u − v/η) − E_K = −e0
        Aeq(1, 2 * K + 1:3 * K) = -1 / ETA;
        Aeq(1, nv) = -1;
        beq = -e0;
        lb(nv) = EMIN; ub(nv) = EMAX;
    end
    options = optimoptions('linprog', 'Algorithm', 'dual-simplex', 'Display', 'off');
    [z, ~, flag] = linprog(c, A, b, Aeq, beq, lb, ub, options);
    if flag ~= 1
        error('阶段 LP 未收敛：flag=%d', flag);
    end
    y = z(1:K); u = z(K + 1:2 * K); v = z(2 * K + 1:3 * K);
    Esoc = e0 + cumsum(ETA * u - v / ETA);
end
