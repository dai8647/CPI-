# -*- coding: utf-8 -*-
"""
R* (HLW) 検証スクリプト — Python 側グラウンドトゥルース
================================================================
Holston-Laubach-Williams (2017/2023) の3段階モデルを Python に移植し、
Excel/VBA (ModuleRStar.bas) の fixed-coefficient 実装と突合するための
参照値を生成する。配布物ではない（このPC専用の検証スクリプト）。

コンポーネント:
  1. HP フィルタ            hpfilter.R               (λ=1600 / 36000)
  2. カルマンフィルタ/スムーザ kalman.*.R             (KF)
  3. メディアン不偏推定量     median.unbiased.estimator.*.R (LW: Stock-Watson 1998)
  4. HLW stage1/2/3        rstar.stage*.R           (LHW, MLE + fixed coeff)
  5. ZLB                   短期名目金利を床(既定 0.0)で下げ止め

入力:  rstar_data.csv  (rstar_fetch.py が生成する四半期パネル)
出力:  rstar_ground_truth.json  (VBA との突合用参照値)

実行:
  python rstar_check.py            # MLE 込みフル
  python rstar_check.py --fast     # MLE をスキップ (初期値 fixed coeff のみ)
"""
import io, os, sys, csv, json, datetime
import numpy as np

try:
    from scipy.optimize import minimize as _scipy_minimize
except Exception:  # pragma: no cover - scipy は ground truth 生成にのみ必須
    _scipy_minimize = None

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
HERE = os.path.dirname(os.path.abspath(__file__))

# HLW 本体 (run.hlw.R) と同じ制約
A3_CONSTRAINT = -0.0025   # IS 曲線の金利感度 a_r の上限
B2_CONSTRAINT = 0.025     # フィリップス曲線の需給ギャップ係数 b_y の下限
ZLB_FLOOR_DEFAULT = 0.0   # 名目短期金利の下限（ZLB）


# ================================================================
# 0. データ準備
# ================================================================
def load_panel():
    """rstar_data.csv → dict of numpy arrays（pi_defl が定義される q_idx 以降）"""
    rows = list(csv.DictReader(open(os.path.join(HERE, 'rstar_data.csv'), encoding='utf-8-sig')))

    def col(name):
        out = []
        for r in rows:
            v = r[name]
            out.append(float(v) if v not in ('', None) else None)
        return out

    d = {k: col(k) for k in ('q_idx', 'qnum', 'real_gdp', 'defl', 'nom', 'i_q', 'pi_defl', 'pi_cpi')}
    return d


def _seasonal_adjust_quarterly(log_series, qnum):
    """四半期ダミーで季節性を除去（各四半期の平均を引く）。"""
    out = np.array(log_series, float)
    qnum = np.asarray(qnum, int)
    for q in (1, 2, 3, 4):
        m = qnum == q
        if np.any(m):
            out[m] = out[m] - np.nanmean(out[m])
    return out


def build_hlw_inputs(d, zlb_floor=ZLB_FLOOR_DEFAULT, inflation_source='defl_sa_annualized'):
    """
    HLW 入力4系列を構築し、推定サンプルを data.start = 期待インフレが最初に
    定義される四半期 に切り出す（HLW では inflation の Lag をデータ先頭で
    使うため、3期前までのインフレが必要）。

    返り値 dict:
      log_output, inflation, inflation_expectations, nominal_rate, real_rate,
      zlb_floored, q_labels, data_start_idx
    """
    n = len(d['q_idx'])
    log_output = np.array([np.log(v) if v is not None else np.nan for v in d['real_gdp']])
    defl = np.array([v if v is not None else np.nan for v in d['defl']])
    if inflation_source == 'defl_sa_annualized':
        # HLW の定義に最も近い: 季節調整済みデフレータの年率換算四半期変化
        sa = _seasonal_adjust_quarterly(np.log(defl), d['qnum'])
        inflation = np.full(n, np.nan)
        inflation[1:] = 400.0 * np.diff(sa)
    elif inflation_source == 'defl_yoy':
        inflation = np.array([v if v is not None else np.nan for v in d['pi_defl']])
    elif inflation_source == 'cpi_yoy':
        inflation = np.array([v if v is not None else np.nan for v in d['pi_cpi']])
    else:  # defl_q: 季節調整なし（非推奨）
        inflation = np.full(n, np.nan)
        inflation[1:] = 400.0 * np.log(defl[1:] / defl[:-1])

    nominal_raw = np.array([v if v is not None else np.nan for v in d['i_q']])
    nominal_rate = np.where(np.isnan(nominal_raw), np.nan, np.maximum(nominal_raw, zlb_floor))

    # inflation.expectations[t] = mean(inflation[t], t-1, t-2, t-3)  （0-based）
    inflation_expectations = np.full(n, np.nan)
    for t in range(3, n):
        w = inflation[t - 3:t + 1]
        if not np.isnan(w).any():
            inflation_expectations[t] = w.mean()

    real_rate = nominal_rate - inflation_expectations

    # data.start = 期待インフレが最初に定義される四半期
    valid = np.where(~np.isnan(inflation_expectations))[0]
    if len(valid) == 0:
        raise RuntimeError('期待インフレが1期も定義できません（データ不足）')
    i0 = int(valid[0])

    q_labels = [f"{int(d['q_idx'][i])}" for i in range(i0, n)]
    return {
        'log_output': log_output[i0:],
        'inflation': inflation[i0:],
        'inflation_expectations': inflation_expectations[i0:],
        'nominal_rate': nominal_rate[i0:],
        'real_rate': real_rate[i0:],
        'zlb_floored_count': int(np.sum((nominal_raw > -1e9) & (nominal_raw < zlb_floor))),
        'q_labels': q_labels,
        'data_start_idx': i0,
        'q_idx_all': d['q_idx'],
        'inflation_source': inflation_source,
    }


# ================================================================
# 1. HP フィルタ (hpfilter.R と同一: mFilter の lambda 型)
# ================================================================
def hp_filter(x, lam):
    x = np.asarray(x, dtype=float).ravel()
    n = len(x)
    imat = np.eye(n)
    Ln = np.zeros((n, n))
    Ln[1:, :] = np.eye(n - 1, n)
    Ln2 = (imat - Ln) @ (imat - Ln)
    Q = Ln2[2:, :].T                 # n x (n-2)
    SIGMA_R = Q.T @ Q                # (n-2)x(n-2)
    SIGMA_n = np.eye(n - 2)
    g = Q.T @ x
    b = np.linalg.solve(SIGMA_n + lam * SIGMA_R, g)
    cycle = lam * (Q @ b)
    trend = x - cycle
    return trend, cycle


# ================================================================
# 2. カルマンフィルタ / スムーザ
# ================================================================
def kalman_loglik(xi0, P0, F, Q, A, H, R, y, x):
    """kalman.log.likelihood.R の移植。ll.vec (T,), ll.cum。"""
    y = np.atleast_2d(y)
    if y.shape[0] == 1 and y.shape[1] != 1:
        y = y.reshape(-1, 1)
    x = np.atleast_2d(x)
    if x.shape[0] == 1 and x.shape[1] != 1:
        x = x.reshape(-1, 1)
    T = y.shape[0]
    nobs = y.shape[1]
    xi = np.asarray(xi0, float).ravel().copy()
    P = np.asarray(P0, float).copy()
    ll_vec = np.zeros(T)
    ll_cum = 0.0
    for t in range(T):
        xi_tm1 = F @ xi
        P_tm1 = F @ P @ F.T + Q
        pred = y[t] - A.T @ x[t] - H.T @ xi_tm1
        HPHR = H.T @ P_tm1 @ H + R
        sign, logdet = np.linalg.slogdet(HPHR)
        ll = -(nobs / 2.0) * np.log(2.0 * np.pi) - 0.5 * logdet - 0.5 * float(pred @ np.linalg.solve(HPHR, pred))
        ll_vec[t] = ll
        ll_cum += float(ll)
        K = P_tm1 @ H @ np.linalg.inv(HPHR)
        xi = xi_tm1 + K @ pred
        P = P_tm1 - K @ H.T @ P_tm1
    return ll_vec, ll_cum


def _safe_inv(M):
    """特異になりやすい行列の逆行列（対称化 + ジッタ）。"""
    M = 0.5 * (M + M.T)
    try:
        return np.linalg.inv(M)
    except np.linalg.LinAlgError:
        n = M.shape[0]
        return np.linalg.inv(M + 1e-9 * np.eye(n))


def _psd(P):
    """対称化して最小固有値を微小正に補正。"""
    P = 0.5 * (P + P.T)
    w = np.linalg.eigvalsh(P)
    if np.min(w) < 1e-12:
        P = P + (1e-10 - np.min(w)) * np.eye(P.shape[0])
    return P


def kalman_states(xi0, P0, F, Q, A, H, R, y, x):
    """kalman.states.R (filtered + RTS smoother) の移植。"""
    y = np.atleast_2d(y)
    if y.shape[0] == 1 and y.shape[1] != 1:
        y = y.reshape(-1, 1)
    x = np.atleast_2d(x)
    if x.shape[0] == 1 and x.shape[1] != 1:
        x = x.reshape(-1, 1)
    T = y.shape[0]
    n = len(np.asarray(xi0).ravel())
    xi = np.asarray(xi0, float).ravel().copy()
    P = np.asarray(P0, float).copy()

    xi_ttm1 = np.zeros((T, n)); P_ttm1 = np.zeros((T, n, n))
    xi_tt = np.zeros((T, n));   P_tt = np.zeros((T, n, n))
    for t in range(T):
        a = F @ xi
        Pp = F @ P @ F.T + Q
        pred = y[t] - A.T @ x[t] - H.T @ a
        HPHR = H.T @ Pp @ H + R
        at = a + Pp @ H @ np.linalg.solve(HPHR, pred)
        Pt = Pp - Pp @ H @ np.linalg.solve(HPHR, H.T @ Pp)
        Pp = _psd(Pp); Pt = _psd(Pt)
        xi_ttm1[t] = a; P_ttm1[t] = Pp
        xi_tt[t] = at;  P_tt[t] = Pt
        xi = at; P = Pt

    # RTS smoother
    xi_tT = np.zeros((T, n)); P_tT = np.zeros((T, n, n))
    xi_tT[-1] = xi_tt[-1]; P_tT[-1] = P_tt[-1]
    for t in range(T - 2, -1, -1):
        J = P_tt[t] @ F.T @ _safe_inv(P_ttm1[t + 1])
        xi_tT[t] = xi_tt[t] + J @ (xi_tT[t + 1] - xi_ttm1[t + 1])
        P_tT[t] = _psd(P_tt[t] + J @ (P_tT[t + 1] - P_ttm1[t + 1]) @ J.T)
    return {'filtered': {'xi_tt': xi_tt, 'P_tt': P_tt, 'xi_ttm1': xi_ttm1, 'P_ttm1': P_ttm1},
            'smoothed': {'xi_tT': xi_tT, 'P_tT': P_tT}}


# ================================================================
# 3. メディアン不偏推定量 (LW / Stock-Watson 1998)
# ================================================================
_VAL_EW = np.array([0.426, 0.476, 0.516, 0.661, 0.826, 1.111, 1.419, 1.762, 2.355, 2.91,
                    3.413, 3.868, 4.925, 5.684, 6.670, 7.690, 8.477, 9.191, 10.693, 12.024,
                    13.089, 14.440, 16.191, 17.332, 18.699, 20.464, 21.667, 23.851, 25.538,
                    26.762, 27.874])
_VAL_MW = np.array([0.689, 0.757, 0.806, 1.015, 1.234, 1.632, 2.018, 2.390, 3.081, 3.699,
                    4.222, 4.776, 5.767, 6.586, 7.703, 8.683, 9.467, 10.101, 11.639, 13.039,
                    13.900, 15.214, 16.806, 18.330, 19.020, 20.562, 21.837, 24.350, 26.248,
                    27.089, 27.758])
_VAL_QL = np.array([3.198, 3.416, 3.594, 4.106, 4.848, 5.689, 6.682, 7.626, 9.16, 10.66,
                    11.841, 13.098, 15.451, 17.094, 19.423, 21.682, 23.342, 24.920, 28.174,
                    30.736, 33.313, 36.109, 39.673, 41.955, 45.056, 48.647, 50.983, 55.514,
                    59.278, 61.311, 64.016])


def _median_interp(statval, table):
    """Stock-Watson Table 3 補間 (R 実装と同一)。"""
    if statval <= table[0]:
        return 0.0
    for i in range(len(table) - 1):
        if table[i] < statval <= table[i + 1]:
            return float(i + (statval - table[i]) / (table[i + 1] - table[i]))
    return float(len(table) - 1)


def _lw_stats(stats):
    stats = np.asarray(stats, float)
    ew = float(np.log(np.mean(np.exp(stats ** 2 / 2.0))))
    mw = float(np.mean(stats ** 2))
    qlr = float(np.max(stats ** 2))
    return ew, mw, qlr


def median_unbiased_lambda_g(series):
    """median.unbiased.estimator.stage1.R の移植。lambda_g = lame/(T-1)。"""
    series = np.asarray(series, float).ravel()
    T = len(series)
    y = 400.0 * np.diff(series)
    L = T - 1
    stats = []
    for i in range(4, T - 5 + 1):              # 1-based i = 4..T-5
        second = np.array([0.0] * i + [1.0] * (L - i))
        xr = np.column_stack([np.ones(L), second])
        XtX = xr.T @ xr
        b = np.linalg.solve(XtX, xr.T @ y)
        resid = y - xr @ b
        s3 = float((resid @ resid) / (T - 2 - 1))
        val = float(b[1] / np.sqrt(s3 * np.linalg.inv(XtX)[1, 1]))
        stats.append(val)
    ew, mw, qlr = _lw_stats(stats)
    lame = _median_interp(ew, _VAL_EW)
    lamm = _median_interp(mw, _VAL_MW)
    lamq = _median_interp(qlr, _VAL_QL)
    return float(lame / (T - 1)), {'ew': ew, 'mw': mw, 'qlr': qlr, 'lame': lame, 'lamm': lamm, 'lamq': lamq}


def median_unbiased_lambda_z(y, x):
    """median.unbiased.estimator.stage2.R の移植。lambda_z = lame/T。"""
    y = np.asarray(y, float).ravel()
    x = np.asarray(x, float)
    if x.ndim == 1:
        x = x.reshape(-1, 1)
    T = x.shape[0]
    stats = []
    for i in range(4, T - 4 + 1):              # 1-based i = 4..T-4
        second = np.array([0.0] * i + [1.0] * (T - i))
        xr = np.column_stack([x, second])
        ncol = xr.shape[1]
        XtX = xr.T @ xr
        b = np.linalg.solve(XtX, xr.T @ y)
        resid = y - xr @ b
        s3 = float((resid @ resid) / (T - ncol))
        val = float(b[-1] / np.sqrt(s3 * np.linalg.inv(XtX)[ncol - 1, ncol - 1]))
        stats.append(val)
    ew, mw, qlr = _lw_stats(stats)
    lame = _median_interp(ew, _VAL_EW)
    return float(lame / T), {'ew': ew, 'mw': mw, 'qlr': qlr, 'lame': lame}


# ================================================================
# 4. HLW stage1/2/3
# ================================================================
def _output_gap(log_output, T):
    x_og = np.column_stack([np.ones(T + 4), np.arange(1, T + 5, dtype=float)])
    b = np.linalg.solve(x_og.T @ x_og, x_og.T @ log_output)
    return (log_output - x_og @ b) * 100.0


def _unpack_stage1(p, y_data, x_data, xi0, P0):
    y_data = np.array(y_data, float, copy=True)
    x_data = np.array(x_data, float, copy=True)
    # stage1 は転置しない（unpack.parameters.stage1.R に t() が無い）
    A = np.zeros((4, 2))
    A[0:2, 0] = p[0:2]
    A[0, 1] = p[3]
    A[2, 1] = p[2]
    A[3, 1] = 1 - p[2]
    H = np.zeros((3, 2))
    H[0, 0] = 1
    H[1:3, 0] = -p[0:2]
    H[1, 1] = -p[3]
    R = np.diag([p[5] ** 2, p[6] ** 2])
    Q = np.zeros((3, 3)); Q[0, 0] = p[7] ** 2
    F = np.zeros((3, 3)); F[0, 0] = F[1, 0] = F[2, 1] = 1.0
    T = y_data.shape[0]
    g = p[4]
    y_data[:, 0] = y_data[:, 0] - np.arange(1, T + 1) * g
    x_data[:, 0] = x_data[:, 0] - np.arange(0, T) * g
    x_data[:, 1] = x_data[:, 1] - np.arange(-1, T - 1) * g
    return dict(xi0=xi0, P0=P0, F=F, Q=Q, A=A, H=H, R=R, y_data=y_data, x_data=x_data)


def _unpack_stage2(p, y_data, x_data, lam_g, xi0, P0):
    A = np.zeros((2, 7))
    A[0, 0:2] = p[0:2]
    A[0, 2:4] = p[2] / 2.0
    A[0, 6] = p[3]
    A[1, 0] = p[6]
    A[1, 4] = p[5]
    A[1, 5] = 1 - p[5]
    A = A.T
    H = np.zeros((2, 4))
    H[0, 0] = 1
    H[0, 1:3] = -p[0:2]
    H[0, 3] = p[4]
    H[1, 1] = -p[6]
    H = H.T
    R = np.diag([p[7] ** 2, p[8] ** 2])
    Q = np.zeros((4, 4))
    Q[0, 0] = p[9] ** 2
    Q[3, 3] = (lam_g * p[9]) ** 2
    F = np.zeros((4, 4))
    F[0, 0] = F[0, 3] = F[1, 0] = F[2, 1] = F[3, 3] = 1.0
    return dict(xi0=xi0, P0=P0, F=F, Q=Q, A=A, H=H, R=R, y_data=y_data, x_data=x_data)


def _unpack_stage3(p, y_data, x_data, lam_g, lam_z, xi0, P0):
    A = np.zeros((2, 6))
    A[0, 0:2] = p[0:2]
    A[0, 2:4] = p[2] / 2.0
    A[1, 0] = p[4]
    A[1, 4] = p[3]
    A[1, 5] = 1 - p[3]
    A = A.T
    H = np.zeros((2, 7))
    H[0, 0] = 1
    H[0, 1:3] = -p[0:2]
    H[0, 3:5] = -p[2] * 2.0
    H[0, 5:7] = -p[2] / 2.0
    H[1, 1] = -p[4]
    H = H.T
    R = np.diag([p[5] ** 2, p[6] ** 2])
    Q = np.zeros((7, 7))
    Q[0, 0] = (1 + lam_g ** 2) * p[7] ** 2
    Q[0, 3] = Q[3, 0] = Q[3, 3] = (lam_g * p[7]) ** 2
    Q[5, 5] = (lam_z * p[5] / p[2]) ** 2
    F = np.zeros((7, 7))
    F[0, 0] = F[0, 3] = F[1, 0] = F[2, 1] = F[3, 3] = F[4, 3] = F[5, 5] = F[6, 5] = 1.0
    return dict(xi0=xi0, P0=P0, F=F, Q=Q, A=A, H=H, R=R, y_data=y_data, x_data=x_data)


def _gradient(f, x, delta=1e-5):
    """utilities.R gradient() と同じ中心差分。"""
    g = np.zeros_like(np.asarray(x, float))
    for i in range(len(x)):
        x1 = np.array(x, float); x1[i] += delta
        x2 = np.array(x, float); x2[i] -= delta
        g[i] = (f(x1) - f(x2)) / delta / 2.0
    return g


def _mle(fun, p0, lb, ub):
    """nloptr(NLOPT_LD_LBFGS, xtol_rel=1e-8) 相当。失敗時は p0 を返す。"""
    p0 = np.asarray(p0, float)
    if _scipy_minimize is None:
        return p0, False
    bounds = [(None if np.isneginf(lb[i]) else lb[i],
               None if np.isposinf(ub[i]) else ub[i]) for i in range(len(p0))]

    def obj(p):
        return fun(np.asarray(p, float))

    def jac(p):
        return _gradient(obj, np.asarray(p, float))

    try:
        res = _scipy_minimize(obj, p0, jac=jac, method='L-BFGS-B', bounds=bounds,
                              options={'ftol': 1e-12, 'gtol': 1e-8, 'maxiter': 2000, 'maxls': 40})
        x = np.asarray(res.x, float)
        if not np.all(np.isfinite(x)):
            return p0, False
        return x, bool(res.success)
    except Exception:
        return p0, False


def rstar_stage1(log_output, inflation, b2_constraint=B2_CONSTRAINT,
                 xi00=None, P00=None, estimate=True):
    T = len(log_output) - 4
    output_gap = _output_gap(log_output, T)

    trend, _ = hp_filter(log_output, 36000.0)
    g_pot = trend[1:]                                   # g.pot.start.index = 2 (1-based)
    if xi00 is None:
        xi00 = np.array([100 * g_pot[2], 100 * g_pot[1], 100 * g_pot[0]])
    if P00 is None:
        P00 = np.eye(3)

    y_is = output_gap[4:(T + 4)]
    x_is = np.column_stack([output_gap[3:(T + 3)], output_gap[2:(T + 2)]])
    b_is = np.linalg.solve(x_is.T @ x_is, x_is.T @ y_is)
    r_is = y_is - x_is @ b_is
    s_is = np.sqrt(np.sum(r_is ** 2) / (len(r_is) - x_is.shape[1]))

    y_ph = inflation[4:(T + 4)]
    x_ph = np.column_stack([inflation[3:(T + 3)],
                            (inflation[2:(T + 2)] + inflation[1:(T + 1)] + inflation[0:T]) / 3.0,
                            output_gap[3:(T + 3)]])
    b_ph = np.linalg.solve(x_ph.T @ x_ph, x_ph.T @ y_ph)
    r_ph = y_ph - x_ph @ b_ph
    s_ph = np.sqrt(np.sum(r_ph ** 2) / (len(r_ph) - x_ph.shape[1]))

    y_data = np.column_stack([100 * log_output[4:(T + 4)], inflation[4:(T + 4)]])
    x_data = np.column_stack([100 * log_output[3:(T + 3)], 100 * log_output[2:(T + 2)],
                              inflation[3:(T + 3)],
                              (inflation[2:(T + 2)] + inflation[1:(T + 1)] + inflation[0:T]) / 3.0])

    p0 = np.array([b_is[0], b_is[1], b_ph[0], b_ph[2], 0.85, s_is, s_ph, 0.5])
    lb = np.full(8, -np.inf); ub = np.full(8, np.inf)
    if b2_constraint is not None and not np.isnan(b2_constraint):
        if p0[3] < b2_constraint:
            p0[3] = b2_constraint
        lb[3] = b2_constraint
    for j in (5, 6, 7):           # sigma_y~, sigma_pi, sigma_y* > 0
        lb[j] = 0.01; ub[j] = 100.0

    # calculate.covariance 相当（P00 の精緻化）
    P_diag = 0.2 * np.eye(3)
    if estimate and _scipy_minimize is not None:
        p_cov, _ = _mle(lambda p: -_ll_wrap(1, p, y_data, x_data, None, None, xi00, P_diag), p0, lb, ub)
        st = _states_wrap(1, p_cov, y_data, x_data, None, None, xi00, P_diag)
        P00 = st['filtered']['P_ttm1'][0]

    if estimate:
        theta, ok = _mle(lambda p: -_ll_wrap(1, p, y_data, x_data, None, None, xi00, P00), p0, lb, ub)
    else:
        theta, ok = p0, False

    ll = _ll_wrap(1, theta, y_data, x_data, None, None, xi00, P00)
    states = _states_wrap(1, theta, y_data, x_data, None, None, xi00, P00)
    potential_smoothed = states['smoothed']['xi_tT'][:, 0] / 100.0
    return dict(theta=theta, converged=ok, log_likelihood=ll, states=states,
                xi00=xi00, P00=P00, potential_smoothed=potential_smoothed,
                output_gap=output_gap, kf=dict(y_data=y_data, x_data=x_data))


def rstar_stage2(log_output, inflation, real_rate, lambda_g,
                 a3_constraint=A3_CONSTRAINT, b2_constraint=B2_CONSTRAINT,
                 xi00=None, P00=None, estimate=True):
    T = len(log_output) - 4
    output_gap = _output_gap(log_output, T)
    trend, _ = hp_filter(log_output, 36000.0)
    g_pot = trend[1:]
    g_pot_diff = np.diff(g_pot)
    if xi00 is None:
        xi00 = np.array([100 * g_pot[2], 100 * g_pot[1], 100 * g_pot[0], 100 * g_pot_diff[1]])
    if P00 is None:
        P00 = np.eye(4)

    y_is = output_gap[4:(T + 4)]
    x_is = np.column_stack([output_gap[3:(T + 3)], output_gap[2:(T + 2)],
                            (real_rate[3:(T + 3)] + real_rate[2:(T + 2)]) / 2.0,
                            np.ones(T)])
    b_is = np.linalg.solve(x_is.T @ x_is, x_is.T @ y_is)
    r_is = y_is - x_is @ b_is
    s_is = np.sqrt(np.sum(r_is ** 2) / (len(r_is) - x_is.shape[1]))

    y_ph = inflation[4:(T + 4)]
    x_ph = np.column_stack([inflation[3:(T + 3)],
                            (inflation[2:(T + 2)] + inflation[1:(T + 1)] + inflation[0:T]) / 3.0,
                            output_gap[3:(T + 3)]])
    b_ph = np.linalg.solve(x_ph.T @ x_ph, x_ph.T @ y_ph)
    r_ph = y_ph - x_ph @ b_ph
    s_ph = np.sqrt(np.sum(r_ph ** 2) / (len(r_ph) - x_ph.shape[1]))

    y_data = np.column_stack([100 * log_output[4:(T + 4)], inflation[4:(T + 4)]])
    x_data = np.column_stack([100 * log_output[3:(T + 3)], 100 * log_output[2:(T + 2)],
                              real_rate[3:(T + 3)], real_rate[2:(T + 2)],
                              inflation[3:(T + 3)],
                              (inflation[2:(T + 2)] + inflation[1:(T + 1)] + inflation[0:T]) / 3.0,
                              np.ones(T)])

    p0 = np.array([b_is[0], b_is[1], b_is[2], b_is[3], -b_is[2], b_ph[0], b_ph[2], s_is, s_ph, 0.5])
    lb = np.full(10, -np.inf); ub = np.full(10, np.inf)
    if b2_constraint is not None and not np.isnan(b2_constraint):
        if p0[6] < b2_constraint:
            p0[6] = b2_constraint
        lb[6] = b2_constraint
    if a3_constraint is not None and not np.isnan(a3_constraint):
        if p0[2] > a3_constraint:
            p0[2] = a3_constraint
        ub[2] = a3_constraint
    for j in (7, 8, 9):           # sigma_y~, sigma_pi, sigma_y* > 0
        lb[j] = 0.01; ub[j] = 100.0

    P_diag = 0.2 * np.eye(4)
    if estimate and _scipy_minimize is not None:
        p_cov, _ = _mle(lambda p: -_ll_wrap(2, p, y_data, x_data, lambda_g, None, xi00, P_diag), p0, lb, ub)
        st = _states_wrap(2, p_cov, y_data, x_data, lambda_g, None, xi00, P_diag)
        P00 = st['filtered']['P_ttm1'][0]
    if estimate:
        theta, ok = _mle(lambda p: -_ll_wrap(2, p, y_data, x_data, lambda_g, None, xi00, P00), p0, lb, ub)
    else:
        theta, ok = p0, False
    ll = _ll_wrap(2, theta, y_data, x_data, lambda_g, None, xi00, P00)
    states = _states_wrap(2, theta, y_data, x_data, lambda_g, None, xi00, P00)

    trend_smoothed = states['smoothed']['xi_tT'][:, 3] * 4.0
    potential_smoothed = np.concatenate([[states['smoothed']['xi_tT'][0, 2], states['smoothed']['xi_tT'][0, 1]],
                                         states['smoothed']['xi_tT'][:, 0]])
    output_gap_smoothed = 100 * log_output[2:(T + 4)] - potential_smoothed

    y_lw = output_gap_smoothed[2:]
    x_lw = np.column_stack([output_gap_smoothed[1:(len(output_gap_smoothed) - 1)],
                            output_gap_smoothed[0:(len(output_gap_smoothed) - 2)],
                            (x_data[:, 2] + x_data[:, 3]) / 2.0,
                            states['smoothed']['xi_tT'][:, 3],
                            np.ones(T)])
    return dict(theta=theta, converged=ok, log_likelihood=ll, states=states,
                xi00=xi00, P00=P00, y=y_lw, x=x_lw, kf=dict(y_data=y_data, x_data=x_data))


def rstar_stage3(log_output, inflation, real_rate, nominal_rate, lambda_g, lambda_z,
                 a3_constraint=A3_CONSTRAINT, b2_constraint=B2_CONSTRAINT,
                 xi00=None, P00=None, estimate=True):
    T = len(log_output) - 4
    output_gap = _output_gap(log_output, T)
    trend, _ = hp_filter(log_output, 36000.0)
    g_pot = trend[1:]
    g_pot_diff = np.diff(g_pot)
    if xi00 is None:
        xi00 = np.array([100 * g_pot[2], 100 * g_pot[1], 100 * g_pot[0],
                         100 * g_pot_diff[1], 100 * g_pot_diff[0], 0.0, 0.0])
    if P00 is None:
        P00 = np.eye(7)

    y_is = output_gap[4:(T + 4)]
    x_is = np.column_stack([output_gap[3:(T + 3)], output_gap[2:(T + 2)],
                            (real_rate[3:(T + 3)] + real_rate[2:(T + 2)]) / 2.0,
                            np.ones(T)])
    b_is = np.linalg.solve(x_is.T @ x_is, x_is.T @ y_is)
    r_is = y_is - x_is @ b_is
    s_is = np.sqrt(np.sum(r_is ** 2) / (len(r_is) - x_is.shape[1]))

    y_ph = inflation[4:(T + 4)]
    x_ph = np.column_stack([inflation[3:(T + 3)],
                            (inflation[2:(T + 2)] + inflation[1:(T + 1)] + inflation[0:T]) / 3.0,
                            output_gap[3:(T + 3)]])
    b_ph = np.linalg.solve(x_ph.T @ x_ph, x_ph.T @ y_ph)
    r_ph = y_ph - x_ph @ b_ph
    s_ph = np.sqrt(np.sum(r_ph ** 2) / (len(r_ph) - x_ph.shape[1]))

    y_data = np.column_stack([100 * log_output[4:(T + 4)], inflation[4:(T + 4)]])
    x_data = np.column_stack([100 * log_output[3:(T + 3)], 100 * log_output[2:(T + 2)],
                              real_rate[3:(T + 3)], real_rate[2:(T + 2)],
                              inflation[3:(T + 3)],
                              (inflation[2:(T + 2)] + inflation[1:(T + 1)] + inflation[0:T]) / 3.0])

    p0 = np.array([b_is[0], b_is[1], b_is[2], b_ph[0], b_ph[2], s_is, s_ph, 0.7])
    lb = np.full(8, -np.inf); ub = np.full(8, np.inf)
    if b2_constraint is not None and not np.isnan(b2_constraint):
        if p0[4] < b2_constraint:
            p0[4] = b2_constraint
        lb[4] = b2_constraint
    if a3_constraint is not None and not np.isnan(a3_constraint):
        if p0[2] > a3_constraint:
            p0[2] = a3_constraint
        ub[2] = a3_constraint
    for j in (5, 6, 7):           # sigma_y~, sigma_pi, sigma_y* > 0
        lb[j] = 0.01; ub[j] = 100.0

    P_diag = 0.2 * np.eye(7)
    if estimate and _scipy_minimize is not None:
        p_cov, _ = _mle(lambda p: -_ll_wrap(3, p, y_data, x_data, lambda_g, lambda_z, xi00, P_diag), p0, lb, ub)
        st = _states_wrap(3, p_cov, y_data, x_data, lambda_g, lambda_z, xi00, P_diag)
        P00 = st['filtered']['P_ttm1'][0]
    if estimate:
        theta, ok = _mle(lambda p: -_ll_wrap(3, p, y_data, x_data, lambda_g, lambda_z, xi00, P00), p0, lb, ub)
    else:
        theta, ok = p0, False
    ll = _ll_wrap(3, theta, y_data, x_data, lambda_g, lambda_z, xi00, P00)
    states = _states_wrap(3, theta, y_data, x_data, lambda_g, lambda_z, xi00, P00)

    trend_f = states['filtered']['xi_tt'][:, 3] * 4.0
    z_f = states['filtered']['xi_tt'][:, 5]
    rstar_f = trend_f + z_f
    trend_s = states['smoothed']['xi_tT'][:, 3] * 4.0
    z_s = states['smoothed']['xi_tT'][:, 5]
    rstar_s = trend_s + z_s
    return dict(theta=theta, converged=ok, log_likelihood=ll, states=states,
                xi00=xi00, P00=P00, rstar_filtered=rstar_f, rstar_smoothed=rstar_s,
                trend_filtered=trend_f, z_filtered=z_f,
                trend_smoothed=trend_s, z_smoothed=z_s, kf=dict(y_data=y_data, x_data=x_data))


def _ll_wrap(stage, p, y_data, x_data, lam_g, lam_z, xi0, P0):
    if stage == 1:
        m = _unpack_stage1(p, y_data, x_data, xi0, P0)
    elif stage == 2:
        m = _unpack_stage2(p, y_data, x_data, lam_g, xi0, P0)
    else:
        m = _unpack_stage3(p, y_data, x_data, lam_g, lam_z, xi0, P0)
    return kalman_loglik(m['xi0'], m['P0'], m['F'], m['Q'], m['A'], m['H'], m['R'],
                         m['y_data'], m['x_data'])[1]


def _states_wrap(stage, p, y_data, x_data, lam_g, lam_z, xi0, P0):
    if stage == 1:
        m = _unpack_stage1(p, y_data, x_data, xi0, P0)
    elif stage == 2:
        m = _unpack_stage2(p, y_data, x_data, lam_g, xi0, P0)
    else:
        m = _unpack_stage3(p, y_data, x_data, lam_g, lam_z, xi0, P0)
    st = kalman_states(m['xi0'], m['P0'], m['F'], m['Q'], m['A'], m['H'], m['R'], m['y_data'], m['x_data'])
    if stage == 1:
        T = m['y_data'].shape[0]
        adj = np.column_stack([np.arange(1, T + 1), np.arange(0, T), np.arange(-1, T - 1)]) * p[4]
        st['filtered']['xi_tt'] = st['filtered']['xi_tt'] + adj
        st['smoothed']['xi_tT'] = st['smoothed']['xi_tT'] + adj
    return st


def run_hlw(inp, estimate=True, zlb_floor=ZLB_FLOOR_DEFAULT):
    lo = inp['log_output']; inf = inp['inflation']
    rr = inp['real_rate']; nr = inp['nominal_rate']
    s1 = rstar_stage1(lo, inf, estimate=estimate)
    lam_g, lam_g_info = median_unbiased_lambda_g(s1['potential_smoothed'])
    s2 = rstar_stage2(lo, inf, rr, lam_g, estimate=estimate)
    lam_z, lam_z_info = median_unbiased_lambda_z(s2['y'], s2['x'])
    s3 = rstar_stage3(lo, inf, rr, nr, lam_g, lam_z, estimate=estimate)
    return {'stage1': s1, 'stage2': s2, 'stage3': s3,
            'lambda_g': lam_g, 'lambda_g_info': lam_g_info,
            'lambda_z': lam_z, 'lambda_z_info': lam_z_info}


# ================================================================
# 5. 自己整合チェック
# ================================================================
def self_checks(inp):
    checks = []

    def add(name, ok, detail):
        checks.append({'name': name, 'ok': bool(ok), 'detail': detail})

    x = np.asarray(inp['log_output'], float)
    tr, cyc = hp_filter(x, 1600.0)
    add('hp_identity_1600', np.max(np.abs((tr + cyc) - x)) < 1e-9,
        float(np.max(np.abs((tr + cyc) - x))))
    tr2, cyc2 = hp_filter(x, 36000.0)
    add('hp_identity_36000', np.max(np.abs((tr2 + cyc2) - x)) < 1e-9,
        float(np.max(np.abs((tr2 + cyc2) - x))))
    add('hp_36000_smoother_than_1600',
        np.var(np.diff(tr2, 2)) < np.var(np.diff(tr, 2)),
        {'var_2diff_1600': float(np.var(np.diff(tr, 2))), 'var_2diff_36000': float(np.var(np.diff(tr2, 2)))})

    # Kalman: 単位根状態 + ノイズ観測の解析的サニティ（T 増加で ll が総和則を満たす）
    rng = np.random.default_rng(0)
    T = 30
    phi, q, r = 0.8, 0.25, 0.09
    xi = np.zeros(T); y = np.zeros((T, 1)); s = 0.0
    for t in range(T):
        s = phi * s + rng.normal(0, np.sqrt(q))
        xi[t] = s
        y[t, 0] = s + rng.normal(0, np.sqrt(r))
    F = np.array([[phi]]); Q = np.array([[q]]); A = np.zeros((1, 1))
    H = np.array([[1.0]]); R = np.array([[r]])
    llv, llc = kalman_loglik(np.array([0.0]), np.array([[1.0]]), F, Q, A, H, R, y, np.ones((T, 1)))
    add('kf_loglik_sum', abs(np.sum(llv) - llc) < 1e-9, float(np.sum(llv) - llc))
    st = kalman_states(np.array([0.0]), np.array([[1.0]]), F, Q, A, H, R, y, np.ones((T, 1)))
    add('kf_shapes', st['filtered']['xi_tt'].shape == (T, 1) and st['smoothed']['xi_tT'].shape == (T, 1),
        [list(st['filtered']['xi_tt'].shape), list(st['smoothed']['xi_tT'].shape)])
    # smoother は終端で filter と一致する
    add('kf_smoother_endpoint',
        abs(st['smoothed']['xi_tT'][-1, 0] - st['filtered']['xi_tt'][-1, 0]) < 1e-9,
        float(st['smoothed']['xi_tT'][-1, 0] - st['filtered']['xi_tt'][-1, 0]))

    # LW: 定義域チェック（値は [0, 30/(T-1)] に収まる、有限）
    rng2 = np.random.default_rng(1)
    rw = np.arange(80) * 0.1 + (np.arange(80) > 40) * 5.0 + rng2.normal(0, 0.05, 80)
    lam0, info0 = median_unbiased_lambda_g(rw)
    add('lw_lambda_g_range', np.isfinite(lam0) and 0.0 <= lam0 <= 30.0 / (80 - 1), lam0)
    x2 = rng2.normal(size=(60, 3))
    y2 = x2 @ np.array([0.5, -0.2, 0.1]) + np.concatenate([[0] * 30 + [1] * 30]) * 2.0
    lam_z0, info_z0 = median_unbiased_lambda_z(y2, x2)
    add('lw_lambda_z_range', np.isfinite(lam_z0) and 0.0 <= lam_z0 <= 30.0 / 60, lam_z0)

    # ZLB: 負の名目金利が床に張り付く
    add('zlb_floor_active', inp['zlb_floored_count'] > 0,
        {'floored_quarters': inp['zlb_floored_count'], 'floor': ZLB_FLOOR_DEFAULT})
    add('zlb_no_negative_real',
        bool(np.nanmin(np.where(np.isnan(inp['real_rate']), np.inf, inp['real_rate'])) > -100),
        float(np.nanmin(np.where(np.isnan(inp['real_rate']), np.inf, inp['real_rate']))))
    return checks


def _tail(a, k=5):
    return [float(v) for v in np.asarray(a, float).ravel()[-k:]]


def write_vba_artifacts(inp, est, s3):
    """ModuleRStar.bas との突合用CSVを書き出す（突合の要点を分離）。"""
    th = s3['theta']
    lines = [('lambda_g', est['lambda_g']), ('lambda_z', est['lambda_z'])]
    for i, v in enumerate(th, 1):
        lines.append((f'theta{i}', float(v)))
    for i, v in enumerate(s3['xi00'], 1):
        lines.append((f'xi00_{i}', float(v)))
    P00 = np.asarray(s3['P00'], float)
    for i in range(P00.shape[0]):
        for j in range(P00.shape[1]):
            lines.append((f'P00_{i+1}_{j+1}', float(P00[i, j])))
    with open(os.path.join(HERE, 'rstar_fixed_coeff.csv'), 'w', newline='', encoding='ascii') as f:
        w = csv.writer(f)
        for k, v in lines:
            w.writerow([k, f'{v:.12g}'])

    with open(os.path.join(HERE, 'rstar_inputs.csv'), 'w', newline='', encoding='ascii') as f:
        w = csv.writer(f)
        w.writerow(['label', 'log_output', 'inflation', 'inflation_expectations', 'real_rate', 'nominal_rate'])
        for i in range(len(inp['log_output'])):
            w.writerow([inp['q_labels'][i]] + [f'{float(v):.12g}' for v in
                       (inp['log_output'][i], inp['inflation'][i], inp['inflation_expectations'][i],
                        inp['real_rate'][i], inp['nominal_rate'][i])])

    rf = np.asarray(s3['rstar_filtered'], float)
    rs = np.asarray(s3['rstar_smoothed'], float)
    with open(os.path.join(HERE, 'rstar_vba_check.csv'), 'w', newline='', encoding='ascii') as f:
        w = csv.writer(f)
        w.writerow(['label', 'rstar_filtered', 'rstar_smoothed'])
        for i in range(len(rf)):
            w.writerow([inp['q_labels'][i + 4], f'{rf[i]:.12g}', f'{rs[i]:.12g}'])
    print('wrote rstar_fixed_coeff.csv / rstar_inputs.csv / rstar_vba_check.csv')


def main():
    fast = '--fast' in sys.argv
    d = load_panel()
    inp = build_hlw_inputs(d, zlb_floor=ZLB_FLOOR_DEFAULT)
    print('HLW inputs: n=%d  q_idx %s..%s  (data_start_idx=%d, inflation=%s)'
          % (len(inp['log_output']), inp['q_labels'][0], inp['q_labels'][-1],
             inp['data_start_idx'], inp['inflation_source']))
    print('ZLB floor=%s  floored quarters=%d' % (ZLB_FLOOR_DEFAULT, inp['zlb_floored_count']))

    checks = self_checks(inp)
    for c in checks:
        print('  [%s] %s -> %s' % ('OK ' if c['ok'] else 'NG ', c['name'], c['detail']))

    hp1600, _ = hp_filter(inp['log_output'], 1600.0)
    hp36000, _ = hp_filter(inp['log_output'], 36000.0)
    lam_g_hp, lam_g_hp_info = median_unbiased_lambda_g(hp36000)

    print('\ncomputing HLW stages (estimate=%s) ...' % (not fast))
    est = run_hlw(inp, estimate=not fast, zlb_floor=ZLB_FLOOR_DEFAULT)
    s3 = est['stage3']
    print('  lambda_g=%.6f  lambda_z=%.6f' % (est['lambda_g'], est['lambda_z']))
    print('  stage3 loglik=%.4f converged=%s' % (s3['log_likelihood'], s3['converged']))
    print('  r* smoothed tail:', [round(v, 3) for v in _tail(s3['rstar_smoothed'])])

    r_star_last = float(s3['rstar_smoothed'][-1])
    r_star_filt_last = float(s3['rstar_filtered'][-1])
    print('  r* last: smoothed=%.3f  filtered=%.3f' % (r_star_last, r_star_filt_last))

    out = {
        'as_of': datetime.datetime.now().isoformat(),
        'source': 'rstar_data.csv (Japan quarterly panel)',
        'estimate_mle': not fast,
        'inflation_source': inp['inflation_source'],
        'zlb': {'floor': ZLB_FLOOR_DEFAULT, 'floored_quarters': inp['zlb_floored_count']},
        'sample': {'n': len(inp['log_output']), 'first_q_idx': int(inp['data_start_idx']),
                   'q_labels_tail': inp['q_labels'][-5:]},
        'hp': {
            'lambda1600': {'lambda': 1600.0, 'trend_head': _tail(hp1600[::-1], 5)[::-1], 'trend_tail': _tail(hp1600)},
            'hlw': {'lambda': 36000.0, 'trend_head': _tail(hp36000[::-1], 5)[::-1], 'trend_tail': _tail(hp36000)},
        },
        'lambda_g_from_hp_trend': {'value': lam_g_hp, 'info': lam_g_hp_info},
        'lambda_g': {'value': est['lambda_g'], 'info': est['lambda_g_info']},
        'lambda_z': {'value': est['lambda_z'], 'info': est['lambda_z_info']},
        'stage1': {'theta': [float(v) for v in est['stage1']['theta']],
                   'log_likelihood': float(est['stage1']['log_likelihood']),
                   'converged': bool(est['stage1']['converged'])},
        'stage2': {'theta': [float(v) for v in est['stage2']['theta']],
                   'log_likelihood': float(est['stage2']['log_likelihood']),
                   'converged': bool(est['stage2']['converged'])},
        'stage3': {'theta': [float(v) for v in s3['theta']],
                   'log_likelihood': float(s3['log_likelihood']),
                   'converged': bool(s3['converged']),
                   'xi00': [float(v) for v in s3['xi00']],
                   'P00': s3['P00'].tolist(),
                   'rstar_filtered_tail': _tail(s3['rstar_filtered']),
                   'rstar_smoothed_tail': _tail(s3['rstar_smoothed']),
                   'rstar_last_filtered': r_star_filt_last,
                   'rstar_last_smoothed': r_star_last},
        # ---- VBA (ModuleRStar.bas) の fixed-coefficient 突合用ブロック ----
        'fixed_coeff': {
            'stage3_theta': [float(v) for v in s3['theta']],
            'lambda_g': float(est['lambda_g']),
            'lambda_z': float(est['lambda_z']),
            'xi00': [float(v) for v in s3['xi00']],
            'P00': s3['P00'].tolist(),
            'C_map': 1.0,              # stage3 は C=0（観測に定数なし）
            'rstar_filtered': [float(v) for v in s3['rstar_filtered']],
            'rstar_smoothed': [float(v) for v in s3['rstar_smoothed']],
        },
        'self_checks': checks,
    }
    dest = os.path.join(HERE, 'rstar_ground_truth.json')
    with open(dest, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print('\nwrote', dest)
    write_vba_artifacts(inp, est, s3)
    ng = [c for c in checks if not c['ok']]
    print('self-checks: %d/%d OK' % (len(checks) - len(ng), len(checks)))
    return 0 if not ng else 2


if __name__ == '__main__':
    sys.exit(main())
