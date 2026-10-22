# -*- coding: utf-8 -*-
# lw_check.py: VBA の LW モデル (加重成長率 + mu マッチング) と同じ計算を Python で行い突合する
import csv, os, numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))

def hp_filter(x, lam):
    x = np.asarray(x, dtype=float).ravel()
    n = len(x)
    imat = np.eye(n)
    Ln = np.zeros((n, n))
    Ln[1:, :] = np.eye(n - 1, n)
    Ln2 = (imat - Ln) @ (imat - Ln)
    Q = Ln2[2:, :].T
    SIGMA_R = Q.T @ Q
    g = Q.T @ x
    b = np.linalg.solve(np.eye(n - 2) + lam * SIGMA_R, g)
    cycle = lam * (Q @ b)
    trend = x - cycle
    return trend

def main():
    rows = list(csv.DictReader(open(os.path.join(HERE, 'rstar_inputs.csv'), encoding='ascii')))
    lo = np.array([float(r['log_output']) for r in rows])
    rr = np.array([float(r['real_rate']) for r in rows])
    n = len(lo)
    lam = 1600.0
    w = [0.5, 0.3, 0.2, 0.1]
    tau = hp_filter(lo, lam)
    g = np.zeros(n)
    g[1:] = 400.0 * np.diff(tau)
    T = n - 4
    raw = np.zeros(T)
    res = np.zeros(T)
    for ti in range(T):
        ii = ti + 4
        s = 0.0
        ws2 = 0.0
        for j in range(4):
            if ii + j <= n - 1:
                s += w[j] * g[ii + j]
                ws2 += w[j]
        res[ti] = s / ws2 if ws2 > 0 else 0.0
        raw[ti] = res[ti]
    # mu window: 1995Q1..2007Q4 -> q_idx (1994Q1=0)
    mS = (1995 - 1994) * 4 + 0
    mE = (2007 - 1994) * 4 + 3
    mask = np.array([(ti + 4 >= mS) and (ti + 4 <= mE) for ti in range(T)])
    mu = float(np.mean(rr[4:4 + T][mask] - raw[mask]))
    res = res + mu
    out = os.path.join(HERE, 'rstar_lw_check.csv')
    with open(out, 'w', newline='', encoding='ascii') as f:
        wtr = csv.writer(f)
        wtr.writerow(['label', 'rstar_lw'])
        for ti in range(T):
            wtr.writerow([rows[ti + 4]['label'], f'{res[ti]:.12g}'])
    print('lw_check: n=%d  first=%.4f  last=%.4f  mu=%.4f' % (T, res[0], res[-1], mu))
    print('wrote', out)

if __name__ == '__main__':
    main()
