# -*- coding: utf-8 -*-
"""
grid2d.py: (λ_z, σ_y* 下限) の2D較正。
判定基準:
  - 尤度 (平坦なら λ_z・σ_y* はデータで識別されない → 外生較正が正当)
  - r* の動学 (サンプル内標準偏差) — 平坦 (0.05pp以下) は非現実的
  - r* 水準 (日本の文献レンジ 0.5〜1.5%)
  - 古典的 output gap の HP ギャップとの RMSE
  - σ_y* が下限に張り付かないこと
"""
import sys, json, numpy as np, openpyxl
sys.stdout.reconfigure(encoding='utf-8')
import importlib.util
spec = importlib.util.spec_from_file_location('rc', r'C:\Users\dai86\Downloads\CPI予想ツール\開発用\rstar_check.py')
rc = importlib.util.module_from_spec(spec); spec.loader.exec_module(rc)

XL = r'C:\Users\dai86\Downloads\CPI予想ツール\CPI.xlsm'
wb = openpyxl.load_workbook(XL, read_only=True)
ws = wb['Rstar']
rows = []
for row in ws.iter_rows(min_row=13, max_row=200, max_col=7, values_only=True):
    if row[0] is None: break
    rows.append(row)
wb.close()
lo = np.array([r[2] for r in rows], float)
infl = np.array([r[3] for r in rows], float)
rr = np.array([r[5] for r in rows], float)
nr = np.array([r[6] for r in rows], float)
hp_trend, _ = rc.hp_filter(lo, 36000.0)
gap_hp = (lo - hp_trend) * 100.0

LAMBDA_G, A3, B2 = 0.192798, -0.0025, 0.025
def qidx(y, q): return (y - 1995) * 4 + (q - 1) - 4
KEY = [('2008Q4',2008,4), ('2013Q4',2013,4), ('2019Q4',2019,4), ('2020Q2',2020,2),
       ('2022Q2',2022,2), ('2024Q4',2024,4), ('2026Q2',2026,2)]

def run(lz, sy_floor):
    """σ_y* に下限を設けた stage3 MLE"""
    T = len(lo) - 4
    trend, _ = rc.hp_filter(lo, 36000.0)
    g_pot = trend[1:]; gd = np.diff(g_pot)
    xi00 = np.array([100*g_pot[2], 100*g_pot[1], 100*g_pot[0], 100*gd[1], 100*gd[0], 0.0, 0.0])
    y = np.column_stack([100*lo[4:T+4], infl[4:T+4]])
    x = np.column_stack([100*lo[3:T+3], 100*lo[2:T+2], rr[3:T+3], rr[2:T+2],
                         infl[3:T+3], (infl[2:T+2]+infl[1:T+1]+infl[0:T])/3.0])
    # 初期値と境界 (FRBNY と同じ形, σ_y* に下限を追加)
    x_og = np.column_stack([np.ones(T+4), np.arange(1, T+5, dtype=float)])
    og = (lo - x_og @ np.linalg.solve(x_og.T @ x_og, x_og.T @ lo)) * 100
    y_is = og[4:T+4]
    x_is = np.column_stack([og[3:T+3], og[2:T+2], (rr[3:T+3]+rr[2:T+2])/2, np.ones(T)])
    b_is = np.linalg.solve(x_is.T @ x_is, x_is.T @ y_is)
    r_is = y_is - x_is @ b_is
    s_is = np.sqrt(np.sum(r_is**2)/(len(r_is)-x_is.shape[1]))
    y_ph = infl[4:T+4]
    x_ph = np.column_stack([infl[3:T+3], (infl[2:T+2]+infl[1:T+1]+infl[0:T])/3, og[3:T+3]])
    b_ph = np.linalg.solve(x_ph.T @ x_ph, x_ph.T @ y_ph)
    r_ph = y_ph - x_ph @ b_ph
    s_ph = np.sqrt(np.sum(r_ph**2)/(len(r_ph)-x_ph.shape[1]))
    p0 = np.array([b_is[0], b_is[1], b_is[2], b_ph[0], b_ph[2], s_is, s_ph, 0.7])
    lb = np.full(8, -np.inf); ub = np.full(8, np.inf)
    lb[4] = B2; ub[2] = A3
    lb[5] = 0.01; ub[5] = 100.0
    lb[6] = 0.01; ub[6] = 100.0
    lb[7] = sy_floor; ub[7] = 100.0
    if p0[4] < B2: p0[4] = B2
    if p0[2] > A3: p0[2] = A3
    if p0[7] < sy_floor: p0[7] = sy_floor
    # 1st pass (P_diag=0.2I) → P00
    P_diag = 0.2*np.eye(7)
    if rc._scipy_minimize is not None:
        p_cov, _ = rc._mle(lambda p: -rc._ll_wrap(3, p, y, x, LAMBDA_G, lz, xi00, P_diag), p0, lb, ub)
        st1 = rc._states_wrap(3, p_cov, y, x, LAMBDA_G, lz, xi00, P_diag)
        P00 = st1['filtered']['P_ttm1'][0]
        theta, ok = rc._mle(lambda p: -rc._ll_wrap(3, p, y, x, LAMBDA_G, lz, xi00, P00), p0, lb, ub)
    else:
        theta, P00 = p0, P_diag
    ll = rc._ll_wrap(3, theta, y, x, LAMBDA_G, lz, xi00, P00)
    st = rc._states_wrap(3, theta, y, x, LAMBDA_G, lz, xi00, P00)
    rs = st['smoothed']['xi_tT'][:, 3]*4 + st['smoothed']['xi_tT'][:, 5]
    rf = st['filtered']['xi_tt'][:, 3]*4 + st['filtered']['xi_tt'][:, 5]
    z_s = st['smoothed']['xi_tT'][:, 5]
    pot = st['smoothed']['xi_tT'][:, 0]/100.0
    gap_cls = y[:, 0] - pot*100.0
    return dict(lz=lz, sy_floor=sy_floor, ll=ll, theta=theta, P00=P00, xi00=xi00,
                rs=rs, rf=rf, z=z_s, gap_cls=gap_cls)

print('=== 2D グリッド: λ_z × σ_y*下限 ===')
print(f'{"lam_z":>7} {"sy_floor":>8} {"logL":>10} {"sig_y*":>8} {"r*std":>7} {"r*min":>7} {"r*max":>7} '
      f'{"z_rms":>7} {"gapcls_rms":>10} {"RMSE_HP":>8} {"r*tail":>7}')
best_rows = []
for sy_floor in (0.01, 0.2, 0.4, 0.5):
    for lz in (0.0, 0.001, 0.002, 0.0035, 0.005):
        r = run(lz, sy_floor)
        rmse = float(np.sqrt(np.mean((r['gap_cls'] - gap_hp[4:])**2)))
        rstd = float(np.std(r['rs']))
        print(f'{lz:7.4f} {sy_floor:8.2f} {r["ll"]:10.2f} {r["theta"][7]:8.4f} {rstd:7.3f} '
              f'{r["rs"].min():7.3f} {r["rs"].max():7.3f} {np.sqrt(np.mean(r["z"]**2)):7.3f} '
              f'{np.sqrt(np.mean(r["gap_cls"]**2)):10.3f} {rmse:8.3f} {r["rs"][-1]:7.3f}')
        best_rows.append((lz, sy_floor, rstd, rmse, r))
np.save(r'C:\Users\dai86\Temp\rstar_lab\grid2d.npy', np.array(best_rows, dtype=object), allow_pickle=True)
print('\nsaved grid2d.npy')
