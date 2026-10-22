# -*- coding: utf-8 -*-
"""
finalize.py: 採用構成 (λ_g=0.192798, λ_z=0, σ_y* 下限=0.20) の
  theta / P00 を生成し、VBA FixCoeff に貼り付ける形で出力する。
  あわせて 産出ギャップ(古典) と HP ギャップの整合・r* 経路を検証する。
"""
import sys, json, numpy as np, openpyxl
sys.stdout.reconfigure(encoding='utf-8')
import importlib.util
spec = importlib.util.spec_from_file_location('rc', r'C:\Users\dai86\Downloads\CPI予想ツール\開発用\rstar_check.py')
rc = importlib.util.module_from_spec(spec); spec.loader.exec_module(rc)

XL = r'C:\Users\dai86\Downloads\CPI予想ツール\CPI.xlsm'
wb = openpyxl.load_workbook(XL, read_only=True); ws = wb['Rstar']
rows = []
for row in ws.iter_rows(min_row=13, max_row=200, max_col=7, values_only=True):
    if row[0] is None: break
    rows.append(row)
wb.close()
lo = np.array([r[2] for r in rows], float)
infl = np.array([r[3] for r in rows], float)
rr = np.array([r[5] for r in rows], float)
nr = np.array([r[6] for r in rows], float)
n = len(lo); T = n - 4

LAMBDA_G, LAMBDA_Z = 0.192798, 0.0
A3, B2, SY_FLOOR = -0.0025, 0.025, 0.20

# ---- Python 側と同じ MLE 手順 (σ_y* 下限付き) ----
trend, _ = rc.hp_filter(lo, 36000.0)
g_pot = trend[1:]; gd = np.diff(g_pot)
xi00 = np.array([100*g_pot[2], 100*g_pot[1], 100*g_pot[0], 100*gd[1], 100*gd[0], 0.0, 0.0])
y = np.column_stack([100*lo[4:T+4], infl[4:T+4]])
x = np.column_stack([100*lo[3:T+3], 100*lo[2:T+2], rr[3:T+3], rr[2:T+2],
                     infl[3:T+3], (infl[2:T+2]+infl[1:T+1]+infl[0:T])/3.0])
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
lb[7] = SY_FLOOR; ub[7] = 100.0
if p0[4] < B2: p0[4] = B2
if p0[2] > A3: p0[2] = A3
if p0[7] < SY_FLOOR: p0[7] = SY_FLOOR

P_diag = 0.2*np.eye(7)
p_cov, _ = rc._mle(lambda p: -rc._ll_wrap(3, p, y, x, LAMBDA_G, LAMBDA_Z, xi00, P_diag), p0, lb, ub)
st1 = rc._states_wrap(3, p_cov, y, x, LAMBDA_G, LAMBDA_Z, xi00, P_diag)
P00 = st1['filtered']['P_ttm1'][0]
theta, ok = rc._mle(lambda p: -rc._ll_wrap(3, p, y, x, LAMBDA_G, LAMBDA_Z, xi00, P00), p0, lb, ub)
ll = rc._ll_wrap(3, theta, y, x, LAMBDA_G, LAMBDA_Z, xi00, P00)
st = rc._states_wrap(3, theta, y, x, LAMBDA_G, LAMBDA_Z, xi00, P00)

names = ['a_y1', 'a_y2', 'a_3', 'b_pi', 'b_y', 'sigma_y~', 'sigma_pi', 'sigma_y*']
print('=== theta (最終) ===')
for nm, v in zip(names, theta):
    print(f'  {nm:10} = {v:.12f}')
print(f'\nlogL = {ll:.6f}   P00 shape {P00.shape}')

# 検証系列
z_s = st['smoothed']['xi_tT'][:, 5]
z_f = st['filtered']['xi_tt'][:, 5]
rs = st['smoothed']['xi_tT'][:, 3]*4 + st['smoothed']['xi_tT'][:, 5]
rf = st['filtered']['xi_tt'][:, 3]*4 + st['filtered']['xi_tt'][:, 5]
gap_s = y[:, 0] - st['smoothed']['xi_tT'][:, 0]
gap_f = y[:, 0] - st['filtered']['xi_tt'][:, 0]
hp, _ = rc.hp_filter(lo, 36000.0)
gap_hp = (lo - hp)*100.0
rmse = float(np.sqrt(np.mean((gap_s - gap_hp[4:])**2)))
print(f'\n産出ギャップ(古典): rms={np.sqrt(np.mean(gap_s**2)):.3f}%  max|.|={np.max(np.abs(gap_s)):.2f}%  RMSE vs HP36={rmse:.4f}')
print(f'r*: range {rs.min():.3f}%..{rs.max():.3f}%  std={rs.std():.3f}  tail={rs[-1]:.3f}%')
print(f'z状態 (r*分解): rms={np.sqrt(np.mean(z_s**2)):.5f}')

def qidx(y_, q): return (y_-1995)*4 + (q-1) - 4
KEY = [('1996Q4',1996,4),('2000Q4',2000,4),('2008Q4',2008,4),('2013Q4',2013,4),('2019Q4',2019,4),
       ('2020Q2',2020,2),('2021Q2',2021,2),('2022Q2',2022,2),('2024Q4',2024,4),('2026Q2',2026,2)]
print('\n' + ' '.join(f'{k[0]:>8}' for k in KEY))
print('model  ' + ' '.join(f'{gap_s[qidx(y_,q)]:8.2f}' for _,y_,q in KEY))
print('HP     ' + ' '.join(f'{gap_hp[qidx(y_,q)+4]:8.2f}' for _,y_,q in KEY))
print('r*     ' + ' '.join(f'{rs[qidx(y_,q)]:8.2f}' for _,y_,q in KEY))

# VBA 貼り付け用テキスト
def fmt(v): return repr(float(v))
out = ['--- VBA FixCoeff 置換用 ---']
out.append('    p(0) = ' + fmt(theta[0]))
out.append('    p(1) = ' + fmt(theta[1]))
out.append('    p(2) = ' + fmt(theta[2]))
out.append('    p(3) = ' + fmt(theta[3]))
out.append('    p(4) = ' + fmt(theta[4]))
out.append('    p(5) = ' + fmt(theta[5]))
out.append('    p(6) = ' + fmt(theta[6]))
out.append('    p(7) = ' + fmt(theta[7]))
out.append('--- P00 ---')
for i in range(7):
    out.append('    ' + ' : '.join(fmt(P00[i, j]) for j in range(7)))
out.append('--- 検証系列 (Excel 検証用) ---')
out.append('rstar_smoothed, rstar_filtered, gap_classic_smoothed, gap_classic_filtered, z_smoothed')
for i in range(len(rs)):
    out.append(f'{i},{rs[i]:.12f},{rf[i]:.12f},{gap_s[i]:.12f},{gap_f[i]:.12f},{z_s[i]:.12f}')
txt = '\n'.join(out)
open(r'C:\Users\dai86\Temp\rstar_lab\vba_constants.txt', 'w', encoding='utf-8').write(txt)
print('\nsaved vba_constants.txt (VBA 貼り付け用)')

json.dump(dict(lam_g=LAMBDA_G, lam_z=LAMBDA_Z, sy_floor=SY_FLOOR, theta=theta.tolist(),
               P00=P00.tolist(), logL=float(ll), rmse_hp=rmse,
               rstar_smoothed=rs.tolist(), rstar_filtered=rf.tolist(),
               gap_classic_smoothed=gap_s.tolist(), gap_classic_filtered=gap_f.tolist(),
               z_smoothed=z_s.tolist()),
          open(r'C:\Users\dai86\Temp\rstar_lab\final_v5.json', 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)
print('saved final_v5.json')
