# -*- coding: utf-8 -*-
"""
lz_grid.py: 日本の HLW stage3 で λ_z を走査し、
  (a) 尤度プロファイル (b) ギャップの識別可否 (c) r* の動学
を測定する。入力は CPI.xlsm の Rstar シート（VBA と同じ入力列）。
"""
import sys, time, numpy as np, openpyxl
sys.stdout.reconfigure(encoding='utf-8')
import importlib.util
spec = importlib.util.spec_from_file_location('rc', r'C:\Users\dai86\Downloads\CPI予想ツール\開発用\rstar_check.py')
rc = importlib.util.module_from_spec(spec); spec.loader.exec_module(rc)

XL = r'C:\Users\dai86\Downloads\CPI予想ツール\CPI.xlsm'

def load_panel():
    wb = openpyxl.load_workbook(XL, read_only=True)
    ws = wb['Rstar']
    rows = []
    for row in ws.iter_rows(min_row=13, max_row=200, max_col=7, values_only=True):
        if row[0] is None: break
        rows.append(row)
    wb.close()
    lo = np.array([r[2] for r in rows], float)
    inf = np.array([r[3] for r in rows], float)
    rr = np.array([r[5] for r in rows], float)
    nr = np.array([r[6] for r in rows], float)
    return lo, inf, rr, nr

lo, inf, rr, nr = load_panel()
print(f'n = {len(lo)}  ({2019+0})  T = {len(lo)-4}')

LAMBDA_G = 0.192798
A3 = -0.0025
B2 = 0.025

def run(lz, estimate=True, lg=LAMBDA_G):
    t0 = time.time()
    est = rc.rstar_stage3(lo, inf, rr, nr, lg, lz,
                          a3_constraint=A3, b2_constraint=B2,
                          xi00=None, P00=None, estimate=estimate)
    s3 = est
    th = s3['theta']
    zf = s3['states']['filtered']['xi_tt'][:, 5]
    zs = s3['states']['smoothed']['xi_tT'][:, 5]
    rf = s3['rstar_filtered']
    rs = s3['rstar_smoothed']
    return dict(lz=lz, ll=s3['log_likelihood'], theta=th, zf=zf, zs=zs, rf=rf, rs=rs,
                sigma_ystar=th[7], sigma_ytilde=th[5], sigma_pi=th[6],
                a3=th[2], b_pi=th[3], b_y=th[4], a_y1=th[0], a_y2=th[1], secs=time.time()-t0)

print('\n=== λ_z グリッド × stage3 MLE (a3<=-0.0025, b2>=0.025) ===')
print(f'{"lam_z":>8} {"logL":>10} {"sigma_y*":>9} {"sigma_y~":>9} {"gap_rms":>8} {"gap_max":>8} '
      f'{"r*2019Q4":>9} {"r*2020Q2":>9} {"r*tail":>8} {"secs":>5}')
results = []
for lz in (0.0, 0.0005, 0.001, 0.002, 0.005, 0.01, 0.02):
    r = run(lz, estimate=True)
    results.append(r)
    print(f'{lz:8.4f} {r["ll"]:10.2f} {r["sigma_ystar"]:9.4f} {r["sigma_ytilde"]:9.4f} '
          f'{np.sqrt(np.mean(r["zs"]**2)):8.3f} {np.max(np.abs(r["zs"])):8.3f} '
          f'{r["rs"][95]:9.3f} {r["rs"][98]:9.3f} {r["rs"][-1]:8.3f} {r["secs"]:5.1f}')

print('\n=== 最良 λ_z の θ ===')
best = max(results, key=lambda r: r['ll'])
names = ['a_y1', 'a_y2', 'a_3', 'b_pi', 'b_y', 'sigma_y~', 'sigma_pi', 'sigma_y*']
for n, v in zip(names, best['theta']):
    print(f'  {n:10} = {v:.9f}')
print(f'  logL = {best["ll"]:.4f}  (lam_z = {best["lz"]})')

# 参考: λ_z=0 で θ だけ再推定した場合（v4相当）
print('\n=== 参考: λ_z=0 の MLE 結果 ===')
r0 = results[0]
for n, v in zip(names, r0['theta']):
    print(f'  {n:10} = {v:.9f}')

np.save(r'C:\Users\dai86\Temp\rstar_lab\lz_grid_results.npy',
        np.array([{k: (v if not isinstance(v, np.ndarray) else v.tolist()) for k, v in r.items()} for r in results],
                 dtype=object), allow_pickle=True)
print('\nsaved lz_grid_results.npy')
