# -*- coding: utf-8 -*-
"""
lz_calib.py: λ_z の較正。
  目標: モデルの産出ギャップを、古典的な HP フィルタのギャップ
        (log_output - HPトレンド) の分散に整合させる。
  λ_z のみを走査し、θ は各 λ_z で stage3 MLE により再推計する。
"""
import sys, numpy as np, openpyxl
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
n = len(lo); T = n - 4

# ---- ベンチマーク1: HP フィルタのギャップ (log_output - HPトレンド) ----
hp_trend, hp_cycle = rc.hp_filter(lo, 36000.0)
hp_trend1600, hp_cycle1600 = rc.hp_filter(lo, 1600.0)
gap_hp36 = (lo - hp_trend) * 100.0
gap_hp16 = (lo - hp_trend1600) * 100.0
print(f'HP(36000) ギャップ: rms={np.sqrt(np.mean(gap_hp36**2)):.3f}%  max|.|={np.max(np.abs(gap_hp36)):.2f}%  '
      f'(2020Q2={gap_hp36[99]:.2f}%, 2022Q2={gap_hp36[107]:.2f}%)')
print(f'HP(1600)  ギャップ: rms={np.sqrt(np.mean(gap_hp16**2)):.3f}%  max|.|={np.max(np.abs(gap_hp16)):.2f}%  '
      f'(2020Q2={gap_hp16[99]:.2f}%, 2022Q2={gap_hp16[107]:.2f}%)')

# ---- ベンチマーク2: OLS トレンドとの乖離 (HLW 内部の output_gap 定義) ----
xog = np.column_stack([np.ones(T+4), np.arange(1, T+5, dtype=float)])
b = np.linalg.solve(xog.T @ xog, xog.T @ lo)
gap_ols = (lo - xog @ b) * 100.0
print(f'OLSトレンド乖離: rms={np.sqrt(np.mean(gap_ols**2)):.3f}%  max|.|={np.max(np.abs(gap_ols)):.2f}%')

LAMBDA_G = 0.192798
A3, B2 = -0.0025, 0.025

print(f'\n=== λ_z 較正 (θ は各 λ_z で MLE 再推計) ===')
print(f'{"lam_z":>8} {"logL":>10} {"sig_y*":>7} {"gap_rms":>8} {"gap_max":>8} {"gap_2020Q2":>10} '
      f'{"gap_2022Q2":>10} {"RMSE vs HP36":>12} {"r*tail":>8} {"r*2019Q4":>9} {"r*2008Q4":>9}')
best = None
for lz in (0.0, 0.001, 0.002, 0.003, 0.004, 0.005, 0.006, 0.008, 0.010, 0.015):
    est = rc.rstar_stage3(lo, infl, rr, nr, LAMBDA_G, lz, a3_constraint=A3, b2_constraint=B2, estimate=True)
    th = est['theta']
    zs = est['states']['smoothed']['xi_tT'][:, 5]
    zf = est['states']['filtered']['xi_tt'][:, 5]
    rmse = float(np.sqrt(np.mean((zs - gap_hp36[4:])**2)))
    print(f'{lz:8.4f} {est["log_likelihood"]:10.2f} {th[7]:7.4f} {np.sqrt(np.mean(zs**2)):8.3f} '
          f'{np.max(np.abs(zs)):8.3f} {zs[101]:10.3f} {zs[109]:10.3f} {rmse:12.3f} '
          f'{est["rstar_smoothed"][-1]:8.3f} {est["rstar_smoothed"][99]:9.3f} {est["rstar_smoothed"][55]:9.3f}')
    if lz == 0.0:
        print('   θ(λz=0) =', np.round(th, 6))
print('\n注: インデックスは results 開始が 1996Q1(=0) → 2020Q2=index 99, 2022Q2=index 107, 2019Q4=index 95, 2008Q4=index 51')
