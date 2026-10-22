# -*- coding: utf-8 -*-
"""
exp_variants.py: 期待インフレ定義を変えたときの
  (a) stage2 median-unbiased λ_z  (b) stage3 尤度プロファイルの平坦さ
  (c) σ_y* が下限から動くか  (d) ギャップのRMS
を比較する。log_output と inflation は Rstar シートの実データを使用。
"""
import sys, time, numpy as np, openpyxl
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
# 期待CPI (core) forecast: メインパネル or 推計履歴 から取得できるか確認
wb.close()

lo = np.array([r[2] for r in rows], float)
infl = np.array([r[3] for r in rows], float)
exp_infl_cur = np.array([r[4] for r in rows], float)
rr_cur = np.array([r[5] for r in rows], float)
nr = np.array([r[6] for r in rows], float)
n = len(lo)
print(f'n={n}  T={n-4}')

# ---- 期待インフレの候補 ----
def exp_4q_avg(infl):
    out = np.full(len(infl), np.nan)
    for t in range(3, len(infl)):
        w = infl[t-3:t+1]
        if not np.isnan(w).any(): out[t] = w.mean()
    return out

def exp_adaptive(infl, k=0.5):
    out = np.full(len(infl), np.nan)
    e = np.nan
    for t in range(len(infl)):
        if np.isnan(infl[t]): continue
        if np.isnan(e):
            e = infl[t]
        else:
            e = e + k*(infl[t]-e)
        out[t] = e
    return out

def exp_ma2(infl):
    out = np.full(len(infl), np.nan)
    for t in range(1, len(infl)):
        w = infl[t-1:t+1]
        if not np.isnan(w).any(): out[t] = w.mean()
    return out

def exp_forward4(infl):
    """4四半期先までの平均（事前の期待の代替、academically 'rational' proxy）"""
    out = np.full(len(infl), np.nan)
    for t in range(len(infl)-4):
        w = infl[t+1:t+5]
        if not np.isnan(w).any(): out[t] = w.mean()
    return out

variants = {
    '4Q平均 (現行)': exp_4q_avg(infl),
    '2期移動平均': exp_ma2(infl),
    '適応的 (k=0.5)': exp_adaptive(infl, 0.5),
    '適応的 (k=0.3)': exp_adaptive(infl, 0.3),
    '前方4期平均': exp_forward4(infl),
}

LAMBDA_G = 0.192798
A3, B2 = -0.0025, 0.025

def stage2_lam_z(lo, infl, rr):
    """Python 実装の stage2 (median unbiased) 相当"""
    s2 = rc.rstar_stage2(lo, infl, rr, LAMBDA_G, estimate=False)
    r = rc.median_unbiased_lambda_z(s2['y'], s2['x'])
    return r[0] if isinstance(r, tuple) else r

def profile(lo, infl, rr, nr, lz_list):
    out = []
    for lz in lz_list:
        est = rc.rstar_stage3(lo, infl, rr, nr, LAMBDA_G, lz,
                              a3_constraint=A3, b2_constraint=B2, estimate=True)
        th = est['theta']
        zs = est['states']['smoothed']['xi_tT'][:, 5]
        out.append((lz, est['log_likelihood'], th[7], th[5], np.sqrt(np.mean(zs**2)),
                    np.max(np.abs(zs)), est['rstar_smoothed'][-1], est['rstar_smoothed'][95]))
    return out

lz_list = (0.0, 0.0005, 0.001, 0.002, 0.005, 0.01)
for name, exp in variants.items():
    rr = nr - exp
    if np.isnan(rr).any():
        rr = np.where(np.isnan(rr), 0.0, rr)
    lz2 = stage2_lam_z(lo, infl, rr)
    print(f'\n### {name}  (stage2 median-unbiased λ_z = {lz2:.6f})')
    print(f'{"lam_z":>8} {"logL":>10} {"sigma_y*":>9} {"sigma_y~":>9} {"gap_rms":>8} {"gap_max":>8} {"r*tail":>8} {"r*2019Q4":>9}')
    res = profile(lo, infl, rr, nr, lz_list)
    for lz, ll, sy, sy2, gr, gm, rt, r19 in res:
        print(f'{lz:8.4f} {ll:10.2f} {sy:9.4f} {sy2:9.4f} {gr:8.3f} {gm:8.3f} {rt:8.3f} {r19:9.3f}')
    lls = [r[1] for r in res]
    print(f'  → 尤度平坦度 (max-min) = {max(lls)-min(lls):.3f}')
