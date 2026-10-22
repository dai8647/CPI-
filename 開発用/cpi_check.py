# -*- coding: utf-8 -*-
"""
CPI予想ツール 検証スクリプト（Python側グラウンドトゥルース）
- e-Stat「消費者物価指数（2025年基準）」(statsDataId=0004052037) から
  東京都区部(13100)・全国(00000) の 前年比(tab=3)・指数(tab=1) を取得
- 東京→全国 同月前年比の線形回帰（系列別）を複数窓でバックテストし最良窓を確認
- Excel側(VBA)の計算結果とのクロスチェックに使う
実行: python cpi_check.py
"""
import json, sys, io, os, urllib.request, datetime
import statistics

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# e-Stat APIキーは環境変数 ESTAT_API_KEY で渡す (ハードコードしない)
KEY = os.environ.get('ESTAT_API_KEY', '')
if not KEY:
    try:
        KEY = open(os.path.join(os.path.dirname(__file__), '..', 'CPI.xlsm'), 'rb').read() and ''
    except Exception:
        KEY = ''
if not KEY:
    import openpyxl
    _wb = openpyxl.load_workbook(os.path.join(os.path.dirname(__file__), '..', 'CPI.xlsm'), read_only=True)
    KEY = str(_wb['データ']['B2'].value or '').strip()
    _wb.close()
if not KEY:
    print('ERROR: e-Stat APIキーがありません。環境変数 ESTAT_API_KEY を設定するか、'
          'CPI.xlsm の データシート B2 にキーを入力してください。')
    sys.exit(2)
BASE = 'https://api.e-stat.go.jp/rest/3.0/app/json/getStatsData'
SID = '0004052037'          # 2025年基準（2026-10公開、基準改定後の本命表）
AREA_TOKYO = '13100'
AREA_NATL = '00000'
CATS = [('0001', '総合'), ('0161', 'コア'), ('0178', 'コアコア')]


def call(params):
    url = BASE + '?' + '&'.join(f'{k}={v}' for k, v in params.items())
    with urllib.request.urlopen(url, timeout=60) as r:
        return json.load(r)


def fetch_series(area, cat, tab):
    """月次の系列を {YYYYMM: float} で返す（年平均等は除外）"""
    d = call({'appId': KEY, 'statsDataId': SID, 'cdArea': area, 'cdCat01': cat,
              'cdTab': tab, 'metaGetFlg': 'N', 'cntGetFlg': 'N', 'limit': '100000'})
    res = d['GET_STATS_DATA']['RESULT']
    if res['STATUS'] != 0:
        raise RuntimeError(f'API error: {res["ERROR_MSG"]}')
    vals = d['GET_STATS_DATA']['STATISTICAL_DATA']['DATA_INF'].get('VALUE', [])
    if isinstance(vals, dict):
        vals = [vals]
    out = {}
    for v in vals:
        t = v.get('@time', ''); val = v.get('$', '')
        y, mm = t[:4], t[6:8]
        if '01' <= mm <= '12' and val not in ('', '***', '-'):
            try:
                out[int(y + mm)] = float(val)
            except ValueError:
                pass
    return out


def ols(xs, ys):
    """最小二乗法 y = a*x + b → (a, b)"""
    n = len(xs)
    mx, my = sum(xs) / n, sum(ys) / n
    sxx = sum((x - mx) ** 2 for x in xs)
    if sxx == 0:
        return None
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    a = sxy / sxx
    return a, my - a * mx


def backtest(months, ty, ny, window, span):
    """expanding-window バックテスト: 直近 window か月で回帰→翌月を1点予想"""
    errs = []
    for m in span:                       # 予想対象月
        hist = [k for k in months if k < m and k in ty and k in ny]
        hist = hist[-window:] if window else hist   # None=全履歴
        if len(hist) < 12:
            continue
        fit = ols([ty[k] for k in hist], [ny[k] for k in hist])
        if fit is None or m not in ny or m not in ty:
            continue
        pred = fit[0] * ty[m] + fit[1]
        errs.append(ny[m] - pred)
    if not errs:
        return None
    mae = sum(abs(e) for e in errs) / len(errs)
    rmse = (sum(e * e for e in errs) / len(errs)) ** 0.5
    bias = sum(errs) / len(errs)
    return mae, rmse, bias, len(errs), errs


def main():
    print('e-Stat 2025年基準表から取得中...', flush=True)
    data = {}
    for area, aname in [(AREA_TOKYO, '東京'), (AREA_NATL, '全国')]:
        for cat, cname in CATS:
            yoy = fetch_series(area, cat, '3')
            data[f'{aname}_{cname}'] = yoy
            ks = sorted(yoy)
            print(f'  {aname}_{cname}: {len(ks)}か月  {ks[0]//100}/{ks[0]%100:02d} ～ {ks[-1]//100}/{ks[-1]%100:02d}')
    # 指数レベル（連鎖推計用）
    natl_idx = fetch_series(AREA_NATL, '0001', '1')
    tokyo_idx = fetch_series(AREA_TOKYO, '0001', '1')
    print(f'  全国指数: {len(natl_idx)}か月, 東京指数: {len(tokyo_idx)}か月')

    months = sorted(set(data['東京_総合']) & set(data['全国_総合'])
                    & set(data['東京_コア']) & set(data['全国_コア'])
                    & set(data['東京_コアコア']) & set(data['全国_コアコア']))
    print(f'  東京・全国の共通月: {len(months)}か月 ({months[0]}～{months[-1]})')

    # ---- バックテスト: 窓の選定（直近10年分を1点先予想で検証）----
    span = [m for m in months if m >= 201701]
    print('\n=== バックテスト（2017/01以降を1点先予想で検証、MAE %pt）===')
    best = {}
    for cname in ['総合', 'コア', 'コアコア']:
        ty, ny = data[f'東京_{cname}'], data[f'全国_{cname}']
        print(f'--- {cname}')
        for w in [12, 24, 36, 60, None]:
            r = backtest(months, ty, ny, w, span)
            if r:
                tag = f'{w}か月' if w else '全履歴'
                print(f'  {tag:6s} MAE={r[0]:.3f} RMSE={r[1]:.3f} bias={r[2]:+.3f} n={r[3]}')
                best.setdefault(cname, []).append((r[0], w))
    print('\n最良窓:', {c: min(v)[1] for c, v in best.items()})

    # ---- 最新1点の予想（全自動更新と同じ手順）----
    # 予想対象 = 東京の最新月（全国は1か月遅れで公表されるため）
    target = max(data['東京_総合'])
    WINDOW = 12
    print(f'\n=== 最新予想（東京{target//100}/{target%100:02d} → 全国{target//100}/{target%100:02d} 予想）===')
    results = {}
    for cname in ['総合', 'コア', 'コアコア']:
        ty, ny = data[f'東京_{cname}'], data[f'全国_{cname}']
        hist = [k for k in months if k < target and k in ty and k in ny][-WINDOW:]
        a, b = ols([ty[k] for k in hist], [ny[k] for k in hist])
        tx = ty[target]
        pred = a * tx + b
        results[cname] = (a, b, tx, pred)
        print(f'  {cname}: α={a:.4f} β={b:+.4f} 東京={tx} → 全国予想 {pred:.2f}%')

    # 指数の連鎖推計: 全国指数(昨年同月) × (1+予想前年比/100)
    ym = target
    yoy_prev = ym - 100                     # 昨年同月 (202509)
    pred_idx = None
    if yoy_prev in natl_idx:
        base_idx = natl_idx[yoy_prev]
        pred_idx = base_idx * (1 + results['総合'][3] / 100)
        print(f'  総合指数推計: {base_idx}(2025/{yoy_prev%100:02d}) × (1+予想YoY/100) = {pred_idx:.1f}')

    # ---- JSONで保存（Excel側とのクロスチェック用）----
    out = {
        'as_of': datetime.datetime.now().isoformat(),
        'statsDataId': SID,
        'target_month': target,
        'window': WINDOW,
        'ols': {c: {'alpha': results[c][0], 'beta': results[c][1],
                    'tokyo': results[c][2], 'pred': results[c][3]} for c in results},
        'pred_index': pred_idx,
        'series_tail': {k: {m: v[m] for m in months[-6:]} for k, v in data.items()},
    }
    with open(__file__.replace('cpi_check.py', 'cpi_ground_truth.json'), 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print('\ncpi_ground_truth.json に保存しました')


if __name__ == '__main__':
    main()
