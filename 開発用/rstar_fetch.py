# -*- coding: utf-8 -*-
"""
R*推計 開発用グラウンドトゥルース生成（配布物ではない・このPC専用）
  1. 内閣府 年次推計 四半期表（名目/実質/デフレーター）を xlsx ダウンロード→1994Q1〜の四半期パネル
  2. 日銀 stat-search Code API で無担保コール翌日物 月次平均 STRACLUCON を取得→四半期平均
  3. CPI.xlsx の データ!F列（全国コア前年比・2013-09〜）を四半期平均で併記
出力: rstar_data.csv  (quarter_start_serial, real_gdp, defl, nom, i_q, pi_defl, pi_cpi)
実行: python rstar_fetch.py [--force]
"""
import io, os, sys, csv, datetime, urllib.request
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
HDR = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126 Safari/537.36'}

GDP_BASE = 'https://www.esri.cao.go.jp/jp/sna/data/data_list/kakuhou/files/2024/tables/'
GDP_FILES = {'nom': '2024qom1n_jp.xlsx', 'real': '2024qom1rn_jp.xlsx', 'defl': '2024qom1dn_jp.xlsx'}
BOJ_URL = ('https://www.stat-search.boj.or.jp/api/v1/getDataCode?format=csv&lang=jp'
           '&db=FM02&code=STRACLUCON&startDate=199401&endDate=202612')
FORCE = '--force' in sys.argv


def fetch(url, dest, binary=True):
    if os.path.exists(dest) and not FORCE:
        return
    req = urllib.request.Request(url, headers=HDR)
    with urllib.request.urlopen(req, timeout=120) as r:
        b = r.read()
    if b[:2] == b'\x1f\x8b':
        import gzip; b = gzip.decompress(b)
    with open(dest, 'wb' if binary else 'wb') as f:
        f.write(b)


def parse_gdp_xlsx(path):
    """実数シートから 1994Q1 起点の四半期系列を {q_idx: value} で返す。
    規則: 年ラベル行=col1が非空、col1 から順に四半期が並ぶ(平成6年 4列, 平成7年 4列…)。
    対象行=ラベルに「国内総生産」を含み数値セル数が最大の行（=GDP合計行、行5）。"""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb['実数']
    rows = list(ws.iter_rows(values_only=True))
    # 年ラベル行の特定
    yr_row = None
    for i, r in enumerate(rows[:12]):
        if r[1] and r[0] is None and isinstance(r[1], str):
            yr_row = i
            break
    # 1列目のラベルから年を拾い年ラベルを四半期indexへ展開
    first_year = None
    year_of_col = {}
    col = 1
    labels = rows[yr_row]
    while col < len(labels) and labels[col] is not None:
        lab = str(labels[col])
        import re
        m = re.match(r'(平成|令和)(\d+)年', lab)
        if m:
            y = (1988 if m.group(1) == '平成' else 2018) + int(m.group(2))
            if first_year is None:
                first_year = y
        col += 1
    # 対象行
    best, best_n = None, 0
    for i, r in enumerate(rows):
        lab = str(r[0]) if r[0] else ''
        if '国内総生産' in lab:
            n = sum(1 for x in r[1:] if isinstance(x, (int, float)) and x not in (None, ''))
            if n > best_n:
                best, best_n = i, n
    vals = {}
    q = 0
    for j, x in enumerate(rows[best][1:], start=1):
        if isinstance(x, (int, float)) and x is not None:
            vals[q] = float(x)
            q += 1
    return first_year, vals, best_n


def parse_boj_call(path):
    """BOJ Code API csv(cp932) → {YYYYMM: 平均}。ヘッダ行はSTATUS等、データ部は DATE/VALUE 列"""
    with open(path, 'rb') as f:
        raw = f.read()
    txt = raw.decode('cp932', 'ignore')
    lines = txt.splitlines()
    # データ部ヘッダ行（"DATE","VALUE" 等）を探す
    hdr_i = None
    for i, l in enumerate(lines):
        if l.startswith('SERIES_CODE,'):
            hdr_i = i
            break
    out = {}
    if hdr_i is None:
        return out
    # データ部: SERIES_CODE,NAME,UNIT,FREQ,CAT,LAST_UPDATE,SURVEY_DATES(YYYYMM),VALUES(value)
    header = [h.strip().strip('"') for h in lines[hdr_i].split(',')]
    di = next(i for i, h in enumerate(header) if h.upper() == 'SURVEY_DATES')
    vi = next(i for i, h in enumerate(header) if h.upper() == 'VALUES')
    for l in lines[hdr_i + 1:]:
        parts = l.split(',')
        if len(parts) < vi + 1:
            continue
        d, v = parts[di].strip().strip('"'), parts[vi].strip().strip('"')
        if len(d) == 6 and d.isdigit() and v.replace('.', '', 1).replace('-', '', 1).isdigit():
            out[d[:6]] = float(v)
    return out


def quarter_from_month(ym):
    y, m = int(ym[:4]), int(ym[4:])
    return y, (m - 1) // 3 + 1


def main():
    # 1. GDP
    gdp = {}
    first_years = {}
    for key, fn in GDP_FILES.items():
        p = os.path.join(HERE, key + '.xlsx')
        try:
            fetch(GDP_BASE + fn, p)
            fy, vals, n = parse_gdp_xlsx(p)
            gdp[key] = vals
            first_years[key] = fy
            print('GDP', key, 'from', fy, 'n=', n)
        except Exception as e:
            print('GDP parse fail', key, repr(e))
            return 1
    start = first_years['real']
    nq = min(max(gdp[k]) + 1 for k in gdp)
    # 2. BOJ call
    bojp = os.path.join(HERE, 'boj_call.csv')
    fetch(BOJ_URL, bojp, binary=False)
    call = parse_boj_call(bojp)
    print('call months:', len(call))
    # 3. CPI (workbook)
    cpi = {}
    try:
        from openpyxl import load_workbook
        _cpi = os.path.join(HERE, 'CPI.xlsm')
        if not os.path.exists(_cpi):
            _cpi = os.path.join(os.path.dirname(HERE), 'CPI.xlsm')  # ルートの配布物
        wb = load_workbook(_cpi, data_only=True)
        ws = wb['データ']
        for r in ws.iter_rows(min_row=6, values_only=True):
            if r[0] and isinstance(r[0], datetime.datetime) and isinstance(r[5], (int, float)):
                cpi[r[0].strftime('%Y%m')] = float(r[5])
        print('cpi months:', len(cpi))
    except Exception as e:
        print('CPI.xlsm not available:', repr(e))

    # quarterly panel
    out = []
    q = 0
    y, qq = start, 1
    while q < nq:
        real = gdp['real'].get(q)
        defl = gdp['defl'].get(q)
        nom = gdp['nom'].get(q)
        months = [f'{y}{(qq-1)*3+m+1:02d}' for m in range(3)]
        iv = [call[m] for m in months if m in call]
        i_q = sum(iv) / len(iv) if len(iv) >= 2 else (iv[0] if iv else None)
        pv = [cpi[m] for m in months if m in cpi]
        pi_cpi = sum(pv) / len(pv) if len(pv) >= 2 else None
        pi_defl = (defl / gdp['defl'][q - 4] - 1) * 100 if (defl and q >= 4 and gdp['defl'].get(q - 4)) else None
        if real is not None and defl is not None:
            out.append([q, y, qq, real, defl, nom, i_q, pi_defl, pi_cpi])
        q += 1
        qq += 1
        if qq > 4:
            qq = 1; y += 1
    dest = os.path.join(HERE, 'rstar_data.csv')
    with open(dest, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['q_idx', 'year', 'qnum', 'real_gdp', 'defl', 'nom', 'i_q', 'pi_defl', 'pi_cpi'])
        w.writerows(out)
    print('wrote', dest, len(out), 'quarters, last', out[-1][1:3] if out else None)
    return 0


if __name__ == '__main__':
    sys.exit(main())
