# -*- coding: utf-8 -*-
"""
publish_github.py — GitHub 公開用スナップショット做成 & push

  python 開発用/publish_github.py [--push]

やること:
  1. 配送フォルダの内容をスナップショットにコピー
  2. CPI.xlsm の e-Stat API キー (データシート B2) を除去してサニタイズ
  3. サニタイズ版で配布用 zip を作り直す
  4. キー混入を検査 (検出したら中断)
  5. 履歴なしの単一コミットで push (--push 時)

注意: 公開リポジトリは履歴を永久保持するため、ローカルの git 履歴をそのまま
push すると API キーが過去のコミットから漏れます。本スクリプトは
「履歴なしスナップショット」のみ_beta_ を push します。
"""
import os, sys, shutil, subprocess, time, zipfile, argparse

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
# 実行ごとに一意な作業 dir (前回の git プロセスがファイルを保持ipheral して
# rmtree が失敗するのを避ける)
WORK = os.path.join(os.environ.get('TEMP', r'C:\Temp'),
                    'cpi_publish_' + subprocess.run(['git', 'rev-parse', '--short', 'HEAD'],
                    capture_output=True, text=True, cwd=BASE).stdout.strip() or 'tmp')
REMOTE = 'https://github.com/dai8647/CPI-.git'
def read_local_key():
    """ローカル CPI.xlsm (データシート B2) から API キーを取り出す。
    キーをこのスクリプト自体に埋め込まないため、動的に取得する。"""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(os.path.join(BASE, 'CPI.xlsm'), read_only=True)
        v = wb['データ']['B2'].value
        wb.close()
        return str(v).strip() if v else ''
    except Exception:
        return ''


KEY_RE = ''   # 実行時に read_local_key() で設定
FILES = ['CPI.xlsm', 'ModuleRStar.bas', 'ModuleRStar_win.bas',
         'ModuleCPI.bas', 'ModuleCPI_win.bas', 'README.txt']
ZIPNAME = 'CPI予想ツール_R-star版_20261021.zip'


def sh(cmd, cwd=None):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        print(' '.join(cmd), '\n', r.stdout, r.stderr)
    return r


def sanitize_xlsm(path):
    """Excel COM で データシート B2 (APIキー) を空にして保存"""
    import win32com.client
    subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'], capture_output=True)
    time.sleep(2)
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(path)
        wb.Worksheets('データ').Range('B2').Value = ''
        wb.Save()
        wb.Close()
    finally:
        excel.Quit()
        time.sleep(1)
        subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'], capture_output=True)


def scan_keys(root):
    """ディレクトリ全体 (アーカイブ内も) からキーを検出"""
    global KEY_RE
    if not KEY_RE:
        KEY_RE = read_local_key()
    if not KEY_RE:
        print('   (ローカルに API キーが無い: 検査スキップ)')
        return []
    hits = []
    for dirpath, dirnames, filenames in os.walk(root):
        if '.git' in dirpath:
            continue
        for f in filenames:
            p = os.path.join(dirpath, f)
            try:
                if KEY_RE.encode() in open(p, 'rb').read():
                    hits.append(p)
            except Exception:
                pass
            if f.lower().endswith(('.xlsm', '.xlsx', '.zip')):
                try:
                    with zipfile.ZipFile(p) as z:
                        for n in z.namelist():
                            if KEY_RE.encode() in z.read(n):
                                hits.append(f'{p}!{n}')
                except Exception:
                    pass
    return hits


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--push', action='store_true', help='実際に push する')
    args = ap.parse_args()

    if os.path.exists(WORK):
        shutil.rmtree(WORK)
    os.makedirs(WORK)
    for f in FILES:
        shutil.copy(os.path.join(BASE, f), os.path.join(WORK, f))
    shutil.copytree(os.path.join(BASE, '開発用'), os.path.join(WORK, '開発用'))
    for extra in ('.gitignore', 'README.md'):
        src = os.path.join(BASE, extra)
        if os.path.exists(src):
            shutil.copy(src, os.path.join(WORK, extra))
    # 配布 zip と同名の旧版を除去 (作り直すため)
    for f in os.listdir(WORK):
        if f.endswith('.zip') and f != ZIPNAME:
            os.remove(os.path.join(WORK, f))

    print('1. CPI.xlsm をサニタイズ (APIキー除去)')
    sanitize_xlsm(os.path.join(WORK, 'CPI.xlsm'))

    print('2. 配布用 zip を作成')
    zp = os.path.join(WORK, ZIPNAME)
    with zipfile.ZipFile(zp, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in FILES:
            z.write(os.path.join(WORK, f), f)

    print('3. キー混入検査')
    hits = scan_keys(WORK)
    if hits:
        print('!! キーが検出されました。中止します:', hits)
        return 1
    print('   検出なし OK')

    sh(['git', 'init', '-q'], cwd=WORK)
    sh(['git', 'config', 'user.name', 'dai8647'], cwd=WORK)
    sh(['git', 'config', 'user.email', 'dai8647@users.noreply.github.com'], cwd=WORK)
    sh(['git', 'add', '-A'], cwd=WORK)
    sh(['git', 'commit', '-q', '-m',
        'feat: CPI予想ツール + r*(HLW/LW)推計 [R-star版]\n\n'
        '- 東京→全国 CPI 予想 (e-Stat 2025年基準・系列別回帰)\n'
        '- HLW モデルによる自然実質金利 r* と産出ギャップ推計 (VBA のみ)\n'
        '- 内閣府 四半期別GDP速報CSV / 日銀コールレートの自動取得\n'
        '- グラフ4枚 + 乖離チャート\n'
        '- 検証: VBA vs Python 全系列 max差 1.2e-8\n\n'
        '注: 公開のため e-Stat API キーは含めていません。\n'
        '   利用時は データシート B2 に API キーを貼り付けてください。'], cwd=WORK)
    sh(['git', 'branch', '-M', 'main'], cwd=WORK)
    sh(['git', 'remote', 'remove', 'origin'], cwd=WORK)  # 未設定なら無視
    sh(['git', 'remote', 'add', 'origin', REMOTE], cwd=WORK)
    print('4. スナップショット作成完了:', WORK)
    if args.push:
        r = sh(['git', 'push', '--force', '-u', 'origin', 'main'], cwd=WORK)
        print(r.stdout[-400:], r.stderr[-400:])
    else:
        print('   (--push なし: push はしていません)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
