========================================================
 開発用フォルダの説明 (配布物には不要です)
========================================================

配布するのは 1 つ上のフォルダにある次の 5 つだけです:
    CPI.xlsm … 本体 (Rstar シート・グラフ・ボタン込み)
    ModuleRStar.bas / ModuleRStar_win.bas … R*推計モジュール (UTF-8 / CP932)
    ModuleCPI.bas  / ModuleCPI_win.bas   … CPI本体モジュール (同上)
    README.txt … 使い方・データ源・再推定手順
このフォルダ (開発用) の中身は、推計をやり直すときや検証に使うもので、
Excel に入れて配る必要はありません。

--------------------------------------------------------
1. Python スクリプト (この PC でのみ使用)
--------------------------------------------------------
rstar_check.py       HLW の基準実装 (FRBNY の R コードの移植)。stage1/2/3、
                     メディアン不偏推定量、HP、カルマン、自己チェック
rstar_fetch.py       生データ取得 (内閣府 GDP xlsx / 日銀コール) → rstar_data.csv
lw_check.py          LW モデル (加重成長率+mu) の独立計算 → rstar_lw_check.csv
cpi_check.py         既存 CPI ツールの回帰の基準実装 (R* とは独立)
（gen_fixed_coeff_v2.py 相当の処理は rstar_check.py の関数を呼んで実行）

--------------------------------------------------------
2. 突合用 CSV / JSON (VBA と Python の一致確認用)
--------------------------------------------------------
rstar_fixed_coeff.csv      旧構成 (GDPデフレーター) の固定係数
rstar_fixed_coeff_v2.csv   ★現行 (民間消費デフレーター) の固定係数
                           → VBA FixCoeff と LAMBDA_G/LAMBDA_Z の根拠
rstar_inputs.csv           旧構成の入力テーブル 121 四半期
rstar_inputs_v2.csv        ★現行の入力テーブル (Fetch の突合先)
rstar_vba_check.csv        旧構成の r* 参照系列
rstar_vba_check_v2.csv     ★現行の r* 参照系列 (VBA と max差 8.8e-6)
rstar_lw_check.csv         LW モデルの参照系列 (VBA と完全一致)
rstar_data.csv             四半期パネル (GDP/デフレーター/コール/CPI, q_idx 付き)
rstar_ground_truth.json    旧構成の MLE 結果 (theta, xi00, P00, r* 系列)
cpi_ground_truth.json      既存 CPI ツールの回帰の期待値

--------------------------------------------------------
3. 元データ (取得済みキャッシュ)
--------------------------------------------------------
qom1rn.xlsx / qom1dn.xlsx / qom1n.xlsx / qom2.xlsx
    内閣府 四半期別GDP (実質・デフレーター・名目・2次速報)
    ※ VBA は実行時に TEMP へ再取得します (キー不要)
s2rn.xlsx / s2dn.xlsx / s22rn.xlsx   参考: 別表 (使用していない)
real.xlsx / defl.xlsx / nom.xlsx     取得スクリプトの作業コピー
boj_call.csv                         日銀 STRACLUCON 月次 (393ヶ月)
probe_call.xlsx / probe_sbs.xlsx     調査時の試行ファイル
boj_api_manual.pdf                   日銀 時系列検索 API の説明書

--------------------------------------------------------
4. FRBNY 本家コード (参照用)
--------------------------------------------------------
hlw_code/HLW_Code/   ニューヨーク連銀の HLW 実装 (R)
    rstar.stage1-3.R / unpack.parameters.*.R / kalman.*.R /
    median.unbiased.estimator.*.R / prepare.rstar.data.us.R
    → rstar_check.py と ModuleRStar.bas はこのコードの移植です

--------------------------------------------------------
5. 推計をやり直す手順 (四半期ごと)
--------------------------------------------------------
1) このフォルダで  python rstar_fetch.py --force    (パネル再取得)
2) trace: rstar_check.py の関数を使い stage1/2 (estimate=False) で
   λg, λz を算出 → stage3 を MLE → rstar_fixed_coeff_v2.csv を更新
3) 1つ上の ModuleRStar.bas の FixCoeff (theta, xi00, P00) と
   LAMBDA_G / LAMBDA_Z 定数を新値に張り替える
4) ModuleRStar_win.bas を再生成 (UTF-8 → CP932) して CPI.xlsm に再インポート
5) CPI.xlsm の Rstar シートで「R*推計更新」を押し、rstar_vba_check_v2.csv と突合
   (最大差が 1e-4 未満なら一致)

--------------------------------------------------------
4. 2026-10-21 追加 (v5 較正・産出ギャップ識別)
--------------------------------------------------------
rstar_fixed_coeff_v5.json   ★現行の固定係数 (theta/P00/logL/r*・gap 系列)
                           → VBA ModuleRStar の FixCoeff に張り付けた値
lz_grid.py                 λ_z グリッド × stage3 MLE (尤度プロファイル・gap 識別確認)
grid2d.py                  (λ_z, sigma_y*下限) の2D較正 (RMSE vs HP ギャップ)
exp_variants.py            期待インフレ定義 5 通りの比較 (識別改善せずを実証)
lz_calib.py                HP フィルタ ギャップとの整合をみた λ_z 較正
finalize.py                採用構成で theta/P00 を生成し VBA 貼り付け用テキストを出力
boj_gap.xlsx               日銀「需給ギャップ・潜在成長率」取得結果 (キャッシュ)
  再推定手順:
   1. CPI.xlsm の Rstar シートから入力列 (C〜G) を取り出すか、
      ModuleRStar の RStar_FetchData を実行して最新系列を取得
   2. finalize.py を実行 (lam_z / sigma_y* 下限は表の設定値に合わせる)
   3. 出た theta/P00 を ModuleRStar.bas の FixCoeff に張り替え
   4. ModuleRStar_win.bas を再生成 (CP932 化) し CPI.xlsm に再インポート
  較正の判断基準:
   - セルフテストの「gap vs HP RMSE」が 0.5% を大きく超えないこと
   - r* の水準が 0.5〜1.5% のレンジに収まること
   - 日銀レビュー 2026-J-5 の r* 推計レンジ (▲0.9%〜+0.5%) との位置関係
