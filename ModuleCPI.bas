Option Explicit

'================================================================
' 東京CPI → 全国CPI 予想ツール（2025年基準・全自動版）
'----------------------------------------------------------------
' 【データ源】e-Stat「消費者物価指数（2025年基準）」statsDataId=0004052037
'   ・東京都区部(13100)は毎月末に速報公表（全国に約3週間先行）
'   ・全国(00000)は翌月中旬公表
'   ・前年比は基準改定の接続調整済み系列（2020年基準の数値と多少異なる）
' 【ロジック】
'   ① 東京と全国の「同じ月の前年比」を回帰: 全国 = α×東京 + β（系列別に独立推定）
'   ② 回帰窓（直近Nか月）はメインパネルで変更可（既定12か月）
'   ③ 精度は過去10年を「1点先予想」で追いかけたバックテストMAEで検証
'   ④ 総合指数の推計 = 昨年同月の全国指数 × (1+予想前年比/100)（連鎖法）
' 旧Module1の単一α/β流用・指数×0.998決め打ちは廃止。
'================================================================

Public Const SH_MAIN  As String = "メインパネル"
Public Const SH_DATA  As String = "データ"
Public Const SH_HIST  As String = "推計履歴"
Public Const SH_DIAG  As String = "回帰診断"
Public Const SH_ABOUT As String = "説明"

Public Const ESTAT_BASE As String = "https://api.e-stat.go.jp/rest/3.0/app/json/getStatsData"

'--- モジュール内共有データ（FetchAllData が格納） ---
Private dTokyoS As Object, dTokyoC As Object, dTokyoCC As Object
Private dNatS As Object, dNatC As Object, dNatCC As Object
Private dNatIdx As Object, dTokyoIdx As Object
Private mMonths() As Long           ' 東京・全国の共通月（昇順）
Private mNCommon As Long
Private mTarget As Long             ' 予想対象月 = 東京の最新月 (YYYYMM)

Private Const COL_LAST As Long = 9  ' データシート I列=全国総合指数
Private Const NM_TAYLOR As Long = 12   ' テイラールールの予想期間 (ヶ月)

'================================================================
' 0. 画面の再構築（一度だけ実行）
'================================================================
Public Sub RebuildPanels()
    Call BuildMainV2
    Call BuildDiagV2
    Call BuildAboutV2
    Call BuildDataSettings
    Worksheets(SH_HIST).Cells(2, 10).Value = "メモ"
    Worksheets(SH_HIST).Cells(2, 10).Font.name = "メイリオ"
    Worksheets(SH_HIST).Cells(2, 10).Font.Bold = True
    Call SetStatus("ステータス: 画面再構築済み（▶ 全自動更新を実行してください）", RGB(0, 120, 215))
End Sub

'----------------------------------------------------------------
Private Sub BuildMainV2()
    Dim ws As Worksheet
    Set ws = Worksheets(SH_MAIN)
    ws.Unprotect Password:=""
    ws.Cells.Clear

    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 26
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 3
    ws.Rows("1").RowHeight = 8
    ws.Rows("2").RowHeight = 42

    With ws.Range("B2:E2")
        .Merge
        .Value = "東京CPI → 全国CPI 予想ツール（2025年基準・全自動版）"
        .Font.name = "メイリオ"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 84, 166)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Rows("4").RowHeight = 18
    With ws.Range("B4:E4")
        .Merge
        .Value = "出典: 総務省統計局 消費者物価指数（2025年基準・接続済み前年比） | e-Stat API 自動取得"
        .Font.name = "メイリオ"
        .Font.Size = 9
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(230, 240, 255)
    End With

    Call DrawHeader(ws, "B6:E6", "① 東京都区部CPI（速報値・自動取得）", RGB(0, 120, 215))

    Dim lbl1 As Variant, fmt1 As Variant, note1 As Variant
    lbl1 = Array("対象月", "総合（前年比 %）", "生鮮食品除く コア（前年比 %）", _
                 "生鮮食品・エネルギー除く コアコア（前年比 %）", "総合指数（2025=100）")
    note1 = Array("", "% (前年比)", "% (前年比)", "% (前年比)", "指数値")
    Dim r As Long
    For r = 0 To 4
        With ws.Cells(7 + r, 2)
            .Value = lbl1(r)
            .Font.name = "メイリオ": .Font.Size = 10
            .Interior.Color = RGB(240, 246, 255)
        End With
        With ws.Cells(7 + r, 3)
            .NumberFormat = IIf(r = 0, "yyyy/mm", IIf(r = 4, "0.0", "0.00"))
            .Font.name = "メイリオ": .Font.Size = 11: .Font.Bold = True
            .Font.Color = RGB(0, 84, 166)
            .Interior.Color = RGB(255, 255, 220)
            .HorizontalAlignment = xlCenter
        End With
        With ws.Cells(7 + r, 4)
            .Value = note1(r)
            .Font.name = "メイリオ": .Font.Size = 9
            .Font.Color = RGB(120, 120, 120)
            .Interior.Color = RGB(248, 248, 248)
        End With
    Next r

    ws.Rows("12").RowHeight = 10
    Call DrawHeader(ws, "B13:E13", "② 回帰パラメータ（系列別に自動推定）", RGB(70, 130, 80))

    ws.Rows("14").RowHeight = 26
    ws.Cells(14, 2).Value = "回帰窓（直近何か月でフィットするか）"
    ws.Cells(14, 2).Font.name = "メイリオ": ws.Cells(14, 2).Font.Size = 10
    ws.Cells(14, 2).Interior.Color = RGB(240, 250, 240)
    With ws.Cells(14, 3)
        .Value = 12
        .NumberFormat = "0"
        .Font.name = "メイリオ": .Font.Size = 11: .Font.Bold = True
        .Font.Color = RGB(70, 130, 80)
        .Interior.Color = RGB(240, 255, 240)
        .HorizontalAlignment = xlCenter
    End With
    ws.Cells(14, 4).Value = "推奨 12～24（バックテストで12が最良）"
    ws.Cells(14, 4).Font.name = "メイリオ": ws.Cells(14, 4).Font.Size = 9
    ws.Cells(14, 4).Font.Color = RGB(120, 120, 120)

    Dim sLbl As Variant
    sLbl = Array("総合", "コア（生鮮除く）", "コアコア（生鮮・エネ除く）")
    For r = 0 To 2
        ws.Cells(15 + r, 2).Value = sLbl(r) & "：全国 = α × 東京 + β"
        ws.Cells(15 + r, 2).Font.name = "メイリオ"
        ws.Cells(15 + r, 2).Font.Size = 10
        ws.Cells(15 + r, 2).Interior.Color = RGB(240, 250, 240)
        ws.Cells(15 + r, 3).NumberFormat = "0.0000"
        ws.Cells(15 + r, 4).NumberFormat = "0.0000"
        Dim c As Long
        For c = 3 To 4
            With ws.Cells(15 + r, c)
                .Font.name = "メイリオ": .Font.Size = 11
                .Font.Color = RGB(70, 130, 80)
                .Interior.Color = RGB(240, 255, 240)
                .HorizontalAlignment = xlCenter
            End With
        Next c
    Next r
    ws.Cells(14, 3).Interior.Color = RGB(255, 255, 220)   ' 窓は入力セル

    ws.Rows("18").RowHeight = 10
    Call DrawHeader(ws, "B19:E19", "③ 全国CPI 予想（対象月の速報値を使用）", RGB(180, 60, 0))

    Dim lbl3 As Variant, fmt3 As Variant
    lbl3 = Array("総合（前年比 % 予想）", "コア（前年比 % 予想）", "コアコア（前年比 % 予想）", "総合指数（2025=100 予想）")
    For r = 0 To 3
        With ws.Cells(20 + r, 2)
            .Value = lbl3(r)
            .Font.name = "メイリオ": .Font.Size = 10
            .Interior.Color = RGB(255, 245, 240)
        End With
        With ws.Cells(20 + r, 3)
            .NumberFormat = IIf(r = 3, "0.0", "0.00")
            .Font.name = "メイリオ": .Font.Size = 14: .Font.Bold = True
            .Font.Color = RGB(180, 60, 0)
            .Interior.Color = RGB(255, 235, 225)
            .HorizontalAlignment = xlCenter
        End With
        With ws.Cells(20 + r, 4)
            .Font.name = "メイリオ": .Font.Size = 9
            .Font.Color = RGB(120, 120, 120)
            .Interior.Color = RGB(248, 248, 248)
        End With
    Next r

    ws.Cells(25, 2).Font.name = "メイリオ"
    ws.Cells(25, 2).Font.Size = 9
    ws.Cells(25, 2).Font.Color = RGB(100, 100, 100)
    ws.Cells(26, 2).Font.name = "メイリオ"
    ws.Cells(26, 2).Font.Size = 9
    ws.Cells(26, 2).Font.Color = RGB(100, 100, 100)

    Call AddButton(ws, "B28", "▶ 全自動更新（取得→推定→予想）", "ModuleCPI.AutoUpdate", RGB(0, 120, 215))
    Call AddButton(ws, "D28", "データ取得のみ", "ModuleCPI.FetchAllData", RGB(90, 90, 90))
    Call AddButton(ws, "B30", "R*推計更新（LW/LHW）", "ModuleRStar.RStar_Run", RGB(120, 60, 160))

    ws.Cells(33, 2).Font.name = "メイリオ"
    ws.Cells(33, 2).Font.Size = 9
    ws.Cells(33, 2).Font.Color = RGB(120, 120, 120)
    ws.Cells(35, 2).Font.name = "メイリオ"
    ws.Cells(35, 2).Font.Size = 10
    ws.Cells(35, 2).Font.Bold = True
    ws.Cells(35, 2).Value = "ステータス: 待機中"
End Sub

'----------------------------------------------------------------
Private Sub BuildDiagV2()
    Dim ws As Worksheet
    Set ws = Worksheets(SH_DIAG)
    ws.Cells.Clear
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 22
    Dim c As Long
    For c = 3 To 9
        ws.Columns(c).ColumnWidth = 13
    Next c

    With ws.Range("B2:I2")
        .Merge
        .Value = "回帰診断（系列別・同月前年比回帰の精度検証）"
        .Font.name = "メイリオ": .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(70, 130, 80)
        .HorizontalAlignment = xlCenter
    End With

    Dim hdr As Variant
    hdr = Array("系列", "バックテスト件数", "MAE (%pt)", "RMSE (%pt)", "バイアス (%pt)", "最大 過大", "最大 過小", "窓内 R2")
    For c = 0 To 7
        With ws.Cells(4, 2 + c)
            .Value = hdr(c)
            .Font.name = "メイリオ": .Font.Size = 10: .Font.Bold = True
            .Interior.Color = RGB(220, 240, 220)
            .HorizontalAlignment = xlCenter
            .WrapText = True
        End With
    Next c

    Dim sLbl As Variant
    sLbl = Array("総合", "コア（生鮮除く）", "コアコア（生鮮・エネ除く）")
    Dim r As Long
    For r = 0 To 2
        With ws.Cells(5 + r, 2)
            .Value = sLbl(r)
            .Font.name = "メイリオ": .Font.Size = 10
            .Interior.Color = RGB(240, 250, 240)
        End With
        For c = 3 To 9
            With ws.Cells(5 + r, c)
                .NumberFormat = "0.000"
                .Font.name = "メイリオ": .Font.Size = 10
                .HorizontalAlignment = xlCenter
            End With
        Next c
    Next r

    ws.Cells(9, 2).Value = "※ MAE = 過去10年を「1点先予想」で追いかけた実績との平均誤差（%ポイント）。"
    ws.Cells(10, 2).Value = "※ 毎回その時点より過去の窓だけで回帰し直す厳密な検証（期外）です。"
    ws.Cells(9, 2).Font.name = "メイリオ": ws.Cells(9, 2).Font.Size = 9: ws.Cells(9, 2).Font.Color = RGB(100, 100, 100)
    ws.Cells(10, 2).Font.name = "メイリオ": ws.Cells(10, 2).Font.Size = 9: ws.Cells(10, 2).Font.Color = RGB(100, 100, 100)

    ' 予想 vs 実績 の明細（直近24か月）
    Dim row0 As Long
    row0 = 12
    With ws.Range(ws.Cells(row0, 2), ws.Cells(row0, 7))
        .Merge
        .Value = "■ 直近24か月の予想 vs 実績（明細）"
        .Font.name = "メイリオ": .Font.Size = 11: .Font.Bold = True
        .Interior.Color = RGB(230, 240, 255)
    End With

    Dim dHdr As Variant
    dHdr = Array("年月", "東京", "全国 実績", "全国 予想", "誤差 (実績-予想)")
    Dim blk As Long, rr As Long
    For blk = 0 To 2
        Dim base As Long
        base = row0 + 1 + blk * 27
        ws.Cells(base, 2).Value = "【" & sLbl(blk) & "】"
        ws.Cells(base, 2).Font.name = "メイリオ"
        ws.Cells(base, 2).Font.Bold = True
        ws.Cells(base, 2).Font.Size = 10
        For c = 0 To 4
            With ws.Cells(base + 1, 2 + c)
                .Value = dHdr(c)
                .Font.name = "メイリオ": .Font.Size = 9: .Font.Bold = True
                .Interior.Color = RGB(235, 235, 235)
                .HorizontalAlignment = xlCenter
            End With
        Next c
        For rr = 0 To 23
            For c = 0 To 4
                With ws.Cells(base + 2 + rr, 2 + c)
                    .Font.name = "メイリオ": .Font.Size = 9
                    .NumberFormat = IIf(c = 0, "yyyy/mm", "0.00")
                    .HorizontalAlignment = IIf(c = 0, xlCenter, xlRight)
                End With
            Next c
        Next rr
    Next blk
End Sub

'----------------------------------------------------------------
Private Sub AboutLine(ws As Worksheet, ByRef r As Long, txt As String)
    With ws.Cells(r, 2)
        .Value = txt
        .Font.name = "メイリオ": .Font.Size = 10
        If Left$(txt, 3) = "【" Then
            .Font.Bold = True
            .Font.Color = RGB(0, 84, 166)
        End If
    End With
    r = r + 1
End Sub

Private Sub BuildAboutV2()
    Dim ws As Worksheet
    Set ws = Worksheets(SH_ABOUT)
    ws.Cells.Clear
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 118

    ws.Cells(1, 2).Value = "ツール説明（2025年基準・全自動版）"
    ws.Cells(1, 2).Font.name = "メイリオ": ws.Cells(1, 2).Font.Size = 16: ws.Cells(1, 2).Font.Bold = True

    Dim r As Long
    r = 3
    Call AboutLine(ws, r, "")
    Call AboutLine(ws, r, "【このツールでできること】")
    Call AboutLine(ws, r, "・毎月末に公表される「東京都区部CPI速報」から、約3週間後に公表される「全国CPI」を予想します。")
    Call AboutLine(ws, r, "・「▶ 全自動更新」ボタン1つで、最新データ取得 → 回帰係数の推定 → 予想計算 → 履歴保存まで実行します。")
    Call AboutLine(ws, r, "")
    Call AboutLine(ws, r, "【データ源】")
    Call AboutLine(ws, r, "・e-Stat「消費者物価指数（2025年基準）」（統計表ID 0004052037）。2026年10月に基準改定された最新の統計表です。")
    Call AboutLine(ws, r, "・前年比は基準改定をまたいで接続調整された系列なので、2020年基準の旧数値と多少異なりますが一貫しています。")
    Call AboutLine(ws, r, "・東京都区部は毎月末（その月の速報）、全国は翌月17日前後に公表され、データは全自動更新で追いつきます。")
    Call AboutLine(ws, r, "・速報値は翌月に小幅改訂されることがあります（全国分公表時）。毎回全履歴を取り直すので自動で反映されます。")
    Call AboutLine(ws, r, "")
    Call AboutLine(ws, r, "【予想ロジック（正しい理由）】")
    Call AboutLine(ws, r, "1. 東京と全国の同じ月の前年比のペアを過去12か月分集める（回帰窓はメインパネル②で変更可）。")
    Call AboutLine(ws, r, "2. 系列ごと（総合・コア・コアコア）に独立した直線 全国 = α × 東京 + β を最小二乗法で当てはめる。")
    Call AboutLine(ws, r, "3. 最新の東京速報を式に入れて全国の前年比を予想する。")
    Call AboutLine(ws, r, "4. 総合指数は「昨年同月の全国指数 × (1 + 予想前年比/100)」で連鎖計算（×0.998のような決め打ちはしない）。")
    Call AboutLine(ws, r, "5. 精度は「過去10年を1点先予想で追いかけた誤差（MAE）」として回帰診断シートに毎回表示する。")
    Call AboutLine(ws, r, "")
    Call AboutLine(ws, r, "【精度の目安（2026/10時点のバックテスト）】")
    Call AboutLine(ws, r, "・総合 MAE 0.15%pt / コア 0.16%pt / コアコア 0.13%pt ― 予想はだいたい±0.2%ptの中で当たる。")
    Call AboutLine(ws, r, "・窓を長く（36か月等）すると景気局面の変化に追いつかず誤差が増えることがバックテストで確認済み。")
    Call AboutLine(ws, r, "")
    Call AboutLine(ws, r, "【使い方（毎月のサイクル）】")
    Call AboutLine(ws, r, "① 月末: 東京速報が公表されたら「▶ 全自動更新」→ ③に全国の予想が出る。")
    Call AboutLine(ws, r, "② 翌月17日頃: 全国実績が公表されたら「▶ 全自動更新」→ データシートの履歴で予想と実績を比べられる。")
    Call AboutLine(ws, r, "③ 黄色いセル（①の東京値・②の窓）は手で書き換えてもOK。その後「全国CPI予想を再計算」で即やり直せる。")
    Call AboutLine(ws, r, "")
    Call AboutLine(ws, r, "【旧バージョンからの変更点】")
    Call AboutLine(ws, r, "・データ源を2020年基準から2025年基準の新表に変更（基準改定対応。旧表は2026/08で更新終了）。")
    Call AboutLine(ws, r, "・回帰係数を系列ごとに独立推定（旧版は総合の係数をコア・コアコアに流用していた）。")
    Call AboutLine(ws, r, "・指数推計を連鎖法に変更（旧版は ×0.998 の決め打ちだった）。")
    Call AboutLine(ws, r, "・回帰診断を期外バックテスト（MAE）に変更（旧版のR2は当てはまりだけで誤差の大きさが分からなかった）。")
    Call AboutLine(ws, r, "・旧Module1のマクロは残しているが、新しいボタンはすべて本モジュール（ModuleCPI）を使う。")
End Sub

'----------------------------------------------------------------
Private Sub BuildDataSettings()
    Dim ws As Worksheet
    Set ws = Worksheets(SH_DATA)
    ws.Cells(4, 1).Value = "--- 過去データ（2025年基準・接続済み前年比・回帰推計用）---"
    Dim hdr As Variant
    hdr = Array("年月", "東京_総合前年比", "東京_コア前年比", "東京_コアコア前年比", _
                "全国_総合前年比", "全国_コア前年比", "全国_コアコア前年比", "備考", "全国_総合指数(2025=100)")
    Dim c As Long
    For c = 0 To 8
        With ws.Cells(5, 1 + c)
            .Value = hdr(c)
            .Font.name = "メイリオ": .Font.Size = 9: .Font.Bold = True
            .Interior.Color = RGB(220, 235, 250)
            .WrapText = True
            .HorizontalAlignment = xlCenter
        End With
    Next c

    ' 設定ブロック（J:K）
    ws.Cells(1, 10).Value = "設定（通常は変更不要）"
    ws.Cells(1, 10).Font.Bold = True
    ws.Cells(1, 10).Font.name = "メイリオ"
    Dim names As Variant, vals As Variant
    names = Array("統計表ID", "東京地域コード", "全国地域コード", "回帰窓(月)", "履歴保持(月)", "バックテスト(月)")
    vals = Array("0004052037", "13100", "00000", 12, 156, 120)
    Dim r As Long
    For r = 0 To 5
        ws.Cells(2 + r, 10).Value = names(r)
        ws.Cells(2 + r, 11).NumberFormat = "@"
        ws.Cells(2 + r, 11).Value = CStr(vals(r))
        ws.Cells(2 + r, 10).Font.name = "メイリオ"
        ws.Cells(2 + r, 10).Font.Size = 9
        ws.Cells(2 + r, 10).Interior.Color = RGB(245, 245, 245)
        ws.Cells(2 + r, 11).Font.name = "メイリオ"
        ws.Cells(2 + r, 11).Font.Size = 9
    Next r
    ws.Columns(10).ColumnWidth = 16
    ws.Columns(11).ColumnWidth = 14
End Sub

'================================================================
' 1. 共通ユーティリティ
'================================================================
Private Sub DrawHeader(ws As Worksheet, rng As String, title As String, col As Long)
    With ws.Range(rng)
        .Merge
        .Value = title
        .Font.name = "メイリオ": .Font.Size = 11: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = col
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(ws.Range(rng).row).RowHeight = 22
End Sub

Private Sub AddButton(ws As Worksheet, cellAddr As String, cap As String, mac As String, col As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range(cellAddr).Left, _
                                 ws.Range(cellAddr).Top, 150, 32)
    With shp
        .Fill.ForeColor.RGB = col
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = cap
        .TextFrame2.TextRange.Font.Name = "メイリオ"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = mac
    End With
End Sub

'----------------------------------------------------------------
' 実行ログ（CPI予想ツール\run.log に追記）— ハング/エラーの切り分け用
'----------------------------------------------------------------
Public Sub Log(msg As String)
    On Error Resume Next
    Dim f As Long
    f = FreeFile
    Open ThisWorkbook.Path & "\run.log" For Append As #f
    Print #f, Format$(Now, "hh:mm:ss") & "  " & msg
    Close #f
End Sub

Private Sub SetStatus(msg As String, col As Long)
    With Worksheets(SH_MAIN).Range("B35")
        .Value = msg
        .Font.Color = col
    End With
    Call Log(msg)
End Sub

Private Function GetSettingStr(nm As String, dflt As String) As String
    Dim ws As Worksheet
    Set ws = Worksheets(SH_DATA)
    Dim r As Long
    For r = 2 To 8
        If Trim$(CStr(ws.Cells(r, 10).Value)) = nm Then
            GetSettingStr = Trim$(CStr(ws.Cells(r, 11).Value))
            Exit Function
        End If
    Next r
    GetSettingStr = dflt
End Function

Private Function GetSettingNum(nm As String, dflt As Double) As Double
    Dim s As String
    s = GetSettingStr(nm, "")
    If IsNumeric(s) Then
        GetSettingNum = CDbl(s)
    Else
        GetSettingNum = dflt
    End If
End Function

Public Function HttpGet(url As String) As String
    Call Log("HTTP-SEND " & Left$(url, 70))
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 10000, 10000, 30000, 60000   ' 解決/接続/送信/受信 (ms)
    http.Open "GET", url, False
    http.Send
    Call Log("HTTP-RECV status=" & http.Status & " len=" & Len(http.responseText))
    If http.Status <> 200 Then
        Err.Raise vbObjectError + 513, , "e-Stat サーバーエラー HTTP " & http.Status
    End If
    HttpGet = http.responseText
End Function

Private Function SeriesUrl(apiKey As String, sid As String, area As String, cat As String, tabS As String) As String
    SeriesUrl = ESTAT_BASE & "?appId=" & apiKey & "&statsDataId=" & sid & _
                "&cdArea=" & area & "&cdCat01=" & cat & "&cdTab=" & tabS & _
                "&metaGetFlg=N&cntGetFlg=N&limit=100000"
End Function

'----------------------------------------------------------------
' JSON の VALUE 配列を文字列走査して {月: 値} の辞書にする
' 月コードは YYYY00MMMM 形式。年平均(YYYY10...)・欠損(***)は除外
'----------------------------------------------------------------
Public Function ParseSeries(resp As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim vStart As Long, pos As Long, recStart As Long, recEnd As Long
    vStart = InStr(resp, """VALUE"":[")
    If vStart = 0 Then
        Call Log("PARSE: VALUE配列なし（統計表ID・APIキー・通信を確認）")
        Set ParseSeries = d
        Exit Function
    End If
    pos = vStart
    Dim rec As String, tPos As Long, tCode As String
    Dim mm As String, mm2 As String, dPos As Long, valStr As String
    Do
        recStart = InStr(pos, resp, "{""@tab""")
        If recStart = 0 Then Exit Do
        recEnd = InStr(recStart, resp, "}")
        If recEnd = 0 Then Exit Do
        rec = Mid$(resp, recStart, recEnd - recStart + 1)
        tPos = InStr(rec, """@time"":""")
        If tPos > 0 Then
            tCode = Mid$(rec, tPos + 9, 10)
            If Mid$(tCode, 5, 2) = "00" Then
                mm = Mid$(tCode, 7, 2)
                mm2 = Mid$(tCode, 9, 2)
                If mm >= "01" And mm <= "12" And mm = mm2 Then
                    dPos = InStr(rec, """$"":""")
                    If dPos > 0 Then
                        valStr = Mid$(rec, dPos + 5)
                        valStr = Left$(valStr, InStr(valStr, """") - 1)
                        If valStr <> "" And valStr <> "***" Then
                            d(CLng(Left$(tCode, 4) & mm)) = CDbl(valStr)
                        End If
                    End If
                End If
            End If
        End If
        pos = recEnd + 1
    Loop
    Set ParseSeries = d
End Function

'----------------------------------------------------------------
' 最小二乗法 y = a*x + b（x(), y() は 1..n）
'----------------------------------------------------------------
Private Sub FitOLS(x() As Double, y() As Double, n As Long, ByRef a As Double, ByRef b As Double, ByRef r2 As Double)
    Dim mx As Double, my As Double
    Dim i As Long
    For i = 1 To n
        mx = mx + x(i): my = my + y(i)
    Next i
    mx = mx / n: my = my / n
    Dim sxx As Double, sxy As Double
    For i = 1 To n
        sxx = sxx + (x(i) - mx) ^ 2
        sxy = sxy + (x(i) - mx) * (y(i) - my)
    Next i
    If sxx = 0 Then
        a = 0: b = my: r2 = 0
    Else
        a = sxy / sxx
        b = my - a * mx
        Dim ssT As Double, ssR As Double, p As Double
        For i = 1 To n
            p = a * x(i) + b
            ssT = ssT + (y(i) - my) ^ 2
            ssR = ssR + (y(i) - p) ^ 2
        Next i
        r2 = 0
        If ssT > 0 Then r2 = 1 - ssR / ssT
    End If
End Sub

'================================================================
' 2. データ取得（e-Stat 2025年基準）
'================================================================
'----------------------------------------------------------------
' トレンドチャートシート更新 (直近24ヶ月) + 乖離チャート再構築
'   ・A:年月 B:東京総合 C:東京コア D:東京コアコア E:全国総合
'     F:全国コア G:全国コアコア H:全国-東京 乖離(pp)
'   ・2枚目のチャートは「乖離のみ」(タイトルと内容を一致させる)
'----------------------------------------------------------------
Public Sub BuildTrendSheet()
    On Error GoTo E
    Dim wsT As Worksheet
    On Error Resume Next
    Set wsT = Worksheets("トレンドチャート")
    On Error GoTo E
    If wsT Is Nothing Then Exit Sub
    Const NR As Long = 24
    If mNCommon < NR Then Exit Sub
    Dim i As Long, r As Long, ym As Long
    r = 4
    For i = mNCommon - NR + 1 To mNCommon
        ym = mMonths(i)
        wsT.Cells(r, 1).Value = DateSerial(ym \ 100, ym Mod 100, 1)
        wsT.Cells(r, 2).Value = dTokyoS(ym)
        wsT.Cells(r, 3).Value = dTokyoC(ym)
        wsT.Cells(r, 4).Value = dTokyoCC(ym)
        wsT.Cells(r, 5).Value = dNatS(ym)
        wsT.Cells(r, 6).Value = dNatC(ym)
        wsT.Cells(r, 7).Value = dNatCC(ym)
        wsT.Cells(r, 8).Value = dNatS(ym) - dTokyoS(ym)
        r = r + 1
    Next i
    wsT.Cells(3, 8).Value = "全国-東京乖離"
    wsT.Range("A4:A" & (3 + NR)).NumberFormat = "yyyy/mm"
    wsT.Range("B4:H" & (3 + NR)).NumberFormat = "0.0"
    wsT.Cells(3, 8).Font.Bold = True
    wsT.Cells(3, 8).Font.Color = RGB(112, 48, 160)

    Dim co As ChartObject, ch As Chart
    If wsT.ChartObjects.Count >= 2 Then
        Set co = wsT.ChartObjects(2)
        Set ch = co.Chart
        Do While ch.SeriesCollection.Count > 0
            ch.SeriesCollection(1).Delete
        Loop
        ch.ChartType = xlLine
        ch.HasTitle = True
        ch.ChartTitle.Text = "全国−東京 総合CPI乖離 (pp)"
        ch.ChartTitle.Font.Name = "メイリオ"
        ch.ChartTitle.Font.Size = 12
        ch.ChartTitle.Font.Bold = True
        ch.ChartTitle.Font.Color = RGB(31, 78, 121)
        ch.HasLegend = True
        ch.Legend.Position = xlLegendPositionBottom
        ch.Legend.Font.Name = "メイリオ"
        ch.Legend.Font.Size = 9
        ch.ChartArea.Font.Name = "メイリオ"
        ch.ChartArea.Font.Size = 9
        ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        ch.ChartArea.Format.Line.Visible = msoFalse
        With ch.Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "乖離 (pp)"
            .AxisTitle.Font.Name = "メイリオ"
            .AxisTitle.Font.Size = 9
            .AxisTitle.Font.Bold = True
            .TickLabels.Font.Name = "メイリオ"
            .TickLabels.Font.Size = 9
            .TickLabels.NumberFormat = "0.0"
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
            .MajorGridlines.Format.Line.Weight = 0.5
            .Format.Line.ForeColor.RGB = RGB(191, 191, 191)
            .Format.Line.Weight = 0.75
        End With
        With ch.Axes(xlCategory)
            .TickLabels.Font.Name = "メイリオ"
            .TickLabels.Font.Size = 8
            .TickLabelSpacing = 3
            .TickLabels.Orientation = -45
            .HasMajorGridlines = False
            .Format.Line.ForeColor.RGB = RGB(191, 191, 191)
            .Format.Line.Weight = 0.75
        End With
        ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(252, 252, 252)
        ch.PlotArea.Format.Line.Visible = msoFalse
        Dim s As Series
        Set s = ch.SeriesCollection.NewSeries
        s.Name = "全国−東京 (総合)"
        s.Values = wsT.Range("H4:H" & (3 + NR))
        s.XValues = wsT.Range("A4:A" & (3 + NR))
        s.ChartType = xlLine
        s.Format.Line.ForeColor.RGB = RGB(112, 48, 160)
        s.Format.Line.Weight = 2.25
        s.MarkerStyle = xlMarkerStyleCircle
        s.MarkerSize = 4
        s.MarkerBackgroundColor = RGB(255, 255, 255)
        s.MarkerForegroundColor = RGB(112, 48, 160)
    End If
    Exit Sub
E:
    Call Log("BuildTrendSheet error: " & Err.Number & " - " & Err.Description)
End Sub

Public Sub FetchAllData()
    Call Log("FETCH-ENTER")
    Dim wsData As Worksheet
    Set wsData = Worksheets(SH_DATA)
    Dim apiKey As String
    apiKey = Trim$(CStr(wsData.Range("B2").Value))
    Call Log("FETCH-KEY len=" & Len(apiKey))
    If apiKey = "" Then
        Call SetStatus("エラー: データシート B2 に e-Stat APIキーがありません", RGB(200, 0, 0))
        Exit Sub
    End If

    Dim sid As String, aTok As String, aNat As String
    sid = GetSettingStr("統計表ID", "0004052037")
    aTok = GetSettingStr("東京地域コード", "13100")
    aNat = GetSettingStr("全国地域コード", "00000")
    sid = Right$("0000000000" & sid, 10)
    aNat = Right$("00000" & aNat, 5)
    If aTok <> "00000" Then aTok = Right$("00000" & aTok, 5)
    Call Log("FETCH-SETTING sid=" & sid)

    Call SetStatus("ステータス: e-Stat からデータ取得中…（8リクエスト）", RGB(0, 120, 215))

    Call Log("HTTP 1/8 東京総合YoY")
    Set dTokyoS = ParseSeries(HttpGet(SeriesUrl(apiKey, sid, aTok, "0001", "3")))
    Call Log("HTTP 2/8 東京コアYoY")
    Set dTokyoC = ParseSeries(HttpGet(SeriesUrl(apiKey, sid, aTok, "0161", "3")))
    Call Log("HTTP 3/8 東京コアコアYoY")
    Set dTokyoCC = ParseSeries(HttpGet(SeriesUrl(apiKey, sid, aTok, "0178", "3")))
    Call Log("HTTP 4/8 全国総合YoY")
    Set dNatS = ParseSeries(HttpGet(SeriesUrl(apiKey, sid, aNat, "0001", "3")))
    Call Log("HTTP 5/8 全国コアYoY")
    Set dNatC = ParseSeries(HttpGet(SeriesUrl(apiKey, sid, aNat, "0161", "3")))
    Call Log("HTTP 6/8 全国コアコアYoY")
    Set dNatCC = ParseSeries(HttpGet(SeriesUrl(apiKey, sid, aNat, "0178", "3")))
    Call Log("HTTP 7/8 全国総合指数")
    Set dNatIdx = ParseSeries(HttpGet(SeriesUrl(apiKey, sid, aNat, "0001", "1")))
    Call Log("HTTP 8/8 東京総合指数")
    Set dTokyoIdx = ParseSeries(HttpGet(SeriesUrl(apiKey, sid, aTok, "0001", "1")))
    Call Log("HTTP 完了")

    If dTokyoS.Count < 60 Or dNatS.Count < 60 Or dNatIdx.Count < 12 Then
        Call SetStatus("エラー: e-Statからの取得に失敗（データ件数不足）", RGB(200, 0, 0))
        Exit Sub
    End If

    ' 共通月の計算（6系列がそろう月のみ・昇順）
    mNCommon = 0
    Erase mMonths
    ReDim mMonths(1 To dTokyoS.Count)
    Dim k As Variant, j As Long
    For Each k In dTokyoS.Keys
        If dTokyoC.Exists(k) And dTokyoCC.Exists(k) And dNatS.Exists(k) And dNatC.Exists(k) And dNatCC.Exists(k) Then
            mNCommon = mNCommon + 1
            mMonths(mNCommon) = CLng(k)
        End If
    Next k
    Call SortLongs(mMonths, mNCommon)

    Call WriteHistory(wsData)
    Call BuildTrendSheet
    Call SetStatus("ステータス: データ取得完了 " & mNCommon & "か月（最新 " & mMonths(mNCommon) \ 100 & "/" & Format$(mMonths(mNCommon) Mod 100, "00") & "）", RGB(0, 150, 0))
End Sub

'----------------------------------------------------------------
Private Sub SortLongs(arr() As Long, n As Long)
    Dim i As Long, j As Long, t As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(j) < arr(i) Then
                t = arr(i): arr(i) = arr(j): arr(j) = t
            End If
        Next j
    Next i
End Sub

'----------------------------------------------------------------
Private Sub WriteHistory(wsData As Worksheet)
    Dim keepN As Long
    keepN = CLng(GetSettingNum("履歴保持(月)", 156))
    If keepN > mNCommon Then keepN = mNCommon

    ' 旧データをクリア（設定ブロック J:K は触らない）
    wsData.Range("A6:I2000").ClearContents
    wsData.Range("A6:I2000").Interior.ColorIndex = xlNone

    Dim startK As Long
    startK = mNCommon - keepN + 1
    Dim i As Long, r As Long
    Dim ym As Long
    For i = 0 To keepN - 1
        ym = mMonths(mNCommon - i)          ' 新しい順
        r = 6 + i
        wsData.Cells(r, 1).Value = DateSerial(ym \ 100, ym Mod 100, 1)
        wsData.Cells(r, 1).NumberFormat = "yyyy/mm"
        wsData.Cells(r, 2).Value = dTokyoS(ym)
        wsData.Cells(r, 3).Value = dTokyoC(ym)
        wsData.Cells(r, 4).Value = dTokyoCC(ym)
        wsData.Cells(r, 5).Value = dNatS(ym)
        wsData.Cells(r, 6).Value = dNatC(ym)
        wsData.Cells(r, 7).Value = dNatCC(ym)
        wsData.Cells(r, 8).Value = "API取得(2025基準)"
        If dNatIdx.Exists(ym) Then
            wsData.Cells(r, 9).Value = dNatIdx(ym)
            wsData.Cells(r, 9).NumberFormat = "0.0"
        End If
        If i Mod 2 = 0 Then
            wsData.Range(wsData.Cells(r, 1), wsData.Cells(r, 9)).Interior.Color = RGB(248, 252, 248)
        End If
    Next i
    wsData.Cells(3, 1).Value = "最終取得: " & Format$(Now, "yyyy/mm/dd hh:mm") & _
        "（" & keepN & "か月表示・全" & mNCommon & "か月）"
    wsData.Cells(3, 1).Font.name = "メイリオ"
    wsData.Cells(3, 1).Font.Size = 9
End Sub

'================================================================
' 3. 回帰推定（系列別）＋バックテスト
'================================================================
Public Sub FitAllRegressions()
    Dim wsMain As Worksheet
    Set wsMain = Worksheets(SH_MAIN)
    Call SetStatus("ステータス: 回帰係数を推定中…", RGB(70, 130, 80))

    mTarget = 0
    Dim k As Variant
    For Each k In dTokyoS.Keys
        If CLng(k) > mTarget Then mTarget = CLng(k)
    Next k

    Dim w As Long
    If IsNumeric(wsMain.Range("C14").Value) And Val(wsMain.Range("C14").Value) >= 6 Then
        w = CLng(wsMain.Range("C14").Value)
    Else
        w = CLng(GetSettingNum("回帰窓(月)", 12))
        wsMain.Range("C14").Value = w
    End If

    Dim seriesName As Variant
    seriesName = Array("総合", "コア（生鮮除く）", "コアコア（生鮮・エネ除く）")

    Dim sIdx As Long
    For sIdx = 0 To 2
        Dim dT As Object, dN As Object
        Select Case sIdx
            Case 0: Set dT = dTokyoS: Set dN = dNatS
            Case 1: Set dT = dTokyoC: Set dN = dNatC
            Case 2: Set dT = dTokyoCC: Set dN = dNatCC
        End Select

        ' --- 窓内フィット（対象月より過去のみ） ---
        Dim xs() As Double, ys() As Double
        Dim cnt As Long, i As Long, ym As Long
        ReDim xs(1 To w): ReDim ys(1 To w)
        cnt = 0
        For i = mNCommon To 1 Step -1
            ym = mMonths(i)
            If ym < mTarget Then
                cnt = cnt + 1
                xs(cnt) = dT(ym): ys(cnt) = dN(ym)
                If cnt = w Then Exit For
            End If
        Next i
        If cnt < 6 Then
            Err.Raise vbObjectError + 516, , "回帰に必要な過去データが足りません（" & cnt & "か月）"
        End If

        Dim a As Double, b As Double, r2 As Double
        Call FitOLS(xs, ys, cnt, a, b, r2)

        wsMain.Cells(15 + sIdx, 3).Value = a
        wsMain.Cells(15 + sIdx, 3).NumberFormat = "0.0000"
        wsMain.Cells(15 + sIdx, 4).Value = b
        wsMain.Cells(15 + sIdx, 4).NumberFormat = "0.0000"

        ' --- バックテスト: 過去10年を1点先予想で検証 ---
        Dim spanN As Long
        spanN = CLng(GetSettingNum("バックテスト(月)", 120))
        Dim nBT As Long, errs() As Double
        ReDim errs(1 To spanN)
        nBT = 0
        Dim mae As Double, rmse As Double, bias As Double
        Dim maxO As Double, maxU As Double
        maxO = -999: maxU = 999
        Dim startI As Long
        startI = mNCommon - spanN + 1
        If startI < 1 Then startI = 1
        Dim p As Long
        For p = startI To mNCommon
            ym = mMonths(p)
            If ym < mTarget Then
                Dim hx() As Double, hy() As Double
                Dim cn As Long
                ReDim hx(1 To w): ReDim hy(1 To w)
                cn = 0
                Dim q As Long
                For q = p - 1 To 1 Step -1
                    cn = cn + 1
                    hx(cn) = dT(mMonths(q)): hy(cn) = dN(mMonths(q))
                    If cn = w Then Exit For
                Next q
                If cn >= 6 Then
                    Dim a2 As Double, b2 As Double, rr2 As Double
                    Call FitOLS(hx, hy, cn, a2, b2, rr2)
                    If a2 <> 0 Or b2 <> 0 Then
                        nBT = nBT + 1
                        errs(nBT) = dN(ym) - (a2 * dT(ym) + b2)
                    End If
                End If
            End If
        Next p
        If nBT > 0 Then
            For i = 1 To nBT
                mae = mae + Abs(errs(i))
                rmse = rmse + errs(i) ^ 2
                bias = bias + errs(i)
                If errs(i) > maxO Then maxO = errs(i)
                If errs(i) < maxU Then maxU = errs(i)
            Next i
            mae = mae / nBT
            rmse = Sqr(rmse / nBT)
            bias = bias / nBT
        End If

        ' --- 回帰診断シートへ ---
        Dim wsDiag As Worksheet
        Set wsDiag = Worksheets(SH_DIAG)
        wsDiag.Cells(5 + sIdx, 2).Value = seriesName(sIdx)
        wsDiag.Cells(5 + sIdx, 3).Value = nBT
        wsDiag.Cells(5 + sIdx, 3).NumberFormat = "0"
        wsDiag.Cells(5 + sIdx, 4).Value = Round(mae, 3)
        wsDiag.Cells(5 + sIdx, 5).Value = Round(rmse, 3)
        wsDiag.Cells(5 + sIdx, 6).Value = Round(bias, 3)
        wsDiag.Cells(5 + sIdx, 7).Value = IIf(maxO = -999, "-", Round(maxO, 3))
        wsDiag.Cells(5 + sIdx, 8).Value = IIf(maxU = 999, "-", Round(maxU, 3))
        wsDiag.Cells(5 + sIdx, 9).Value = Round(r2, 3)
    Next sIdx

    Call WriteDiagDetail
    Call SetStatus("ステータス: 回帰推定完了（窓 " & w & "か月）", RGB(70, 130, 80))
End Sub

'----------------------------------------------------------------
' 回帰診断に「予想 vs 実績」の直近24か月明細を書く
'----------------------------------------------------------------
Private Sub WriteDiagDetail()
    Dim wsDiag As Worksheet
    Set wsDiag = Worksheets(SH_DIAG)
    Dim w As Long
    If IsNumeric(Worksheets(SH_MAIN).Range("C14").Value) And Val(Worksheets(SH_MAIN).Range("C14").Value) >= 6 Then
        w = CLng(Worksheets(SH_MAIN).Range("C14").Value)
    Else
        w = 12
    End If

    Dim sIdx As Long
    For sIdx = 0 To 2
        Dim dT As Object, dN As Object
        Select Case sIdx
            Case 0: Set dT = dTokyoS: Set dN = dNatS
            Case 1: Set dT = dTokyoC: Set dN = dNatC
            Case 2: Set dT = dTokyoCC: Set dN = dNatCC
        End Select
        Dim base As Long
        base = 13 + sIdx * 27
        Dim p As Long, rr As Long
        rr = 0
        Dim hx() As Double, hy() As Double
        Dim a As Double, b As Double, r2 As Double
        For p = mNCommon To 1 Step -1
            Dim ym As Long
            ym = mMonths(p)
            If ym < mTarget Then
                Dim cn As Long, q As Long
                ReDim hx(1 To w): ReDim hy(1 To w)
                cn = 0
                For q = p - 1 To 1 Step -1
                    cn = cn + 1
                    hx(cn) = dT(mMonths(q)): hy(cn) = dN(mMonths(q))
                    If cn = w Then Exit For
                Next q
                If cn >= 6 Then
                    Call FitOLS(hx, hy, cn, a, b, r2)
                    wsDiag.Cells(base + 2 + rr, 2).Value = DateSerial(ym \ 100, ym Mod 100, 1)
                    wsDiag.Cells(base + 2 + rr, 3).Value = dT(ym)
                    wsDiag.Cells(base + 2 + rr, 4).Value = dN(ym)
                    wsDiag.Cells(base + 2 + rr, 5).Value = Round(a * dT(ym) + b, 2)
                    wsDiag.Cells(base + 2 + rr, 6).Value = Round(dN(ym) - (a * dT(ym) + b), 2)
                    rr = rr + 1
                    If rr = 24 Then Exit For
                End If
            End If
        Next p
    Next sIdx
End Sub

'================================================================
' 4. 全国CPI予想計算
'================================================================
Public Sub CalcNational()
    Dim wsMain As Worksheet
    Set wsMain = Worksheets(SH_MAIN)
    Dim ym As Long
    ym = ReadTargetMonth()
    If ym <= 0 Then
        Call SetStatus("エラー: 対象月が不正です（例: 2026/09）", RGB(200, 0, 0))
        Exit Sub
    End If
    Dim v8 As Variant
    Dim v9 As Variant
    Dim v10 As Variant
    v8 = wsMain.Range("C8").Value
    v9 = wsMain.Range("C9").Value
    v10 = wsMain.Range("C10").Value
    If (IsEmpty(v8) = True) Or (IsNumeric(v8) = False) Then
        Call SetStatus("エラー: 対象月の東京CPI値が空です", RGB(200, 0, 0))
        Exit Sub
    End If
    If (IsEmpty(v9) = True) Or (IsNumeric(v9) = False) Then
        Call SetStatus("エラー: 対象月の東京CPI値が空です", RGB(200, 0, 0))
        Exit Sub
    End If
    If (IsEmpty(v10) = True) Or (IsNumeric(v10) = False) Then
        Call SetStatus("エラー: 対象月の東京CPI値が空です", RGB(200, 0, 0))
        Exit Sub
    End If
    Dim tS As Double
    Dim tC As Double
    Dim tCC As Double
    tS = CDbl(v8)
    tC = CDbl(v9)
    tCC = CDbl(v10)
    Dim alphaS As Double
    Dim bS As Double
    Dim aC As Double
    Dim bC As Double
    Dim aCC As Double
    Dim bCC As Double
    alphaS = Val(wsMain.Range("C15").Value)
    bS = Val(wsMain.Range("D15").Value)
    aC = Val(wsMain.Range("C16").Value)
    bC = Val(wsMain.Range("D16").Value)
    aCC = Val(wsMain.Range("C17").Value)
    bCC = Val(wsMain.Range("D17").Value)
    If (alphaS = 0) And (bS = 0) Then
        Call SetStatus("エラー: 回帰係数が未計算です。先に全自動更新を実行してください", RGB(200, 0, 0))
        Exit Sub
    End If
    Dim pS As Double
    Dim pC As Double
    Dim pCC As Double
    pS = alphaS * tS + bS
    pC = aC * tC + bC
    pCC = aCC * tCC + bCC
    wsMain.Range("C20").Value = Round(pS, 2)
    wsMain.Range("C21").Value = Round(pC, 2)
    wsMain.Range("C22").Value = Round(pCC, 2)
    Dim ymPrev As Long
    ymPrev = ym - 100
    Dim idxPrev As Double
    Dim idxOK As Boolean
    idxOK = FindNatIndex(ymPrev, idxPrev)
    If idxOK Then
        wsMain.Range("C23").Value = Round(idxPrev * (1 + pS / 100), 1)
        wsMain.Range("D23").Value = "昨年同月 " & idxPrev & " に予想前年比を掛けた値"
    Else
        wsMain.Range("C23").Value = ""
        wsMain.Range("D23").Value = "昨年同月の指数がデータ範囲外"
    End If
    Dim wsDiag As Worksheet
    Set wsDiag = Worksheets(SH_DIAG)
    wsMain.Range("D20").Value = "バックテストMAE 約" & Format(wsDiag.Cells(5, 4).Value, "0.00") & "pt"
    wsMain.Range("D21").Value = "バックテストMAE 約" & Format(wsDiag.Cells(6, 4).Value, "0.00") & "pt"
    wsMain.Range("D22").Value = "バックテストMAE 約" & Format(wsDiag.Cells(7, 4).Value, "0.00") & "pt"
    Dim yrN As Long
    Dim moN As Long
    yrN = Int(ym / 100)
    moN = ym Mod 100
    wsMain.Range("B25").Value = "計算式: 全国 = 係数 × 東京 + 切片（系列ごとに独立推定）"
    wsMain.Range("B26").Value = "予想対象: 全国 " & CStr(yrN) & "年" & Format(moN, "00") & "月（翌月17日頃公表）"
    Call SaveHistoryRow(ym, tS, tC, tCC, pS, pC, pCC)
    Call SetStatus("ステータス: 予想完了 " & CStr(yrN) & "/" & Format(moN, "00"), RGB(180, 60, 0))
End Sub

'----------------------------------------------------------------
Private Function ReadTargetMonth() As Long
    Dim v As Variant
    v = Worksheets(SH_MAIN).Range("C7").Value
    If IsDate(v) Then
        ReadTargetMonth = Year(CDate(v)) * 100 + Month(CDate(v))
    Else
        Dim s As String
        s = Replace(Trim$(CStr(v)), "年", "/")
        s = Replace(s, "月", "")
        Dim parts() As String
        parts = Split(s, "/")
        If UBound(parts) = 1 Then
            If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
                ReadTargetMonth = CLng(parts(0)) * 100 + CLng(parts(1))
            End If
        End If
    End If
End Function

'----------------------------------------------------------------
' データシート I列から指定月の全国指数を探す
'----------------------------------------------------------------
Private Function FindNatIndex(ym As Long, ByRef idx As Double) As Boolean
    Dim wsData As Worksheet
    Set wsData = Worksheets(SH_DATA)
    Dim r As Long
    Dim lastR As Long
    lastR = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row
    For r = 6 To lastR
        If Not IsEmpty(wsData.Cells(r, 1).Value) And IsDate(wsData.Cells(r, 1).Value) Then
            If Year(CDate(wsData.Cells(r, 1).Value)) * 100 + Month(CDate(wsData.Cells(r, 1).Value)) = ym Then
                If IsNumeric(wsData.Cells(r, 9).Value) And wsData.Cells(r, 9).Value <> "" Then
                    idx = CDbl(wsData.Cells(r, 9).Value)
                    FindNatIndex = True
                    Exit Function
                End If
            End If
        End If
    Next r
    FindNatIndex = False
End Function

'----------------------------------------------------------------
Private Sub SaveHistoryRow(ym As Long, tS As Double, tC As Double, tCC As Double, _
                           pS As Double, pC As Double, pCC As Double)
    Dim wsHist As Worksheet
    Set wsHist = Worksheets(SH_HIST)
    Dim r As Long
    r = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).row + 1
    If r < 3 Then r = 3
    wsHist.Cells(r, 1).Value = Now
    wsHist.Cells(r, 1).NumberFormat = "yyyy/mm/dd hh:mm"
    wsHist.Cells(r, 2).Value = ym \ 100 & "/" & Format$(ym Mod 100, "00")
    wsHist.Cells(r, 3).Value = tS
    wsHist.Cells(r, 4).Value = tC
    wsHist.Cells(r, 5).Value = tCC
    wsHist.Cells(r, 6).Value = Round(pS, 2)
    wsHist.Cells(r, 7).Value = Round(pC, 2)
    wsHist.Cells(r, 8).Value = Round(pCC, 2)
    wsHist.Cells(r, 9).Value = "α=" & Format$(Worksheets(SH_MAIN).Range("C15").Value, "0.0000") & _
        " β=" & Format$(Worksheets(SH_MAIN).Range("D15").Value, "0.0000") & _
        " 窓=" & Worksheets(SH_MAIN).Range("C14").Value
    wsHist.Cells(r, 10).Value = "実績対比は全国公表後に全自動更新で確認"
    Dim c As Long
    For c = 1 To 10
        wsHist.Cells(r, c).Font.name = "メイリオ"
        wsHist.Cells(r, c).Font.Size = 10
    Next c
End Sub

'================================================================
' 5. 全自動更新（メインのボタン）
'================================================================
'----------------------------------------------------------------
' テイラールールによる政策金利の目安 (金利予想シート)
'   i_eq(t) = r* + pi(t) + b_pi*(pi(t) - pi*) + b_y*gap
'   i(t)    = i_eq(t) + rho*( i(t-1) - i_eq(t-1) )
'----------------------------------------------------------------
' パラメータセルの数値取得 (空・非数値は既定値。IsNumeric(Empty) は True なので空判定を明示)
Private Function RateParam(ws As Worksheet, ByVal addr As String, ByVal dfl As Double) As Double
    Dim v As Variant
    v = ws.Range(addr).Value
    RateParam = dfl
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    RateParam = CDbl(v)
End Function

' パラメータセルが空なら既定値を書き込む
Private Sub SetRateParam(ws As Worksheet, ByVal addr As String, ByVal dfl As Variant)
    Dim v As Variant
    v = ws.Range(addr).Value
    If IsEmpty(v) Then
        ws.Range(addr).Value = dfl
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        ws.Range(addr).Value = dfl
    End If
End Sub

' 最新の市場無担保コールレート (月平均, %) を返す。
' Rstar シートの入力テーブル G列 (nom_rate) の最終行から取得。取得できていなければ 0。
Private Function LastCallRate() As Double
    On Error GoTo TErr
    Dim wsR As Worksheet
    Dim r As Long, v As Double
    Set wsR = Nothing
    On Error Resume Next
    Set wsR = ThisWorkbook.Worksheets("Rstar")
    On Error GoTo TErr
    If wsR Is Nothing Then
        LastCallRate = 0#
        Exit Function
    End If
    r = 431
    Do While r < 831
        If Trim$(CStr(wsR.Cells(r, 1).Value)) = "" Then Exit Do
        r = r + 1
    Loop
    r = r - 1
    If r >= 431 Then
        If IsNumeric(wsR.Cells(r, 7).Value) Then
            v = CDbl(wsR.Cells(r, 7).Value)
            LastCallRate = v
        End If
    End If
    Exit Function
TErr:
    LastCallRate = 0#
End Function

Public Sub BuildTaylorPath()
    On Error GoTo TErr
    Dim wsM As Worksheet, wsR As Worksheet, wsT As Worksheet, wsM2 As Worksheet
    Dim ymT As Long, ym As Long, mm As Long, r As Long, hc As Long, rLast As Long
    Dim piStar As Double, bPi As Double, bY As Double, rhoS As Double, iNow As Double
    Dim sIdx As Long, rstarMode As Long
    Dim rStar As Double, gap As Double, rho As Double
    Dim aCore As Double, bCore As Double, aTot As Double, bTot As Double, aCoreCC As Double, bCoreCC As Double
    Dim xT As Double, xPrev As Double, pNow As Double
    Dim iPrev As Double, iEqPrev As Double
    Dim ymArr() As Long, piArr() As Double, iEqArr() As Double, iPath() As Double
    Dim msg As String

    Set wsM = Worksheets(SH_MAIN)
    On Error Resume Next
    Set wsR = Worksheets("Rstar")
    On Error GoTo TErr
    If wsR Is Nothing Then Exit Sub

    ymT = ReadTargetMonth()
    If ymT <= 0 Then Exit Sub

    Set wsT = GetRateSheet()

    SetRateParam wsT, "B4", 2#
    SetRateParam wsT, "B5", 0.5
    SetRateParam wsT, "B6", 0.5
    SetRateParam wsT, "B7", 0.5
    SetRateParam wsT, "B8", 1.25
    SetRateParam wsT, "B9", 2
    SetRateParam wsT, "B10", 0
    SetRateParam wsT, "B11", 0.75
    SetRateParam wsT, "B12", 0
    piStar = RateParam(wsT, "B4", 2#) / 100#
    bPi = RateParam(wsT, "B5", 0.5)
    bY = RateParam(wsT, "B6", 0.5)
    rhoS = RateParam(wsT, "B7", 0.5)
    iNow = RateParam(wsT, "B8", 1.25) / 100#
    sIdx = CLng(RateParam(wsT, "B9", 2#))
    rstarMode = CLng(RateParam(wsT, "B10", 0#))

    rStar = 0#
    gap = 0#
    rLast = 431
    Do While rLast < 831
        If Trim$(CStr(wsR.Cells(rLast, 1).Value)) = "" Then Exit Do
        rLast = rLast + 1
    Loop
    rLast = rLast - 1
    If rLast >= 431 Then
        If IsNumeric(wsR.Cells(rLast, 3).Value) Then rStar = CDbl(wsR.Cells(rLast, 3).Value) / 100#
        If IsNumeric(wsR.Cells(rLast, 7).Value) Then gap = CDbl(wsR.Cells(rLast, 7).Value) / 100#
    End If
    If rstarMode = 1 Then
        If IsNumeric(wsT.Range("B11").Value) Then rStar = CDbl(wsT.Range("B11").Value) / 100#
    End If
    If IsNumeric(wsT.Range("B12").Value) Then
        If CDbl(wsT.Range("B12").Value) <> 0# Then gap = CDbl(wsT.Range("B12").Value) / 100#
    End If

    aTot = Val(wsM.Range("C15").Value)
    bTot = Val(wsM.Range("D15").Value)
    aCore = Val(wsM.Range("C16").Value)
    bCore = Val(wsM.Range("D16").Value)
    aCoreCC = Val(wsM.Range("C17").Value)
    bCoreCC = Val(wsM.Range("D17").Value)
    If aTot = 0# And bTot = 0# Then Exit Sub

    Dim dT As Object
    If sIdx = 1 Then
        Set dT = dTokyoS
    ElseIf sIdx = 3 Then
        Set dT = dTokyoCC
    Else
        Set dT = dTokyoC
    End If
    If dT Is Nothing Then
        Call Log("BuildTaylorPath: CPI データが未取得 (先に全自動更新/データ取得を実行してください)")
        Exit Sub
    End If

    rho = EstimateAr1(dT, 36)
    If rho < 0# Then rho = 0#
    If rho > 1# Then rho = 1#

    ReDim ymArr(0 To NM_TAYLOR - 1)
    ReDim piArr(0 To NM_TAYLOR - 1)
    ReDim iEqArr(0 To NM_TAYLOR - 1)
    ReDim iPath(0 To NM_TAYLOR - 1)

    ym = ymT
    For mm = 0 To NM_TAYLOR - 1
        ymArr(mm) = ym
        ym = NextYm(ym)
    Next mm

    If sIdx = 1 Then
        pNow = Val(wsM.Range("C20").Value)
    ElseIf sIdx = 3 Then
        pNow = Val(wsM.Range("C22").Value)
    Else
        pNow = Val(wsM.Range("C21").Value)
    End If
    piArr(0) = pNow

    xPrev = Val(wsM.Range("C9").Value)
    If xPrev = 0# Then xPrev = Val(wsM.Range("C8").Value)
    For mm = 1 To NM_TAYLOR - 1
        xT = piStar * 100# + rho * (xPrev - piStar * 100#)
        xPrev = xT
        If sIdx = 1 Then
            piArr(mm) = aTot * xT + bTot
        ElseIf sIdx = 3 Then
            piArr(mm) = aCoreCC * xT + bCoreCC
        Else
            piArr(mm) = aCore * xT + bCore
        End If
    Next mm

    ' --- 原油シナリオ (CPI パススルー経由で政策金利シミュレーションに反映) ---
    Dim oilBase As Double, passThru As Double, selSc As Long
    Dim scS(1 To 3) As Long, scP(1 To 3) As Long
    Dim scA(1 To 3) As Double, scR(1 To 3) As Double
    Dim oilIdx(1 To 3, 1 To NM_TAYLOR) As Double
    Dim scPath(1 To 3, 1 To NM_TAYLOR) As Double
    Dim piTmp(1 To NM_TAYLOR) As Double
    Dim idxTmp As Variant
    Dim s As Long, mm2 As Long
    Dim iEqTmp As Double, iPathTmp As Double, iPathPrevS As Double, iEqPrevS As Double

    SetRateParam wsT, "B37", 100#
    SetRateParam wsT, "B38", 0.3
    SetRateParam wsT, "B39", 2
    SetRateParam wsT, "B42", 3
    SetRateParam wsT, "C42", 6
    SetRateParam wsT, "D42", 25
    SetRateParam wsT, "E42", 8
    SetRateParam wsT, "B43", 0
    SetRateParam wsT, "C43", 0
    SetRateParam wsT, "D43", 0
    SetRateParam wsT, "E43", 0
    SetRateParam wsT, "B44", 6
    SetRateParam wsT, "C44", 9
    SetRateParam wsT, "D44", -20
    SetRateParam wsT, "E44", -8
    oilBase = RateParam(wsT, "B37", 100#)
    passThru = RateParam(wsT, "B38", 0.3)
    selSc = CLng(RateParam(wsT, "B39", 2#))
    If selSc < 1 Then selSc = 1
    If selSc > 3 Then selSc = 3
    For s = 1 To 3
        scS(s) = CLng(RateParam(wsT, "B" & (41 + s), 0#))
        scP(s) = CLng(RateParam(wsT, "C" & (41 + s), 0#))
        scA(s) = RateParam(wsT, "D" & (41 + s), 0#)
        scR(s) = RateParam(wsT, "E" & (41 + s), 0#)
        idxTmp = OilIndexPath(scS(s), scP(s), scA(s), scR(s), NM_TAYLOR)
        For mm2 = 1 To NM_TAYLOR
            oilIdx(s, mm2) = CDbl(idxTmp(mm2))
        Next mm2
    Next s
    ' 選択シナリオで将来月の CPI 予想を調整 (対象月を除く)
    For mm2 = 2 To NM_TAYLOR
        piArr(mm2 - 1) = piArr(mm2 - 1) + passThru * oilIdx(selSc, mm2) / 10#
    Next mm2

    ' --- テイラールール ---
    iPrev = iNow
    iEqPrev = rStar + piArr(0) / 100# + bPi * (piArr(0) / 100# - piStar) + bY * gap
    For mm = 0 To NM_TAYLOR - 1
        iEqArr(mm) = rStar + piArr(mm) / 100# + bPi * (piArr(mm) / 100# - piStar) + bY * gap
        If mm = 0 Then
            iPath(mm) = iEqArr(0) + rhoS * (iPrev - iEqPrev)
        Else
            iPath(mm) = iEqArr(mm) + rhoS * (iPath(mm - 1) - iEqArr(mm - 1))
        End If
    Next mm

    wsT.Range("A15:I34").ClearContents   ' シナリオ定義 (36行目以降) は消さない
    wsT.Range("A15").Value = "年月"
    wsT.Range("B15").Value = "全国CPI予想(%)"
    wsT.Range("C15").Value = "目標(%)"
    wsT.Range("D15").Value = "産出ギャップ(%)"
    wsT.Range("E15").Value = "r*(%)"
    wsT.Range("F15").Value = "均衡政策金利(%)"
    wsT.Range("G15").Value = "政策金利の目安(%)"
    wsT.Range("H15").Value = "現行との差(pt)"
    wsT.Range("I15").Value = "現行政策金利(%)"
    wsT.Range("A15:I15").Font.Name = "メイリオ"
    wsT.Range("A15:I15").Font.Bold = True
    wsT.Range("A15:I15").Font.Size = 10
    wsT.Range("A15:I15").Font.Color = RGB(31, 78, 121)
    wsT.Range("A15:I15").Interior.Color = RGB(217, 232, 247)
    wsT.Range("A15:I15").HorizontalAlignment = xlCenter
    wsT.Range("A15:I15").Borders.LineStyle = xlContinuous
    wsT.Range("A15:I15").Borders.Color = RGB(146, 176, 209)

    For mm = 0 To NM_TAYLOR - 1
        r = 16 + mm
        wsT.Cells(r, 1).Value = CStr(ymArr(mm) \ 100) & "/" & Format$(ymArr(mm) Mod 100, "00")
        wsT.Cells(r, 2).Value = piArr(mm)
        wsT.Cells(r, 3).Value = piStar * 100#
        wsT.Cells(r, 4).Value = gap * 100#
        wsT.Cells(r, 5).Value = rStar * 100#
        wsT.Cells(r, 6).Value = iEqArr(mm) * 100#
        wsT.Cells(r, 7).Value = iPath(mm) * 100#
        wsT.Cells(r, 8).Value = (iPath(mm) - iNow) * 100#
        wsT.Cells(r, 9).Value = iNow * 100#
    Next mm
    wsT.Range("B16:I27").NumberFormat = "0.00"
    wsT.Range("A16:A27").HorizontalAlignment = xlCenter

    wsT.Range("A4").Value = "目標インフレ率 π* (%)"
    wsT.Range("A5").Value = "インフレ乖離の係数 b_π"
    wsT.Range("A6").Value = "産出ギャップの係数 b_y"
    wsT.Range("A7").Value = "短期金利の平滑化係数 ρ"
    wsT.Range("A8").Value = "現行政策金利 (%)"
    wsT.Range("A9").Value = "CPI系列 (1=総合 2=コア 3=コアコア)"
    wsT.Range("A10").Value = "r* の扱い (0=モデル値 1=手動)"
    wsT.Range("A11").Value = "r* 手動上書き (%)"
    wsT.Range("A12").Value = "産出ギャップ 手動上書き (%) (0=モデル値)"
    wsT.Range("A4:B12").Font.Name = "メイリオ"
    wsT.Range("A4:B12").Font.Size = 10
    wsT.Range("A4:B12").Borders.LineStyle = xlContinuous
    wsT.Range("A4:B12").Borders.Color = RGB(191, 191, 191)
    wsT.Range("A4:B12").Borders.Weight = xlThin
    wsT.Range("B4:B12").Interior.Color = RGB(255, 247, 214)
    wsT.Range("B4:B12").Font.Bold = True
    wsT.Range("B4:B12").HorizontalAlignment = xlRight
    wsT.Range("B4").NumberFormat = "0.0"
    wsT.Range("B8").NumberFormat = "0.00"
    wsT.Range("B11:B12").NumberFormat = "0.00"

    msg = "前提: r* = " & Format$(rStar * 100#, "0.00") & "% / 産出ギャップ = " & Format$(gap * 100#, "0.00") & "% / AR(1)係数 = " & Format$(rho, "0.00")
    msg = msg & " / 現行政策金利 = " & Format$(iNow * 100#, "0.00") & "% (日銀の目標金利・決定会合後に手動更新) / 市場無担保コール = " & Format$(LastCallRate(), "0.00") & "% (直近四半期平均・自動取得)"
    msg = msg & " / 原油シナリオ = " & scName(selSc) & " (振幅 " & Format$(scA(selSc), "0.0") & "%・パススルー " & Format$(passThru, "0.00") & "pp)"
    If rstarMode = 1 Then msg = msg & " (r*は手動値)"
    wsT.Range("A13").Value = msg
    wsT.Range("A13").Font.Name = "メイリオ"
    wsT.Range("A13").Font.Size = 9
    wsT.Range("A13").Font.Color = RGB(120, 120, 120)

    ' --- r* 感応度: 日銀推計レンジ (▲0.9% / 中心0% / 上限+0.5%) での i_eq ---
    Dim rLo As Double, rMid As Double, rHi As Double
    Dim iLo As Double, iMid As Double, iHi As Double, iMod As Double
    Dim piLast As Double
    piLast = piArr(NM_TAYLOR - 1) / 100#
    rLo = -0.009: rMid = 0#: rHi = 0.005
    iLo = rLo + piLast + bPi * (piLast - piStar) + bY * gap
    iMid = rMid + piLast + bPi * (piLast - piStar) + bY * gap
    iHi = rHi + piLast + bPi * (piLast - piStar) + bY * gap
    iMod = iEqArr(NM_TAYLOR - 1)
    wsT.Range("A29").Value = "r* 感応度 (2027/08時点の均衡政策金利 i_eq のレンジ)"
    wsT.Range("A29").Font.Name = "メイリオ"
    wsT.Range("A29").Font.Size = 10
    wsT.Range("A29").Font.Bold = True
    wsT.Range("A29").Font.Color = RGB(31, 78, 121)
    wsT.Range("A30").Value = "日銀推計レンジ下限 r* = -0.9%"
    wsT.Range("B30").Value = Round(iLo * 100#, 2)
    wsT.Range("A31").Value = "同 中心 r* = 0.0%"
    wsT.Range("B31").Value = Round(iMid * 100#, 2)
    wsT.Range("A32").Value = "同 上限 r* = +0.5%"
    wsT.Range("B32").Value = Round(iHi * 100#, 2)
    wsT.Range("A33").Value = "本ツールのモデル値 r* = " & Format$(rStar * 100#, "0.00") & "%"
    wsT.Range("B33").Value = Round(iMod * 100#, 2)
    wsT.Range("A34").Value = "現行政策金利 (手動入力)"
    wsT.Range("B34").Value = Round(iNow * 100#, 2)
    wsT.Range("A30:B34").Font.Name = "メイリオ"
    wsT.Range("A30:B34").Font.Size = 10
    wsT.Range("B30:B34").NumberFormat = "0.00"
    wsT.Range("A30:B34").Borders.LineStyle = xlContinuous
    wsT.Range("A30:B34").Borders.Color = RGB(191, 191, 191)
    wsT.Range("B30:B34").Interior.Color = RGB(242, 242, 242)
    wsT.Range("B30:B34").Font.Bold = True
    wsT.Range("A33:B33").Font.Color = RGB(31, 78, 121)
    wsT.Range("A34:B34").Font.Color = RGB(120, 120, 120)

    ' --- 原油シナリオの定義・指数パス・比較パスの書き出し ---
    wsT.Range("A36").Value = "原油シナリオ (指数ベース 100・変化%)"
    wsT.Range("A36").Font.Name = "メイリオ"
    wsT.Range("A36").Font.Size = 10
    wsT.Range("A36").Font.Bold = True
    wsT.Range("A36").Font.Color = RGB(31, 78, 121)
    wsT.Range("A37").Value = "基準指数"
    wsT.Range("A38").Value = "CPIパススルー (指数+10 → CPI +pp)"
    wsT.Range("A39").Value = "選択シナリオ (1=上昇 2=基準 3=下落)"
    wsT.Range("A41").Value = "シナリオ"
    wsT.Range("B41").Value = "開始月"
    wsT.Range("C41").Value = "ピーク月"
    wsT.Range("D41").Value = "振幅(%)"
    wsT.Range("E41").Value = "残存(%)"
    wsT.Range("A42").Value = "上昇"
    wsT.Range("A43").Value = "基準"
    wsT.Range("A44").Value = "下落"
    wsT.Range("A41:E41").Font.Name = "メイリオ"
    wsT.Range("A41:E41").Font.Size = 9
    wsT.Range("A41:E41").Font.Bold = True
    wsT.Range("A41:E44").Borders.LineStyle = xlContinuous
    wsT.Range("A41:E44").Borders.Color = RGB(191, 191, 191)
    wsT.Range("A37:B39").Font.Name = "メイリオ"
    wsT.Range("A37:B39").Font.Size = 10
    wsT.Range("A37:B39").Borders.LineStyle = xlContinuous
    wsT.Range("A37:B39").Borders.Color = RGB(191, 191, 191)
    wsT.Range("B37").NumberFormat = "0.0"
    wsT.Range("B38").NumberFormat = "0.00"
    wsT.Range("B39").NumberFormat = "0"
    wsT.Range("B42:E44").NumberFormat = "0.0"
    wsT.Range("B37:B39").Interior.Color = RGB(255, 247, 214)
    wsT.Range("B42:E44").Interior.Color = RGB(255, 247, 214)
    wsT.Range("B37:B39").Font.Bold = True
    wsT.Range("B42:E44").Font.Bold = True
    wsT.Range("A46").Value = "原油指数パス (変化%) とシナリオ別の政策金利の目安 (%)"
    wsT.Range("A46").Font.Name = "メイリオ"
    wsT.Range("A46").Font.Size = 9
    wsT.Range("A46").Font.Bold = True
    For mm2 = 1 To NM_TAYLOR
        wsT.Cells(47, 1 + mm2).Value = CStr(ymArr(mm2 - 1) \ 100) & "/" & Format$(ymArr(mm2 - 1) Mod 100, "00")
        For s = 1 To 3
            wsT.Cells(47 + s, 1 + mm2).Value = oilIdx(s, mm2)
        Next s
    Next mm2
    wsT.Range("A48").Value = "指数 上昇"
    wsT.Range("A49").Value = "指数 基準"
    wsT.Range("A50").Value = "指数 下落"
    wsT.Range("B48:M50").NumberFormat = "0.0"
    wsT.Range("A48:M50").Font.Name = "メイリオ"
    wsT.Range("A48:M50").Font.Size = 9
    ' シナリオ別の政策金利パス (比較用)
    iEqPrevS = rStar + piArr(0) / 100# + bPi * (piArr(0) / 100# - piStar) + bY * gap
    For s = 1 To 3
        iPathPrevS = iNow
        For mm2 = 1 To NM_TAYLOR
            piTmp(mm2) = piArr(mm2 - 1)
            If mm2 >= 2 Then piTmp(mm2) = piTmp(mm2) + passThru * oilIdx(s, mm2) / 10#
            iEqTmp = rStar + piTmp(mm2) / 100# + bPi * (piTmp(mm2) / 100# - piStar) + bY * gap
            iPathTmp = iEqTmp + rhoS * (iPathPrevS - iEqPrevS)
            scPath(s, mm2) = iPathTmp
            iPathPrevS = iPathTmp
            iEqPrevS = iEqTmp
        Next mm2
    Next s
    wsT.Range("P15").Value = "シナリオ比較: 政策金利の目安 (%)"
    wsT.Range("P15").Font.Name = "メイリオ"
    wsT.Range("P15").Font.Size = 9
    wsT.Range("P15").Font.Bold = True
    wsT.Range("P16").Value = "年月"
    wsT.Range("P17").Value = "上昇"
    wsT.Range("P18").Value = "基準"
    wsT.Range("P19").Value = "下落"
    For mm2 = 1 To NM_TAYLOR
        wsT.Cells(16, 16 + mm2).Value = CStr(ymArr(mm2 - 1) \ 100) & "/" & Format$(ymArr(mm2 - 1) Mod 100, "00")
        For s = 1 To 3
            wsT.Cells(16 + s, 16 + mm2).Value = scPath(s, mm2) * 100#
        Next s
    Next mm2
    wsT.Range("Q17:AB19").NumberFormat = "0.00"
    wsT.Range("P16:AB19").Font.Name = "メイリオ"
    wsT.Range("P16:AB19").Font.Size = 9

    Call BuildRateChart(wsT, NM_TAYLOR)
    Call BuildRateChart2(wsT)
    Call BuildRateChart3(wsT, NM_TAYLOR)

    Set wsM2 = Worksheets(SH_MAIN)
    wsM2.Range("B31").Value = "④ 政策金利の目安（テイラールール）"
    wsM2.Range("B31").Font.Name = "メイリオ"
    wsM2.Range("B31").Font.Size = 10
    wsM2.Range("B31").Font.Bold = True
    wsM2.Range("B31").Font.Color = RGB(31, 78, 121)
    wsM2.Range("B32").Value = "12ヶ月先の政策金利の目安 / 現行政策金利との差"
    wsM2.Range("C32").Value = Round(iPath(NM_TAYLOR - 1) * 100#, 2)
    wsM2.Range("D32").Value = Round((iPath(NM_TAYLOR - 1) - iNow) * 100#, 2) & " pt"
    wsM2.Range("C32").NumberFormat = "0.00"
    wsM2.Range("D32").NumberFormat = "0.00"
    wsM2.Range("E32").Value = "→ 金利予想シート参照"
    wsM2.Range("E32").Font.Size = 9
    wsM2.Range("E32").Font.Color = RGB(120, 120, 120)
    Exit Sub
TErr:
    Call Log("BuildTaylorPath error: " & Err.Number & " - " & Err.Description)
End Sub

' 金利予想シート (無ければ作成)
Private Function GetRateSheet() As Worksheet
    Dim ws As Worksheet
    Dim i As Long
    On Error Resume Next
    Set ws = Worksheets("金利予想")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        ws.Name = "金利予想"
        For i = 1 To 40
            ws.Cells(i, 1).Value = ""
        Next i
        ws.Range("A1").Value = "テイラールールによる政策金利の目安"
        ws.Range("A1").Font.Name = "メイリオ"
        ws.Range("A1").Font.Size = 14
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Color = RGB(31, 78, 121)
        ws.Range("A2").Value = "i_eq = r* + pi + b_pi*(pi - pi*) + b_y*gap   /   i = i_eq + rho*(i-1 - i_eq-1)"
        ws.Range("A2").Font.Name = "メイリオ"
        ws.Range("A2").Font.Size = 9
        ws.Range("A2").Font.Color = RGB(120, 120, 120)
        ws.Columns("A").ColumnWidth = 30
        ws.Columns("B:I").ColumnWidth = 15
    End If
    Set GetRateSheet = ws
End Function

' AR(1) 係数: x(t) = pi* + rho*(x(t-1) - pi*) の rho を標本推定
Private Function EstimateAr1(d As Object, ByVal nMax As Long) As Double
    On Error GoTo TErr
    Dim n As Long, i As Long, ym As Long, nUse As Long
    Dim xs() As Double, ys() As Double
    Dim mx As Double, my As Double, sxy As Double, sxx As Double
    nUse = mNCommon
    If nUse > nMax Then nUse = nMax
    If nUse < 12 Then
        EstimateAr1 = 0.8
        Exit Function
    End If
    ReDim xs(1 To nUse - 1)
    ReDim ys(1 To nUse - 1)
    For i = mNCommon - nUse + 1 To mNCommon - 1
        ym = mMonths(i)
        If d.Exists(ym) And d.Exists(mMonths(i + 1)) Then
            n = n + 1
            xs(n) = CDbl(d(ym))
            ys(n) = CDbl(d(mMonths(i + 1)))
        End If
    Next i
    If n < 8 Then
        EstimateAr1 = 0.8
        Exit Function
    End If
    For i = 1 To n
        mx = mx + xs(i)
        my = my + ys(i)
    Next i
    mx = mx / n
    my = my / n
    For i = 1 To n
        sxy = sxy + (xs(i) - mx) * (ys(i) - my)
        sxx = sxx + (xs(i) - mx) * (xs(i) - mx)
    Next i
    If sxx = 0# Then
        EstimateAr1 = 0.8
    Else
        EstimateAr1 = sxy / sxx
    End If
    Exit Function
TErr:
    EstimateAr1 = 0.8
End Function

' 翌月 (YYYYMM)
Private Function NextYm(ByVal ymIn As Long) As Long
    Dim y As Long, m As Long
    y = ymIn \ 100
    m = ymIn Mod 100
    m = m + 1
    If m > 12 Then
        m = 1
        y = y + 1
    End If
    NextYm = y * 100 + m
End Function

' 金利予想シートのグラフ
Private Sub BuildRateChart(ws As Worksheet, ByVal nM As Long)
    Dim co As ChartObject
    Dim ch As Chart
    Dim s As Series
    Dim xs As Range
    On Error Resume Next
    For Each co In ws.ChartObjects
        co.Delete
    Next co
    On Error GoTo 0
    Set co = ws.ChartObjects.Add(ws.Range("A37").Left, ws.Range("A37").Top, 700, 320)
    co.Name = "RateChart"
    Set ch = co.Chart
    ch.ChartType = xlLine
    ch.HasTitle = True
    ch.ChartTitle.Text = "テイラールールによる政策金利の目安と全国CPI予想"
    ch.ChartTitle.Font.Name = "メイリオ"
    ch.ChartTitle.Font.Size = 12
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Font.Color = RGB(31, 78, 121)
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.Legend.Font.Name = "メイリオ"
    ch.Legend.Font.Size = 9
    ch.ChartArea.Font.Name = "メイリオ"
    ch.ChartArea.Font.Size = 9
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.Axes(xlValue).HasTitle = True
    ch.Axes(xlValue).AxisTitle.Text = "％"
    ch.Axes(xlValue).AxisTitle.Font.Name = "メイリオ"
    ch.Axes(xlValue).AxisTitle.Font.Size = 9
    ch.Axes(xlValue).TickLabels.Font.Name = "メイリオ"
    ch.Axes(xlValue).TickLabels.Font.Size = 9
    ch.Axes(xlValue).TickLabels.NumberFormat = "0.0"
    ch.Axes(xlValue).HasMajorGridlines = True
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
    ch.Axes(xlValue).MajorGridlines.Format.Line.Weight = 0.5
    ch.Axes(xlValue).Format.Line.ForeColor.RGB = RGB(191, 191, 191)
    ch.Axes(xlValue).Format.Line.Weight = 0.75
    ch.Axes(xlCategory).TickLabels.Font.Name = "メイリオ"
    ch.Axes(xlCategory).TickLabels.Font.Size = 8
    ch.Axes(xlCategory).TickLabelSpacing = 1
    ch.Axes(xlCategory).TickLabels.Orientation = -45
    ch.Axes(xlCategory).HasMajorGridlines = False
    ch.Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(191, 191, 191)
    ch.Axes(xlCategory).Format.Line.Weight = 0.75
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(252, 252, 252)
    ch.PlotArea.Format.Line.Visible = msoFalse
    Set xs = ws.Range("A16:A" & (15 + nM))
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "全国CPI予想"
    s.Values = ws.Range("B16:B" & (15 + nM))
    s.XValues = xs
    s.ChartType = xlLine
    s.Format.Line.ForeColor.RGB = RGB(112, 173, 71)
    s.Format.Line.Weight = 2#
    s.MarkerStyle = xlMarkerStyleCircle
    s.MarkerSize = 4
    s.MarkerBackgroundColor = RGB(255, 255, 255)
    s.MarkerForegroundColor = RGB(112, 173, 71)
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "政策金利の目安 (テイラールール)"
    s.Values = ws.Range("G16:G" & (15 + nM))
    s.XValues = xs
    s.ChartType = xlLine
    s.Format.Line.ForeColor.RGB = RGB(192, 80, 77)
    s.Format.Line.Weight = 2.5
    s.MarkerStyle = xlMarkerStyleCircle
    s.MarkerSize = 4
    s.MarkerBackgroundColor = RGB(255, 255, 255)
    s.MarkerForegroundColor = RGB(192, 80, 77)
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "現行政策金利"
    s.Values = ws.Range("I16:I" & (15 + nM))
    s.XValues = xs
    s.ChartType = xlLine
    s.Format.Line.ForeColor.RGB = RGB(120, 120, 120)
    s.Format.Line.Weight = 1.25
    s.Format.Line.DashStyle = msoLineSysDot
    s.MarkerStyle = xlMarkerStyleNone
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "均衡政策金利 i_eq"
    s.Values = ws.Range("F16:F" & (15 + nM))
    s.XValues = xs
    s.ChartType = xlLine
    s.Format.Line.ForeColor.RGB = RGB(192, 80, 77)
    s.Format.Line.Weight = 1.25
    s.Format.Line.DashStyle = msoLineDash
    s.MarkerStyle = xlMarkerStyleNone
End Sub

Public Sub AutoUpdate()
    On Error GoTo ErrH
    Dim wsMain As Worksheet
    Set wsMain = Worksheets(SH_MAIN)

    Call SetStatus("ステータス: ①データ取得中…", RGB(0, 120, 215))
    Call FetchAllData

    ' --- 対象月 = 東京の最新月（回帰推定より前に確定させる） ---
    Dim kT As Variant
    mTarget = 0
    For Each kT In dTokyoS.Keys
        If CLng(kT) > mTarget Then mTarget = CLng(kT)
    Next kT
    If mTarget = 0 Then
        Call SetStatus("エラー: 東京のデータが取得できませんでした", RGB(200, 0, 0))
        Exit Sub
    End If

    ' --- 東京最新値を①へ ---
    Dim ym As Long
    ym = mTarget
    wsMain.Range("C7").Value = DateSerial(ym \ 100, ym Mod 100, 1)
    wsMain.Range("C7").NumberFormat = "yyyy/mm"
    wsMain.Range("C8").Value = dTokyoS(ym)
    wsMain.Range("C9").Value = dTokyoC(ym)
    wsMain.Range("C10").Value = dTokyoCC(ym)
    If dTokyoIdx.Exists(ym) Then
        wsMain.Range("C11").Value = dTokyoIdx(ym)
        wsMain.Range("C11").NumberFormat = "0.0"
    End If

    Call SetStatus("ステータス: ②回帰推定中…", RGB(70, 130, 80))
    Call FitAllRegressions

    Call SetStatus("ステータス: ③予想計算中…", RGB(180, 60, 0))
    Call CalcNational

    ' --- R*推計（オプショナル、失敗してもCPI結果は影響されない） ---
    On Error Resume Next
    Dim wsR As Worksheet
    Set wsR = Nothing: Set wsR = Worksheets("Rstar")
    If Not wsR Is Nothing Then
        Dim rstarAuto As String: rstarAuto = Trim$(CStr(wsR.Cells(10, 2).Value))
        If rstarAuto = "ON" Then
            On Error GoTo RStarErr
            Call ModuleRStar.RStar_Run
            On Error GoTo ErrH
        End If
    End If
    On Error GoTo ErrH

    ' --- テイラールールによる政策金利の目安 (R* の結果を使うので R* 計算の後) ---
    On Error Resume Next
    Call BuildTaylorPath
    On Error GoTo ErrH
    wsMain.Range("B33").Value = "最終更新: " & Format$(Now, "yyyy/mm/dd hh:mm:ss") & "  (e-Stat 2025年基準)"
    Call SetStatus("ステータス: 全自動更新完了 ✓", RGB(0, 150, 0))
    Exit Sub

RStarErr:
    ' R*失敗はログに記録するが、CPI結果はロールバックしない
    Call SetStatus("ステータス: R*推計エラー " & Err.Number & ": " & Err.Description, RGB(200, 60, 0))
    Call Log("RStar error: " & Err.Number & " - " & Err.Description)
    Resume Next

ErrH:
    Call SetStatus("ステータス: エラー " & Err.Number & " - " & Err.Description, RGB(200, 0, 0))
End Sub

' r* 感応度のグラフ (A30:B34 の5行: 下限/中心/上限/モデル値/現行政策金利)
Private Sub BuildRateChart2(ws As Worksheet)
    Dim co As ChartObject
    Dim ch As Chart
    Dim s As Series
    On Error Resume Next
    ws.ChartObjects("RateChart2").Delete
    On Error GoTo 0
    Set co = ws.ChartObjects.Add(ws.Range("A58").Left, ws.Range("A58").Top, 700, 300)
    co.Name = "RateChart2"
    Set ch = co.Chart
    ch.ChartType = xlColumnClustered
    ch.HasTitle = True
    ch.ChartTitle.Text = "r* の根拠による均衡政策金利 i_eq のレンジ (2027/08時点)"
    ch.ChartTitle.Font.Name = "メイリオ"
    ch.ChartTitle.Font.Size = 12
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Font.Color = RGB(31, 78, 121)
    ch.HasLegend = False
    ch.ChartArea.Font.Name = "メイリオ"
    ch.ChartArea.Font.Size = 9
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.Axes(xlValue).HasTitle = True
    ch.Axes(xlValue).AxisTitle.Text = "％"
    ch.Axes(xlValue).AxisTitle.Font.Name = "メイリオ"
    ch.Axes(xlValue).AxisTitle.Font.Size = 9
    ch.Axes(xlValue).TickLabels.Font.Name = "メイリオ"
    ch.Axes(xlValue).TickLabels.Font.Size = 9
    ch.Axes(xlValue).TickLabels.NumberFormat = "0.0"
    ch.Axes(xlValue).HasMajorGridlines = True
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
    ch.Axes(xlValue).MajorGridlines.Format.Line.Weight = 0.5
    ch.Axes(xlCategory).TickLabels.Font.Name = "メイリオ"
    ch.Axes(xlCategory).TickLabels.Font.Size = 9
    ch.Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(191, 191, 191)
    ch.Axes(xlCategory).Format.Line.Weight = 0.75
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(252, 252, 252)
    ch.PlotArea.Format.Line.Visible = msoFalse
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "i_eq"
    s.Values = ws.Range("B30:B34")
    s.XValues = ws.Range("A30:A34")
    s.Points(1).Format.Fill.ForeColor.RGB = RGB(112, 48, 160)
    s.Points(2).Format.Fill.ForeColor.RGB = RGB(91, 155, 213)
    s.Points(3).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
    s.Points(4).Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
    s.Points(5).Format.Fill.ForeColor.RGB = RGB(150, 150, 150)
    s.HasDataLabels = True
    s.DataLabels.Font.Name = "メイリオ"
    s.DataLabels.Font.Size = 9
    s.DataLabels.NumberFormat = "0.00"
End Sub



' シナリオ名の取得
Private Function scName(ByVal s As Long) As String
    If s = 1 Then
        scName = "上昇"
    ElseIf s = 3 Then
        scName = "下落"
    Else
        scName = "基準"
    End If
End Function

' 原油指数パス: 開始月まで横ばい → ピーク月へ線形上昇 → 残存へ線形回落
' 戻り値は 1始まりの配列 (変化率 %)
Private Function OilIndexPath(ByVal startM As Long, ByVal peakM As Long, _
                             ByVal amp As Double, ByVal resid As Double, ByVal nM As Long) As Variant
    Dim out() As Double
    Dim t As Long, rest As Long
    ReDim out(1 To nM)
    For t = 1 To nM
        If t <= startM Then
            out(t) = 0#
        ElseIf t <= peakM Then
            If peakM = startM Then
                out(t) = amp
            Else
                out(t) = amp * (t - startM) / (peakM - startM)
            End If
        Else
            rest = nM - peakM
            If rest <= 0 Then
                out(t) = amp
            Else
                out(t) = amp + (resid - amp) * (t - peakM) / rest
            End If
        End If
    Next t
    OilIndexPath = out
End Function

' シナリオ比較グラフ (上昇/基準/下落 の政策金利の目安)
Private Sub BuildRateChart3(ws As Worksheet, ByVal nM As Long)
    Dim co As ChartObject
    Dim ch As Chart
    Dim s As Series
    Dim xs As Range
    On Error Resume Next
    ws.ChartObjects("RateChart3").Delete
    On Error GoTo 0
    Set co = ws.ChartObjects.Add(ws.Range("A100").Left, ws.Range("A100").Top, 700, 300)
    co.Name = "RateChart3"
    Set ch = co.Chart
    ch.ChartType = xlLine
    ch.HasTitle = True
    ch.ChartTitle.Text = "原油シナリオ別の政策金利の目安 (テイラールール)"
    ch.ChartTitle.Font.Name = "メイリオ"
    ch.ChartTitle.Font.Size = 12
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Font.Color = RGB(31, 78, 121)
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.Legend.Font.Name = "メイリオ"
    ch.Legend.Font.Size = 9
    ch.ChartArea.Font.Name = "メイリオ"
    ch.ChartArea.Font.Size = 9
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.Axes(xlValue).HasTitle = True
    ch.Axes(xlValue).AxisTitle.Text = "％"
    ch.Axes(xlValue).AxisTitle.Font.Name = "メイリオ"
    ch.Axes(xlValue).AxisTitle.Font.Size = 9
    ch.Axes(xlValue).TickLabels.Font.Name = "メイリオ"
    ch.Axes(xlValue).TickLabels.Font.Size = 9
    ch.Axes(xlValue).TickLabels.NumberFormat = "0.0"
    ch.Axes(xlValue).HasMajorGridlines = True
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
    ch.Axes(xlValue).MajorGridlines.Format.Line.Weight = 0.5
    ch.Axes(xlValue).Format.Line.ForeColor.RGB = RGB(191, 191, 191)
    ch.Axes(xlValue).Format.Line.Weight = 0.75
    ch.Axes(xlCategory).TickLabels.Font.Name = "メイリオ"
    ch.Axes(xlCategory).TickLabels.Font.Size = 8
    ch.Axes(xlCategory).TickLabelSpacing = 1
    ch.Axes(xlCategory).TickLabels.Orientation = -45
    ch.Axes(xlCategory).HasMajorGridlines = False
    ch.Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(191, 191, 191)
    ch.Axes(xlCategory).Format.Line.Weight = 0.75
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(252, 252, 252)
    ch.PlotArea.Format.Line.Visible = msoFalse
    Set xs = ws.Range("Q16:AB16")
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "原油上昇シナリオ"
    s.Values = ws.Range("Q17:AB17")
    s.XValues = xs
    s.ChartType = xlLine
    s.Format.Line.ForeColor.RGB = RGB(192, 80, 77)
    s.Format.Line.Weight = 2.25
    s.MarkerStyle = xlMarkerStyleCircle
    s.MarkerSize = 4
    s.MarkerBackgroundColor = RGB(255, 255, 255)
    s.MarkerForegroundColor = RGB(192, 80, 77)
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "基準シナリオ"
    s.Values = ws.Range("Q18:AB18")
    s.XValues = xs
    s.ChartType = xlLine
    s.Format.Line.ForeColor.RGB = RGB(31, 78, 121)
    s.Format.Line.Weight = 2.25
    s.MarkerStyle = xlMarkerStyleCircle
    s.MarkerSize = 4
    s.MarkerBackgroundColor = RGB(255, 255, 255)
    s.MarkerForegroundColor = RGB(31, 78, 121)
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "原油下落シナリオ"
    s.Values = ws.Range("Q19:AB19")
    s.XValues = xs
    s.ChartType = xlLine
    s.Format.Line.ForeColor.RGB = RGB(112, 173, 71)
    s.Format.Line.Weight = 2.25
    s.MarkerStyle = xlMarkerStyleCircle
    s.MarkerSize = 4
    s.MarkerBackgroundColor = RGB(255, 255, 255)
    s.MarkerForegroundColor = RGB(112, 173, 71)
End Sub
