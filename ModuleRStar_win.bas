Attribute VB_Name = "ModuleRStar"
'================================================================
' ModuleRStar.bas  --  HLW (Holston-Laubach-Williams) r* VBA port
'----------------------------------------------------------------
' rstar_check.py (_unpack_stage3 + kalman_states + RTS smoother)
' と完全一致させる移植。固定係数 theta/xi00/P00 は
' rstar_fixed_coeff.csv (Python MLE 結果) の値をハードコード。
' データ取得は RStar_FetchData (内閣府GDP xlsx + 日銀コール, キー不要)。
' 状態(0-based): [y100_t, y100_t-1, y100_t-2, g*, pi*, z, r~]
'   r*_t = 4*g*_t + z_t  (trend + gap)
'================================================================
Option Explicit
Public Const SH_RSTAR As String = "Rstar"
'--- シートレイアウト (行番号はここで一元管理) ---
' 入力テーブル: ヘッダ I_HDR 行 / データ I_TOP..I_BOTTOM
' 結果テーブル: ヘッダ O_HDR 行 / データ O_TOP..O_TOP+3   (入力と絶対に重ならない)
' 入力列: A=q_idx, B=label(YYYYQn), C=log_output, D=inflation, E=exp_infl, F=real_rate, G=nom_rate
Private Const I_HDR As Long = 12
Private Const I_TOP As Long = 13
Private Const I_BOTTOM As Long = 412        ' 最大 400 四半期分 (100 年)
Private Const O_HDR As Long = 430
Private Const O_TOP As Long = 431
Private m_rf() As Double, m_rs() As Double, m_lab() As String, m_T As Long
Private m_bojErr As String

' lambda_g / lambda_z (v4): 非MLE版 stage1/2 のメディアン不偏推定値 (季節調整済データ)
' (MLE版stage1の潜在出力が過度に平滑になり lam=0 に退化するため OLS 初期化版を使用)
Private Const LAMBDA_G As Double = 0.192798
Private Const LAMBDA_Z As Double = 0
Private Const SOKUHO_BASE As String = "https://www.esri.cao.go.jp/jp/sna/data/data_list/sokuhou/files/"
Private Const SOKUHO_LIST_URL As String = "https://www.esri.cao.go.jp/jp/sna/data/data_list/sokuhou/files/files_sokuhou.html"
Private Const BOJ_GAP_URL As String = "https://www.boj.or.jp/research/research_data/gap/gap.xlsx"
Private Const BOJ_CALL_URL As String = "https://www.stat-search.boj.or.jp/api/v1/getDataCode?format=csv&lang=jp&db=FM02&code=STRACLUCON&startDate=199401&endDate=202612"

'--- matrix utils (0-based Variant 2-D) ---
Private Function MZ(r As Long, c As Long) As Variant
    Dim M() As Double: ReDim M(0 To r - 1, 0 To c - 1): MZ = M
End Function
Private Function MI(n As Long) As Variant
    Dim M() As Double, i As Long: ReDim M(0 To n - 1, 0 To n - 1)
    For i = 0 To n - 1: M(i, i) = 1#: Next i: MI = M
End Function
Private Function MM(ByVal Av As Variant, ByVal Bv As Variant) As Variant
    Dim A() As Double, B() As Double
    A = Av: B = Bv
    Dim a1 As Long, a2 As Long, b1 As Long, b2 As Long
    a1 = UBound(A, 1) - LBound(A, 1) + 1: a2 = UBound(A, 2) - LBound(A, 2) + 1
    b1 = UBound(B, 1) - LBound(B, 1) + 1: b2 = UBound(B, 2) - LBound(B, 2) + 1
    If a2 <> b1 Then Err.Raise vbObjectError + 601, , "MM: dim mismatch"
    Dim Rout() As Double: ReDim Rout(0 To a1 - 1, 0 To b2 - 1)
    Dim i As Long, j As Long, k As Long, s As Double
    For i = 0 To a1 - 1: For j = 0 To b2 - 1
        s = 0#: For k = 0 To a2 - 1: s = s + A(i, k) * B(k, j): Next k
        Rout(i, j) = s: Next j: Next i
    MM = Rout
End Function
Private Function MT(ByVal Av As Variant) As Variant
    Dim A() As Double: A = Av
    Dim r As Long, c As Long, i As Long, j As Long
    r = UBound(A, 1) + 1: c = UBound(A, 2) + 1
    Dim Rout() As Double: ReDim Rout(0 To c - 1, 0 To r - 1)
    For i = 0 To r - 1: For j = 0 To c - 1: Rout(j, i) = A(i, j): Next j: Next i
    MT = Rout
End Function
Private Function MA(ByVal Av As Variant, ByVal Bv As Variant) As Variant
    Dim A() As Double, B() As Double
    A = Av: B = Bv
    Dim r As Long, c As Long, i As Long, j As Long
    r = UBound(A, 1) + 1: c = UBound(A, 2) + 1
    Dim Rout() As Double: ReDim Rout(0 To r - 1, 0 To c - 1)
    For i = 0 To r - 1: For j = 0 To c - 1: Rout(i, j) = A(i, j) + B(i, j): Next j: Next i
    MA = Rout
End Function
Private Function MS(ByVal Av As Variant, ByVal Bv As Variant) As Variant
    Dim A() As Double, B() As Double
    A = Av: B = Bv
    Dim r As Long, c As Long, i As Long, j As Long
    r = UBound(A, 1) + 1: c = UBound(A, 2) + 1
    Dim Rout() As Double: ReDim Rout(0 To r - 1, 0 To c - 1)
    For i = 0 To r - 1: For j = 0 To c - 1: Rout(i, j) = A(i, j) - B(i, j): Next j: Next i
    MS = Rout
End Function
Private Function MSc(ByVal sc As Double, ByVal Av As Variant) As Variant
    Dim A() As Double: A = Av
    Dim r As Long, c As Long, i As Long, j As Long
    r = UBound(A, 1) + 1: c = UBound(A, 2) + 1
    Dim Rout() As Double: ReDim Rout(0 To r - 1, 0 To c - 1)
    For i = 0 To r - 1: For j = 0 To c - 1: Rout(i, j) = sc * A(i, j): Next j: Next i
    MSc = Rout
End Function
Private Function Sym(ByVal Av As Variant) As Variant
    ' 0.5*(A+A') -- Python _psd/_safe_inv の対称化に相当
    Sym = MSc(0.5#, MA(Av, MT(Av)))
End Function
Private Function MN(ByVal Av As Variant) As Variant
    Dim A() As Double: A = Av
    Dim n As Long, i As Long, j As Long, k As Long
    n = UBound(A, 1) + 1
    Dim M() As Double: ReDim M(0 To n - 1, 0 To 2 * n - 1)
    For i = 0 To n - 1: For j = 0 To n - 1
        M(i, j) = A(i, j): M(i, j + n) = IIf(i = j, 1#, 0#): Next j: Next i
    Dim pivRow As Long, best As Double, tmp As Double
    For k = 0 To n - 1
        best = Abs(M(k, k)): pivRow = k
        For i = k + 1 To n - 1
            If Abs(M(i, k)) > best Then
                best = Abs(M(i, k)): pivRow = i
            End If
        Next i
        If best < 0.000000000000001 Then MN = MZ(n, n): Exit Function
        If pivRow <> k Then
            For j = 0 To 2 * n - 1
                tmp = M(k, j): M(k, j) = M(pivRow, j): M(pivRow, j) = tmp
            Next j
        End If
        Dim pv As Double: pv = M(k, k)
        For j = 0 To 2 * n - 1: M(k, j) = M(k, j) / pv: Next j
        For i = 0 To n - 1
            If i <> k Then
                Dim factor As Double: factor = M(i, k)
                If factor <> 0# Then
                    For j = 0 To 2 * n - 1
                        M(i, j) = M(i, j) - factor * M(k, j)
                    Next j
                End If
            End If
        Next i
    Next k
    Dim Rout() As Double: ReDim Rout(0 To n - 1, 0 To n - 1)
    For i = 0 To n - 1: For j = 0 To n - 1: Rout(i, j) = M(i, j + n): Next j: Next i
    MN = Rout
End Function
Private Function SymInv(ByVal Av As Variant) As Variant
    ' _safe_inv 相当: 対称化 -> 逆行列(部分ピボット) -> 特異ならジッタ再試行
    Dim inv As Variant: inv = MN(Sym(Av))
    Dim n As Long, i As Long, j As Long, allzero As Boolean
    n = UBound(inv, 1) + 1
    allzero = True
    For i = 0 To n - 1: For j = 0 To n - 1
        If Abs(CDbl(inv(i, j))) > 0# Then allzero = False
    Next j: Next i
    If allzero Then
        inv = MN(MA(Sym(Av), MSc(0.000000001, MI(n))))
    End If
    SymInv = inv
End Function

'--- HP Filter (hpfilter.R / mFilter lambda 型) ---
Private Function HP(ByVal data As Variant, Optional Lambda As Double = 36000#) As Variant
    Dim x() As Double, n As Long, i As Long, j As Long
    If IsArray(data) Then
        n = UBound(data) - LBound(data) + 1
        ReDim x(0 To n - 1)
        For i = 0 To n - 1: x(i) = CDbl(data(LBound(data) + i)): Next i
    Else
        ReDim x(0 To 0): x(0) = CDbl(data): n = 1
    End If
    If n < 4 Then HP = data: Exit Function
    Dim imat As Variant: imat = MI(n)
    Dim Ln As Variant: Ln = MZ(n, n)
    For i = 1 To n - 1: Ln(i, i - 1) = 1#: Next i
    Dim M1 As Variant: M1 = MS(imat, Ln)
    Dim M2 As Variant: M2 = MM(M1, M1)
    Dim qR As Long: qR = n - 2
    If qR < 1 Then HP = imat: Exit Function
    Dim Qm As Variant: Qm = MZ(n, qR)
    For i = 2 To n - 1: For j = 0 To n - 1: Qm(j, i - 2) = CDbl(M2(i, j)): Next j: Next i
    Dim Qt As Variant: Qt = MT(Qm)
    Dim SigR As Variant: SigR = MM(Qt, Qm)
    Dim SigN As Variant: SigN = MI(qR)
    Dim xv As Variant: xv = MZ(n, 1)
    For i = 0 To n - 1: xv(i, 0) = x(i): Next i
    Dim gvec As Variant: gvec = MM(Qt, xv)
    Dim Ac As Variant: Ac = MZ(qR, qR)
    For i = 0 To qR - 1: For j = 0 To qR - 1
        Ac(i, j) = CDbl(SigN(i, j)) + Lambda * CDbl(SigR(i, j)): Next j: Next i
    Dim b As Variant: b = MM(MN(Ac), gvec)
    Dim Qb As Variant: Qb = MM(Qm, b)
    Dim tr As Variant: tr = MZ(n, 1)
    For i = 0 To n - 1: tr(i, 0) = x(i) - Lambda * CDbl(Qb(i, 0)): Next i
    HP = tr
End Function

'--- stage3 fixed coefficients (v5: 速報CSV 1994Q1-2026Q2 で再推計) ---
' 較正の根拠 (2026-10-21):
'   (1) lambda_g / lambda_z は FRBNY 手順どおり Stage1/2 の median-unbiased 推定値
'       (lambda_g=0.192798, lambda_z=0)。日本のデータでは lambda_z は 0 と区別不能。
'   (2) sigma_y* は MLE が下限(0.01)に張り付く(潜在成長の innovation をゼロにする =
'       潜在成長率が決定論的になり r* が 30 年ほぼ一定 0.73% になる) ため
'       識別制約として下限 0.20 を課す(尤度差は 2.4 点で無視可能)。
'       このときモデルの産出ギャップは HP フィルタのギャップ(日本の標準推計)と
'       RMSE 0.147% で整合する(2020Q2 = -8.31%)。
'   (3) a_3 = -0.0025 (上限制約), b_2 = 0.025 (下限制約) は HLW/FRBNY の仕様どおり固定。
Private Sub FixCoeff(p() As Double, xi0() As Double, P0() As Double, Optional ByVal syStar As Double = -1#)
    ' theta (v5)
    p(0) = 0.552558496348379
    p(1) = 0.2161789109138686
    p(2) = -0.0025
    p(3) = 0.24753497876637343
    p(4) = 0.025
    p(5) = 1.6405209095703452
    p(6) = 1.8866123589543204
    p(7) = 0.2
    If syStar >= 0# Then p(7) = syStar
    ' xi00 は標本から動的計算 (HP トレンド) するためここでは未使用
    xi0(0, 0) = 0#
    xi0(1, 0) = 0#
    xi0(2, 0) = 0#
    xi0(3, 0) = 0#
    xi0(4, 0) = 0#
    xi0(5, 0) = 0#
    xi0(6, 0) = 0#
    ' P00 (stage3 MLE 1st pass filtered P_ttm1[0], v5)
    Dim i As Long, j As Long
    For i = 0 To 6: For j = 0 To 6: P0(i, j) = 0#: Next j: Next i
    P0(0, 0) = 0.44148684285216006: P0(0, 1) = 0.2: P0(0, 3) = 0.20148684275216: P0(0, 4) = 0.2
    P0(1, 0) = 0.2: P0(1, 1) = 0.20000000010000002
    P0(2, 2) = 0.20000000010000002
    P0(3, 0) = 0.20148684275216: P0(3, 3) = 0.20148684285216: P0(3, 4) = 0.2
    P0(4, 0) = 0.2: P0(4, 3) = 0.2: P0(4, 4) = 0.20000000010000002
    P0(5, 5) = 0.20000000010000002: P0(5, 6) = 0.2
    P0(6, 5) = 0.2: P0(6, 6) = 0.20000000010000002
End Sub

'--- stage3 unpack (rstar_check.py _unpack_stage3 と同一) ---
Private Sub Unpack3(p() As Double, lg As Double, lz As Double, ByRef F As Variant, ByRef Q As Variant, _
                    ByRef A As Variant, ByRef H As Variant, ByRef Rv As Variant)
    Dim FArr() As Double: ReDim FArr(0 To 6, 0 To 6)
    FArr(0, 0) = 1#: FArr(0, 3) = 1#: FArr(1, 0) = 1#: FArr(2, 1) = 1#
    FArr(3, 3) = 1#: FArr(4, 3) = 1#: FArr(5, 5) = 1#: FArr(6, 5) = 1#
    F = FArr

    Dim QArr() As Double: ReDim QArr(0 To 6, 0 To 6)
    QArr(0, 0) = (1# + lg * lg) * p(7) * p(7)
    QArr(0, 3) = (lg * p(7)) * (lg * p(7))
    QArr(3, 0) = QArr(0, 3)
    QArr(3, 3) = QArr(0, 3)
    If p(2) <> 0 Then QArr(5, 5) = (lz * p(5) / p(2)) * (lz * p(5) / p(2))
    Q = QArr

    ' A: 6x2 (Python では 2x6 を転置して使用)
    Dim AArr() As Double: ReDim AArr(0 To 5, 0 To 1)
    AArr(0, 0) = p(0): AArr(1, 0) = p(1)
    AArr(2, 0) = p(2) / 2#: AArr(3, 0) = p(2) / 2#
    AArr(0, 1) = p(4)
    AArr(4, 1) = p(3): AArr(5, 1) = 1# - p(3)
    A = AArr

    ' H: 7x2
    Dim HArr() As Double: ReDim HArr(0 To 6, 0 To 1)
    HArr(0, 0) = 1#
    HArr(1, 0) = -p(0): HArr(2, 0) = -p(1)
    HArr(3, 0) = -p(2) * 2#: HArr(4, 0) = -p(2) * 2#
    HArr(5, 0) = -p(2) / 2#: HArr(6, 0) = -p(2) / 2#
    HArr(1, 1) = -p(4)
    H = HArr

    Dim RArr() As Double: ReDim RArr(0 To 1, 0 To 1)
    RArr(0, 0) = p(5) * p(5): RArr(1, 1) = p(6) * p(6)
    Rv = RArr
End Sub

'--- Kalman filter + RTS smoother (kalman_states.R 移植) ---
Private Function KRun(xi0 As Variant, P0 As Variant, F As Variant, Q As Variant, _
                      At As Variant, Ht As Variant, Rm As Variant, _
                      y As Variant, x As Variant) As Object
    Dim T As Long: T = UBound(y, 1) + 1
    Dim n As Long: n = UBound(F, 1) + 1
    Dim xi() As Double: ReDim xi(0 To n - 1)
    Dim P() As Double: ReDim P(0 To n - 1, 0 To n - 1)
    Dim i As Long, j As Long, k As Long, o As Long, l As Long, ti As Long
    For i = 0 To n - 1: xi(i) = xi0(i, 0): Next i
    For i = 0 To n - 1: For j = 0 To n - 1: P(i, j) = P0(i, j): Next j: Next i

    Dim xiTT() As Double: ReDim xiTT(0 To T - 1, 0 To n - 1)
    Dim xiTM1() As Variant: ReDim xiTM1(0 To T - 1)
    Dim xiTS() As Double: ReDim xiTS(0 To T - 1, 0 To n - 1)
    Dim PTTM1() As Variant: ReDim PTTM1(0 To T - 1)
    Dim PTT() As Variant: ReDim PTT(0 To T - 1)

    Dim a() As Double, Ppv() As Double, PtSt() As Double
    Dim PHt() As Double, Kg() As Double, PS() As Double
    Dim pred(0 To 1) As Double, S(0 To 1, 0 To 1) As Double
    Dim Sinv As Variant
    For ti = 0 To T - 1
        ' prediction: a = F xi ; Pp = F P F' + Q
        ReDim a(0 To n - 1)
        For i = 0 To n - 1
            a(i) = 0#
            For j = 0 To n - 1: a(i) = a(i) + F(i, j) * xi(j): Next j
        Next i
        ReDim Ppv(0 To n - 1, 0 To n - 1)
        For i = 0 To n - 1
            For j = 0 To n - 1
                Ppv(i, j) = Q(i, j)
                For k = 0 To n - 1
                    For l = 0 To n - 1
                        Ppv(i, j) = Ppv(i, j) + F(i, k) * P(k, l) * F(j, l)
                    Next l
                Next k
            Next j
        Next i
        ' prediction error: pred = y - A'x - H'a
        For o = 0 To 1
            pred(o) = y(ti, o)
            For i = 0 To 5
                pred(o) = pred(o) - At(i, o) * x(ti, i)
            Next i
            For i = 0 To n - 1
                pred(o) = pred(o) - Ht(i, o) * a(i)
            Next i
        Next o
        ' PHt = P H (7x2)
        ReDim PHt(0 To n - 1, 0 To 1)
        For i = 0 To n - 1
            For o = 0 To 1
                PHt(i, o) = 0#
                For j = 0 To n - 1: PHt(i, o) = PHt(i, o) + Ppv(i, j) * Ht(j, o): Next j
            Next o
        Next i
        ' S = H' P H + R (2x2)
        For o = 0 To 1
            For l = 0 To 1
                S(o, l) = Rm(o, l)
                For i = 0 To n - 1
                    S(o, l) = S(o, l) + Ht(i, o) * PHt(i, l)
                Next i
            Next l
        Next o
        ' Kg = (P H) S^-1 ; xi = a + Kg pred ; P = Pp - Kg (P H)'
        Sinv = SymInv(S)
        ReDim Kg(0 To n - 1, 0 To 1)
        For i = 0 To n - 1
            For o = 0 To 1
                Kg(i, o) = PHt(i, 0) * Sinv(0, o) + PHt(i, 1) * Sinv(1, o)
            Next o
        Next i
        For i = 0 To n - 1
            xi(i) = a(i) + Kg(i, 0) * pred(0) + Kg(i, 1) * pred(1)
        Next i
        ReDim PS(0 To n - 1, 0 To n - 1)
        For i = 0 To n - 1
            For j = 0 To n - 1
                PS(i, j) = Ppv(i, j) - (Kg(i, 0) * PHt(j, 0) + Kg(i, 1) * PHt(j, 1))
            Next j
        Next i
        ' symmetrize (Python _psd の対称化)
        For i = 0 To n - 1
            For j = 0 To n - 1
                P(i, j) = 0.5 * (PS(i, j) + PS(j, i))
            Next j
        Next i
        ' store
        Dim av() As Double: ReDim av(0 To n - 1)
        For i = 0 To n - 1: av(i) = a(i): Next i
        xiTM1(ti) = av
        PTTM1(ti) = Ppv
        For i = 0 To n - 1: xiTT(ti, i) = xi(i): Next i
        ReDim PtSt(0 To n - 1, 0 To n - 1)
        For i = 0 To n - 1: For j = 0 To n - 1: PtSt(i, j) = P(i, j): Next j: Next i
        PTT(ti) = PtSt
    Next ti

    ' RTS smoother
    For i = 0 To n - 1: xiTS(T - 1, i) = xiTT(T - 1, i): Next i
    Dim Jm As Variant, jdv As Variant
    Dim xiTnext() As Double, xiTprev() As Double, dlt() As Double
    For ti = T - 2 To 0 Step -1
        Jm = MM(PTT(ti), MM(MT(F), SymInv(PTTM1(ti + 1))))
        ReDim xiTnext(0 To n - 1): ReDim xiTprev(0 To n - 1)
        For i = 0 To n - 1
            xiTnext(i) = xiTS(ti + 1, i)
            xiTprev(i) = xiTM1(ti + 1)(i)
        Next i
        ReDim dlt(0 To n - 1)
        For i = 0 To n - 1: dlt(i) = xiTnext(i) - xiTprev(i): Next i
        jdv = MM(Jm, ColVec(dlt))
        For i = 0 To n - 1
            xiTS(ti, i) = xiTT(ti, i) + jdv(i, 0)
        Next i
    Next ti

    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("xi_tt") = xiTT: d("xi_tT") = xiTS
    Set KRun = d
End Function

Private Function ColVec(v() As Double) As Variant
    Dim n As Long, i As Long
    n = UBound(v) + 1
    Dim M() As Double: ReDim M(0 To n - 1, 0 To 0)
    For i = 0 To n - 1: M(i, 0) = v(i): Next i
    ColVec = M
End Function

'--- helpers ---
' セルの安全な数値読み取り (IsNumeric で分岐してから CDbl: IIf は両評価なので不可)
Private Function NumCell(v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then NumCell = CDbl(v)
End Function
Private Function GetSheet() As Worksheet
    On Error Resume Next: Set GetSheet = ThisWorkbook.Worksheets(SH_RSTAR): On Error GoTo 0
    If GetSheet Is Nothing Then
        Set GetSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetSheet.Name = SH_RSTAR
    End If
End Function
Private Sub Status(ws As Worksheet, msg As String)
    ws.Range("A1").Value = msg: ws.Range("A1").Font.Color = RGB(0, 90, 160)
    Dim f As Long: f = FreeFile: On Error Resume Next
    Open ThisWorkbook.Path & "\run.log" For Append As #f
    Print #f, Format$(Now, "hh:mm:ss") & " RStar: " & msg: Close #f
End Sub

'--- グラフ生成 (4枚: LHW r*, LW r*, LHW g*, LHW output gap) ---
Private Sub BuildCharts(ws As Worksheet, ByVal T As Long)
    Dim co As ChartObject
    On Error Resume Next
    For Each co In ws.ChartObjects
        co.Delete
    Next co
    On Error GoTo 0
    Dim d1 As Long, d2 As Long
    d1 = O_TOP: d2 = O_TOP + T - 1
    ' チャート1-3 は 2000Q1 から表示する (カルマンフィルタの初期過渡を避けるため)。
    ' 2000Q1 = 1996Q1 の 16 四半期後
    Dim d1c As Long: d1c = O_TOP + 16
    If d1c > d2 Then d1c = d1
    Dim xs As Range: Set xs = ws.Range(ws.Cells(d1, 1), ws.Cells(d2, 1))
    Dim xsc As Range: Set xsc = ws.Range(ws.Cells(d1c, 1), ws.Cells(d2, 1))
    
    ' チャート配置: 入力テーブル下 (row 140頃) から 2x2 グリッド
    ' 幅 640, 高さ 300, 余白 20
    Const CH_W As Double = 640
    Const CH_H As Double = 300
    Const CH_GAP As Double = 20
    Dim leftCol As Double, topRow As Double
    leftCol = ws.Range("A1").Left + 10
    topRow = ws.Range("A140").Top  ' 入力テーブル (133行目終了) の少し下
    
    Dim i As Long
    
    '=== Chart 1: LHW r* (Kalman filter & smoother + 観測実質金利) ===
    '   桁が大きく異なるため、観測系列=主軸(左) / モデル系列=二次軸(右) に分離
    Set co = ws.ChartObjects.Add(leftCol, topRow, CH_W, CH_H)
    co.Name = "RStarChart_LHW_rstar"
    Call StyleChart(co.Chart, "LHW r* (Kalman フィルタ/スムーザ) と実質観測金利 (2000Q1～)", "実質金利 (%)")
    Call AddSeriesStyled(co.Chart, "r* Kalman (フィルタ)", ws.Range(ws.Cells(d1c, 2), ws.Cells(d2, 2)), xsc, _
                         RGB(31, 78, 121), 2.0, True, False)        ' B列: filter (青破線)
    Call AddSeriesStyled(co.Chart, "r* Kalman (スムーザ)", ws.Range(ws.Cells(d1c, 3), ws.Cells(d2, 3)), xsc, _
                         RGB(31, 78, 121), 2.75, False, False)      ' C列: smoother (青実線)
    Call AddSeriesStyled(co.Chart, "実質観測金利 (i - π)", ws.Range(ws.Cells(d1c, 8), ws.Cells(d2, 8)), xsc, _
                         RGB(150, 150, 150), 1.0, False, True)     ' H列: 観測 (灰・マーカーのみ・主軸)
    Call SetSeriesAxis(co.Chart, "r* Kalman (フィルタ)", 2)
    Call SetSeriesAxis(co.Chart, "r* Kalman (スムーザ)", 2)
    Call StyleSecondaryAxis(co.Chart, "r* (%)", "0.00")

    '=== Chart 2: LW r* (r*_LW + 観測実質金利) ===
    Set co = ws.ChartObjects.Add(leftCol + CH_W + CH_GAP, topRow, CH_W, CH_H)
    co.Name = "RStarChart_LW_rstar"
    Call StyleChart(co.Chart, "LW r* (加重潜在成長率+μ) と実質観測金利 (2000Q1～)", "実質金利 (%)")
    Call AddSeriesStyled(co.Chart, "r*_LW", ws.Range(ws.Cells(d1c, 9), ws.Cells(d2, 9)), xsc, _
                         RGB(237, 125, 49), 2.75, False, False)      ' I列: LW (橙実線)
    Call AddSeriesStyled(co.Chart, "実質観測金利 (i - π)", ws.Range(ws.Cells(d1c, 8), ws.Cells(d2, 8)), xsc, _
                         RGB(150, 150, 150), 1.0, False, True)     ' H列: 観測 (主軸)
    Call SetSeriesAxis(co.Chart, "r*_LW", 2)
    Call StyleSecondaryAxis(co.Chart, "r* (%)", "0.00")

    '=== Chart 3: LHW 潜在成長率 g* (フィルタ/スムーザ) + 実質GDP成長率 ===
    Set co = ws.ChartObjects.Add(leftCol, topRow + CH_H + CH_GAP, CH_W, CH_H)
    co.Name = "RStarChart_LHW_gstar"
    Call StyleChart(co.Chart, "LHW 潜在成長率 g* (フィルタ/スムーザ) と実質GDP成長率 (2000Q1～)", "実質GDP成長率 (%(年率))")
    Call AddSeriesStyled(co.Chart, "g* (フィルタ)", ws.Range(ws.Cells(d1c, 4), ws.Cells(d2, 4)), xsc, _
                         RGB(31, 78, 121), 2.0, True, False)         ' D列: filter (青破線)
    Call AddSeriesStyled(co.Chart, "g* (スムーザ)", ws.Range(ws.Cells(d1c, 5), ws.Cells(d2, 5)), xsc, _
                         RGB(31, 78, 121), 2.75, False, False)       ' E列: smoother (青実線)
    Call AddSeriesStyled(co.Chart, "実質GDP成長率 (YoY)", ws.Range(ws.Cells(d1c, 10), ws.Cells(d2, 10)), xsc, _
                         RGB(112, 173, 71), 1.25, False, False)      ' J列: 実質GDP成長率 (緑・主軸)
    Call SetSeriesAxis(co.Chart, "g* (フィルタ)", 2)
    Call SetSeriesAxis(co.Chart, "g* (スムーザ)", 2)
    Call StyleSecondaryAxis(co.Chart, "g* (%(年率))", "0.00")

    '=== Chart 4: 産出ギャップ (フィルタ/スムーザ) + HPフィルタベンチマーク ===
    Set co = ws.ChartObjects.Add(leftCol + CH_W + CH_GAP, topRow + CH_H + CH_GAP, CH_W, CH_H)
    co.Name = "RStarChart_LHW_gap"
    Call StyleChart(co.Chart, "LHW 産出ギャップ (フィルタ/スムーザ) と HPフィルタ・日銀需給ギャップ", "%", "0.0")
    Call AddSeriesStyled(co.Chart, "産出ギャップ (フィルタ)", ws.Range(ws.Cells(d1, 6), ws.Cells(d2, 6)), xs, _
                         RGB(192, 80, 77), 2.0, True, False)       ' F列: filter (赤破線)
    Call AddSeriesStyled(co.Chart, "産出ギャップ (スムーザ)", ws.Range(ws.Cells(d1, 7), ws.Cells(d2, 7)), xs, _
                         RGB(192, 80, 77), 2.75, False, False)     ' G列: smoother (赤実線)
    Call AddSeriesStyled(co.Chart, "HPフィルタ (参考)", ws.Range(ws.Cells(d1, 12), ws.Cells(d2, 12)), xs, _
                         RGB(112, 173, 71), 1.25, False, False)    ' L列: HP ギャップ (緑)
    Call AddSeriesStyled(co.Chart, "日銀需給ギャップ (参考)", ws.Range(ws.Cells(d1, 13), ws.Cells(d2, 13)), xs, _
                         RGB(112, 48, 160), 1.25, False, True)      ' M列: 日銀 需給ギャップ (紫・マーカー)
End Sub

' グラフ全体の体裁
Private Sub StyleChart(ch As Chart, title As String, yTitle As String, Optional yFmt As String = "0.0")
    With ch
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = title
        .ChartTitle.Font.Name = "メイリオ"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(31, 78, 121)
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Name = "メイリオ"
        .Legend.Font.Size = 9
        .ChartArea.Font.Name = "メイリオ"
        .ChartArea.Font.Size = 9
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Line.Visible = msoFalse
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = yTitle
            .AxisTitle.Font.Name = "メイリオ"
            .AxisTitle.Font.Size = 9
            .AxisTitle.Font.Bold = True
            .TickLabels.Font.Name = "メイリオ"
            .TickLabels.Font.Size = 9
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
            .MajorGridlines.Format.Line.Weight = 0.5
            .Format.Line.ForeColor.RGB = RGB(191, 191, 191)
            .Format.Line.Weight = 0.75
            .TickLabels.NumberFormat = yFmt
        End With
        With .Axes(xlCategory)
            .TickLabels.Font.Name = "メイリオ"
            .TickLabels.Font.Size = 8
            .TickLabelSpacing = 8
            .TickLabels.Orientation = -45
            .HasMajorGridlines = False
            .Format.Line.ForeColor.RGB = RGB(191, 191, 191)
            .Format.Line.Weight = 0.75
        End With
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(252, 252, 252)
        .PlotArea.Format.Line.Visible = msoFalse
    End With
End Sub

' 系列の追加と体裁
Private Sub AddSeriesStyled(ch As Chart, nm As String, vals As Range, xs As Range, _
                            colr As Long, wt As Double, dash As Boolean, noMarker As Boolean)
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nm
    s.Values = vals
    s.XValues = xs
    s.ChartType = xlLine
    With s.Format.Line
        .ForeColor.RGB = colr
        .Weight = wt
        If dash Then .DashStyle = msoLineDash Else .DashStyle = msoLineSolid
    End With
    If noMarker Then
        s.MarkerStyle = xlMarkerStyleNone
    Else
        s.MarkerStyle = xlMarkerStyleCircle
        s.MarkerSize = 4
        s.MarkerBackgroundColor = RGB(255, 255, 255)
        s.MarkerForegroundColor = colr
    End If
    If dash Then
        s.Smooth = False
    Else
        s.Smooth = True
    End If
End Sub

' 系列を二次軸(右)へ移す。系列追加と同時に設定すると Excel が系列を並べ替えるため、
' 全系列追加後に名前で指定して移す。
Private Sub SetSeriesAxis(ch As Chart, ByVal seriesName As String, ByVal axisGroup As Long)
    On Error Resume Next
    ch.SeriesCollection(seriesName).AxisGroup = axisGroup
    On Error GoTo 0
End Sub

' 二次軸 (右) の体裁。カテゴリ軸は非表示にして主軸と揃える。
Private Sub StyleSecondaryAxis(ch As Chart, yTitle As String, Optional yFmt As String = "0.00")
    On Error Resume Next
    With ch.Axes(xlValue, xlSecondary)
        .HasTitle = True
        .AxisTitle.Text = yTitle
        .AxisTitle.Font.Name = "メイリオ"
        .AxisTitle.Font.Size = 9
        .AxisTitle.Font.Bold = True
        .TickLabels.Font.Name = "メイリオ"
        .TickLabels.Font.Size = 9
        .TickLabels.NumberFormat = yFmt
        .HasMajorGridlines = False
        .Format.Line.ForeColor.RGB = RGB(191, 191, 191)
        .Format.Line.Weight = 0.75
    End With
    ch.Axes(xlCategory, xlSecondary).TickLabelPosition = xlNone
    On Error GoTo 0
End Sub

'=== PUBLIC API ===
Public Sub RStar_SetupSheet()
    ' R* シートを作成し、書式 (色・罫線・表示形式・枠固定) を設定する
    Dim ws As Worksheet: Set ws = GetSheet()
    On Error GoTo E
    Application.ScreenUpdating = False
    ws.Cells.Clear
    ws.Cells.Interior.ColorIndex = xlNone

    '--- タイトル帯 ---
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 9))
        .Merge
        .Value = "R*(自然利子率)推計シート  －  HLW Kalman と LW モデル"
        .Font.Name = "メイリオ": .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 84, 166)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(1).RowHeight = 28

    '--- パラメータ節 ---
    ws.Cells(2, 1).Value = "パラメータ": ws.Cells(2, 2).Value = "値"
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, 2))
        .Font.Name = "メイリオ": .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
    End With
    Dim i As Long
    ' (使用) LW モデルのパラメータ
    ws.Cells(3, 1).Value = "lw_lambda": ws.Cells(3, 2).Value = 1600#
    ws.Cells(4, 1).Value = "lw_weights": ws.Cells(4, 2).Value = "0.5/0.3/0.2/0.1"
    ws.Cells(5, 1).Value = "mu_start": ws.Cells(5, 2).Value = "1995Q1"
    ws.Cells(6, 1).Value = "mu_end": ws.Cells(6, 2).Value = "2007Q4"
    ' (情報) 推計の前提
    ws.Cells(7, 1).Value = "theta_source (情報)"
    ws.Cells(7, 2).Value = "固定係数 v5 (速報CSV 1994Q1-2026Q2 で再推計 / sig_y*=0.20 下限制約) / xi0はHPトレンドから動的計算 / 産出ギャップ=実際のy100-モデル潜在y100"
    ws.Cells(8, 1).Value = "zlb_floor": ws.Cells(8, 2).Value = 0#
    ws.Cells(9, 1).Value = "data_source (情報)"
    ws.Cells(9, 2).Value = "内閣府 四半期別GDP速報CSV (実質GDP・民間消費デフレーター) + 日銀コール翌日物 (版番号自動検出)"
    ws.Cells(10, 1).Value = "rstar_auto": ws.Cells(10, 2).Value = "ON"

    ' HLW パラメータ (D列ラベル / E列値) -- 「R*推計更新」で反映
    ws.Cells(2, 4).Value = "HLWパラメータ"
    ws.Cells(2, 4).Font.Bold = True: ws.Cells(2, 4).Font.Color = RGB(31, 78, 121)
    ws.Cells(3, 4).Value = "sigma_y* 下限": ws.Cells(3, 5).Value = 0.2
    ws.Cells(4, 4).Value = "lambda_g": ws.Cells(4, 5).Value = 0.192798
    ws.Cells(5, 4).Value = "lambda_z": ws.Cells(5, 5).Value = 0#
    ws.Cells(6, 4).Value = "a_3 制約": ws.Cells(6, 5).Value = -0.0025
    ws.Cells(7, 4).Value = "b_2 制約": ws.Cells(7, 5).Value = 0.025
    ws.Cells(8, 4).Value = "sigma_y* は MLE の識別失敗 (下限0.01) を補う下限値。0.20 で HPフィルタの産出ギャップと RMSE 0.15% で整合"
    ws.Cells(8, 4).Font.Size = 8: ws.Cells(8, 4).Font.Color = RGB(120, 120, 120)
    With ws.Range(ws.Cells(3, 4), ws.Cells(7, 5))
        .Font.Name = "メイリオ": .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
        .Borders.Weight = xlThin
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(3, 5), ws.Cells(7, 5)).Interior.Color = RGB(255, 247, 214)
    ws.Range(ws.Cells(3, 5), ws.Cells(7, 5)).Font.Bold = True
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E").ColumnWidth = 12

    With ws.Range(ws.Cells(2, 1), ws.Cells(10, 2))
        .Font.Name = "メイリオ"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
        .Borders.Weight = xlThin
        .VerticalAlignment = xlCenter
    End With
    ' 入力できるセルは淡い黄色、参照専用 (情報) はグレー
    Dim rUsed As Variant: rUsed = Array(3, 4, 5, 6, 8, 10)
    For i = 0 To UBound(rUsed)
        ws.Cells(rUsed(i), 2).Interior.Color = RGB(255, 247, 214)
        ws.Cells(rUsed(i), 2).Font.Bold = True
    Next i
    Dim rInfo As Variant: rInfo = Array(7, 9)
    For i = 0 To UBound(rInfo)
        ws.Cells(rInfo(i), 1).Font.Color = RGB(120, 120, 120)
        ws.Cells(rInfo(i), 2).Interior.Color = RGB(242, 242, 242)
        ws.Cells(rInfo(i), 2).Font.Color = RGB(120, 120, 120)
    Next i
    ws.Cells(3, 2).NumberFormat = "#,##0"
    ws.Cells(8, 2).NumberFormat = "0.0"

    '--- 入力テーブル節 ---
    ws.Cells(11, 1).Value = "■ 入力データ(内閣府GDP・消費デフレーター・日銀コールから自動取得)"
    With ws.Range(ws.Cells(11, 1), ws.Cells(11, 9))
        .Merge
        .Font.Name = "メイリオ": .Font.Size = 10: .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlLeft
    End With
    Dim hdr As Variant: hdr = Array("q_idx", "四半期", "log_output", "inflation", "exp_infl", "real_rate", "nom_rate")
    For i = 0 To 6
        ws.Cells(I_HDR, i + 1).Value = hdr(i)
    Next i
    With ws.Range(ws.Cells(I_HDR, 1), ws.Cells(I_HDR, 7))
        .Font.Name = "メイリオ": .Font.Size = 10: .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
        .Interior.Color = RGB(217, 232, 247)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(146, 176, 209)
        .Borders.Weight = xlThin
    End With
    ' 入力データ範囲の表示形式 (数値)
    ws.Range(ws.Cells(I_TOP, 3), ws.Cells(I_BOTTOM, 3)).NumberFormat = "0.0000"
    ws.Range(ws.Cells(I_TOP, 4), ws.Cells(I_BOTTOM, 7)).NumberFormat = "0.000"
    ws.Range(ws.Cells(I_TOP, 2), ws.Cells(I_BOTTOM, 2)).HorizontalAlignment = xlCenter

    '--- 結果テーブル節 ---
    ws.Cells(O_HDR - 1, 1).Value = "■ 推計結果(四半期ごと)"
    With ws.Range(ws.Cells(O_HDR - 1, 1), ws.Cells(O_HDR - 1, 9))
        .Merge
        .Font.Name = "メイリオ": .Font.Size = 10: .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlLeft
    End With
    ws.Range(ws.Cells(O_TOP, 2), ws.Cells(O_TOP + 400, 10)).NumberFormat = "0.00"
    ws.Range(ws.Cells(O_TOP, 10), ws.Cells(O_TOP + 400, 10)).NumberFormat = "0.0"
    ws.Range(ws.Cells(O_TOP, 11), ws.Cells(O_TOP + 400, 11)).NumberFormat = "0.000000"
    ws.Range(ws.Cells(O_TOP, 12), ws.Cells(O_TOP + 400, 12)).NumberFormat = "0.00"
    ws.Range(ws.Cells(O_TOP, 1), ws.Cells(O_TOP + 400, 1)).HorizontalAlignment = xlCenter

    '--- 列幅・枠固定 ---
    ws.Columns("A").ColumnWidth = 9
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 13
    ws.Columns("D:G").ColumnWidth = 11
    ws.Columns("H:J").ColumnWidth = 11
    ws.Columns("K").ColumnWidth = 12
    ws.Columns("L").ColumnWidth = 14
    ws.Columns("M").ColumnWidth = 14
    ws.Columns("N").ColumnWidth = 2
    ws.Columns("A:I").VerticalAlignment = xlCenter
    ws.Activate
    ws.Range("A13").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

    Application.ScreenUpdating = True
    Status ws, "Sheet created: " & SH_RSTAR
    Exit Sub
E:
    Application.ScreenUpdating = True
    Status ws, "Error: " & Err.Number & ": " & Err.Description
End Sub

Public Sub RStar_Run()
    Dim ws As Worksheet: Set ws = GetSheet()
    On Error GoTo E
    ' lambda_g / lambda_z は Stage1/2 の median-unbiased 推定値 (Rstar シート E列の
    ' パラメータで上書き可)。sigma_y* は MLE の識別失敗を補う下限 (既定 0.20)。
    ' ZLB フロア (B8) は FetchData 側で nom_rate に適用済み
    ' 注意: IsNumeric(Empty) は True (0 に変換される) なので空判定を明示する。
    Dim lamg As Double, lamz As Double, syStar As Double
    lamg = ParamCell(ws, 4, 5, LAMBDA_G)
    lamz = ParamCell(ws, 5, 5, LAMBDA_Z)
    syStar = ParamCell(ws, 3, 5, 0.2)
    Dim n As Long, r As Long: n = 0
    For r = I_TOP To I_BOTTOM
        If Trim$(CStr(ws.Cells(r, 1).Value)) = "" Then Exit For
        n = n + 1
    Next r
    If n < 10 Then
        Call RStar_FetchData   ' 入力テーブルが空なら VBA 内蔵 fetch
        n = 0
        For r = I_TOP To I_BOTTOM
            If Trim$(CStr(ws.Cells(r, 1).Value)) = "" Then Exit For
            n = n + 1
        Next r
    End If
    If n < 10 Then
        Status ws, "Error: need min 10 data points (got " & n & ")"
        Exit Sub
    End If
    ' 0-based 入力配列 (Python と同一インデックス)
    ' 入力列: A=q_idx, B=label(YYYYQn), C=log_output, D=inflation, E=exp_infl, F=real_rate, G=nom_rate
    Dim lab() As String: ReDim lab(0 To n - 1)
    Dim lo() As Double: ReDim lo(0 To n - 1)
    Dim inf() As Double: ReDim inf(0 To n - 1)
    Dim rr() As Double: ReDim rr(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        lab(i) = Trim$(CStr(ws.Cells(I_TOP + i, 2).Value))
        lo(i) = NumCell(ws.Cells(I_TOP + i, 3).Value)
        inf(i) = NumCell(ws.Cells(I_TOP + i, 4).Value)
        rr(i) = NumCell(ws.Cells(I_TOP + i, 6).Value)
    Next i
    Dim p(0 To 7) As Double, xi0(0 To 6, 0 To 0) As Double, P0(0 To 6, 0 To 6) As Double
    Call FixCoeff(p, xi0, P0, syStar)
    ' xi0 はサンプルから動的に初期化 (Python xi00=None 既定と同一ロジック)。
    '   xi00 = [100*trend(3), 100*trend(2), 100*trend(1),
    '           100*(trend(3)-trend(2)), 100*(trend(2)-trend(1)), 0, 0]
    '   (FixCoeff の凍結値は標本延長で SA 平均がずれるため使用しない)
    Dim trV As Variant: trV = HP(lo, 36000#)
    Dim trA() As Double: trA = trV
    xi0(0, 0) = 100# * trA(3, 0)
    xi0(1, 0) = 100# * trA(2, 0)
    xi0(2, 0) = 100# * trA(1, 0)
    xi0(3, 0) = 100# * (trA(3, 0) - trA(2, 0))
    xi0(4, 0) = 100# * (trA(2, 0) - trA(1, 0))
    Dim F As Variant, Q As Variant, A As Variant, H As Variant, Rv As Variant
    Call Unpack3(p, lamg, lamz, F, Q, A, H, Rv)
    Dim T As Long: T = n - 4
    Dim y() As Double: ReDim y(0 To T - 1, 0 To 1)
    Dim x() As Double: ReDim x(0 To T - 1, 0 To 5)
    Dim ti As Long
    For ti = 0 To T - 1
        y(ti, 0) = 100# * lo(ti + 4): y(ti, 1) = inf(ti + 4)
        x(ti, 0) = 100# * lo(ti + 3): x(ti, 1) = 100# * lo(ti + 2)
        x(ti, 2) = rr(ti + 3): x(ti, 3) = rr(ti + 2)
        x(ti, 4) = inf(ti + 3): x(ti, 5) = (inf(ti + 2) + inf(ti + 1) + inf(ti)) / 3#
    Next ti
    Dim st As Object: Set st = KRun(xi0, P0, F, Q, A, H, Rv, y, x)
    Dim xiT() As Double: xiT = st("xi_tt")
    Dim xiS() As Double: xiS = st("xi_tT")
    ReDim m_rf(0 To T - 1): ReDim m_rs(0 To T - 1): ReDim m_lab(0 To T - 1)
    m_T = T
    For ti = 0 To T - 1
        m_rf(ti) = xiT(ti, 3) * 4# + xiT(ti, 5)
        m_rs(ti) = xiS(ti, 3) * 4# + xiS(ti, 5)
        m_lab(ti) = lab(ti + 4)
    Next ti
    Dim sr As Long: sr = O_TOP
    ws.Range(ws.Cells(O_HDR, 1), ws.Cells(O_HDR + 300, 13)).ClearContents
    Dim hRes As Variant: hRes = Array("四半期", "r* Kalman (フィルタ)", "r* Kalman (スムーザ)", "g* (フィルタ)", "g* (スムーザ)", "産出ギャップ (フィルタ)", "産出ギャップ (スムーザ)", "実質金利 (観測)", "r*_LW", "実質GDP (前年同期比)", "z状態 (r*分解・参考)", "産出ギャップ (HPフィルタ・参考)", "産出ギャップ (日銀需給・参考)")
    Dim hc As Long
    For hc = 0 To 12
        ws.Cells(O_HDR, 1 + hc).Value = hRes(hc)
        ws.Cells(O_HDR, 1 + hc).Font.Bold = True
    Next hc
    With ws.Range(ws.Cells(O_HDR, 1), ws.Cells(O_HDR, 13))
        .Font.Name = "メイリオ": .Font.Bold = True: .Font.Size = 10
        .Font.Color = RGB(31, 78, 121)
        .Interior.Color = RGB(217, 232, 247)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(146, 176, 209)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    For ti = 0 To T - 1
        ws.Cells(sr + ti, 1).Value = CStr(m_lab(ti))
        ws.Cells(sr + ti, 2).Value = m_rf(ti)
        ws.Cells(sr + ti, 3).Value = m_rs(ti)
        ws.Cells(sr + ti, 4).Value = xiT(ti, 3) * 4#
        ws.Cells(sr + ti, 5).Value = xiS(ti, 3) * 4#
        ' 産出ギャップ (古典的定義) = 実際のy100 - モデル潜在y100 (状態0)
        '   ※ FBNY の output.gap.smoothed と同じ量。HPフィルタのギャップと整合する。
        ws.Cells(sr + ti, 6).Value = y(ti, 0) - xiT(ti, 0)
        ws.Cells(sr + ti, 7).Value = y(ti, 0) - xiS(ti, 0)
        ws.Cells(sr + ti, 8).Value = rr(ti + 4)
        ' 前年同期比 (4 四半期前との比較, %)
        ws.Cells(sr + ti, 10).Value = 100# * (lo(ti + 4) - lo(ti))
        ' z 状態 (r* = 4*g* + z の分解項。lam_z=0 のためほぼ 0)
        ws.Cells(sr + ti, 11).Value = xiT(ti, 5)
        ' HP フィルタの産出ギャップ (独立推計との比較用ベンチマーク)
        ws.Cells(sr + ti, 12).Value = 100# * (lo(ti + 4) - trA(ti + 4, 0))
    Next ti
    ' 日銀「需給ギャップ」(生産関数ベース・別口径の独立推計)
    Dim gapBoj As Object
    Set gapBoj = FetchBojGap()
    If Not gapBoj Is Nothing Then
        For ti = 0 To T - 1
            If gapBoj.Exists(CStr(m_lab(ti))) Then
                ws.Cells(sr + ti, 13).Value = gapBoj(CStr(m_lab(ti)))
            End If
        Next ti
    End If
    ' --- LW モデル: r*_LW = Sum w_j g*_(t+j) + mu (加重成長率 + mu マッチング) ---
    Dim lw As Variant
    lw = RunLW(ws, lo, rr, n)
    For ti = 0 To T - 1
        ws.Cells(sr + ti, 9).Value = lw(ti)
    Next ti
    ' 列幅は SetupSheet で設定済み (AutoFit するとヘッダ長で列が異常に広がるため使わない)
    ' 数値書式 (結果表)
    ws.Range(ws.Cells(O_TOP, 2), ws.Cells(O_TOP + 400, 9)).NumberFormat = "0.00"
    ws.Range(ws.Cells(O_TOP, 10), ws.Cells(O_TOP + 400, 10)).NumberFormat = "0.0"
    ws.Range(ws.Cells(O_TOP, 11), ws.Cells(O_TOP + 400, 11)).NumberFormat = "0.000000"
    ws.Range(ws.Cells(O_TOP, 12), ws.Cells(O_TOP + 400, 13)).NumberFormat = "0.00"
    Call BuildCharts(ws, T)
    Dim msgSt As String
    msgSt = "Done: T=" & T & " (HLW + LW)"
    If m_bojErr <> "" Then msgSt = msgSt & " | " & m_bojErr
    Status ws, msgSt
    ' 参考: 固定係数は σ_y* に下限 0.20 を課した MLE 解 (日本のデータでは
    ' sigma_y*・lambda_z が識別できないため)。産出ギャップ (F/G列) は
    ' 実際のy100 - モデル潜在y100 (L列の HP フィルタ推計と RMSE 0.15% で整合)。
    Exit Sub
E:
    Status ws, "Error " & Err.Number & ": " & Err.Description
End Sub

'--- LW モデル (Laubach-Williams 簡易式) ---
' g*_t = 400*(HP trend 差分),  r*_LW,t = Sum_{j=0..3} w_j * g*_(t+j) + mu
' mu = 期間内の平均実質金利 - 期間内の平均(加重成長率)   [計画書の mu マッチング]
' サンプル末尾 3 四半期は利用可能な項のみで重みを再正規化
Private Function RunLW(ws As Worksheet, lo() As Double, rr() As Double, ByVal n As Long) As Variant
    Dim w(0 To 3) As Double, i As Long
    Dim parts() As String
    parts = Split(Replace(CStr(ws.Cells(4, 2).Value), " ", ""), "/")
    For i = 0 To 3
        If i <= UBound(parts) Then w(i) = CDbl(parts(i)) Else w(i) = 0#
    Next i
    Dim lwLam As Double
    lwLam = 1600#
    If IsNumeric(ws.Cells(3, 2).Value) Then lwLam = CDbl(ws.Cells(3, 2).Value)
    ' HP trend (0-based n x 1)
    Dim tr As Variant
    tr = HP(lo, lwLam)
    ' g*(k) = 400 * (tau_k - tau_{k-1}), k = 1..n-1
    Dim g() As Double: ReDim g(0 To n - 1)
    For i = 1 To n - 1
        g(i) = 400# * (CDbl(tr(i, 0)) - CDbl(tr(i - 1, 0)))
    Next i
    Dim T As Long: T = n - 4
    ' 生の加重成長率 (mu マッチング用) と結果
    Dim raw() As Double: ReDim raw(0 To T - 1)
    Dim res() As Double: ReDim res(0 To T - 1)
    Dim ti As Long, j As Long, s As Double, ws2 As Double
    For ti = 0 To T - 1
        Dim ii As Long: ii = ti + 4
        s = 0#: ws2 = 0#
        For j = 0 To 3
            If ii + j <= n - 1 Then
                s = s + w(j) * g(ii + j)
                ws2 = ws2 + w(j)
            End If
        Next j
        If ws2 > 0 Then res(ti) = s / ws2 Else res(ti) = 0#
        raw(ti) = res(ti)
    Next ti
    ' mu マッチング: 期間内 (mu_start..mu_end, q_idx 基準 1994Q1=0) の
    ' 平均実質金利 - 平均加重成長率
    Dim mS As Long, mEnd As Long
    mS = QidxOf(CStr(ws.Cells(5, 2).Value))
    mEnd = QidxOf(CStr(ws.Cells(6, 2).Value))
    Dim muVal As Double, nmu As Long
    muVal = 0#: nmu = 0
    For ti = 0 To T - 1
        ii = ti + 4
        If ii >= mS And ii <= mEnd Then
            muVal = muVal + (rr(ii) - raw(ti))
            nmu = nmu + 1
        End If
    Next ti
    If nmu > 0 Then muVal = muVal / nmu
    For ti = 0 To T - 1
        res(ti) = res(ti) + muVal
    Next ti
    RunLW = res
End Function

' パラメータセルの数値取得 (空・非数値は既定値)
' 注意: IsNumeric(Empty) は True を返すため IsEmpty と空文字を明示的に除外する。
Private Function ParamCell(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal dfl As Double) As Double
    Dim v As Variant
    v = ws.Cells(r, c).Value
    ParamCell = dfl
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    ParamCell = CDbl(v)
End Function

' q_idx -> "1995Q1" 表示用ラベル (1994Q1 = 0)
Private Function QLabelOf(ByVal qidx As Long) As String
    Dim yy As Long, qq As Long
    yy = 1994 + qidx \ 4
    qq = (qidx Mod 4) + 1
    QLabelOf = CStr(yy) & "Q" & CStr(qq)
End Function

' "1995Q1" 形式 -> q_idx (1994Q1 = 0)
Private Function QidxOf(ByVal lab As String) As Long
    lab = AscDigits(Trim$(lab))
    Dim p As Long
    p = InStr(lab, "Q")
    If p < 2 Then QidxOf = -1: Exit Function
    QidxOf = (CLng(Val(Left$(lab, p - 1))) - 1994) * 4 + (CLng(Val(Mid$(lab, p + 1))) - 1)
End Function

Public Sub RStar_SelfTest()
    Dim ws As Worksheet: Set ws = GetSheet()
    Dim r As Long: r = 3
    ws.Cells(2, 10).Value = "R* self-test": ws.Cells(2, 10).Font.Bold = True
    Dim x(0 To 9) As Double, i As Long
    For i = 0 To 9: x(i) = 10# + 0.1 * i + Sin(i): Next i
    Dim tr As Variant: tr = HP(x, 1600#)
    Dim trArr() As Double: trArr = tr
    Dim errv As Double
    For i = 0 To 9: errv = errv + Abs(trArr(i, 0) + (x(i) - trArr(i, 0)) - x(i)): Next i
    ws.Cells(r, 10).Value = "HP trend+cycle=x": ws.Cells(r, 11).Value = IIf(errv < 0.000001, "PASS", "FAIL")
    r = r + 1
    ' Kalman 1-step: 全状態が有限レンジに収まること
    Dim p(0 To 7) As Double, xi0(0 To 6, 0 To 0) As Double, P0(0 To 6, 0 To 6) As Double
    Call FixCoeff(p, xi0, P0, 0.2)
    ' xi0 は RStar_Run と同じく HP トレンドから動的初期化
    Dim tr36 As Variant: tr36 = HP(x, 36000#)
    Dim tr36a() As Double: tr36a = tr36
    xi0(0, 0) = 100# * tr36a(3, 0)
    xi0(1, 0) = 100# * tr36a(2, 0)
    xi0(2, 0) = 100# * tr36a(1, 0)
    xi0(3, 0) = 100# * (tr36a(3, 0) - tr36a(2, 0))
    xi0(4, 0) = 100# * (tr36a(2, 0) - tr36a(1, 0))
    Dim F As Variant, Q As Variant, A As Variant, H As Variant, Rv As Variant
    Call Unpack3(p, 0#, 0#, F, Q, A, H, Rv)
    Dim yy(0 To 0, 0 To 1) As Double, xx(0 To 0, 0 To 5) As Double
    yy(0, 0) = 1000#: yy(0, 1) = 1#
    Dim j As Long
    For j = 0 To 5: xx(0, j) = 1000#: Next j
    Dim st As Object: Set st = KRun(xi0, P0, F, Q, A, H, Rv, yy, xx)
    Dim xiT1() As Double: xiT1 = st("xi_tt")
    Dim finite As Boolean: finite = True
    For j = 0 To 6
        If Abs(xiT1(0, j)) > 1000000# Then finite = False
    Next j
    ws.Cells(r, 10).Value = "Kalman 1-step finite": ws.Cells(r, 11).Value = IIf(finite, "PASS", "FAIL")
    r = r + 1
    ' レイアウト不変条件: 入力テーブル最終行 < 結果ヘッダ行
    Dim rIn As Long: rIn = I_TOP
    Do While Trim$(CStr(ws.Cells(rIn, 1).Value)) <> "" And rIn < 10000
        rIn = rIn + 1
    Loop
    ws.Cells(r, 10).Value = "layout: input end " & (rIn - 1) & " < results " & O_HDR
    ws.Cells(r, 11).Value = IIf(rIn <= O_HDR, "PASS", "FAIL")
    r = r + 1
    ' 産出ギャップの整合チェック: モデル gap (G列) と HP フィルタ gap (L列) の RMSE
    Dim rs2 As Long, rmse As Double, nG As Long, d2 As Double
    rs2 = O_TOP
    Do While Trim$(CStr(ws.Cells(rs2, 1).Value)) <> "" And rs2 < O_HDR + 300
        If IsNumeric(ws.Cells(rs2, 7).Value) And IsNumeric(ws.Cells(rs2, 12).Value) Then
            d2 = CDbl(ws.Cells(rs2, 7).Value) - CDbl(ws.Cells(rs2, 12).Value)
            rmse = rmse + d2 * d2: nG = nG + 1
        End If
        rs2 = rs2 + 1
    Loop
    If nG > 0 Then
        rmse = VBA.Sqr(rmse / nG)
        ws.Cells(r, 10).Value = "gap vs HP RMSE: " & Format$(rmse, "0.000") & "% (n=" & nG & ")"
        ws.Cells(r, 11).Value = IIf(rmse < 0.5, "PASS", "WARN")
    Else
        ws.Cells(r, 10).Value = "gap vs HP RMSE": ws.Cells(r, 11).Value = "no data"
    End If
    Status ws, "Self-test complete"
End Sub

Public Function RStar_GetValue(name As String, idx As Long) As Variant
    On Error GoTo E
    If name = "rstar_filtered" Then
        If idx >= 0 And idx < m_T Then RStar_GetValue = m_rf(idx): Exit Function
    ElseIf name = "rstar_smoothed" Then
        If idx >= 0 And idx < m_T Then RStar_GetValue = m_rs(idx): Exit Function
    ElseIf name = "label" Then
        If idx >= 0 And idx < m_T Then RStar_GetValue = m_lab(idx): Exit Function
    Else: Err.Raise vbObjectError + 702, , "Unknown: " & name
    End If
    Exit Function
E:
    RStar_GetValue = CVErr(xlErrNA)
End Function

Public Function RStar_GetCount() As Long
    On Error GoTo E
    If m_T > 0 Then RStar_GetCount = m_T
    Exit Function
E:
    RStar_GetCount = 0
End Function

' 開発用ヘルパ: Python rstar_inputs.csv をそのまま流し込む (配布環境では不要)
Public Sub RStar_FillInputsFromCsv(path As String)
    Dim ws As Worksheet: Set ws = GetSheet()
    On Error GoTo E
    Dim f As Long: f = FreeFile
    Open path For Input As #f
    Dim lineTxt As String, parts() As String, r As Long, c As Long
    r = 12
    Do Until EOF(f)
        Line Input #f, lineTxt
        lineTxt = Replace(lineTxt, vbCr, "")
        parts = Split(lineTxt, ",")
        For c = 0 To UBound(parts)
            ws.Cells(r, c + 1).Value = Trim$(parts(c))
        Next c
        r = r + 1
    Loop
    Close #f
    Status ws, "CSV imported: " & path
    Exit Sub
E:
    Status ws, "CSV error: " & Err.Description
End Sub


'================================================================
' VBA 内蔵データ取得 (配布先 Python 不要版)
'   GDP: 内閣府 四半期別GDP xlsx (実質/デフレーター, キー不要)
'   Call: 日銀 時系列検索 Code API STRACLUCON (キー不要)
'   rstar_fetch.py / build_hlw_inputs と同一ロジック
'================================================================


Private Function AscDigits(ByVal s As String) As String
    Dim i As Long, c As Long, out As String
    For i = 1 To Len(s)
        c = CLng(AscW(Mid$(s, i, 1)))
        If c < 0 Then c = c + 65536
        If c >= &HFF10 And c <= &HFF19 Then
            out = out & ChrW(c - &HFF10 + 48)
        Else
            out = out & Mid$(s, i, 1)
        End If
    Next i
    AscDigits = out
End Function

Private Sub DownloadBin(url As String, dest As String)
    Dim http As Object, st As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 10000, 10000, 30000, 120000
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.Send
    If http.Status <> 200 Then Err.Raise vbObjectError + 514, , "download fail HTTP " & http.Status & " " & url
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write http.responseBody
    st.SaveToFile dest, 2
    st.Close
End Sub

Private Function DownloadText(url As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 10000, 10000, 30000, 120000
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.setRequestHeader "Accept-Encoding", "identity"
    http.Send
    If http.Status <> 200 Then Err.Raise vbObjectError + 515, , "download fail HTTP " & http.Status & " " & url
    Dim b() As Byte
    b = http.responseBody
    If UBound(b) >= 1 Then
        If b(0) = &H1F And b(1) = &H8B Then
            ' gzip の場合は WinINet 系 XMLHTTP に切り替え (自動展開)
            Dim http2 As Object
            Set http2 = CreateObject("MSXML2.XMLHTTP.6.0")
            http2.Open "GET", url, False
            http2.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
            http2.Send
            If http2.Status <> 200 Then Err.Raise vbObjectError + 516, , "download fail HTTP " & http2.Status & " " & url
            b = http2.responseBody
        End If
    End If
    ' BOJ CSV は CP932
    DownloadText = StrConv(b, vbUnicode)
End Function

' 内閣府 四半期別GDP速報の最新版を検出する
'   files_sokuhou.html -> 最新年 (2026) -> toukei_YYYY.html -> 最新 qeNNN(_K) ディレクトリ
'   出力: dirRel = "2026/qe262_2" / verCode = "2622"  (CSVファイル名の4桁番号)
Private Sub DetectSokuhou(ByRef dirRel As String, ByRef verCode As String)
    Dim re As Object, m As Object, html As String
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    ' 最新年ディレクトリ (一覧の最初の年が最新)
    html = DownloadText(SOKUHO_LIST_URL)
    re.Pattern = "/sna/data/data_list/sokuhou/files/(\d{4})/toukei_\d{4}\.html"
    Set m = re.Execute(html)
    If m.Count = 0 Then Err.Raise vbObjectError + 521, , "sokuhou year not found"
    Dim yr As String: yr = m(0).SubMatches(0)
    ' 年内の最新四半期 (ページ上から順に新しい: 2次 -> 1次)
    html = DownloadText(SOKUHO_BASE & yr & "/toukei_" & yr & ".html")
    re.Pattern = "qe(\d{3})_(\d)/gdemenuja\.html|qe(\d{3})/gdemenuja\.html"
    Set m = re.Execute(html)
    If m.Count = 0 Then Err.Raise vbObjectError + 522, , "sokuhou qe not found: " & yr
    Dim n1 As Variant, n2 As Variant, n3 As Variant
    n1 = m(0).SubMatches(0): n2 = m(0).SubMatches(1): n3 = m(0).SubMatches(2)
    If Not IsEmpty(n2) And CStr(n2) <> "" Then
        verCode = CStr(n1) & CStr(n2)
        dirRel = yr & "/qe" & CStr(n1) & "_" & CStr(n2)
    Else
        verCode = CStr(n3) & "1"
        dirRel = yr & "/qe" & CStr(n3)
    End If
End Sub

' 内閣府速報CSV 1行パース (クォート内カンマ対応)
Private Function SplitCsvLine(ln As String) As Collection
    Dim out As New Collection
    Dim i As Long, n As Long, inQ As Boolean, cur As String
    n = Len(ln): i = 1
    Do While i <= n
        Dim c As String: c = Mid$(ln, i, 1)
        If inQ Then
            If c = """" Then
                If i < n And Mid$(ln, i + 1, 1) = """" Then
                    cur = cur & """": i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & c
            End If
        Else
            If c = """" Then
                inQ = True
            ElseIf c = "," Then
                out.Add cur: cur = ""
            Else
                cur = cur & c
            End If
        End If
        i = i + 1
    Loop
    out.Add cur
    Set SplitCsvLine = out
End Function

' 内閣府速報CSV -> Dictionary q_idx(0-based, 1994Q1起点) -> Double
' colIdx は1-based列位置。行ラベル "1994/ 1- 3." / "4- 6." / "10-12." を解釈、*** は欠損扱い
Private Function ParseEsriCsv(txt As String, ByVal colIdx As Long, ByRef lastQ As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\s*(?:(\d{4})/)?\s*(\d+)\s*-\s*\d+\s*\."
    Dim lines() As String
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    Dim curYear As Long: curYear = 0
    Dim i As Long, f As Collection
    lastQ = -1
    For i = 0 To UBound(lines)
        Dim ln As String: ln = lines(i)
        If Len(ln) = 0 Then GoTo NextI
        If Left$(ln, 1) = "*" Then GoTo NextI       ' 注記行
        If Not re.Test(ln) Then GoTo NextI
        Dim m As Object: Set m = re.Execute(ln)(0)
        If Not IsEmpty(m.SubMatches(0)) And CStr(m.SubMatches(0)) <> "" Then
            curYear = CLng(m.SubMatches(0))
        End If
        If curYear = 0 Then GoTo NextI
        Dim mm As Long: mm = CLng(m.SubMatches(1))
        If mm < 1 Or mm > 12 Then GoTo NextI
        Dim qidx As Long
        qidx = (curYear - 1994) * 4 + ((mm - 1) \ 3)
        If qidx < 0 Then GoTo NextI
        Set f = SplitCsvLine(ln)
        If f.Count < colIdx Then GoTo NextI
        Dim v As String: v = Trim$(f(colIdx))
        v = Replace(v, ",", "")
        If v = "" Or v = "***" Then GoTo NextI
        If Not IsNumeric(v) Then GoTo NextI
        d(qidx) = CDbl(v)
        If qidx > lastQ Then lastQ = qidx
NextI:
    Next i
    Set ParseEsriCsv = d
End Function

' 日銀「需給ギャップ・潜在成長率」gap.xlsx -> Dictionary "YYYYQn" -> 需給ギャップ(%)
' data1 シートの 6 行目以降、A列が "1983.1Q" 形式・B列が需給ギャップ(%)
Private Function ParseBojGap(path As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim wb As Workbook, ws As Worksheet
    On Error GoTo E
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, False, True)
    Application.DisplayAlerts = True
    On Error Resume Next
    Set ws = wb.Worksheets("data1")
    On Error GoTo 0
    If ws Is Nothing Then
        wb.Close False
        Set ParseBojGap = d
        Exit Function
    End If
    Dim r As Long, lastR As Long, lab As String, v As Variant
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 6 To lastR
        lab = Trim$(CStr(ws.Cells(r, 1).Value))
        If InStr(lab, "Q") > 0 And InStr(lab, ".") > 0 Then
            v = ws.Cells(r, 2).Value
            If IsNumeric(v) Then d(BojQLab(lab)) = CDbl(v)
        End If
    Next r
    wb.Close False
    Set ParseBojGap = d
    Exit Function
E:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Set ParseBojGap = d
End Function

' "1983.1Q" -> "1983Q1"
Private Function BojQLab(ByVal s As String) As String
    Dim p As Long, y As String, q As String
    p = InStr(s, ".")
    If p = 0 Then
        BojQLab = s
        Exit Function
    End If
    y = Left$(s, p - 1)
    q = Replace(Mid$(s, p + 1), "Q", "")
    BojQLab = y & "Q" & q
End Function

' 日銀需給ギャップの取得 (TEMP に7日キャッシュ・失敗しても致命的ではない)
Private Function FetchBojGap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo E
    Dim tmp As String
    tmp = Environ$("TEMP") & "\rstar_boj_gap.xlsx"
    Dim stale As Boolean: stale = True
    ' 注意: FileExists は VBA 標準関数ではない (Dir$ を使う)
    If Len(Dir$(tmp)) > 0 Then
        If DateDiff("d", FileDateTime(tmp), Now) < 7 Then stale = False
    End If
    If stale Then DownloadBin BOJ_GAP_URL, tmp
    Set d = ParseBojGap(tmp)
    If d.Count = 0 Then
        If m_bojErr = "" Then m_bojErr = "BOJ gap: parse returned 0 rows"
    End If
    Set FetchBojGap = d
    Exit Function
E:
    m_bojErr = "BOJ gap: " & Err.Number & " " & Err.Description
    On Error Resume Next
    Set FetchBojGap = d
End Function

Private Function StripQ(ByVal s As String) As String
    StripQ = Replace(Replace(Trim$(s), """", ""), vbCr, "")
End Function

Private Function IsNumCell(v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    If IsNumeric(v) Then IsNumCell = True
End Function

' GDP xlsx (実数シート) -> Dictionary q_idx(0-based,1994Q1起点) -> 値, firstYear
Private Function ParseGdpXlsx(path As String, rowKeyword As String, ByRef firstYear As Long) As Object
    Dim wb As Workbook, ws As Worksheet
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, False, True)
    Application.DisplayAlerts = True
    On Error Resume Next
    Set ws = wb.Worksheets("実数")
    On Error GoTo 0
    If ws Is Nothing Then Set ws = wb.Worksheets(1)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    ' 年ラベル行: 1..12行目で A列空 & B列が文字列
    Dim yrRow As Long, r As Long
    yrRow = 0
    For r = 1 To 12
        If Trim$(CStr(ws.Cells(r, 1).Value)) = "" And VarType(ws.Cells(r, 2).Value) = vbString Then
            yrRow = r
            Exit For
        End If
    Next r
    ' 年ラベル -> firstYear (平成=1988, 令和=2018)
    firstYear = 0
    If yrRow > 0 Then
        Dim c As Long, lab As String
        c = 2
        Do While Trim$(CStr(ws.Cells(yrRow, c).Value)) <> ""
            lab = AscDigits(CStr(ws.Cells(yrRow, c).Value))
            Dim base As Long, pos As Long, ye As Long, yv As Long
            base = 0
            If InStr(lab, "平成") > 0 Then base = 1988: pos = InStr(lab, "平成") + 2
            If InStr(lab, "令和") > 0 Then base = 2018: pos = InStr(lab, "令和") + 2
            If base > 0 Then
                ye = InStr(pos, lab, "年")
                If ye > pos Then
                    yv = CLng(Val(Mid$(lab, pos, ye - pos)))
                    If firstYear = 0 Then firstYear = base + yv
                End If
            End If
            c = c + 1
        Loop
    End If
    If firstYear = 0 Then firstYear = 1994
    ' 対象行: A列に rowKeyword を含み数値セル数が最大の行。
    ' 表題行 (「…連鎖方式」を含む) は数値を持たないため自動的に除外される。
    Dim bestR As Long, bestN As Long, lastR As Long, cnt As Long, cc As Long, v As Variant, lb As String
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    bestR = 0: bestN = 0
    For r = 1 To lastR
        lb = AscDigits(Trim$(CStr(ws.Cells(r, 1).Value)))
        If InStr(lb, rowKeyword) > 0 Then
            If InStr(lb, "連鎖方式") = 0 Then
                cnt = 0
                For cc = 2 To 320
                    If IsNumCell(ws.Cells(r, cc).Value) Then cnt = cnt + 1
                Next cc
                If cnt > bestN Then
                    bestN = cnt: bestR = r
                End If
            End If
        End If
    Next r
    If bestR = 0 Then Err.Raise vbObjectError + 517, , "GDP row not found in " & path
    ' 数値セルを列順に q=0,1,2.. として格納 (数値以外はスキップ)
    Dim q As Long
    q = 0
    For cc = 2 To 320
        v = ws.Cells(bestR, cc).Value
        If IsNumCell(v) Then
            d(q) = CDbl(v)
            q = q + 1
        End If
    Next cc
    wb.Close False
    Set ParseGdpXlsx = d
End Function

' BOJ CSV -> Dictionary "YYYYMM" -> 月次コールレート
Private Function ParseBojCall(txt As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim lines() As String, i As Long, j As Long
    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)
    Dim di As Long, vi As Long
    di = -1: vi = -1
    For i = 0 To UBound(lines)
        If InStr(lines(i), "SERIES_CODE") = 1 Then
            Dim hdr() As String
            hdr = Split(lines(i), ",")
            For j = 0 To UBound(hdr)
                If StripQ(hdr(j)) = "SURVEY_DATES" Then di = j
                If StripQ(hdr(j)) = "VALUES" Then vi = j
            Next j
            Exit For
        End If
    Next i
    If di < 0 Or vi < 0 Then Err.Raise vbObjectError + 518, , "BOJ CSV header not found"
    Dim parts() As String, dv As String, vv As String
    For i = i + 1 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            parts = Split(lines(i), ",")
            If UBound(parts) >= vi Then
                dv = StripQ(parts(di))
                vv = StripQ(parts(vi))
                If Len(dv) = 6 And IsNumeric(dv) And Len(vv) > 0 Then
                    If IsNumeric(vv) Then d(dv) = CDbl(vv)
                End If
            End If
        End If
    Next i
    Set ParseBojCall = d
End Function

' メイン: 内閣府 xlsx 2種 + BOJ CSV から Rstar シート入力テーブルを構築
Public Sub RStar_FetchData()
    Dim ws As Worksheet: Set ws = GetSheet()
    On Error GoTo E
    Dim zf As Double
    zf = 0#   ' 既定 0.0 = Python ZLB_FLOOR_DEFAULT
    If IsNumeric(ws.Cells(8, 2).Value) Then zf = CDbl(ws.Cells(8, 2).Value)
    Dim tmpDir As String
    tmpDir = Environ$("TEMP") & "\"
    ' 1. 内閣府 四半期別GDP速報 CSV (最新版を自動検出, 確報xlsxより新四半期まで更新される)
    '    実質GDP: gaku-jgXXXX.csv 原系列 (2020暦年連鎖価格・10億円, 列2=国内総生産(支出側))
    '    デフレーター: def-qgXXXX.csv 原系列 (2020暦年=100, 列3=民間最終消費支出)
    '    FRBNY がコアPCE (消費ベースの平滑な指数) を使うのと同じ発想で、
    '    インフレ源は GDPデフレーターでなく 民間最終消費支出デフレーター
    Dim dirRel As String, verCode As String
    Call DetectSokuhou(dirRel, verCode)
    Dim gReal As Object, gDefl As Object
    Dim maxR As Long, maxD As Long
    Set gReal = ParseEsriCsv(DownloadText(SOKUHO_BASE & dirRel & "/tables/gaku-jg" & verCode & ".csv"), 2, maxR)
    Set gDefl = ParseEsriCsv(DownloadText(SOKUHO_BASE & dirRel & "/tables/def-qg" & verCode & ".csv"), 3, maxD)
    If maxR < 40 Or maxD < 40 Then Err.Raise vbObjectError + 518, , "sokuhou CSV too short: real=" & maxR & " defl=" & maxD
    Status ws, "GDP速報 " & dirRel & " (実質1994Q1~" & QLabelOf(maxR) & ", デフ~" & QLabelOf(maxD) & ")"
    ' 2. BOJ
    Dim months As Object
    Set months = ParseBojCall(DownloadText(BOJ_CALL_URL))
    If months.Count < 12 Then Err.Raise vbObjectError + 519, , "BOJ months too few: " & months.Count
    ' 3. 四半期パネル
    Dim nq As Long, k As Variant
    maxR = 0: maxD = 0
    For Each k In gReal.Keys
        If CLng(k) > maxR Then maxR = CLng(k)
    Next k
    For Each k In gDefl.Keys
        If CLng(k) > maxD Then maxD = CLng(k)
    Next k
    nq = maxR
    If maxD < nq Then nq = maxD
    nq = nq + 1
    Dim qnum() As Long: ReDim qnum(0 To nq - 1)
    Dim realv() As Double: ReDim realv(0 To nq - 1)
    Dim deflv() As Double: ReDim deflv(0 To nq - 1)
    Dim iq() As Variant: ReDim iq(0 To nq - 1)
    Dim y As Long, qq As Long, q As Long
    Dim m1 As String, m2 As String, m3 As String
    Dim iv As Double, nc As Long
    y = 1994: qq = 1
    For q = 0 To nq - 1
        qnum(q) = qq
        If gReal.Exists(q) Then
            realv(q) = gReal(q)
        Else
            If q > 0 Then realv(q) = realv(q - 1) Else realv(q) = 0#
        End If
        If gDefl.Exists(q) Then
            deflv(q) = gDefl(q)
        Else
            If q > 0 Then deflv(q) = deflv(q - 1) Else deflv(q) = 100#
        End If
        ' コール: 四半期内の月次 -> 2ヶ月以上は平均, 1ヶ月はその値, 0は欠損
        m1 = y & Right$("0" & ((qq - 1) * 3 + 1), 2)
        m2 = y & Right$("0" & ((qq - 1) * 3 + 2), 2)
        m3 = y & Right$("0" & ((qq - 1) * 3 + 3), 2)
        iv = 0#: nc = 0
        If months.Exists(m1) Then iv = iv + months(m1): nc = nc + 1
        If months.Exists(m2) Then iv = iv + months(m2): nc = nc + 1
        If months.Exists(m3) Then iv = iv + months(m3): nc = nc + 1
        If nc >= 2 Then
            iq(q) = iv / nc
        ElseIf nc = 1 Then
            iq(q) = iv
        Else
            iq(q) = Empty
        End If
        qq = qq + 1
        If qq > 4 Then
            qq = 1
            y = y + 1
        End If
    Next q
    ' 4. 系列 (build_hlw_inputs と同一)
    '   注意: 内閣府 xlsx の実質GDPは原系列なので log_output にも季節調整をかける
    Dim logOut() As Double: ReDim logOut(0 To nq - 1)
    Dim sa() As Double: ReDim sa(0 To nq - 1)
    For q = 0 To nq - 1
        logOut(q) = VBA.Log(realv(q))
        sa(q) = VBA.Log(deflv(q))
    Next q

    ' 季節調整: 各四半期(qnum)の平均を引く
    '   sa        = log デフレーター (民間最終消費支出)
    '   logOut    = log 実質GDP (内閣府 xlsx は原系列のため要調整)
    Dim qs As Long, s As Double, nq4 As Long, sl As Double, nl As Long
    For qs = 1 To 4
        s = 0#: nq4 = 0: sl = 0#: nl = 0
        For q = 0 To nq - 1
            If qnum(q) = qs Then
                s = s + sa(q)
                nq4 = nq4 + 1
                sl = sl + logOut(q)
                nl = nl + 1
            End If
        Next q
        If nq4 > 0 Then
            s = s / nq4
            If nl > 0 Then sl = sl / nl
            For q = 0 To nq - 1
                If qnum(q) = qs Then
                    sa(q) = sa(q) - s
                    logOut(q) = logOut(q) - sl
                End If
            Next q
        End If
    Next qs
    Dim infl() As Variant: ReDim infl(0 To nq - 1)
    For q = 1 To nq - 1
        infl(q) = 400# * (sa(q) - sa(q - 1))
    Next q
    Dim nomR() As Variant: ReDim nomR(0 To nq - 1)
    For q = 0 To nq - 1
        If IsEmpty(iq(q)) Then
            nomR(q) = Empty
        Else
            If CDbl(iq(q)) < zf Then
                nomR(q) = zf
            Else
                nomR(q) = CDbl(iq(q))
            End If
        End If
    Next q
    Dim expI() As Variant: ReDim expI(0 To nq - 1)
    For q = 3 To nq - 1
        If Not IsEmpty(infl(q)) And Not IsEmpty(infl(q - 1)) And Not IsEmpty(infl(q - 2)) And Not IsEmpty(infl(q - 3)) Then
            expI(q) = (CDbl(infl(q)) + CDbl(infl(q - 1)) + CDbl(infl(q - 2)) + CDbl(infl(q - 3))) / 4#
        End If
    Next q
    ' 5. i0 = expI が最初に定義される四半期
    Dim i0 As Long
    i0 = -1
    For q = 0 To nq - 1
        If Not IsEmpty(expI(q)) Then
            i0 = q
            Exit For
        End If
    Next q
    If i0 < 0 Then Err.Raise vbObjectError + 520, , "no valid exp_infl"
    ' 6. シートへ書き込み
    Dim row As Long
    ' クリアは入力データ行から (ヘッダ行 I_HDR は残す)
    ws.Range(ws.Cells(I_TOP, 1), ws.Cells(O_HDR + 200, 12)).ClearContents
    ' ヘッダを書く (SetupSheet を経由せず Fetch 単独で呼ばれても成立するように)
    Dim hdrI As Variant: hdrI = Array("q_idx", "label", "log_output", "inflation", "exp_infl", "real_rate", "nom_rate")
    Dim hI As Long
    For hI = 0 To 6
        ws.Cells(I_HDR, hI + 1).Value = hdrI(hI)
        ws.Cells(I_HDR, hI + 1).Font.Bold = True
    Next hI
    row = I_TOP
    For q = i0 To nq - 1
        ws.Cells(row, 1).Value = q
        ws.Cells(row, 2).Value = QLabelOf(q)
        ws.Cells(row, 3).Value = logOut(q)
        If Not IsEmpty(infl(q)) Then ws.Cells(row, 4).Value = CDbl(infl(q))
        If Not IsEmpty(expI(q)) Then ws.Cells(row, 5).Value = CDbl(expI(q))
        If Not IsEmpty(expI(q)) And Not IsEmpty(nomR(q)) Then
            ws.Cells(row, 6).Value = CDbl(nomR(q)) - CDbl(expI(q))
            ws.Cells(row, 7).Value = CDbl(nomR(q))
        End If
        row = row + 1
    Next q
    Status ws, "Fetch done: n=" & (nq - i0) & " quarters (実質GDP~" & QLabelOf(nq - 1) & ", BOJ " & months.Count & "months)"
    Exit Sub
E:
    Status ws, "Fetch error " & Err.Number & ": " & Err.Description
End Sub
