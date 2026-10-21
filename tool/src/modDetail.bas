Option Explicit

'==========================================================
' 進捗（詳細）・スキルマップ・事由比率
'   業務マスタの「事由」列で業務IDを 届出×事由×工程 にまとめ、
'   稼働記録から 人ごとの件/時（スキルマップ）と 直近4週の比率（事由比率）を作る。
'   進捗（詳細）は 事由×工程ごとに 流入・処理・残数 を日別に計算する
'   （人の行に時間を入れると 処理見込み＝時間×件/時）。
'   手で入れた値（時間・上限・件/時・比率・対応表）は、作り直しても残す。
'==========================================================

Private Const SH_DET As String = "進捗（詳細）"
Private Const SH_SKL As String = "スキルマップ"
Private Const SH_RAT As String = "事由比率"
Private Const SH_INP As String = "計算用_割当"
Private Const SH_MST As String = "業務マスタ"
Private Const SH_WRK As String = "稼働記録_検索結果_出力データ"
Private Const SH_SUM As String = "【給与手当】集計"
Private Const SH_SNAP As String = "現在の状況（積み上げ）"
Private Const SH_DAILY As String = "日別処理件数（積み上げ）"
Private Const SH_PLAN As String = "見込み申請（昨年実績）"
Private Const SH_SIMPLE As String = "進捗（簡易版）"
Private Const SH_MENU As String = "操作"

Private Const MST_HDR As Long = 4
Private Const MST_COL_REASON As Long = 11       ' 業務マスタ K列：事由
Private Const KIND_CASE As String = "案件処理"

Private Const DAYS_BACK As Long = 14            ' 詳細に出す過去の日数
Private Const DAYS_AHEAD As Long = 60           ' 詳細に出す先の日数
Private Const RATIO_DAYS As Long = 28           ' 比率に使う直近の日数（4週）

Private Const D_HDR As Long = 4                 ' 詳細：日付の行
Private Const D_REST As Long = 5                ' 詳細：営業日フラグの行（非表示）
Private Const D_TOP As Long = 7                 ' 詳細：届出別合計の最初の行
Private Const D_TOP_N As Long = 13              ' 詳細：届出ごとの行数
Private Const D_FIRST As Long = 6               ' 詳細：最初の日付の列（F。簡易版と同じ）
' 届出別合計の行の並び（DetailTopRow の k）。実績（青）→ 予定（黄）
Private Const K_C1 As Long = 0                  ' 実績：確認中（1次）
Private Const K_C2 As Long = 1                  ' 実績：確認中（2次）
Private Const K_A1 As Long = 2                  ' 実績：認定件数（1次）
Private Const K_A2 As Long = 3                  ' 実績：認定件数（2次）
Private Const K_PLAN As Long = 4                ' 予定：見込み申請件数
Private Const K_P1 As Long = 5                  ' 予定：見込み認定件数（1次）
Private Const K_P2 As Long = 6                  ' 予定：見込み認定件数（2次）
Private Const K_R1 As Long = 7                  ' 予定：残数（1次）
Private Const K_R2 As Long = 8                  ' 予定：残数（2次）
Private Const K_RT As Long = 9                  ' 予定：残数（1次＋2次）。E列に上限（超えたら赤）
Private Const K_H1 As Long = 10                 ' 予定：時間（1次）
Private Const K_H2 As Long = 11                 ' 予定：時間（2次）
Private Const K_U As Long = 12                  ' 見込み認定件数（1次）のうち2次の対なし（非表示）

' 進捗（簡易版）と同じ色・線
Private Const CLR_ACT As Long = 16247773        ' 実績：RGB(221,235,247)
Private Const CLR_PLAN As Long = 13431551       ' 予定：RGB(255,242,204)
Private Const CLR_LINE As Long = 5263440        ' 太線・破線：RGB(80,80,80)
Private Const CLR_GRID As Long = 12566463       ' 細線：RGB(191,191,191)

Private Const S_PEOPLE As Long = 7              ' スキルマップ：人の一覧の最初の行
Private Const S_FIRST As Long = 6               ' スキルマップ：件/時の最初の人の列（F）
Private Const R_FIRST As Long = 7               ' 事由比率：表の最初の行

Private Const CLR_PURPLE As Long = 10498160     ' RGB(112,48,160)
Private Const CLR_INPUT As Long = 13434879      ' RGB(255,255,204)

' 業務マスタ → 事由×工程（ブロック）
Private mBlk As Object          ' 届出|事由|工程 → 番号
Private mIdMap As Object           ' 業務ID → 届出|事由|工程
Private mReasonSort As Object   ' 届出|事由 → 並び順（最小の業務ID）
Private nBlk As Long
Private bTodoke() As String, bReason() As String, bKotei() As String
Private bRecent() As Double, bAuto() As Double, bOv() As Variant, bAdopt() As Double, bAlloc() As Variant, bShow() As Boolean

' 人
Private mP As Object            ' EID → 番号
Private nP As Long
Private pEid() As String, pName() As Variant, pDisp() As Variant, pLimit() As Variant, pNote() As Variant, pEx() As Boolean

' 稼働記録の集計
Private skCnt As Object, skMin As Object        ' キー|EID → 件数, 分（期間全体）
Private pCase As Object                         ' EID → 件数（案件処理、期間全体）
Private actCnt As Object                        ' キー|日付 → 件数（詳細の過去の列）
Private actHrs As Object                        ' キー|EID|日付 → 時間（詳細の過去の列）
Private skillVal As Object                      ' キー|EID → 件/時（採用値、0より大きいもの）

' 詳細の手入力（計算用_割当）
Private mInp As Object          ' 種別|EID|キー → Collection of Array(日付, 値)
Private mFutureH As Object      ' キー → True（今日以降に時間の入力がある）
Private mFutureHP As Object     ' キー|EID → True
Private mDetNew As Boolean      ' 進捗（詳細）を今回新しく作った

'----------------------------------------------------------
' 簡易版から参照する、届出別合計の行（届出番号 0～4、合計 5）
'   k: K_C1/K_C2 確認中 K_A1/K_A2 認定件数 K_PLAN 見込み申請件数 K_P1/K_P2 見込み認定件数 K_R1/K_R2 残数 K_H1/K_H2 時間
'----------------------------------------------------------
Public Function DetailTopRow(ByVal TodokeIdx As Long, ByVal k As Long) As Long
    DetailTopRow = D_TOP + TodokeIdx * D_TOP_N + k
End Function

' 業務名「C1：【増】出生」→ 事由「【増】出生」
Public Function ReasonFromName(ByVal nm As String) As String
    Dim p As Long
    p = InStr(nm, "：")
    If p > 0 Then ReasonFromName = Trim$(Mid$(nm, p + 1)) Else ReasonFromName = Trim$(nm)
End Function

'----------------------------------------------------------
' 操作シートのボタンから呼ぶ
'----------------------------------------------------------
Public Sub 詳細を更新()
    MsgBox RefreshDetail(), vbInformation, "進捗（詳細）"
End Sub

Public Function RefreshDetail() As String
    Dim calcMode As XlCalculation, act As Object
    Set act = ActiveSheet
    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fin
    RefreshDetail = BuildDetail()
Fin:
    On Error Resume Next
    act.Activate
    On Error GoTo 0
    Application.EnableEvents = True
    Application.Calculation = calcMode
    If calcMode = xlCalculationAutomatic Then Application.Calculate
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Function

Private Function BuildDetail() As String
    Dim today As Date, d0 As Date, d1 As Date, nShow As Long, i As Long, msg As String
    today = Date: d0 = today - DAYS_BACK: d1 = today + DAYS_AHEAD
    If FindSheet(SH_MST) Is Nothing Then BuildDetail = "業務マスタがありません。": Exit Function
    If FindSheet(SH_WRK) Is Nothing Then BuildDetail = "稼働記録がまだ取り込まれていません。": Exit Function

    ' シートは 簡易版 → 詳細 → スキルマップ → 事由比率 の順に並べる
    mDetNew = (FindSheet(SH_DET) Is Nothing)
    GetSheet SH_DET: GetSheet SH_SKL: GetSheet SH_RAT

    SaveInputs today
    LoadInputs today
    LoadBlocks
    LoadPeople
    LoadWork today, d0
    BuildSkill
    msg = BuildRatio(today)
    LinkSimple
    BuildDet today, d0, d1

    For i = 0 To nBlk - 1
        If bShow(i) Then nShow = nShow + 1
    Next
    BuildDetail = "進捗（詳細）・スキルマップ・事由比率を更新しました（表示する事由×工程 " & nShow & "／" & nBlk & "、" & nP & "人）。" & msg
End Function

'==========================================================
' 1. 手入力の退避と読み込み（計算用_割当）
'==========================================================
Private Sub SaveInputs(ByVal today As Date)
    Dim ws As Worksheet, inp As Worksheet, f As Variant, hdr As Variant
    Dim lastR As Long, lastC As Long, r As Long, c As Long, i As Long, n As Long, t As Long
    Dim code As String, kind As String, eid As String, key As String, rk As String
    Dim seen As Object, adds As Collection, oldEnd As Long, v As Variant, a As Variant
    Dim old As Variant, outArr() As Variant, dv As Long, pa As Variant

    Set seen = NewDict()
    Set adds = New Collection
    t = CLng(today)
    Set ws = FindSheet(SH_DET)
    If Not ws Is Nothing Then
        lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1     ' 折りたたんだ行も含める
        lastC = ws.Cells(D_HDR, ws.Columns.Count).End(xlToLeft).Column
        If lastC >= 6 And lastR > D_HDR Then
            f = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Formula
            hdr = ws.Range(ws.Cells(D_HDR, 1), ws.Cells(D_HDR, lastC)).Value
            If IsDate(hdr(1, lastC)) Then oldEnd = CLng(CDate(hdr(1, lastC)))
            For r = D_HDR + 1 To lastR
                code = CStr(f(r, 1)): kind = ""
                If code Like "H[12]|*" Then
                    kind = "H": eid = Mid$(code, InStrRev(code, "|") + 1)
                    pa = Split(code, "|")
                    If UBound(pa) >= 4 Then
                        key = BlockKey(CStr(pa(1)), CStr(pa(2)), CStr(pa(3)))
                    Else
                        key = BlockKey(CStr(f(r, 2)), CStr(f(r, 3)), CStr(f(r, 4)))   ' 以前の並び（届出・事由・工程の列）
                    End If
                ElseIf Left$(code, 2) = "U|" Then
                    kind = "U": eid = Mid$(code, 3): key = ""
                ElseIf Left$(code, 3) = "YR|" Then
                    ' 残数の上限：E列に数字を直接入れたものだけ残す（式＝簡易版の上限とつながっている）
                    seen("Z||" & code) = True
                    v = f(r, 5)
                    If Len(CStr(v)) > 0 And Left$(CStr(v), 1) <> "=" And IsNumeric(v) Then adds.Add Array("Z", "", code, t, CDbl(v))
                End If
                If Len(kind) > 0 Then
                    seen(kind & "|" & eid & "|" & key) = True
                    For c = 6 To lastC
                        If IsDate(hdr(1, c)) Then
                            dv = CLng(CDate(hdr(1, c)))
                            If dv >= t Then
                                v = f(r, c)
                                If Len(CStr(v)) > 0 And Left$(CStr(v), 1) <> "=" And IsNumeric(v) Then
                                    adds.Add Array(kind, eid, key, dv, CDbl(v))
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If
    End If

    ' 今回シートから読んだ行・期間の分は入れ替え、それ以外（過去・期間外・表示されていない行）は残す
    Set inp = GetSheet(SH_INP)
    lastR = inp.Cells(inp.Rows.Count, 1).End(xlUp).Row
    ReDim outArr(1 To Application.Max(1, lastR + adds.Count), 1 To 5)
    If lastR >= 2 Then
        old = inp.Range(inp.Cells(2, 1), inp.Cells(lastR, 5)).Value
        For i = 1 To UBound(old, 1)
            rk = CStr(old(i, 1)) & "|" & CStr(old(i, 2)) & "|" & CStr(old(i, 3))
            dv = CLng(old(i, 4))
            If Not (seen.Exists(rk) And ((dv >= t And dv <= oldEnd) Or CStr(old(i, 1)) = "Z")) Then
                n = n + 1
                For c = 1 To 5: outArr(n, c) = old(i, c): Next
            End If
        Next
    End If
    For Each a In adds
        n = n + 1
        For c = 0 To 4: outArr(n, c + 1) = a(c): Next
    Next
    inp.Cells.ClearContents
    inp.Range("A1:E1").Value = Array("種別", "EID", "事由キー", "日付", "値")
    If n > 0 Then
        inp.Range("A2").Resize(n, 5).Value = outArr
        inp.Range("D2").Resize(n, 1).NumberFormat = "yyyy/mm/dd"
    End If
End Sub

Private Sub LoadInputs(ByVal today As Date)
    Dim inp As Worksheet, lastR As Long, v As Variant, i As Long, rk As String, p As Variant
    Set mInp = NewDict(): Set mFutureH = NewDict(): Set mFutureHP = NewDict()
    Set inp = GetSheet(SH_INP)
    lastR = inp.Cells(inp.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub
    v = inp.Range(inp.Cells(2, 1), inp.Cells(lastR, 5)).Value
    For i = 1 To UBound(v, 1)
        If IsNumeric(v(i, 5)) And Not IsEmpty(v(i, 5)) Then
            rk = CStr(v(i, 1)) & "|" & CStr(v(i, 2)) & "|" & CStr(v(i, 3))
            If Not mInp.Exists(rk) Then mInp.Add rk, New Collection
            mInp(rk).Add Array(CLng(v(i, 4)), CDbl(v(i, 5)))
            If CStr(v(i, 1)) = "H" And CLng(v(i, 4)) >= CLng(today) Then
                mFutureH(CStr(v(i, 3))) = True
                mFutureHP(CStr(v(i, 3)) & "|" & CStr(v(i, 2))) = True
            End If
        End If
    Next
End Sub

' その日に入力した値（なければ Empty）
Private Function InputAt(ByVal rk As String, ByVal d As Long) As Variant
    Dim a As Variant
    InputAt = Empty
    If Not mInp.Exists(rk) Then Exit Function
    For Each a In mInp(rk)
        If a(0) = d Then InputAt = a(1)
    Next
End Function

' その日より前で、いちばん新しい入力（なければ Empty）
Private Function InputBefore(ByVal rk As String, ByVal d As Long) As Variant
    Dim a As Variant, best As Long
    InputBefore = Empty
    If Not mInp.Exists(rk) Then Exit Function
    For Each a In mInp(rk)
        If a(0) < d And a(0) > best Then best = a(0): InputBefore = a(1)
    Next
End Function

'==========================================================
' 2. 業務マスタ・人・稼働記録
'==========================================================
Private Sub LoadBlocks()
    Dim ws As Worksheet, lastR As Long, r As Long, v As Variant
    Dim id As String, td As String, kt As String, rs As String, k As String, rk As String, i As Long

    Set ws = FindSheet(SH_MST)
    Set mBlk = NewDict(): Set mIdMap = NewDict(): Set mReasonSort = NewDict()
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    nBlk = 0
    ReDim bTodoke(0 To lastR): ReDim bReason(0 To lastR): ReDim bKotei(0 To lastR)
    If lastR <= MST_HDR Then Exit Sub
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, MST_COL_REASON)).Value
    For r = MST_HDR + 1 To lastR
        id = Trim$(CStr(v(r, 2)))
        td = TodokeDisp(Trim$(CStr(v(r, 4))))
        kt = Trim$(CStr(v(r, 5)))
        rs = Trim$(CStr(v(r, MST_COL_REASON)))
        If Len(rs) = 0 Then rs = ReasonFromName(CStr(v(r, 3)))
        If Len(id) > 0 And Trim$(CStr(v(r, 6))) = KIND_CASE And Len(td) > 0 And (kt = "1次" Or kt = "2次") And Len(rs) > 0 Then
            k = BlockKey(td, rs, kt)
            If Not mBlk.Exists(k) Then
                mBlk.Add k, nBlk
                bTodoke(nBlk) = td: bReason(nBlk) = rs: bKotei(nBlk) = kt
                nBlk = nBlk + 1
            End If
            mIdMap(id) = k
            rk = td & "|" & rs
            If Not mReasonSort.Exists(rk) Then
                mReasonSort.Add rk, id
            ElseIf id < mReasonSort(rk) Then
                mReasonSort(rk) = id
            End If
        End If
    Next
    i = Application.Max(0, nBlk - 1)
    ReDim bRecent(0 To i): ReDim bAuto(0 To i): ReDim bOv(0 To i): ReDim bAdopt(0 To i)
    ReDim bAlloc(0 To i): ReDim bShow(0 To i)
End Sub

Private Sub LoadPeople()
    Dim ws As Worksheet, r As Long, lastR As Long, e As String
    Set mP = NewDict()
    nP = 0
    ReDim pEid(0 To 499): ReDim pName(0 To 499): ReDim pDisp(0 To 499): ReDim pLimit(0 To 499): ReDim pNote(0 To 499): ReDim pEx(0 To 499)
    Set ws = FindSheet(SH_SKL)
    If ws Is Nothing Then Exit Sub
    ' 人の一覧：次の「■」の見出しまで読む（途中の空行・行の追加・削除は自由）
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = S_PEOPLE To lastR
        e = Trim$(CStr(ws.Cells(r, 2).Value))
        If Left$(e, 1) = "■" Then Exit For
        If Len(e) > 0 Then
            AddPerson e, ws.Cells(r, 3).Value, ws.Cells(r, 4).Value, ws.Cells(r, 5).Value, ws.Cells(r, 7).Value, _
                      Len(Trim$(CStr(ws.Cells(r, 8).Value))) > 0
        End If
    Next
End Sub

Private Sub AddPerson(ByVal eid As String, nm As Variant, dsp As Variant, lim As Variant, note As Variant, Optional ByVal ex As Boolean = False)
    If mP.Exists(eid) Or nP >= 500 Then Exit Sub
    mP.Add eid, nP
    pEid(nP) = eid: pName(nP) = nm: pDisp(nP) = dsp: pLimit(nP) = lim: pNote(nP) = note: pEx(nP) = ex
    nP = nP + 1
End Sub

' 表示名（なければ氏名、それもなければEID）。稼働実績・操作シートからも使う
Public Function PersonDisplayName(ByVal eid As String) As String
    Dim ws As Worksheet, r As Long, lastR As Long, e As String
    PersonDisplayName = eid
    Set ws = FindSheet(SH_SKL)
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = S_PEOPLE To lastR
        e = Trim$(CStr(ws.Cells(r, 2).Value))
        If Left$(e, 1) = "■" Then Exit For
        If e = eid Then
            If Len(Trim$(CStr(ws.Cells(r, 4).Value))) > 0 Then
                PersonDisplayName = Trim$(CStr(ws.Cells(r, 4).Value))
            ElseIf Len(Trim$(CStr(ws.Cells(r, 3).Value))) > 0 Then
                PersonDisplayName = Trim$(CStr(ws.Cells(r, 3).Value))
            End If
            Exit Function
        End If
    Next
End Function

' 表示名の式（スキルマップの表示名を変えるとすぐ反映）
Private Function DispF(ByVal eid As String) As String
    DispF = "=IFERROR(INDEX(det_pdisp,MATCH(""" & eid & """,det_peid,0)),""" & eid & """)"
End Function

Private Function Disp(ByVal j As Long) As String
    If Len(Trim$(CStr(pDisp(j)))) > 0 Then
        Disp = Trim$(CStr(pDisp(j)))
    ElseIf Len(Trim$(CStr(pName(j)))) > 0 Then
        Disp = Trim$(CStr(pName(j)))
    Else
        Disp = pEid(j)
    End If
End Function

Private Sub LoadWork(ByVal today As Date, ByVal d0 As Date)
    Dim src As Worksheet, v As Variant, nR As Long, nC As Long, r As Long, i As Long
    Dim cE As Long, cD As Long, cI As Long, cN As Long, cT As Long
    Dim id As String, k As String, eid As String, d As Long, cnt As Double, mins As Double, t As Long
    Dim newE As Object, arr As Variant

    Set skCnt = NewDict(): Set skMin = NewDict(): Set pCase = NewDict()
    Set actCnt = NewDict(): Set actHrs = NewDict(): Set newE = NewDict()
    Set src = FindSheet(SH_WRK)
    nR = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    nC = src.Cells(1, src.Columns.Count).End(xlToLeft).Column
    If nR < 2 Then Exit Sub
    v = src.Range("A1").Resize(nR, nC).Value
    cE = HeadIdx(v, nC, "EID"): cD = HeadIdx(v, nC, "日付"): cI = HeadIdx(v, nC, "業務ID")
    cN = HeadIdx(v, nC, "処理案件数"): cT = HeadIdx(v, nC, "所要時間（分）_")
    If cE * cD * cI * cN * cT = 0 Then Err.Raise vbObjectError + 1, , SH_WRK & " の列（EID・日付・業務ID・処理案件数・所要時間（分）_）が見つかりません"
    t = CLng(today)
    For r = 2 To nR
        eid = Trim$(CStr(v(r, cE)))
        If Len(eid) > 0 Then
            If Not mP.Exists(eid) Then newE(eid) = True
            id = Trim$(CStr(v(r, cI)))
            If mIdMap.Exists(id) And IsDate(v(r, cD)) Then
                k = mIdMap(id): d = CLng(Int(CDbl(CDate(v(r, cD)))))
                cnt = Val(CStr(v(r, cN))): mins = ToMinutes(v(r, cT))
                AddNum skCnt, k & "|" & eid, cnt
                AddNum skMin, k & "|" & eid, mins
                AddNum pCase, eid, cnt
                If d >= t - RATIO_DAYS And d < t Then bRecent(mBlk(k)) = bRecent(mBlk(k)) + cnt
                If d >= CLng(d0) And d < t Then
                    AddNum actCnt, k & "|" & d, cnt
                    AddNum actHrs, k & "|" & eid & "|" & d, mins / 60
                End If
            End If
        End If
    Next
    arr = SortedKeys(newE)
    For i = 0 To UBound(arr)
        AddPerson CStr(arr(i)), Empty, Empty, Empty, Empty
    Next
End Sub

'==========================================================
' 3. スキルマップ
'==========================================================
Private Sub BuildSkill()
    Dim ws As Worksheet, ov As Object, ord As Variant, nb As Long, i As Long, j As Long, b As Long, k As String
    Dim rM As Long, rEid As Long, rH As Long, r0 As Long, r1 As Long, rX As Long, autoC As Long, lastPC As Long
    Dim vals() As Variant, autos() As Variant, cnts() As Variant, keys() As Variant, lab() As Variant
    Dim cnt As Double, mins As Double, pp() As Variant, eids() As Variant, heads() As Variant, x As Variant
    Dim fc As FormatCondition, tl As String, inc() As Long, nI As Long, pe As String, dn() As Variant

    Set ov = ReadSkillOverrides()
    Set skillVal = NewDict()
    ord = SortedBlocks(False, False)
    nb = UBound(ord) + 1
    ' 除外していない人だけ件/時の表に出す
    ReDim inc(1 To Application.Max(1, nP))
    For j = 0 To nP - 1
        If Not pEx(j) Then nI = nI + 1: inc(nI) = j
    Next
    If nb < 1 Or nI < 1 Then Exit Sub

    ReDim vals(1 To nb, 1 To nI): ReDim autos(1 To nb, 1 To nI): ReDim cnts(1 To nb, 1 To nI)
    ReDim keys(1 To nb, 1 To 1): ReDim lab(1 To nb, 1 To 4)
    For i = 1 To nb
        b = ord(i - 1): k = BKey(b)
        keys(i, 1) = k
        lab(i, 1) = bTodoke(b): lab(i, 2) = bReason(b): lab(i, 3) = bKotei(b)
        For j = 1 To nI
            pe = pEid(inc(j))
            cnt = DictNum(skCnt, k & "|" & pe): mins = DictNum(skMin, k & "|" & pe)
            If cnt > 0 And mins > 0 Then autos(i, j) = Round(cnt / (mins / 60), 2) Else autos(i, j) = Empty
            If cnt > 0 Then cnts(i, j) = cnt
            If ov.Exists(k & "|" & pe) Then vals(i, j) = ov(k & "|" & pe) Else vals(i, j) = autos(i, j)
            x = vals(i, j)
            If Not IsEmpty(x) Then
                If IsNumeric(x) Then
                    If CDbl(x) > 0 Then skillVal(k & "|" & pe) = CDbl(x)
                End If
            End If
        Next
    Next

    Set ws = GetSheet(SH_SKL)
    ClearSheet ws
    ws.Range("B1").Value = "スキルマップ"
    ws.Range("B3").Value = "人の一覧：氏名・表示名・1日の上限時間・備考を入力（黄色）。表示名はファイル全体（詳細・稼働実績・操作シート）の表示に使います（空欄なら氏名、氏名も空欄ならEID）。上限時間は進捗（詳細）の割り当てのアラートに使います（空欄はチェックしない）。" & vbLf & _
                           "人を足すときは、一覧に行を足してEIDを入れる（稼働記録にまだいない人も可）。稼働記録に出てきたEIDは自動でも追加されます。外すときは除外の欄に○（行を消しても、稼働記録にいる人は次の更新で戻ります）。" & vbLf & _
                           "件/時：稼働記録の期間全体の 処理件数÷所要時間。数字を直接入れると上書き（黄色）。0か空欄にすると、その人はその事由・工程をしない扱い（進捗（詳細）に行が出ません）。変更は「詳細を更新」で詳細の行に反映。"

    ' 人の一覧
    ws.Cells(5, 2).Value = "■ 人の一覧"
    ws.Range("B6:H6").Value = Array("EID", "氏名", "表示名", "1日の上限時間", "案件件数（期間全体）", "備考", "除外")
    ReDim pp(1 To nP, 1 To 7): ReDim dn(1 To nP, 1 To 1)
    For j = 1 To nP
        pp(j, 1) = pEid(j - 1): pp(j, 2) = pName(j - 1): pp(j, 3) = pDisp(j - 1)
        pp(j, 4) = pLimit(j - 1): pp(j, 5) = DictNum(pCase, pEid(j - 1)): pp(j, 6) = pNote(j - 1)
        If pEx(j - 1) Then pp(j, 7) = "○"
        dn(j, 1) = "=IF(TRIM(D" & (S_PEOPLE + j - 1) & ")<>"""",TRIM(D" & (S_PEOPLE + j - 1) & "),IF(TRIM(C" & (S_PEOPLE + j - 1) & ")<>"""",TRIM(C" & (S_PEOPLE + j - 1) & "),B" & (S_PEOPLE + j - 1) & "))"
    Next
    ws.Cells(S_PEOPLE, 2).Resize(nP, 1).NumberFormat = "@"
    ws.Cells(S_PEOPLE, 2).Resize(nP, 7).Value = pp
    ws.Cells(S_PEOPLE, 1).Resize(nP, 1).Formula = dn     ' 表示に使う名前（A列・非表示）

    ' 件/時
    rM = S_PEOPLE + nP + 1
    rEid = rM + 1: rH = rM + 2: r0 = rH + 1: r1 = r0 + nb - 1
    lastPC = S_FIRST + nI - 1
    autoC = lastPC + 2
    ws.Cells(rM, 2).Value = "■ 件/時（1時間あたりの処理件数）"
    ReDim eids(1 To 1, 1 To nI): ReDim heads(1 To 1, 1 To nI)
    For j = 1 To nI: eids(1, j) = pEid(inc(j)): heads(1, j) = "=$A$" & (S_PEOPLE + inc(j)): Next
    ws.Cells(rEid, S_FIRST).Resize(1, nI).Value = eids
    ws.Cells(rEid, autoC).Resize(1, nI).Value = eids
    ws.Range(ws.Cells(rH, 2), ws.Cells(rH, 5)).Value = Array("届出", "事由", "工程", "できる人数")
    ws.Cells(rH, S_FIRST).Resize(1, nI).Formula = heads
    ws.Cells(rH, autoC).Resize(1, nI).Formula = heads
    ws.Cells(rM, autoC).Value = "自動値（比較用・非表示）"
    ws.Cells(r0, 1).Resize(nb, 1).Value = keys
    ws.Cells(r0, 2).Resize(nb, 3).Value = lab
    ws.Cells(r0, S_FIRST).Resize(nb, nI).Value = vals
    ws.Cells(r0, autoC).Resize(nb, nI).Value = autos
    ws.Cells(r0, 5).Resize(nb, 1).FormulaR1C1 = "=COUNTIF(RC" & S_FIRST & ":RC" & lastPC & ","">0"")"

    ' 経験件数（折りたたみ）
    rX = r1 + 2
    ws.Cells(rX, 2).Value = "■ 経験件数（期間全体の処理件数。＋で開く）"
    ws.Range(ws.Cells(rX + 1, 2), ws.Cells(rX + 1, 4)).Value = Array("届出", "事由", "工程")
    ws.Cells(rX + 1, S_FIRST).Resize(1, nI).Formula = heads
    ws.Cells(rX + 2, 2).Resize(nb, 3).Value = lab
    ws.Cells(rX + 2, S_FIRST).Resize(nb, nI).Value = cnts

    ' 名前（詳細から参照）
    SetName "det_skl_keys", ws.Cells(r0, 1).Resize(nb, 1)
    SetName "det_skl_eids", ws.Cells(rEid, S_FIRST).Resize(1, nI)
    SetName "det_skl_vals", ws.Cells(r0, S_FIRST).Resize(nb, nI)
    SetName "det_skl_auto", ws.Cells(r0, autoC).Resize(nb, nI)
    SetName "det_peid", ws.Cells(S_PEOPLE, 2).Resize(nP, 1)
    SetName "det_plimit", ws.Cells(S_PEOPLE, 5).Resize(nP, 1)
    SetName "det_pdisp", ws.Cells(S_PEOPLE, 1).Resize(nP, 1)

    ' 書式
    With ws.Cells
        .Font.Name = "Meiryo UI": .Font.Size = 10
    End With
    ws.Range("B1").Font.Size = 14: ws.Range("B1").Font.Bold = True
    LinkCell ws, ws.Range("B2"), SH_MENU, "操作へ戻る"
    LinkCell ws, ws.Range("D2"), SH_DET, "進捗（詳細）へ"
    With ws.Range(ws.Cells(3, 2), ws.Cells(3, 20))
        .Merge: .WrapText = True: .VerticalAlignment = xlTop
        .Font.Size = 9: .Font.Color = RGB(89, 89, 89)
    End With
    ws.Rows(3).RowHeight = 56
    SectionStyle ws.Range(ws.Cells(5, 2), ws.Cells(5, 8))
    SectionStyle ws.Range(ws.Cells(rM, 2), ws.Cells(rM, lastPC))
    SectionStyle ws.Range(ws.Cells(rX, 2), ws.Cells(rX, lastPC))
    HeadStyle ws.Range("B6:H6")
    HeadStyle ws.Range(ws.Cells(rH, 2), ws.Cells(rH, lastPC))
    HeadStyle ws.Range(ws.Cells(rX + 1, 2), ws.Cells(rX + 1, lastPC))
    GridStyle ws.Range(ws.Cells(S_PEOPLE, 2), ws.Cells(S_PEOPLE + nP - 1, 8))
    GridStyle ws.Range(ws.Cells(r0, 2), ws.Cells(r1, lastPC))
    GridStyle ws.Range(ws.Cells(rX + 2, 2), ws.Cells(rX + 1 + nb, lastPC))
    ws.Range(ws.Cells(S_PEOPLE, 3), ws.Cells(S_PEOPLE + nP - 1, 5)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(S_PEOPLE, 7), ws.Cells(S_PEOPLE + nP - 1, 8)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(S_PEOPLE, 8), ws.Cells(S_PEOPLE + nP - 1, 8)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(S_PEOPLE, 5), ws.Cells(S_PEOPLE + nP - 1, 5)).NumberFormat = "0.0#"
    ws.Range(ws.Cells(S_PEOPLE, 6), ws.Cells(S_PEOPLE + nP - 1, 6)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(r0, S_FIRST), ws.Cells(r1, lastPC)).NumberFormat = "0.0;-0.0;"
    ws.Range(ws.Cells(r0, autoC), ws.Cells(r1, autoC + nI - 1)).NumberFormat = "0.0;-0.0;"
    ws.Range(ws.Cells(rX + 2, S_FIRST), ws.Cells(rX + 1 + nb, lastPC)).NumberFormat = "#,##0;-#,##0;"
    ws.Range(ws.Cells(r0, S_FIRST), ws.Cells(r1, lastPC)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(rH, S_FIRST), ws.Cells(rH, lastPC)).WrapText = True
    ws.Rows(rH).RowHeight = 30
    ws.Range(ws.Cells(r0, 2), ws.Cells(r1, 3)).Font.Color = RGB(64, 64, 64)
    ' 届出・工程が変わる行に線
    For i = r0 + 1 To r1
        If ws.Cells(i, 2).Value <> ws.Cells(i - 1, 2).Value Or ws.Cells(i, 4).Value <> ws.Cells(i - 1, 4).Value Then
            With ws.Range(ws.Cells(i, 2), ws.Cells(i, lastPC)).Borders(xlEdgeTop)
                .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(128, 128, 128)
            End With
        End If
    Next
    ' 黄色＝自動値と違う（手で上書きした）値
    tl = ws.Cells(r0, S_FIRST).Address(False, False)
    x = ws.Cells(r0, autoC).Address(False, False)
    Set fc = ws.Range(ws.Cells(r0, S_FIRST), ws.Cells(r1, lastPC)).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=NOT(OR(AND(" & tl & "=""""," & x & "=""""),AND(ISNUMBER(" & tl & "),ISNUMBER(" & x & "),ABS(N(" & tl & ")-N(" & x & "))<0.005)))")
    fc.Interior.Color = RGB(255, 230, 153): fc.Font.Bold = True
    ' 上限時間が空欄
    tl = ws.Cells(S_PEOPLE, 5).Address(False, False)
    Set fc = ws.Range(ws.Cells(S_PEOPLE, 5), ws.Cells(S_PEOPLE + nP - 1, 5)).FormatConditions.Add(Type:=xlExpression, Formula1:="=" & tl & "=""""")
    fc.Interior.Color = RGB(252, 228, 214)
    ' 除外した人は薄い文字
    Set fc = ws.Range(ws.Cells(S_PEOPLE, 2), ws.Cells(S_PEOPLE + nP - 1, 7)).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=TRIM($H" & S_PEOPLE & ")<>""""")
    fc.Font.Color = RGB(166, 166, 166): fc.Font.Strikethrough = True
    ' 氏名が空欄（新しいEID）
    tl = ws.Cells(S_PEOPLE, 3).Address(False, False)
    Set fc = ws.Range(ws.Cells(S_PEOPLE, 3), ws.Cells(S_PEOPLE + nP - 1, 3)).FormatConditions.Add(Type:=xlExpression, Formula1:="=" & tl & "=""""")
    fc.Interior.Color = RGB(252, 228, 214)

    ws.Columns(1).ColumnWidth = 2: ws.Columns(1).Hidden = True
    ws.Columns(2).ColumnWidth = 14: ws.Columns(3).ColumnWidth = 30: ws.Columns(4).ColumnWidth = 12
    ws.Columns(5).ColumnWidth = 12: ws.Columns(6).ColumnWidth = 12: ws.Columns(7).ColumnWidth = 20
    ws.Range(ws.Cells(1, S_FIRST + 2), ws.Cells(1, lastPC)).EntireColumn.ColumnWidth = 9
    If nP >= 2 Then ws.Range(ws.Cells(1, S_FIRST), ws.Cells(1, S_FIRST + 1)).EntireColumn.ColumnWidth = 12
    ws.Range(ws.Cells(1, lastPC + 1), ws.Cells(1, autoC + nI - 1)).EntireColumn.Hidden = True
    ws.Rows(rEid).Hidden = True
    ws.Rows((rX + 1) & ":" & (rX + 1 + nb)).Group
    ws.Outline.SummaryRow = xlSummaryAbove
    ws.Outline.ShowLevels RowLevels:=1
End Sub

' 前回のスキルマップで、自動値と違う値（手で上書きした値）を読む
Private Function ReadSkillOverrides() As Object
    Dim d As Object, kv As Variant, ev As Variant, vv As Variant, av As Variant, i As Long, j As Long
    Set d = NewDict()
    Set ReadSkillOverrides = d
    If Not (NameRange("det_skl_keys") And NameRange("det_skl_eids") And NameRange("det_skl_vals") And NameRange("det_skl_auto")) Then Exit Function
    kv = RangeArr(ThisWorkbook.Names("det_skl_keys").RefersToRange)
    ev = RangeArr(ThisWorkbook.Names("det_skl_eids").RefersToRange)
    vv = RangeArr(ThisWorkbook.Names("det_skl_vals").RefersToRange)
    av = RangeArr(ThisWorkbook.Names("det_skl_auto").RefersToRange)
    If UBound(vv, 1) <> UBound(av, 1) Or UBound(vv, 2) <> UBound(av, 2) Then Exit Function
    For i = 1 To UBound(vv, 1)
        For j = 1 To UBound(vv, 2)
            If Not SameVal(vv(i, j), av(i, j)) Then d(CStr(kv(i, 1)) & "|" & CStr(ev(1, j))) = vv(i, j)
        Next
    Next
End Function

'==========================================================
' 4. 事由比率（直近4週の比率・上書き・扶養の対応表）
'==========================================================
Private Function BuildRatio(ByVal today As Date) As String
    Dim ws As Worksheet, ovR As Object, oldMap As Object, tot As Object, grp As Object, allocCnt As Object, allocTot As Object
    Dim ord As Variant, nb As Long, i As Long, j As Long, b As Long, g As String, k As String, gk As Variant
    Dim lst As Collection, rs As Variant, kt As Long, ktn As String, tg As Collection, w As Double, sw As Double
    Dim unm As Object, msg As String, nUnm As Long, x As Variant, cnts As Variant
    Dim shR As Object, isNew As Object, nNew As Object, nAll As Object, sumAuto As Object, nZero As Object, sumShow As Object

    ' 前回の比率・表示と対応表（比率は手で上書きして使う値。前回の表があればそれを使う）
    Set ovR = NewDict(): Set shR = NewDict(): Set oldMap = NewDict()
    Set ws = FindSheet(SH_RAT)
    If Not ws Is Nothing Then ReadRatioSheet ws, ovR, shR, oldMap

    ' 直近4週の比率（初めての事由×工程の初期値に使う）
    Set tot = NewDict()
    For b = 0 To nBlk - 1
        AddNum tot, bTodoke(b) & "|" & bKotei(b), bRecent(b)
    Next
    Set nNew = NewDict(): Set nAll = NewDict(): Set sumAuto = NewDict(): Set nZero = NewDict()
    For b = 0 To nBlk - 1
        g = bTodoke(b) & "|" & bKotei(b)
        If tot(g) > 0 Then bAuto(b) = bRecent(b) / tot(g) Else bAuto(b) = 0
        AddNum nAll, g, 1
        k = BKey(b)
        If ovR.Exists(k) Then
            bOv(b) = ovR(k)
        Else
            bOv(b) = Empty
            AddNum nNew, g, 1
            If bAuto(b) > 0 Then AddNum sumAuto, g, bAuto(b) Else AddNum nZero, g, 1
        End If
        bAlloc(b) = Empty
    Next
    ' 初めての行：直近4週の比率。実績のない事由は0.5%。届出×工程の行がすべて初めてなら、合計が100%になるように調整
    For b = 0 To nBlk - 1
        If IsEmpty(bOv(b)) Then
            g = bTodoke(b) & "|" & bKotei(b)
            If bAuto(b) > 0 Then
                bOv(b) = bAuto(b)
                If DictNum(nNew, g) = DictNum(nAll, g) And DictNum(sumAuto, g) > 0 Then
                    bOv(b) = Round(bAuto(b) / sumAuto(g) * (1 - 0.005 * DictNum(nZero, g)), 4)
                End If
            Else
                bOv(b) = 0.005
                If DictNum(sumAuto, g) = 0 And DictNum(nNew, g) = DictNum(nAll, g) Then bOv(b) = Round(1 / DictNum(nAll, g), 4)
            End If
        End If
    Next
    ' 表示：前回の表示の欄（初めての行は○）。今日以降の時間の入力がある行は必ず出す
    For b = 0 To nBlk - 1
        k = BKey(b)
        If shR.Exists(k) Then bShow(b) = (Len(shR(k)) > 0) Else bShow(b) = True
        If mFutureH.Exists(k) Then bShow(b) = True
    Next
    ' 計算に使う比率：表示する行の中で合計100%に割り戻す
    Set sumShow = NewDict()
    For b = 0 To nBlk - 1
        If bShow(b) Then AddNum sumShow, bTodoke(b) & "|" & bKotei(b), CDbl(bOv(b))
    Next
    For b = 0 To nBlk - 1
        g = bTodoke(b) & "|" & bKotei(b)
        If bShow(b) And DictNum(sumShow, g) > 0 Then bAdopt(b) = CDbl(bOv(b)) / sumShow(g) Else bAdopt(b) = 0
    Next

    ' 扶養：集計のフラグなし残（1次・2次）を 届出事由＋区分 ごとに数える
    Set grp = ReadFuyoCounts()
    ' 対応表：前回の行を残し、新しい 届出事由＋区分 は事由名から自動で対応づける
    For Each gk In grp.keys
        If Not oldMap.Exists(gk) Then oldMap.Add gk, AutoMatch(CStr(gk))
    Next
    ' 対応先の事由へ分ける（同じ届出事由に対応先が複数あるときは採用比率で分ける）
    Set allocCnt = NewDict(): Set allocTot = NewDict(): Set unm = NewDict()
    For Each gk In oldMap.keys
        If grp.Exists(gk) Then
            cnts = grp(gk)
            Set lst = oldMap(gk)
            For kt = 1 To 2
                ktn = kt & "次"
                If cnts(kt - 1) > 0 Then
                    Set tg = New Collection: sw = 0
                    For Each rs In lst
                        If Len(rs) > 0 Then
                            If mBlk.Exists(BlockKey("扶養", CStr(rs), ktn)) Then
                                tg.Add mBlk(BlockKey("扶養", CStr(rs), ktn))
                                sw = sw + bAdopt(mBlk(BlockKey("扶養", CStr(rs), ktn)))
                            End If
                        End If
                    Next
                    If tg.Count = 0 Then
                        AddNum unm, CStr(gk), cnts(kt - 1)
                    Else
                        For Each x In tg
                            If sw > 0 Then w = bAdopt(x) / sw Else w = 1 / tg.Count
                            AddNum allocCnt, CStr(x), cnts(kt - 1) * w
                            AddNum allocTot, ktn, cnts(kt - 1) * w
                        Next
                    End If
                End If
            Next
        End If
    Next
    For b = 0 To nBlk - 1
        If bTodoke(b) = "扶養" Then
            If DictNum(allocTot, bKotei(b)) > 0 Then bAlloc(b) = DictNum(allocCnt, CStr(b)) / allocTot(bKotei(b))
        End If
    Next

    WriteRatio today, grp, oldMap, unm

    ' メッセージ：対応のない届出事由
    For Each gk In unm.keys
        nUnm = nUnm + 1
    Next
    If nUnm > 0 Then msg = msg & vbLf & "・扶養の届出事由で、事由に対応していないものが " & nUnm & " 件あります（事由比率シートの対応表を確認してください）"
    BuildRatio = msg
End Function

' 届出事由|区分 → Array(1次の件数, 2次の件数)（扶養・フラグなし）
Private Function ReadFuyoCounts() As Object
    Dim d As Object, ws As Worksheet, v As Variant, nR As Long, nC As Long, r As Long
    Dim cT As Long, cS As Long, cF As Long, cR As Long, cK As Long, st As String, kt As Long, gk As String, a As Variant
    Set d = NewDict()
    Set ReadFuyoCounts = d
    Set ws = FindSheet(SH_SUM)
    If ws Is Nothing Then Exit Function
    nR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    nC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If nR < 2 Then Exit Function
    v = ws.Range("A1").Resize(nR, nC).Value
    cT = HeadIdx(v, nC, "申請種別"): cS = HeadIdx(v, nC, "ステータス"): cF = HeadIdx(v, nC, "案件状況")
    cR = HeadIdx(v, nC, "届出事由"): cK = HeadIdx(v, nC, "届出事由区分")
    If cT * cS * cF * cR * cK = 0 Then Exit Function
    For r = 2 To nR
        If CStr(v(r, cT)) = "扶養" And IsNoFlag(v(r, cF)) Then
            st = CStr(v(r, cS)): kt = 0
            Select Case st
                Case "確認待ち", "申請受付", "ＯＰ確認中": kt = 1
                Case "ＳＶ確認中": kt = 2
            End Select
            If kt > 0 Then
                gk = Trim$(CStr(v(r, cR))) & "|" & Trim$(CStr(v(r, cK)))
                If d.Exists(gk) Then a = d(gk) Else a = Array(0#, 0#)
                a(kt - 1) = a(kt - 1) + 1
                d(gk) = a
            End If
        End If
    Next
End Function

' 届出事由＋区分 から、扶養の事由を探す（例：離婚|減 → 【減】本人の離婚、本人の結婚|増 → 扶養なし・扶養あり の2つ）
Private Function AutoMatch(ByVal gk As String) As Collection
    Dim c As New Collection, p As Variant, b As Long, pre As String, seen As Object
    p = Split(gk, "|")
    Set seen = NewDict()
    pre = "【" & p(1) & "】"
    For b = 0 To nBlk - 1
        If bTodoke(b) = "扶養" And Not seen.Exists(bReason(b)) Then
            If Left$(bReason(b), Len(pre)) = pre And Len(p(0)) > 0 Then
                If InStr(Mid$(bReason(b), Len(pre) + 1), p(0)) > 0 Then
                    c.Add bReason(b): seen(bReason(b)) = True
                End If
            End If
        End If
    Next
    If c.Count = 0 Then c.Add ""
    Set AutoMatch = c
End Function

Private Sub ReadRatioSheet(ws As Worksheet, ovR As Object, shR As Object, oldMap As Object)
    Dim lastR As Long, v As Variant, r As Long, a As String, gk As String, isNew As Boolean
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastR < R_FIRST Then Exit Sub
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, 7)).Value
    isNew = (CStr(v(R_FIRST - 1, 5)) = "比率")     ' 以前の表（自動比率・上書き比率）なら、比率は今回の初期値で作り直す
    For r = R_FIRST To lastR
        a = CStr(v(r, 1))
        If a = "M" Then
            gk = Trim$(CStr(v(r, 2))) & "|" & Trim$(CStr(v(r, 3)))
            If Not oldMap.Exists(gk) Then oldMap.Add gk, New Collection
            oldMap(gk).Add Trim$(CStr(v(r, 4)))
        ElseIf isNew And Len(a) > 0 And InStr(a, "|") > 0 And Left$(a, 2) <> "T|" Then
            If Not IsEmpty(v(r, 5)) And IsNumeric(v(r, 5)) Then
                If Len(CStr(v(r, 5))) > 0 Then ovR(a) = CDbl(v(r, 5))
            End If
            shR(a) = Trim$(CStr(v(r, 6)))
        End If
    Next
End Sub

Private Sub WriteRatio(ByVal today As Date, grp As Object, oldMap As Object, unm As Object)
    Dim ws As Worksheet, ord As Variant, nb As Long, i As Long, b As Long, r As Long, n As Long
    Dim arr() As Variant, g As String, prevG As String, gk As Variant, rs As Variant, p As Variant
    Dim lastT As Long, rMap As Long, mArr() As Variant, nm As Long, cnts As Variant, fc As FormatCondition
    Dim keysSorted As Variant, st As String

    Set ws = GetSheet(SH_RAT)
    ClearSheet ws
    ord = SortedBlocks(False, False)
    nb = UBound(ord) + 1
    If nb < 1 Then Exit Sub
    ReDim arr(1 To nb, 1 To 11)
    For i = 0 To nb - 1
        b = ord(i): r = i + 1: n = R_FIRST + i
        arr(r, 1) = BKey(b): arr(r, 2) = bTodoke(b): arr(r, 3) = bReason(b): arr(r, 4) = bKotei(b)
        arr(r, 5) = bOv(b)
        If bShow(b) Then arr(r, 6) = "○"
        ' 計算に使う比率（表示○の行の中で合計100%に割り戻す）
        arr(r, 10) = "=IF(F" & n & "="""",0,IFERROR(N(E" & n & ")/SUMIFS($E:$E,$B:$B,B" & n & ",$D:$D,D" & n & ",$F:$F,""<>""),0))"
        ' 確認中の配分（扶養は集計の実数、ほかは比率と同じ）
        If IsEmpty(bAlloc(b)) Then arr(r, 11) = "=J" & n Else arr(r, 11) = bAlloc(b)
    Next
    lastT = R_FIRST + nb - 1

    ws.Range("B1").Value = "事由比率"
    ws.Range("B3").Value = "比率＝届出×工程ごとに、見込み申請件数と確認中を事由に分ける割合。最初は直近4週（" & Format(today - RATIO_DAYS, "yyyy/m/d") & "－" & Format(today - 1, "yyyy/m/d") & _
        "）の実績から（実績のない事由は0.5%）。数字は上書きして使ってください。合計がぴったり100%でなくても、表示○の行の中で割合に直して使います。" & vbLf & _
        "表示○の事由×工程だけ進捗（詳細）に出します（○を消すと出さない）。扶養の確認中だけは、下の対応表で【給与手当】集計の届出事由の実数で分けます。"
    ws.Cells(5, 2).Value = "■ 事由比率"
    ws.Range("A6:F6").Value = Array("キー", "届出", "事由", "工程", "比率", "表示")
    ws.Range("J6:K6").Value = Array("計算用の比率", "確認中の配分")
    ws.Cells(R_FIRST, 1).Resize(nb, 11).Formula = arr
    SetName "det_rat_keys", ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastT, 1))
    SetName "det_rat_adopt", ws.Range(ws.Cells(R_FIRST, 10), ws.Cells(lastT, 10))
    SetName "det_rat_alloc", ws.Range(ws.Cells(R_FIRST, 11), ws.Cells(lastT, 11))

    ' 扶養の対応表（折りたたみ）
    rMap = lastT + 3
    ws.Cells(rMap, 2).Value = "■ 扶養：届出事由 → 事由 の対応（扶養の確認中を事由に分けるための表。＋で開く。ふだんは触らなくてよい）"
    ws.Range(ws.Cells(rMap + 1, 1), ws.Cells(rMap + 1, 7)).Value = Array("", "届出事由", "届出事由区分", "事由（業務マスタ）", "確認中（1次）", "確認中（2次）", "状態")
    keysSorted = SortedKeys(oldMap)
    nm = 0
    For i = 0 To UBound(keysSorted)
        For Each rs In oldMap(keysSorted(i))
            nm = nm + 1
        Next
    Next
    If nm > 0 Then
        ReDim mArr(1 To nm, 1 To 7)
        n = 0
        For i = 0 To UBound(keysSorted)
            gk = keysSorted(i)
            p = Split(gk, "|")
            If grp.Exists(gk) Then cnts = grp(gk) Else cnts = Array(0#, 0#)
            For Each rs In oldMap(gk)
                n = n + 1
                mArr(n, 1) = "M": mArr(n, 2) = p(0): mArr(n, 3) = p(1): mArr(n, 4) = rs
                mArr(n, 5) = cnts(0): mArr(n, 6) = cnts(1)
                If Len(rs) = 0 Then
                    st = "対応なし"
                ElseIf Not (mBlk.Exists(BlockKey("扶養", CStr(rs), "1次")) Or mBlk.Exists(BlockKey("扶養", CStr(rs), "2次"))) Then
                    st = "業務マスタにない事由"
                ElseIf unm.Exists(gk) Then
                    st = "一部の工程の行なし"
                Else
                    st = "OK"
                End If
                mArr(n, 7) = st
            Next
        Next
        ws.Cells(rMap + 2, 1).Resize(nm, 7).Value = mArr
    End If

    ' 書式
    With ws.Cells
        .Font.Name = "Meiryo UI": .Font.Size = 10
    End With
    ws.Range("B1").Font.Size = 14: ws.Range("B1").Font.Bold = True
    LinkCell ws, ws.Range("B2"), SH_MENU, "操作へ戻る"
    LinkCell ws, ws.Range("D2"), SH_DET, "進捗（詳細）へ"
    With ws.Range("B3:G3")
        .Merge: .WrapText = True: .VerticalAlignment = xlTop
        .Font.Size = 9: .Font.Color = RGB(89, 89, 89)
    End With
    ws.Rows(3).RowHeight = 54
    SectionStyle ws.Range("B5:F5")
    HeadStyle ws.Range("B6:F6")
    GridStyle ws.Range(ws.Cells(R_FIRST, 2), ws.Cells(lastT, 6))
    ws.Range(ws.Cells(R_FIRST, 5), ws.Cells(lastT, 5)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(R_FIRST, 5), ws.Cells(lastT, 6)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(R_FIRST, 6), ws.Cells(lastT, 6)).HorizontalAlignment = xlCenter
    ' 届出×工程が変わるところに太線
    prevG = ""
    For i = 0 To nb - 1
        b = ord(i): g = bTodoke(b) & "|" & bKotei(b): n = R_FIRST + i
        If i > 0 And g <> prevG Then
            With ws.Range(ws.Cells(n, 2), ws.Cells(n, 6)).Borders(xlEdgeTop)
                .LineStyle = xlContinuous: .Weight = xlMedium: .Color = CLR_LINE
            End With
        End If
        prevG = g
    Next
    ' 表示しない行は薄い文字
    Set fc = ws.Range(ws.Cells(R_FIRST, 2), ws.Cells(lastT, 5)).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$F" & R_FIRST & "=""""")
    fc.Font.Color = RGB(166, 166, 166)
    SectionStyle ws.Range(ws.Cells(rMap, 2), ws.Cells(rMap, 8))
    HeadStyle ws.Range(ws.Cells(rMap + 1, 2), ws.Cells(rMap + 1, 7))
    If nm > 0 Then
        GridStyle ws.Range(ws.Cells(rMap + 2, 2), ws.Cells(rMap + 1 + nm, 7))
        ws.Range(ws.Cells(rMap + 2, 4), ws.Cells(rMap + 1 + nm, 4)).Interior.Color = CLR_INPUT
        Set fc = ws.Range(ws.Cells(rMap + 2, 7), ws.Cells(rMap + 1 + nm, 7)).FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=G" & (rMap + 2) & "<>""OK""")
        fc.Interior.Color = RGB(255, 0, 0): fc.Font.Color = vbWhite
    End If
    ws.Columns(1).ColumnWidth = 2: ws.Columns(1).Hidden = True
    ws.Columns(2).ColumnWidth = 16: ws.Columns(3).ColumnWidth = 34: ws.Columns(4).ColumnWidth = 30
    ws.Columns(5).ColumnWidth = 11: ws.Columns(6).ColumnWidth = 11
    ws.Columns(7).ColumnWidth = 18
    ws.Range(ws.Cells(R_FIRST, 10), ws.Cells(lastT, 11)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(1, 10), ws.Cells(1, 11)).EntireColumn.Hidden = True
    ws.Rows((rMap + 1) & ":" & (rMap + 1 + Application.Max(nm, 1))).Group
    ws.Outline.SummaryRow = xlSummaryAbove
    ws.Outline.ShowLevels RowLevels:=1
End Sub

'==========================================================
' 5. 進捗（詳細）
'==========================================================
Private Sub BuildDet(ByVal today As Date, ByVal d0 As Date, ByVal d1 As Date)
    Dim ws As Worksheet, nDays As Long, lastC As Long, t As Long
    Dim ord As Variant, i As Long, j As Long, b As Long, b2 As Long, r As Long, c As Long, k As Long, ti As Long
    Dim tdl As Variant, td As String, lastRs As String, first As Boolean
    Dim rKind() As String, rBlk() As Long, rP() As Long, rTd() As Long, rLvl() As Long
    Dim blkC() As Long, blkA() As Long, blkI() As Long, blkQ() As Long, blkP() As Long, blkR() As Long, blkH1() As Long, blkHn() As Long
    Dim pairP() As Long, unpaired() As Boolean, inS() As Boolean, anyS As Boolean
    Dim secStart As Long, secEnd As Long, pStart As Long, ovRow As Long, lastR As Long, maxR As Long
    Dim arr() As Variant, f As String, rng As String, rngA As String, pat As String, key As String, eid As String, rk As String
    Dim v As Variant, seed As Variant, dd As Long, dr As Long, sp As Object, kt As String, isOne As Boolean
    Dim fc As FormatCondition, area As Range, tl As String, rr As Long

    Set ws = GetSheet(SH_DET)
    nDays = CLng(d1) - CLng(d0) + 1
    lastC = D_FIRST + nDays - 1
    t = CLng(today)
    tdl = TodokeList()
    ord = SortedBlocks(True, True)
    Set sp = SimpleRows()

    ' ---- 行の割り付け ----
    maxR = D_TOP + 6 * D_TOP_N + 20 + (UBound(ord) + 1) * (8 + nP) + nBlk + 5 * 4 + nP * 3 + 20
    ReDim rKind(1 To maxR): ReDim rBlk(1 To maxR): ReDim rP(1 To maxR): ReDim rTd(1 To maxR): ReDim rLvl(1 To maxR)
    i = Application.Max(0, nBlk - 1)
    ReDim blkC(0 To i): ReDim blkA(0 To i): ReDim blkI(0 To i): ReDim blkQ(0 To i): ReDim blkP(0 To i): ReDim blkR(0 To i)
    ReDim blkH1(0 To i): ReDim blkHn(0 To i)
    ReDim pairP(0 To i): ReDim unpaired(0 To i): ReDim inS(0 To i)
    For r = 1 To maxR: rLvl(r) = 1: Next
    For i = 0 To 5
        For k = 0 To D_TOP_N - 1
            rKind(DetailTopRow(i, k)) = "T" & k: rTd(DetailTopRow(i, k)) = i
        Next
    Next
    r = D_TOP + 6 * D_TOP_N
    secStart = r + 1: r = secStart: rKind(r) = "SEC"
    For ti = 0 To 4
        td = tdl(ti): first = True: lastRs = ""
        For i = 0 To UBound(ord)
            b = ord(i)
            If bTodoke(b) = td Then
                If first Then r = r + 1: rKind(r) = "TD": rTd(r) = ti: first = False
                If bReason(b) <> lastRs Then
                    r = r + 1: rKind(r) = "RS": rTd(r) = ti: rBlk(r) = b: rLvl(r) = 2
                    lastRs = bReason(b)
                End If
                r = r + 1: rKind(r) = "C": rBlk(r) = b: rTd(r) = ti: rLvl(r) = 3: blkC(b) = r
                r = r + 1: rKind(r) = "A": rBlk(r) = b: rTd(r) = ti: rLvl(r) = 3: blkA(b) = r
                r = r + 1: rKind(r) = "I": rBlk(r) = b: rTd(r) = ti: rLvl(r) = 3: blkI(b) = r
                r = r + 1: rKind(r) = "Q": rBlk(r) = b: rTd(r) = ti: rLvl(r) = 3: blkQ(b) = r
                r = r + 1: rKind(r) = "P": rBlk(r) = b: rTd(r) = ti: rLvl(r) = 3: blkP(b) = r
                blkH1(b) = 0: blkHn(b) = 0
                For j = 0 To nP - 1
                    If pEx(j) Then
                    ElseIf skillVal.Exists(BKey(b) & "|" & pEid(j)) Or mFutureHP.Exists(BKey(b) & "|" & pEid(j)) Then
                        r = r + 1: rKind(r) = "H": rBlk(r) = b: rP(r) = j: rTd(r) = ti: rLvl(r) = 4
                        If blkH1(b) = 0 Then blkH1(b) = r
                        blkHn(b) = r
                    End If
                Next
                r = r + 1: rKind(r) = "R": rBlk(r) = b: rTd(r) = ti: rLvl(r) = 3: blkR(b) = r
            End If
        Next
        If Not first Then
            r = r + 1: rKind(r) = "D1": rTd(r) = ti: rLvl(r) = 3
            r = r + 1: rKind(r) = "D2": rTd(r) = ti: rLvl(r) = 3
        End If
    Next
    secEnd = r
    r = r + 2: rKind(r) = "SEC2"
    r = r + 1: rKind(r) = "OV": ovRow = r
    pStart = r + 1
    For j = 0 To nP - 1
        If Not pEx(j) Then
            r = r + 1: rKind(r) = "U": rP(r) = j
            r = r + 1: rKind(r) = "S": rP(r) = j
            r = r + 1: rKind(r) = "L": rP(r) = j
        End If
    Next
    lastR = r

    ' 1次→2次 の対応（同じ事由の1次と2次があれば対。2次の対なしの1次処理は、対なしの2次（なければ2次全体）へ比率で流す）
    For i = 0 To UBound(ord)
        b = ord(i)
        If bKotei(b) = "2次" Then
            key = BlockKey(bTodoke(b), bReason(b), "1次")
            If mBlk.Exists(key) Then
                b2 = mBlk(key)
                If bShow(b2) Then pairP(b) = blkP(b2)
            End If
        Else
            key = BlockKey(bTodoke(b), bReason(b), "2次")
            unpaired(b) = True
            If mBlk.Exists(key) Then
                If bShow(mBlk(key)) Then unpaired(b) = False
            End If
        End If
    Next
    For ti = 0 To 4
        anyS = False
        For i = 0 To UBound(ord)
            b = ord(i)
            If bTodoke(b) = tdl(ti) And bKotei(b) = "2次" And pairP(b) = 0 And bAdopt(b) > 0 Then inS(b) = True: anyS = True
        Next
        If Not anyS Then
            For i = 0 To UBound(ord)
                b = ord(i)
                If bTodoke(b) = tdl(ti) And bKotei(b) = "2次" And bAdopt(b) > 0 Then inS(b) = True
            Next
        End If
    Next

    ' ---- 中身 ----
    ReDim arr(1 To lastR, 1 To lastC)
    rng = "R" & secStart & "C:R" & lastR & "C"
    rngA = "R" & secStart & "C1:R" & lastR & "C1"
    arr(1, 2) = "進捗（詳細）"
    arr(3, 2) = "進捗（簡易版）と同じ並び：実績（青）＝確認中・認定件数、予定（黄）＝見込み申請件数 → 確認中（見込み含む） → 見込み認定件数（＋で人の行） → 残数。" & _
                "人の行に、その日からの時間を入れると 見込み認定件数＝時間×件/時（スキルマップ）。入れた日からその値が続き、変えたい日にだけ入れる（黄色）。1次で認定した分は翌日の2次へ流れる。" & vbLf & _
                "事由ごとの確認中は今日の列だけ（昨日までの最新の記録×配分比。簡易版の「確認中（見込み含む）」と同じ起点）。2次の時間を入れた日は、見込み認定件数（2次）の合計が簡易版の見込み認定件数（2次）に入ります。人別の割り当てが上限を超えると赤（日付も赤）。"
    arr(D_HDR, 2) = "届出／事由": arr(D_HDR, 3) = "区分": arr(D_HDR, 4) = "項目": arr(D_HDR, 5) = "上限・比率・件/時"
    arr(D_REST, 4) = "営業日"
    For c = D_FIRST To lastC
        arr(D_HDR, c) = CDbl(CLng(d0) + c - D_FIRST)
        arr(D_REST, c) = "=IF(OR(WEEKDAY(R4C,2)>5,COUNTIF(休日一覧!C1,R4C)>0),0,1)"
    Next
    arr(D_TOP - 1, 2) = "■ 届出別の合計（実績＝進捗（簡易版）と同じ数字。予定＝下の事由別の合計）"

    ' 届出別の合計
    For i = 0 To 5
        If i < 5 Then td = tdl(i): pat = td Else td = "合計": pat = "*"
        rr = DetailTopRow(i, 0)
        arr(rr, 2) = td
        arr(DetailTopRow(i, K_C1), 4) = "確認中（1次）"
        arr(DetailTopRow(i, K_C2), 4) = "確認中（2次）"
        arr(DetailTopRow(i, K_A1), 4) = "認定件数（1次）"
        arr(DetailTopRow(i, K_A2), 4) = "認定件数（2次）"
        arr(DetailTopRow(i, K_PLAN), 4) = "見込み申請件数"
        arr(DetailTopRow(i, K_P1), 4) = "見込み認定件数（1次）"
        arr(DetailTopRow(i, K_P2), 4) = "見込み認定件数（2次）"
        arr(DetailTopRow(i, K_R1), 4) = "残数（1次）"
        arr(DetailTopRow(i, K_R2), 4) = "残数（2次）"
        arr(DetailTopRow(i, K_RT), 4) = "残数（1次＋2次）"
        arr(DetailTopRow(i, K_H1), 4) = "時間（1次）"
        arr(DetailTopRow(i, K_H2), 4) = "時間（2次）"
        arr(DetailTopRow(i, K_U), 4) = "見込み認定件数（1次）のうち2次の対なし"
        For k = 0 To D_TOP_N - 1
            If k <= K_A2 Then arr(DetailTopRow(i, k), 1) = "JT" Else arr(DetailTopRow(i, k), 1) = "YT"
        Next
        arr(DetailTopRow(i, K_RT), 1) = "YR|" & td
        ' 残数の上限：手で入れた値があればそれ、なければ簡易版の「残数の上限」
        v = InputBefore("Z||YR|" & td, 999999)
        If Not IsEmpty(v) Then
            arr(DetailTopRow(i, K_RT), 5) = v
        ElseIf sp.Exists(i & "|残数") Then
            arr(DetailTopRow(i, K_RT), 5) = "=IF('" & SH_SIMPLE & "'!R" & sp(i & "|残数") & "C5="""","""",'" & SH_SIMPLE & "'!R" & sp(i & "|残数") & "C5)"
        End If
        For c = D_FIRST To lastC
            arr(DetailTopRow(i, K_C1), c) = SimpleRef(sp, i, "確認中（1次）", True)
            arr(DetailTopRow(i, K_C2), c) = SimpleRef(sp, i, "確認中（2次）", True)
            arr(DetailTopRow(i, K_A1), c) = SimpleRef(sp, i, "認定件数（1次）", False)
            arr(DetailTopRow(i, K_A2), c) = SimpleRef(sp, i, "認定件数（2次）", False)
            arr(DetailTopRow(i, K_PLAN), c) = "=IFERROR(LET(x,XLOOKUP(DATE(YEAR(R4C)-1,MONTH(R4C),DAY(R4C)),'" & SH_PLAN & "'!R4C3:R4C400,'" & SH_PLAN & "'!R" & (5 + i) & "C3:R" & (5 + i) & "C400,""""),IF(ISNUMBER(x),x,"""")),"""")"
            arr(DetailTopRow(i, K_P1), c) = "=IF(R4C<TODAY(),"""",IF(N(R" & DetailTopRow(i, K_H1) & "C)=0,"""",SUMIFS(" & rng & "," & rngA & ",""P1*|" & pat & """)))"
            arr(DetailTopRow(i, K_P2), c) = "=IF(R4C<TODAY(),"""",IF(N(R" & DetailTopRow(i, K_H2) & "C)=0,"""",SUMIFS(" & rng & "," & rngA & ",""P2*|" & pat & """)))"
            arr(DetailTopRow(i, K_R1), c) = "=IF(R4C<TODAY(),"""",SUMIFS(" & rng & "," & rngA & ",""R1|" & pat & """))"
            arr(DetailTopRow(i, K_R2), c) = "=IF(R4C<TODAY(),"""",SUMIFS(" & rng & "," & rngA & ",""R2|" & pat & """))"
            arr(DetailTopRow(i, K_RT), c) = "=IF(R4C<TODAY(),"""",N(R" & DetailTopRow(i, K_R1) & "C)+N(R" & DetailTopRow(i, K_R2) & "C))"
            arr(DetailTopRow(i, K_H1), c) = "=IF(R4C<TODAY(),"""",SUMIFS(" & rng & "," & rngA & ",""H1|" & pat & "|*""))"
            arr(DetailTopRow(i, K_H2), c) = "=IF(R4C<TODAY(),"""",SUMIFS(" & rng & "," & rngA & ",""H2|" & pat & "|*""))"
            arr(DetailTopRow(i, K_U), c) = "=SUMIFS(" & rng & "," & rngA & ",""P1U|" & pat & """)"
        Next
    Next

    ' 事由別
    arr(secStart, 2) = "■ 事由別（＋で開く：事由 → 実績・予定 → 見込み認定件数の＋で人の行）"
    For r = secStart + 1 To secEnd
        Select Case rKind(r)
        Case "TD"
            arr(r, 2) = "■ " & tdl(rTd(r))
        Case "RS"
            b = rBlk(r): td = bTodoke(b)
            arr(r, 2) = bReason(b): arr(r, 4) = "残数（1次＋2次）"
            f = ""
            For i = 0 To UBound(ord)
                b2 = ord(i)
                If bTodoke(b2) = td And bReason(b2) = bReason(b) Then f = f & "+N(R" & blkR(b2) & "C)"
            Next
            For c = D_FIRST To lastC
                arr(r, c) = "=IF(R4C<TODAY(),""""," & Mid$(f, 2) & ")"
            Next
        Case "C", "A", "I", "Q", "P", "R", "H"
            b = rBlk(r): td = bTodoke(b): key = BKey(b): ti = rTd(r)
            kt = bKotei(b): isOne = (kt = "1次")
            Select Case rKind(r)
            Case "C"
                arr(r, 1) = IIf(isOne, "JC1|", "JC2|") & td
                arr(r, 2) = bReason(b) & vbLf & "（" & kt & "）"
                arr(r, 3) = "実績"
                arr(r, 4) = "確認中"
                arr(r, 5) = "=IFERROR(INDEX(det_rat_alloc,MATCH(""" & key & """,det_rat_keys,0)),0)"
                f = "=IF(R4C=TODAY(),N(R" & DetailTopRow(ti, IIf(isOne, K_C1, K_C2)) & "C[-1])*RC5,"""")"
                For c = D_FIRST To lastC: arr(r, c) = f: Next
            Case "A"
                arr(r, 1) = IIf(isOne, "JA1|", "JA2|") & td
                arr(r, 4) = IIf(isOne, "認定件数（1次）", "認定件数（2次）")
                For c = D_FIRST To lastC
                    dd = CLng(d0) + c - D_FIRST
                    If dd < t Then
                        If actCnt.Exists(key & "|" & dd) Then arr(r, c) = actCnt(key & "|" & dd)
                    End If
                Next
            Case "I"
                arr(r, 3) = "予定"
                arr(r, 5) = "=IFERROR(INDEX(det_rat_adopt,MATCH(""" & key & """,det_rat_keys,0)),0)"
                If isOne Then
                    arr(r, 1) = "I1|" & td: arr(r, 4) = "見込み申請件数"
                    f = "=IF(R4C<TODAY(),"""",N(R" & DetailTopRow(ti, K_PLAN) & "C)*RC5)"
                Else
                    arr(r, 1) = IIf(inS(b), "I2S|", "I2|") & td: arr(r, 4) = "1次からの流入"
                    f = ""
                    If pairP(b) > 0 Then f = "N(R" & pairP(b) & "C[-1])"
                    If inS(b) Then f = f & IIf(Len(f) > 0, "+", "") & "IFERROR(N(R" & DetailTopRow(ti, K_U) & "C[-1])*RC5/SUMIFS(R" & secStart & "C5:R" & lastR & "C5," & rngA & ",""I2S|" & td & """),0)"
                    If Len(f) = 0 Then f = "0"
                    f = "=IF(R4C<TODAY(),"""",IF(R4C[-1]<TODAY(),0," & f & "))"
                End If
                For c = D_FIRST To lastC: arr(r, c) = f: Next
            Case "Q"
                arr(r, 1) = IIf(isOne, "Q1|", "Q2|") & td
                arr(r, 4) = "確認中（見込み含む）"
                f = "=IF(R4C<TODAY(),"""",IF(R4C[-1]<TODAY(),N(R" & blkC(b) & "C),N(R" & blkR(b) & "C[-1]))+N(R" & blkI(b) & "C))"
                For c = D_FIRST To lastC: arr(r, c) = f: Next
            Case "P"
                If isOne Then arr(r, 1) = IIf(unpaired(b), "P1U|", "P1|") & td Else arr(r, 1) = "P2|" & td
                arr(r, 4) = IIf(isOne, "見込み認定件数（1次）", "見込み認定件数（2次）")
                If blkH1(b) > 0 Then
                    f = "=IF(R4C<TODAY(),"""",IF(R5C=0,0,MIN(N(R" & blkQ(b) & "C),SUMPRODUCT(R" & blkH1(b) & "C:R" & blkHn(b) & "C,R" & blkH1(b) & "C5:R" & blkHn(b) & "C5))))"
                Else
                    f = "=IF(R4C<TODAY(),"""",0)"
                End If
                For c = D_FIRST To lastC: arr(r, c) = f: Next
            Case "R"
                arr(r, 1) = IIf(isOne, "R1|", "R2|") & td
                arr(r, 4) = "残数"
                f = "=IF(R4C<TODAY(),"""",MAX(0,N(R" & blkQ(b) & "C)-N(R" & blkP(b) & "C)))"
                For c = D_FIRST To lastC: arr(r, c) = f: Next
            Case "H"
                j = rP(r): eid = pEid(j)
                arr(r, 1) = IIf(isOne, "H1|", "H2|") & key & "|" & eid
                arr(r, 4) = DispF(eid)
                arr(r, 5) = "=IFERROR(INDEX(det_skl_vals,MATCH(""" & key & """,det_skl_keys,0),MATCH(""" & eid & """,det_skl_eids,0)),0)"
                rk = "H|" & eid & "|" & key
                For c = D_FIRST To lastC
                    dd = CLng(d0) + c - D_FIRST
                    If dd < t Then
                        If actHrs.Exists(key & "|" & eid & "|" & dd) Then arr(r, c) = Round(actHrs(key & "|" & eid & "|" & dd), 2)
                    Else
                        v = InputAt(rk, dd)
                        If Not IsEmpty(v) Then
                            arr(r, c) = v
                        ElseIf dd = t Then
                            seed = InputBefore(rk, dd)
                            If IsEmpty(seed) Then arr(r, c) = "=0" Else arr(r, c) = "=" & Str$(seed)
                        Else
                            arr(r, c) = "=N(RC[-1])"
                        End If
                    End If
                Next
            End Select
        Case "D1", "D2"
            td = tdl(rTd(r))
            dr = DailyRow(td, IIf(rKind(r) = "D1", "登録", "ＳＶ承認"))
            arr(r, 1) = "JD"
            arr(r, 2) = td
            arr(r, 4) = IIf(rKind(r) = "D1", "確認用：認定件数（1次）の差", "確認用：認定件数（2次）の差")
            arr(r, 5) = "稼働記録－件数一覧"
            If dr > 0 Then
                f = "=IF(R4C>=TODAY(),"""",LET(x,XLOOKUP(R4C,'" & SH_DAILY & "'!R4C1:R4C1000,'" & SH_DAILY & "'!R" & dr & "C1:R" & dr & "C1000,""""),IF(ISNUMBER(x),SUMIFS(" & rng & "," & rngA & ",""" & IIf(rKind(r) = "D1", "JA1|", "JA2|") & td & """)-x,"""")))"
                For c = D_FIRST To lastC: arr(r, c) = f: Next
            End If
        End Select
    Next

    ' 人別の割り当て
    arr(secEnd + 2, 2) = "■ 人別の割り当て時間（上限＝スキルマップの1日の上限時間。その日だけ変えるときは上限の欄に入れる。休み＝0）"
    arr(ovRow, 4) = "上限を超えた人数"
    For c = D_FIRST To lastC
        arr(ovRow, c) = "=IF(R4C<TODAY(),"""",COUNTIFS(R" & pStart & "C:R" & lastR & "C,""<0"",R" & pStart & "C1:R" & lastR & "C1,""L|*""))"
    Next
    For r = pStart To lastR
        j = rP(r): eid = pEid(j)
        Select Case rKind(r)
        Case "U"
            arr(r, 1) = "U|" & eid: arr(r, 2) = DispF(eid): arr(r, 4) = "上限"
            For c = D_FIRST To lastC
                dd = CLng(d0) + c - D_FIRST
                If dd >= t Then
                    v = InputAt("U|" & eid & "|", dd)
                    If IsEmpty(v) Then
                        arr(r, c) = "=IF(R5C=0,"""",IFERROR(IF(INDEX(det_plimit,MATCH(""" & eid & """,det_peid,0))="""","""",INDEX(det_plimit,MATCH(""" & eid & """,det_peid,0))),""""))"
                    Else
                        arr(r, c) = v
                    End If
                End If
            Next
        Case "S"
            arr(r, 1) = "S|" & eid: arr(r, 4) = "割当"
            For c = D_FIRST To lastC
                arr(r, c) = "=IF(R5C=0,0,SUMIFS(" & rng & "," & rngA & ",""H?|*|" & eid & """))"
            Next
        Case "L"
            arr(r, 1) = "L|" & eid: arr(r, 4) = "残り"
            For c = D_FIRST To lastC
                arr(r, c) = "=IF(OR(R4C<TODAY(),R[-2]C=""""),"""",R[-2]C-R[-1]C)"
            Next
        End Select
    Next

    ' ---- 書き出し ----
    ClearSheet ws
    ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Formula2R1C1 = arr

    ' ---- 書式（進捗（簡易版）と同じ見た目） ----
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC))
        .Font.Name = "Meiryo UI": .Font.Size = 10
    End With
    ws.Range("B1").Font.Size = 14: ws.Range("B1").Font.Bold = True
    LinkCell ws, ws.Range("B2"), SH_MENU, "操作へ戻る"
    LinkCell ws, ws.Range("D2"), SH_SIMPLE, "進捗（簡易版）へ"
    LinkCell ws, ws.Range("E2"), SH_SKL, "スキルマップ"
    LinkCell ws, ws.Range("G2"), SH_RAT, "事由比率"
    With ws.Range(ws.Cells(3, 2), ws.Cells(3, D_FIRST + 25))
        .Merge: .WrapText = True: .VerticalAlignment = xlTop
        .Font.Size = 9: .Font.Color = RGB(89, 89, 89)
    End With
    ws.Rows(3).RowHeight = 40
    With ws.Range(ws.Cells(D_HDR, 2), ws.Cells(D_HDR, 4))
        .Font.Bold = True: .Interior.Color = CLR_ACT: .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(D_HDR, 5), ws.Cells(D_HDR, lastC)).Font.Bold = True
    ws.Range(ws.Cells(D_HDR, D_FIRST), ws.Cells(D_HDR, lastC)).NumberFormat = "m/d(aaa)"
    ws.Range(ws.Cells(D_HDR, D_FIRST), ws.Cells(D_HDR, lastC)).HorizontalAlignment = xlCenter
    ws.Rows(D_HDR).RowHeight = 30
    SectionStyle ws.Range(ws.Cells(D_TOP - 1, 2), ws.Cells(D_TOP - 1, lastC))
    SectionStyle ws.Range(ws.Cells(secStart, 2), ws.Cells(secStart, lastC))
    SectionStyle ws.Range(ws.Cells(secEnd + 2, 2), ws.Cells(secEnd + 2, lastC))
    ws.Range(ws.Cells(D_TOP, D_FIRST), ws.Cells(lastR, lastC)).NumberFormat = "#,##0;-#,##0;"

    ' 届出別の合計：簡易版と同じブロック
    For i = 0 To 5
        StyleBlock ws, DetailTopRow(i, 0), DetailTopRow(i, K_A2), DetailTopRow(i, K_H2), lastC, 12
        For k = K_C1 To K_A2
            ws.Range(ws.Cells(DetailTopRow(i, k), D_FIRST), ws.Cells(DetailTopRow(i, k), lastC)).NumberFormat = "#,##0"
        Next
        ws.Range(ws.Cells(DetailTopRow(i, K_RT), 4), ws.Cells(DetailTopRow(i, K_RT), lastC)).Font.Bold = True
        ws.Range(ws.Cells(DetailTopRow(i, K_R1), D_FIRST), ws.Cells(DetailTopRow(i, K_RT), lastC)).NumberFormat = "#,##0"
        With ws.Cells(DetailTopRow(i, K_RT), 5)
            .Interior.Color = CLR_PLAN: .NumberFormat = """上限 ""#,##0": .HorizontalAlignment = xlRight
        End With
        ws.Range(ws.Cells(DetailTopRow(i, K_H1), D_FIRST), ws.Cells(DetailTopRow(i, K_H2), lastC)).NumberFormat = "0.0;-0.0;"
        ws.Rows(DetailTopRow(i, K_U)).Hidden = True
    Next

    ' 事由別
    For r = secStart + 1 To secEnd
        Select Case rKind(r)
        Case "TD"
            With ws.Range(ws.Cells(r, 2), ws.Cells(r, lastC))
                .Interior.Color = RGB(228, 223, 236): .Font.Bold = True: .Font.Size = 11
                .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlMedium: .Borders(xlEdgeTop).Color = CLR_LINE
            End With
            ws.Rows(r).RowHeight = 20
        Case "RS"
            With ws.Range(ws.Cells(r, 2), ws.Cells(r, lastC))
                .Interior.Color = RGB(242, 242, 242): .Font.Bold = True
                .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlMedium: .Borders(xlEdgeTop).Color = CLR_LINE
            End With
            ws.Range(ws.Cells(r, D_FIRST), ws.Cells(r, lastC)).NumberFormat = "#,##0"
        Case "C"
            b = rBlk(r)
            StyleBlock ws, blkC(b), blkA(b), blkR(b), lastC, 10
            ws.Cells(r, 5).NumberFormat = """配分 ""0.0%"
            ws.Range(ws.Cells(r, D_FIRST), ws.Cells(r, lastC)).NumberFormat = "#,##0"
            ws.Cells(blkI(b), 5).NumberFormat = """比率 ""0.0%"
            ws.Range(ws.Cells(blkR(b), 4), ws.Cells(blkR(b), lastC)).Font.Bold = True
            ws.Range(ws.Cells(blkR(b), D_FIRST), ws.Cells(blkR(b), lastC)).NumberFormat = "#,##0"
            ws.Range(ws.Cells(blkQ(b), D_FIRST), ws.Cells(blkQ(b), lastC)).NumberFormat = "#,##0"
            ws.Range(ws.Cells(r, 5), ws.Cells(blkR(b), 5)).Font.Color = RGB(89, 89, 89)
            ws.Range(ws.Cells(r, 5), ws.Cells(blkR(b), 5)).Font.Size = 9
        Case "H"
            With ws.Cells(r, 4)
                .Interior.pattern = xlNone: .Font.Color = RGB(31, 78, 121): .IndentLevel = 1
            End With
            ws.Cells(r, 5).NumberFormat = "0.0""件/h"";;"
            ws.Range(ws.Cells(r, D_FIRST), ws.Cells(r, lastC)).NumberFormat = "0.0;-0.0;"
        Case "D1", "D2"
            With ws.Range(ws.Cells(r, 2), ws.Cells(r, lastC))
                .Font.Color = RGB(128, 128, 128): .Font.Italic = True: .Font.Size = 9
            End With
            ws.Range(ws.Cells(r, D_FIRST), ws.Cells(r, lastC)).NumberFormat = "+#,##0;-#,##0;"
            If rKind(r) = "D1" Then
                With ws.Range(ws.Cells(r, 2), ws.Cells(r, lastC)).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous: .Weight = xlMedium: .Color = CLR_LINE
                End With
            End If
        End Select
    Next

    ' 人別
    ws.Range(ws.Cells(ovRow, 4), ws.Cells(ovRow, lastC)).Font.Bold = True
    For r = pStart To lastR
        Select Case rKind(r)
        Case "U"
            With ws.Range(ws.Cells(r, 2), ws.Cells(r + 2, 2))
                .Merge: .VerticalAlignment = xlCenter: .Interior.Color = CLR_ACT: .Font.Bold = True
            End With
            ws.Range(ws.Cells(r, 4), ws.Cells(r + 2, 4)).Interior.Color = CLR_PLAN
            With ws.Range(ws.Cells(r, 2), ws.Cells(r, lastC)).Borders(xlEdgeTop)
                .LineStyle = xlContinuous: .Weight = xlMedium: .Color = CLR_LINE
            End With
            ws.Range(ws.Cells(r, D_FIRST), ws.Cells(r + 2, lastC)).NumberFormat = "0.0;-0.0;"
            ws.Range(ws.Cells(r + 2, D_FIRST), ws.Cells(r + 2, lastC)).NumberFormat = "0.0;-0.0;0.0"
            ws.Range(ws.Cells(r + 1, 4), ws.Cells(r + 1, lastC)).Font.Bold = True
        End Select
    Next

    ' 日付の列：縦の細線（簡易版と同じ）
    With ws.Range(ws.Cells(D_HDR, D_FIRST), ws.Cells(lastR, lastC))
        .Borders(xlInsideVertical).LineStyle = xlContinuous: .Borders(xlInsideVertical).Color = CLR_GRID
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Color = CLR_GRID
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Color = CLR_GRID
    End With

    ' 条件付き書式（簡易版と同じ色）
    Set area = ws.Range(ws.Cells(D_TOP, D_FIRST), ws.Cells(lastR, lastC))
    tl = ws.Cells(D_TOP, D_FIRST).Address(False, False)
    ' 実績の行の未来・予定の行の過去は薄いグレー
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & ColRef(D_FIRST, D_HDR) & ")," & ColRef(D_FIRST, D_HDR) & ">TODAY(),LEFT($A" & D_TOP & ",1)=""J"")")
    fc.Interior.Color = RGB(231, 231, 231): fc.Font.Color = RGB(166, 166, 166)
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & ColRef(D_FIRST, D_HDR) & ")," & ColRef(D_FIRST, D_HDR) & "<TODAY(),$A" & D_TOP & "<>"""",LEFT($A" & D_TOP & ",1)<>""J"")")
    fc.Interior.Color = RGB(231, 231, 231): fc.Font.Color = RGB(166, 166, 166)
    ' 土日祝
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & ColRef(D_FIRST, D_HDR) & ")," & ColRef(D_FIRST, D_REST) & "=0,LEFT($B" & D_TOP & ",1)<>""■"")")
    fc.Interior.Color = RGB(217, 217, 217): fc.Font.Color = RGB(128, 128, 128)
    fc.SetFirstPriority
    ' 今日
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & ColRef(D_FIRST, D_HDR) & ")," & ColRef(D_FIRST, D_HDR) & "=TODAY(),LEFT($B" & D_TOP & ",1)<>""■"")")
    fc.Interior.Color = RGB(255, 230, 190)
    fc.SetFirstPriority
    ' 日付の見出し：今日までは青、先は黄、今日はオレンジ
    Set fc = ws.Range(ws.Cells(D_HDR, D_FIRST), ws.Cells(D_HDR, lastC)).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=" & ColRef(D_FIRST, D_HDR) & "<=TODAY()")
    fc.Interior.Color = CLR_ACT
    Set fc = ws.Range(ws.Cells(D_HDR, D_FIRST), ws.Cells(D_HDR, lastC)).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=" & ColRef(D_FIRST, D_HDR) & ">TODAY()")
    fc.Interior.Color = RGB(255, 230, 153)
    Set fc = ws.Range(ws.Cells(D_HDR, D_FIRST), ws.Cells(D_HDR, lastC)).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=" & ColRef(D_FIRST, D_HDR) & "=TODAY()")
    fc.Interior.Color = RGB(237, 125, 49): fc.Font.Color = vbWhite
    fc.SetFirstPriority
    ' 手で入れた時間・上限は黄色
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(" & ColRef(D_FIRST, D_HDR) & ">=TODAY(),ISNUMBER(" & tl & "),NOT(ISFORMULA(" & tl & ")),OR(LEFT($A" & D_TOP & ",1)=""H"",LEFT($A" & D_TOP & ",2)=""U|""))")
    fc.Interior.Color = RGB(255, 242, 0): fc.Font.Bold = True
    fc.SetFirstPriority
    ' 件/時がない人に時間
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(LEFT($A" & D_TOP & ",1)=""H""," & ColRef(D_FIRST, D_HDR) & ">=TODAY(),N(" & tl & ")>0,N($E" & D_TOP & ")=0)")
    fc.Interior.Color = RGB(255, 0, 0): fc.Font.Color = vbWhite
    fc.SetFirstPriority
    ' 割当が上限を超えた
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(LEFT($A" & D_TOP & ",2)=""S|""," & ColRef(D_FIRST, D_HDR) & ">=TODAY(),ISNUMBER(" & ws.Cells(D_TOP - 1, D_FIRST).Address(False, False) & "),N(" & tl & ")>" & ws.Cells(D_TOP - 1, D_FIRST).Address(False, False) & ")")
    fc.Interior.Color = RGB(255, 0, 0): fc.Font.Color = vbWhite
    fc.SetFirstPriority
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(LEFT($A" & D_TOP & ",2)=""L|"",ISNUMBER(" & tl & ")," & tl & "<0)")
    fc.Font.Color = RGB(255, 0, 0): fc.Font.Bold = True
    fc.SetFirstPriority
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(LEFT($A" & D_TOP & ",3)=""YR|""," & ColRef(D_FIRST, D_HDR) & ">=TODAY(),ISNUMBER(" & tl & "),ISNUMBER($E" & D_TOP & ")," & tl & ">$E" & D_TOP & ")")
    fc.Interior.Color = 13551615: fc.Font.Color = 393372
    fc.SetFirstPriority
    Set fc = ws.Range(ws.Cells(ovRow, D_FIRST), ws.Cells(ovRow, lastC)).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=N(" & ws.Cells(ovRow, D_FIRST).Address(False, False) & ")>0")
    fc.Interior.Color = RGB(255, 0, 0): fc.Font.Color = vbWhite
    fc.SetFirstPriority
    Set fc = ws.Range(ws.Cells(D_HDR, D_FIRST), ws.Cells(D_HDR, lastC)).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=N(" & ws.Cells(ovRow, D_FIRST).Address(True, False) & ")>0")
    fc.Interior.Color = RGB(255, 0, 0): fc.Font.Color = vbWhite
    fc.SetFirstPriority

    ' 列・行
    ws.Columns(1).ColumnWidth = 2: ws.Columns(1).Hidden = True
    ws.Columns(2).ColumnWidth = 24: ws.Columns(3).ColumnWidth = 5: ws.Columns(4).ColumnWidth = 20
    ws.Columns(5).ColumnWidth = 11
    ws.Range(ws.Cells(1, D_FIRST), ws.Cells(1, lastC)).EntireColumn.ColumnWidth = 9.5
    ws.Rows(D_REST).Hidden = True
    For r = secStart + 1 To secEnd
        If rLvl(r) > 1 Then ws.Rows(r).OutlineLevel = rLvl(r)
    Next
    ws.Outline.SummaryRow = xlSummaryAbove
    ws.Outline.ShowLevels RowLevels:=2
    ws.Rows(D_REST).Hidden = True
    For i = 0 To 5
        ws.Rows(DetailTopRow(i, K_U)).Hidden = True
    Next

    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
    ws.Cells(D_HDR + 2, D_FIRST).Select      ' F6（5行目は非表示）の左上で固定：A～E列と1～5行目
    ActiveWindow.FreezePanes = True
    ActiveWindow.ScrollColumn = D_FIRST + DAYS_BACK - 2     ' 一昨日の列から見せる（今日の列が見える位置）
    ActiveWindow.ScrollRow = D_HDR + 2
    ws.Cells(D_HDR + 2, D_FIRST + DAYS_BACK).Select       ' 選択中のセルを今日の列に（画面の外にあると、折りたたみの開閉で横に飛ぶ）
    ActiveWindow.DisplayGridlines = False
End Sub

' 簡易版と同じブロックの見た目（届出・事由のラベルを縦に結合、実績＝青／予定＝黄、外枠の太線・実績と予定の間の破線・細線）
'   r0：ブロックの最初の行、rA：実績の最後の行、rE：ブロックの最後の行
Private Sub StyleBlock(ws As Worksheet, ByVal r0 As Long, ByVal rA As Long, ByVal rE As Long, ByVal lastC As Long, ByVal labelSize As Long)
    Dim rg As Range
    ' 行の細線
    With ws.Range(ws.Cells(r0, 2), ws.Cells(rE, lastC))
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous: .Borders(xlInsideHorizontal).Color = CLR_GRID
    End With
    ' ラベル列
    With ws.Range(ws.Cells(r0, 2), ws.Cells(rE, 2))
        .Merge: .Interior.Color = CLR_ACT: .Font.Bold = True: .Font.Size = labelSize
        .VerticalAlignment = xlCenter: .WrapText = True
    End With
    With ws.Range(ws.Cells(r0, 3), ws.Cells(rA, 3))
        .Merge: .Value = "実績": .Interior.Color = CLR_ACT: .Font.Bold = True: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(rA + 1, 3), ws.Cells(rE, 3))
        .Merge: .Value = "予定": .Interior.Color = CLR_PLAN: .Font.Bold = True: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(r0, 4), ws.Cells(rA, 4)).Interior.Color = CLR_ACT
    ws.Range(ws.Cells(rA + 1, 4), ws.Cells(rE, 4)).Interior.Color = CLR_PLAN
    ' 縦の細線（ラベル～項目）
    With ws.Range(ws.Cells(r0, 2), ws.Cells(rE, 5))
        .Borders(xlInsideVertical).LineStyle = xlContinuous: .Borders(xlInsideVertical).Color = CLR_GRID
    End With
    ' 外枠：上と左は太線
    Set rg = ws.Range(ws.Cells(r0, 2), ws.Cells(r0, lastC))
    With rg.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = CLR_LINE
    End With
    With ws.Range(ws.Cells(r0, 2), ws.Cells(rE, 2)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = CLR_LINE
    End With
    With ws.Range(ws.Cells(rE, 2), ws.Cells(rE, lastC)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = CLR_LINE
    End With
    ' 実績と予定の間は破線
    With ws.Range(ws.Cells(rA, 3), ws.Cells(rA, lastC)).Borders(xlEdgeBottom)
        .LineStyle = xlDash: .Weight = xlThin: .Color = CLR_LINE
    End With
End Sub

' 簡易版の各ブロックの行（届出番号|項目 → 行）
Private Function SimpleRows() As Object
    Dim d As Object, ws As Worksheet, lastR As Long, r As Long, ti As Long, nm As String, it As String
    Set d = NewDict()
    Set SimpleRows = d
    Set ws = FindSheet(SH_SIMPLE)
    If ws Is Nothing Then Exit Function
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1     ' 折りたたんだ行も含める
    ti = -1
    For r = 5 To lastR
        nm = Trim$(CStr(ws.Cells(r, 2).Value))
        Select Case nm
            Case "給与口座": ti = 0
            Case "住居": ti = 1
            Case "通勤": ti = 2
            Case "扶養": ti = 3
            Case "その他": ti = 4
            Case "合計": ti = 5
        End Select
        it = Trim$(CStr(ws.Cells(r, 4).Value))
        If ti >= 0 And Len(it) > 0 Then
            If Not d.Exists(ti & "|" & it) Then d.Add ti & "|" & it, r
        End If
    Next
End Function

' 簡易版の行を日付で引く式（実績：確認中は今日まで、件数は昨日まで）
Private Function SimpleRef(sp As Object, ByVal ti As Long, ByVal it As String, ByVal inclToday As Boolean) As String
    Dim r As Long, s As String
    If Not sp.Exists(ti & "|" & it) Then SimpleRef = "": Exit Function
    r = sp(ti & "|" & it)
    s = "'" & SH_SIMPLE & "'!"
    SimpleRef = "=IF(R4C" & IIf(inclToday, ">", ">=") & "TODAY(),"""",LET(x,XLOOKUP(R4C," & s & "R4C6:R4C400," & s & "R" & r & "C6:R" & r & "C400,""""),IF(ISNUMBER(x),x,"""")))"
End Function

' 簡易版の言葉を詳細とそろえ（認定件数（1次）／（2次）、見込み認定件数（1次）／（2次）など）、
' 認定件数（2次）と同じ中身の「2次実績」の行を消し、見込み認定件数の内訳の行を詳細の届出別合計の行につなぐ
Private Sub LinkSimple()
    Dim ws As Worksheet, lastR As Long, r As Long, ti As Long, nm As String, k As Long, f As String, it As String, nw As String, t3 As String
    Set ws = FindSheet(SH_SIMPLE)
    If ws Is Nothing Then Exit Sub
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1     ' 折りたたんだ行も含める
    For r = lastR To 5 Step -1
        nw = ""
        Select Case Trim$(CStr(ws.Cells(r, 4).Value))
            Case "2次実績": ws.Rows(r).Delete
            Case "1次残": nw = "確認中（1次）"
            Case "2次残": nw = "確認中（2次）"
            Case "1次実績", "1次処理件数": nw = "認定件数（1次）"
            Case "認定件数": nw = "認定件数（2次）"
            Case "認定見込み": nw = "見込み認定件数（2次）"
            Case "1次処理見込み": nw = "見込み認定件数（1次）"
            Case "2次処理見込み", "認定見込み（詳細）": nw = "見込み認定件数（2次・詳細）"
            Case "認定見込み（手入力）": nw = "見込み認定件数（2次・手入力）"
        End Select
        If Len(nw) > 0 Then ws.Cells(r, 4).Value = nw
    Next
    ' 上の説明文の言葉もそろえる
    t3 = CStr(ws.Range("B3").Value)
    If InStr(t3, "認定見込み") > 0 Or InStr(t3, "認定件数＝") > 0 Then
        t3 = Replace(t3, "認定見込み", "見込み認定件数（2次）")
        t3 = Replace(t3, "認定件数＝", "認定件数（2次）＝")
        ws.Range("B3").Value = t3
    End If
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    ti = -1
    For r = 5 To lastR
        nm = Trim$(CStr(ws.Cells(r, 2).Value))
        Select Case nm
            Case "給与口座": ti = 0
            Case "住居": ti = 1
            Case "通勤": ti = 2
            Case "扶養": ti = 3
            Case "その他": ti = 4
            Case "合計": ti = 5
        End Select
        it = Trim$(CStr(ws.Cells(r, 4).Value))
        k = -1
        Select Case it
            Case "見込み認定件数（1次）": k = K_P1
            Case "見込み認定件数（2次・詳細）": k = K_P2
            Case "認定件数（1次）"
                If CStr(ws.Cells(r, 5).Value) <> "登録" Then ws.Cells(r, 5).Value = "登録"
            Case "見込み認定件数（2次・手入力）"
                If CStr(ws.Cells(r, 5).Value) <> "手で入れた値の引き継ぎ" Then ws.Cells(r, 5).Value = "手で入れた値の引き継ぎ"
        End Select
        If ti >= 0 And k >= 0 Then
            f = "=IF(ISNUMBER($F$4),LET(d,$F$4#,x,XLOOKUP(d,'" & SH_DET & "'!$F$4:$NZ$4,'" & SH_DET & "'!$F$" & DetailTopRow(ti, k) & ":$NZ$" & DetailTopRow(ti, k) & ",""""),IF(d<TODAY(),"""",IF(ISNUMBER(x),x,""""))),"""")"
            If ws.Cells(r, 6).Formula2 <> f Then ws.Cells(r, 6).Formula2 = f
            If k = K_P2 And CStr(ws.Cells(r, 5).Value) <> "進捗（詳細）で2次の時間を入れた日" Then ws.Cells(r, 5).Value = "進捗（詳細）で2次の時間を入れた日"
            If k = K_P1 And CStr(ws.Cells(r, 5).Value) <> "進捗（詳細）で1次の時間を入れた日" Then ws.Cells(r, 5).Value = "進捗（詳細）で1次の時間を入れた日"
        End If
    Next
End Sub

' 起点：昨日までの最新の記録（現在の状況（積み上げ））のステータスの合計
Private Function StartFormula(ByVal lst As String, ByVal st As String) As String
    Dim s As String
    s = "'" & SH_SNAP & "'!"
    StartFormula = "=LET(h," & s & "R4C1:R4C1000,v," & s & "R5C1:R400C1000," & _
        "sel,--(ISNUMBER(MATCH(" & s & "R5C2:R400C2," & lst & ",0))*(" & s & "R5C3:R400C3=""ステータス"")*ISNUMBER(MATCH(" & s & "R5C4:R400C4," & st & ",0)))," & _
        "s,MMULT(TRANSPOSE(sel),IF(ISNUMBER(v),v,0)),n,MMULT(TRANSPOSE(sel),--ISNUMBER(v))," & _
        "XLOOKUP(1,ISNUMBER(h)*(h<TODAY())*(n>0),s,0,0,-1))"
End Function

Private Function SumTopF(ByVal k As Long) As String
    Dim i As Long, s As String
    For i = 0 To 4
        s = s & IIf(i > 0, "+", "") & "N(R" & DetailTopRow(i, k) & "C6)"
    Next
    SumTopF = s
End Function

' 日別処理件数（積み上げ）の 届出×Status の行
Private Function DailyRow(ByVal td As String, ByVal st As String) As Long
    Dim ws As Worksheet, lastR As Long, r As Long
    Set ws = FindSheet(SH_DAILY)
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 5 To lastR
        If CStr(ws.Cells(r, 2).Value) = td And CStr(ws.Cells(r, 3).Value) = st Then DailyRow = r: Exit Function
    Next
End Function

' A1形式の、列だけ相対・行は固定の参照（条件付き書式用）
Private Function ColRef(ByVal c As Long, ByVal r As Long) As String
    ColRef = ThisWorkbook.Worksheets(1).Cells(r, c).Address(True, False)
End Function

'==========================================================
' 並び・キー
'==========================================================
Private Function TodokeList() As Variant
    TodokeList = Array("給与口座", "住居", "通勤", "扶養", "その他")
End Function

' 現在の状況（積み上げ）の申請種別（配列定数）
Private Function TodokeList2(ByVal td As String) As String
    If td = "その他" Then TodokeList2 = "{""児童（認定）"",""児童（改定）""}" Else TodokeList2 = "{""" & td & """}"
End Function

Private Function TodokeDisp(ByVal t As String) As String
    Select Case t
        Case "給与口座", "住居", "通勤", "扶養": TodokeDisp = t
        Case "児童", "その他": TodokeDisp = "その他"
        Case Else: TodokeDisp = ""
    End Select
End Function

Private Function TodokeIdx(ByVal t As String) As Long
    Select Case t
        Case "給与口座": TodokeIdx = 0
        Case "住居": TodokeIdx = 1
        Case "通勤": TodokeIdx = 2
        Case "扶養": TodokeIdx = 3
        Case Else: TodokeIdx = 4
    End Select
End Function

Private Function BlockKey(ByVal td As String, ByVal rs As String, ByVal kt As String) As String
    BlockKey = td & "|" & rs & "|" & kt
End Function

Private Function BKey(ByVal b As Long) As String
    BKey = BlockKey(bTodoke(b), bReason(b), bKotei(b))
End Function

' byReason=True：届出→事由→工程（詳細）、False：届出→工程→事由（スキルマップ・事由比率）
Private Function SortedBlocks(ByVal byReason As Boolean, ByVal onlyShown As Boolean) As Variant
    Dim idx() As Long, sk() As String, n As Long, b As Long, i As Long, j As Long, ti As Long, ts As String, rsk As String
    n = -1
    If nBlk = 0 Then SortedBlocks = Array(): Exit Function
    ReDim idx(0 To nBlk - 1): ReDim sk(0 To nBlk - 1)
    For b = 0 To nBlk - 1
        If (Not onlyShown) Or bShow(b) Then
            n = n + 1: idx(n) = b
            rsk = mReasonSort(bTodoke(b) & "|" & bReason(b)) & "|" & bReason(b)
            If byReason Then
                sk(n) = TodokeIdx(bTodoke(b)) & "|" & rsk & "|" & bKotei(b)
            Else
                sk(n) = TodokeIdx(bTodoke(b)) & "|" & bKotei(b) & "|" & rsk
            End If
        End If
    Next
    If n < 0 Then SortedBlocks = Array(): Exit Function
    For i = 1 To n
        ti = idx(i): ts = sk(i): j = i - 1
        Do While j >= 0
            If sk(j) <= ts Then Exit Do
            idx(j + 1) = idx(j): sk(j + 1) = sk(j): j = j - 1
        Loop
        idx(j + 1) = ti: sk(j + 1) = ts
    Next
    Dim res() As Variant
    ReDim res(0 To n)
    For i = 0 To n: res(i) = idx(i): Next
    SortedBlocks = res
End Function

'==========================================================
' 共通
'==========================================================
Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Private Sub AddNum(d As Object, ByVal k As String, ByVal x As Double)
    If d.Exists(k) Then d(k) = d(k) + x Else d.Add k, x
End Sub

Private Function DictNum(d As Object, ByVal k As String) As Double
    If d.Exists(k) Then DictNum = d(k)
End Function

Private Function SameVal(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim ea As Boolean, eb As Boolean
    ea = IsEmpty(a) Or (CStr(a) = ""): eb = IsEmpty(b) Or (CStr(b) = "")
    If ea Or eb Then SameVal = (ea And eb): Exit Function
    If IsNumeric(a) And IsNumeric(b) Then
        SameVal = (Abs(CDbl(a) - CDbl(b)) < 0.005)
    Else
        SameVal = (CStr(a) = CStr(b))
    End If
End Function

Private Function IsNoFlag(ByVal x As Variant) As Boolean
    If IsEmpty(x) Then IsNoFlag = True: Exit Function
    IsNoFlag = (Trim$(CStr(x)) = "" Or UCase$(Trim$(CStr(x))) = "NULL")
End Function

Private Function HeadIdx(v As Variant, ByVal nC As Long, ByVal nm As String) As Long
    Dim c As Long
    For c = 1 To nC
        If Trim$(CStr(v(1, c))) = nm Then HeadIdx = c: Exit Function
    Next
End Function

Private Function RangeArr(rg As Range) As Variant
    Dim a() As Variant
    If rg.Cells.Count = 1 Then
        ReDim a(1 To 1, 1 To 1): a(1, 1) = rg.Value: RangeArr = a
    Else
        RangeArr = rg.Value
    End If
End Function

Private Function NameRange(ByVal nm As String) As Boolean
    Dim rg As Range
    On Error Resume Next
    Set rg = ThisWorkbook.Names(nm).RefersToRange
    On Error GoTo 0
    NameRange = Not (rg Is Nothing)
End Function

Private Sub SetName(ByVal nm As String, rg As Range)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & rg.Worksheet.Name & "'!" & rg.Address
End Sub

Private Function SortedKeys(d As Object) As Variant
    Dim arr() As String, k As Variant, n As Long, i As Long, j As Long, t As String
    If d.Count = 0 Then SortedKeys = Array(): Exit Function
    ReDim arr(0 To d.Count - 1)
    For Each k In d.keys: arr(n) = CStr(k): n = n + 1: Next
    For i = 1 To n - 1
        t = arr(i): j = i - 1
        Do While j >= 0
            If arr(j) <= t Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = t
    Next
    SortedKeys = arr
End Function

' "00:07:18" や時刻値を分に直す
Private Function ToMinutes(ByVal t As Variant) As Double
    Dim p As Variant
    If IsEmpty(t) Then Exit Function
    If VarType(t) = vbDate Or VarType(t) = vbDouble Then ToMinutes = CDbl(t) * 1440: Exit Function
    p = Split(CStr(t), ":")
    If UBound(p) = 2 Then
        ToMinutes = Val(p(0)) * 60 + Val(p(1)) + Val(p(2)) / 60
    ElseIf UBound(p) = 1 Then
        ToMinutes = Val(p(0)) * 60 + Val(p(1))
    End If
End Function

Private Function FindSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nm Then Set FindSheet = ws: Exit Function
    Next
End Function

' なければ作る（詳細・スキルマップ・事由比率は簡易版の後ろ、計算用_割当は最後に非表示で）
Private Function GetSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet, after As Worksheet
    Set ws = FindSheet(nm)
    If ws Is Nothing Then
        Select Case nm
            Case SH_DET: Set after = FindSheet(SH_SIMPLE)
            Case SH_SKL: Set after = FindSheet(SH_DET)
            Case SH_RAT: Set after = FindSheet(SH_SKL)
        End Select
        If after Is Nothing Then Set after = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set ws = ThisWorkbook.Worksheets.Add(after:=after)
        ws.Name = nm
        If nm = SH_INP Then ws.Visible = xlSheetHidden
    End If
    Set GetSheet = ws
End Function

Private Sub ClearSheet(ws As Worksheet)
    ws.Cells.Hyperlinks.Delete
    ws.Cells.ClearOutline
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Rows.Hidden = False
    ws.Columns.Hidden = False
    ws.Rows.RowHeight = ws.StandardHeight
End Sub

Private Sub LinkCell(ws As Worksheet, rg As Range, ByVal sheetName As String, ByVal caption As String)
    ws.Hyperlinks.Add Anchor:=rg, Address:="", SubAddress:="'" & sheetName & "'!A1", TextToDisplay:=ChrW(&H25B6) & " " & caption
    rg.Font.Size = 9
End Sub

Private Sub SectionStyle(rg As Range)
    With rg
        .Interior.Color = CLR_PURPLE: .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 10
    End With
    rg.Worksheet.Rows(rg.Row).RowHeight = 20
End Sub

Private Sub HeadStyle(rg As Range)
    With rg
        .Font.Bold = True: .Interior.Color = RGB(221, 235, 247): .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(191, 191, 191)
    End With
End Sub

Private Sub GridStyle(rg As Range)
    With rg.Borders
        .LineStyle = xlContinuous: .Color = RGB(217, 217, 217)
    End With
End Sub

