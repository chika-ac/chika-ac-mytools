Option Explicit

'==========================================================
' 稼働実績
'   稼働記録（1人1ファイルを蓄積）を、日付ごとに
'   届出×工程、業務ID×業務名 で集計する（全員分＋個人別）。
'   業務ID→届出・工程・集計区分・事由 は「業務マスタ」で管理する（事由は進捗（詳細）で使う）。
'==========================================================

Private Const WORK_SRC As String = "稼働記録_検索結果_出力データ"
Private Const SH_OUT As String = "稼働実績"
Private Const SH_MST As String = "業務マスタ"
Private Const OUT_HDR As Long = 4       ' 稼働実績：日付の見出し行
Private Const OUT_FIRST As Long = 7     ' 稼働実績：最初の日付の列（B区分 C届出/業務ID D工程/業務名 E指標 F期間計）
Private Const MST_HDR As Long = 4       ' 業務マスタ：見出し行
Private Const ALL_WHO As String = "全員"
Private Const KIND_CASE As String = "案件処理"

' 取込側から、稼働記録の取込先シート名を確かめるのに使う
Public Function WorkSourceSheet() As String
    WorkSourceSheet = WORK_SRC
End Function

'----------------------------------------------------------
' 操作シートのボタンから呼ぶ
'----------------------------------------------------------
Public Sub 稼働実績を更新()
    MsgBox RefreshWork(), vbInformation, "稼働実績"
End Sub

Public Function RefreshWork() As String
    Dim calcMode As XlCalculation
    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo Fin
    RefreshWork = BuildWork()
Fin:
    Application.EnableEvents = True
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Function

Private Function BuildWork() As String
    Dim src As Worksheet, v As Variant, nR As Long, nC As Long, r As Long
    Dim colEid As Long, colDate As Long, colId As Long, colName As Long, colCnt As Long, colTime As Long
    Dim mst As Object, newIds As String, info As Variant
    Dim agg As Object, keys As Object, people As Object
    Dim who As Variant, d As Long, dMin As Long, dMax As Long
    Dim id As String, nm As String, cnt As Double, mins As Double
    Dim kA As String, kB As String

    Set src = FindSheet(WORK_SRC)
    If src Is Nothing Then BuildWork = "稼働記録がまだ取り込まれていません。": Exit Function
    nR = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    nC = src.Cells(1, src.Columns.Count).End(xlToLeft).Column
    If nR < 2 Then BuildWork = "稼働記録にデータがありません。": Exit Function

    Application.ScreenUpdating = False
    v = src.Range("A1").Resize(nR, nC).Value
    colEid = NeedCol(v, nC, "EID"): colDate = NeedCol(v, nC, "日付"): colId = NeedCol(v, nC, "業務ID")
    colName = NeedCol(v, nC, "業務名"): colCnt = NeedCol(v, nC, "処理案件数"): colTime = NeedCol(v, nC, "所要時間（分）_")

    ' 業務マスタ（なければ作る。新しい業務IDは仮の分類で追加）
    Set mst = LoadMaster(v, nR, colId, colName, colDate, newIds)

    ' 集計：agg(担当|段|キー) → Dictionary(日付 → Array(件数, 分))
    Set agg = CreateObject("Scripting.Dictionary")
    Set keys = CreateObject("Scripting.Dictionary")     ' 担当|段|キー → 並び順の文字列
    Set people = CreateObject("Scripting.Dictionary")
    dMin = 0: dMax = 0
    For r = 2 To nR
        If IsDate(v(r, colDate)) And Len(CStr(v(r, colId))) > 0 Then
            d = CLng(Int(CDbl(CDate(v(r, colDate)))))
            If dMin = 0 Or d < dMin Then dMin = d
            If d > dMax Then dMax = d
            id = CStr(v(r, colId)): nm = CStr(v(r, colName))
            cnt = Val(CStr(v(r, colCnt)))
            mins = ToMinutes(v(r, colTime))
            info = mst(id)    ' Array(業務名, 届出, 工程, 区分)
            kA = info(3) & "|" & info(1) & "|" & info(2)
            kB = info(3) & "|" & id & "|" & nm
            people(CStr(v(r, colEid))) = True
            For Each who In Array(ALL_WHO, CStr(v(r, colEid)))
                AddAgg agg, keys, who & "|A|" & kA, SortKeyA(info), d, cnt, mins
                AddAgg agg, keys, who & "|B|" & kB, SortKeyB(info, id), d, cnt, mins
            Next
        End If
    Next
    If dMin = 0 Then Application.ScreenUpdating = True: BuildWork = "稼働記録に日付のある行がありません。": Exit Function

    WriteReport agg, keys, people, dMin, dMax
    Application.ScreenUpdating = True
    BuildWork = "稼働実績を更新しました（" & Format(CDate(dMin), "yyyy/m/d") & "?" & Format(CDate(dMax), "yyyy/m/d") & "、" & people.Count & "人）。" & _
                  IIf(Len(newIds) > 0, vbLf & "業務マスタに新しい業務IDを追加しました（届出・工程を確認してください）：" & newIds, "")
End Function

Private Sub AddAgg(agg As Object, keys As Object, ByVal k As String, ByVal sk As String, ByVal d As Long, ByVal cnt As Double, ByVal mins As Double)
    Dim byDay As Object, a As Variant
    If Not agg.Exists(k) Then
        Set byDay = CreateObject("Scripting.Dictionary")
        agg.Add k, byDay
        keys(k) = sk
    End If
    Set byDay = agg(k)
    If byDay.Exists(d) Then a = byDay(d) Else a = Array(0#, 0#)
    a(0) = a(0) + cnt: a(1) = a(1) + mins
    byDay(d) = a
End Sub

'----------------------------------------------------------
' 稼働実績シートを書き直す
'----------------------------------------------------------
Private Sub WriteReport(agg As Object, keys As Object, people As Object, ByVal dMin As Long, ByVal dMax As Long)
    Dim ws As Worksheet, nDays As Long, nCols As Long, whoList As Variant, w As Long
    Dim rowsNeeded As Long, outArr() As Variant, rowKind() As String, o As Long
    Dim lvl As Variant, ks As Variant, i As Long, k As String, parts As Variant
    Dim byDay As Object, dd As Long, c As Long, a As Variant, tc As Double, tm As Double
    Dim metric As Long, isCase As Boolean, lastR As Long, whoName As String
    Dim idxTitle As Long, linkRow() As Long, secRow() As Long, st As Variant

    Set ws = FindSheet(SH_OUT)
    If ws Is Nothing Then Err.Raise vbObjectError + 1, , "「" & SH_OUT & "」シートがありません"
    nDays = dMax - dMin + 1
    nCols = OUT_FIRST - 1 + nDays

    whoList = SortedPeople(people)
    rowsNeeded = 1 + (UBound(whoList) + 2) * 4 + agg.Count * 4 + 12
    ReDim linkRow(-1 To UBound(whoList) + 1): ReDim secRow(-1 To UBound(whoList) + 1)
    ReDim outArr(1 To rowsNeeded, 1 To nCols)
    ReDim rowKind(1 To rowsNeeded)

    ' 見出し
    outArr(1, 2) = "区分": outArr(1, 3) = "届出 / 業務ID": outArr(1, 4) = "工程 / 業務名": outArr(1, 5) = "指標": outArr(1, 6) = "期間計"
    For dd = 0 To nDays - 1: outArr(1, OUT_FIRST + dd) = CDate(dMin + dd): Next
    rowKind(1) = "H"
    o = 1

    ' 担当者一覧（名前をクリックするとその人の表へ）
    o = o + 1: idxTitle = o
    outArr(o, 2) = "■ 担当者一覧（名前をクリックすると、その人の表へ移動します）"
    rowKind(o) = "I"
    For w = -1 To UBound(whoList)
        o = o + 1
        If w = -1 Then whoName = ALL_WHO Else whoName = whoList(w)
        st = PersonStats(agg, whoName)     ' Array(日数, 初日, 最終日, 案件件数, 案件分)
        outArr(o, 2) = WhoDisp(whoName)
        outArr(o, 3) = "記録 " & st(0) & "日"
        If st(1) > 0 Then outArr(o, 4) = "期間 " & Format(CDate(st(1)), "yyyy/m/d") & "?" & Format(CDate(st(2)), "yyyy/m/d")
        outArr(o, 5) = "件/時間(案件)"
        If st(4) > 0 Then outArr(o, 6) = Round(st(3) / (st(4) / 60), 2)
        rowKind(o) = "L"
        linkRow(w) = o
    Next

    For w = -1 To UBound(whoList)
        o = o + 1
        If w = -1 Then
            whoName = ALL_WHO
        Else
            whoName = whoList(w)
        End If
        outArr(o, 2) = "■ " & WhoDisp(whoName)
        rowKind(o) = "P"
        secRow(w) = o
        For Each lvl In Array("A", "B")
            o = o + 1
            outArr(o, 2) = IIf(lvl = "A", "　届出 × 工程", "　業務別（詳細）")
            rowKind(o) = "S"
            ks = SortedKeysFor(keys, whoName & "|" & lvl & "|")
            For i = 0 To UBound(ks)
                k = ks(i)
                parts = Split(k, "|")      ' 担当|段|区分|届出or業務ID|工程or業務名
                Set byDay = agg(k)
                isCase = (parts(2) = KIND_CASE)
                For metric = 0 To 3
                    o = o + 1
                    outArr(o, 2) = parts(2): outArr(o, 3) = parts(3): outArr(o, 4) = parts(4)
                    outArr(o, 5) = MetricName(metric)
                    rowKind(o) = "M" & metric
                    tc = 0: tm = 0
                    For dd = 0 To nDays - 1
                        If byDay.Exists(dMin + dd) Then
                            a = byDay(dMin + dd)
                            tc = tc + a(0): tm = tm + a(1)
                            outArr(o, OUT_FIRST + dd) = MetricValue(metric, a(0), a(1), isCase)
                        End If
                    Next
                    outArr(o, 6) = MetricValue(metric, tc, tm, isCase)
                Next
            Next
        Next
    Next

    ' 書き出し
    ws.Range(ws.Cells(OUT_HDR, 1), ws.Cells(ws.Rows.Count, ws.Columns.Count)).Hyperlinks.Delete
    ws.Range(ws.Cells(OUT_HDR, 1), ws.Cells(ws.Rows.Count, ws.Columns.Count)).Clear
    ws.Cells(OUT_HDR, 1).Resize(o, nCols).Value = outArr
    ws.Range("B2").Value = "元データ：" & WORK_SRC & "　最終更新 " & Format(Now, "yyyy/mm/dd hh:nn") & _
                           "　期間 " & Format(CDate(dMin), "yyyy/m/d") & "?" & Format(CDate(dMax), "yyyy/m/d")
    lastR = OUT_HDR + o - 1
    FormatReport ws, rowKind, o, nCols, lastR

    ' 一覧 → 各人の表、各人の表 → 一覧 のリンク
    For w = -1 To UBound(whoList)
        If w = -1 Then whoName = ALL_WHO Else whoName = whoList(w)
        ws.Hyperlinks.Add Anchor:=ws.Cells(OUT_HDR + linkRow(w) - 1, 2), Address:="", _
            SubAddress:="'" & SH_OUT & "'!B" & (OUT_HDR + secRow(w) - 1), TextToDisplay:=WhoDisp(whoName)
        ws.Hyperlinks.Add Anchor:=ws.Cells(OUT_HDR + secRow(w) - 1, 5), Address:="", _
            SubAddress:="'" & SH_OUT & "'!B" & (OUT_HDR + idxTitle - 1), TextToDisplay:=ChrW(&H25B2) & " 一覧へ戻る"
        With ws.Cells(OUT_HDR + secRow(w) - 1, 5).Font
            .Color = vbWhite: .Size = 9: .Bold = False
        End With
    Next
End Sub

' 担当者の表示名（スキルマップの表示名 → 氏名 → EID）
Private Function WhoDisp(ByVal whoName As String) As String
    If whoName = ALL_WHO Then WhoDisp = ALL_WHO Else WhoDisp = PersonDisplayName(whoName)
End Function

' 担当者ごとの 記録日数・期間・案件処理の件数と分
Private Function PersonStats(agg As Object, ByVal whoName As String) As Variant
    Dim k As Variant, days As Object, d As Variant, a As Variant, firstD As Long, lastD As Long
    Dim cnt As Double, mins As Double, prefix As String, byDay As Object
    Set days = CreateObject("Scripting.Dictionary")
    prefix = whoName & "|A|"
    For Each k In agg.keys
        If Left$(k, Len(prefix)) = prefix Then
            Set byDay = agg(k)
            For Each d In byDay.keys
                days(d) = True
                If firstD = 0 Or d < firstD Then firstD = d
                If d > lastD Then lastD = d
                If Mid$(k, Len(prefix) + 1, Len(KIND_CASE) + 1) = KIND_CASE & "|" Then
                    a = byDay(d)
                    cnt = cnt + a(0): mins = mins + a(1)
                End If
            Next
        End If
    Next
    PersonStats = Array(days.Count, firstD, lastD, cnt, mins)
End Function

Private Function MetricName(ByVal metric As Long) As String
    Select Case metric
        Case 0: MetricName = "処理件数"
        Case 1: MetricName = "所要時間（分）"
        Case 2: MetricName = "件/時間"
        Case Else: MetricName = "1件あたり分"
    End Select
End Function

Private Function MetricValue(ByVal metric As Long, ByVal cnt As Double, ByVal mins As Double, ByVal isCase As Boolean) As Variant
    Select Case metric
        Case 0: MetricValue = cnt
        Case 1: MetricValue = Round(mins, 1)
        Case 2: If isCase And mins > 0 Then MetricValue = Round(cnt / (mins / 60), 2) Else MetricValue = Empty
        Case 3: If isCase And cnt > 0 Then MetricValue = Round(mins / cnt, 2) Else MetricValue = Empty
    End Select
End Function

Private Sub FormatReport(ws As Worksheet, rowKind() As String, ByVal n As Long, ByVal nCols As Long, ByVal lastR As Long)
    Dim i As Long, r As Long, rg As Range
    With ws.Range(ws.Cells(OUT_HDR, 2), ws.Cells(lastR, nCols))
        .Font.Name = "Meiryo UI": .Font.Size = 10
    End With
    For i = 1 To n
        r = OUT_HDR + i - 1
        Set rg = ws.Range(ws.Cells(r, 2), ws.Cells(r, nCols))
        Select Case rowKind(i)
            Case "H"
                rg.Font.Bold = True: rg.Interior.Color = RGB(221, 235, 247): rg.HorizontalAlignment = xlCenter
                ws.Range(ws.Cells(r, OUT_FIRST), ws.Cells(r, nCols)).NumberFormat = "m/d(aaa)"
            Case "P"
                rg.Interior.Color = RGB(0, 112, 112): rg.Font.Color = vbWhite: rg.Font.Bold = True: rg.Font.Size = 12
                ws.Rows(r).RowHeight = 22
            Case "S"
                rg.Interior.Color = RGB(218, 238, 238): rg.Font.Bold = True: rg.Font.Color = RGB(0, 90, 90)
            Case "I"
                rg.Interior.Color = RGB(255, 242, 204): rg.Font.Bold = True: rg.Font.Color = RGB(0, 90, 90): rg.Font.Size = 11
                ws.Rows(r).RowHeight = 22
            Case "L"
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Borders.LineStyle = xlContinuous
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Borders.Color = RGB(217, 217, 217)
                ws.Cells(r, 6).NumberFormat = "0.0"
            Case "M0", "M1", "M2", "M3"
                rg.Borders.LineStyle = xlContinuous: rg.Borders.Color = RGB(217, 217, 217)
                ws.Range(ws.Cells(r, 6), ws.Cells(r, nCols)).NumberFormat = IIf(rowKind(i) = "M0", "#,##0;-#,##0;""""", IIf(rowKind(i) = "M1", "#,##0;-#,##0;""""", "0.0"))
                If rowKind(i) = "M0" Then
                    rg.Borders(xlEdgeTop).Weight = xlMedium: rg.Borders(xlEdgeTop).Color = RGB(128, 128, 128)
                Else
                    ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Font.Color = RGB(191, 191, 191)
                End If
                If rowKind(i) = "M2" Then ws.Range(ws.Cells(r, 5), ws.Cells(r, nCols)).Font.Bold = True
        End Select
    Next
    ws.Range(ws.Cells(OUT_HDR + 1, 6), ws.Cells(lastR, 6)).Font.Bold = True
    ws.Range(ws.Cells(OUT_HDR + 1, 6), ws.Cells(lastR, 6)).Interior.Color = RGB(242, 242, 242)
    ws.Range(ws.Cells(OUT_HDR, OUT_FIRST), ws.Cells(lastR, nCols)).ColumnWidth = 9
    ShadeRestDays ws, OUT_HDR, OUT_FIRST, lastR
    HighlightToday ws, OUT_HDR, OUT_FIRST, lastR, nCols
End Sub

'----------------------------------------------------------
' 業務マスタ
'----------------------------------------------------------
Private Function LoadMaster(v As Variant, ByVal nR As Long, ByVal colId As Long, ByVal colName As Long, ByVal colDate As Long, ByRef newIds As String) As Object
    Dim ws As Worksheet, d As Object, r As Long, lastR As Long, id As String, nm As String, g As Variant
    Dim firstD As Object, lastD As Object, dt As Long, nNew As Long

    Set ws = FindSheet(SH_MST)
    If ws Is Nothing Then Err.Raise vbObjectError + 1, , "「" & SH_MST & "」シートがありません"
    Set d = CreateObject("Scripting.Dictionary")
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = MST_HDR + 1 To lastR
        id = Trim$(CStr(ws.Cells(r, 2).Value))
        If Len(id) > 0 Then d(id) = r
    Next

    ' 新しい業務IDを追加（届出・区分は業務名から仮に入れる）
    Set firstD = CreateObject("Scripting.Dictionary"): Set lastD = CreateObject("Scripting.Dictionary")
    For r = 2 To nR
        id = Trim$(CStr(v(r, colId)))
        If Len(id) > 0 Then
            If Not d.Exists(id) Then
                lastR = lastR + 1
                If lastR <= MST_HDR Then lastR = MST_HDR + 1
                nm = CStr(v(r, colName))
                g = GuessMaster(id, nm)
                ws.Cells(lastR, 2).NumberFormat = "@"
                ws.Cells(lastR, 2).Value = id
                ws.Cells(lastR, 3).Value = nm
                ws.Cells(lastR, 4).Value = g(0)
                ws.Cells(lastR, 5).Value = g(1)
                ws.Cells(lastR, 6).Value = g(2)
                ws.Cells(lastR, 8).Value = "自動追加 " & Format(Date, "yyyy/m/d")
                If g(2) = KIND_CASE Then ws.Cells(lastR, 11).Value = ReasonFromName(nm)
                d(id) = lastR
                nNew = nNew + 1
                If nNew <= 10 Then newIds = newIds & IIf(Len(newIds) > 0, "、", "") & id
            End If
            If IsDate(v(r, colDate)) Then
                dt = CLng(Int(CDbl(CDate(v(r, colDate)))))
                If Not firstD.Exists(id) Then firstD(id) = dt: lastD(id) = dt
                If dt < firstD(id) Then firstD(id) = dt
                If dt > lastD(id) Then lastD(id) = dt
            End If
        End If
    Next
    If nNew > 10 Then newIds = newIds & " ほか（計" & nNew & "件）"
    If nNew > 0 Then
        FormatMaster ws
        d.RemoveAll
        lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        For r = MST_HDR + 1 To lastR
            id = Trim$(CStr(ws.Cells(r, 2).Value))
            If Len(id) > 0 Then d(id) = r
        Next
    End If

    ' 出現期間を更新し、ID → Array(業務名, 届出, 工程, 区分) を返す
    Dim res As Object, k As Variant, todoke As String, kotei As String, kubun As String
    Set res = CreateObject("Scripting.Dictionary")
    For Each k In d.keys
        r = d(k)
        If firstD.Exists(k) Then
            ws.Cells(r, 9).Value = CDate(firstD(k)): ws.Cells(r, 10).Value = CDate(lastD(k))
        End If
        todoke = Trim$(CStr(ws.Cells(r, 4).Value)): If Len(todoke) = 0 Then todoke = "未設定"
        kotei = Trim$(CStr(ws.Cells(r, 5).Value)): If Len(kotei) = 0 Then kotei = "未設定"
        kubun = Trim$(CStr(ws.Cells(r, 6).Value)): If Len(kubun) = 0 Then kubun = "その他"
        res(k) = Array(CStr(ws.Cells(r, 3).Value), todoke, kotei, kubun)
    Next
    ws.Range(ws.Cells(MST_HDR + 1, 9), ws.Cells(ws.Cells(ws.Rows.Count, 2).End(xlUp).Row, 10)).NumberFormat = "yyyy/mm/dd"
    Set LoadMaster = res
End Function

' 業務名から 届出・工程・集計区分 を仮決め（工程は案件処理なら「未設定」＝入力してもらう）
Private Function GuessMaster(ByVal id As String, ByVal nm As String) As Variant
    Dim kubun As String, todoke As String, kotei As String
    Select Case UCase$(Left$(id, 2))
        Case "C1": kubun = KIND_CASE
        Case "C2": kubun = "間接作業"
        Case "E2", "E1": kubun = "社内業務"
        Case Else
            If UCase$(Left$(id, 1)) = "D" Then kubun = "休憩" Else kubun = "その他"
    End Select
    If kubun = KIND_CASE Then
        If InStr(nm, "通勤") > 0 Then
            todoke = "通勤"
        ElseIf InStr(nm, "給与口座") > 0 Then
            todoke = "給与口座"
        ElseIf InStr(nm, "認定請求") > 0 Or InStr(nm, "額改定") > 0 Or InStr(nm, "消滅") > 0 Or InStr(nm, "児童") > 0 Then
            todoke = "児童"
        ElseIf InStr(nm, "転居") > 0 Or InStr(nm, "住居") > 0 Or InStr(nm, "家賃") > 0 Or InStr(nm, "別棟") > 0 Or InStr(nm, "配同") > 0 Or InStr(nm, "支給額区分") > 0 Then
            todoke = "住居"
        ElseIf InStr(nm, "【増】") > 0 Or InStr(nm, "【減】") > 0 Or InStr(nm, "扶養") > 0 Then
            todoke = "扶養"
        Else
            todoke = "未設定"
        End If
        kotei = "未設定"
    Else
        todoke = "?": kotei = "?"
    End If
    GuessMaster = Array(todoke, kotei, kubun)
End Function

Private Sub FormatMaster(ws As Worksheet)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastR <= MST_HDR Then Exit Sub
    With ws.Range(ws.Cells(MST_HDR + 1, 2), ws.Cells(lastR, 11))
        .Font.Name = "Meiryo UI": .Font.Size = 10
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(191, 191, 191)
    End With
    ws.Range(ws.Cells(MST_HDR + 1, 4), ws.Cells(lastR, 7)).Interior.Color = RGB(255, 242, 204)
    ws.Range(ws.Cells(MST_HDR + 1, 11), ws.Cells(lastR, 11)).Interior.Color = RGB(255, 242, 204)
    ' 業務ID順に並べる
    ws.Range(ws.Cells(MST_HDR + 1, 2), ws.Cells(lastR, 11)).Sort Key1:=ws.Cells(MST_HDR + 1, 2), Order1:=xlAscending, Header:=xlNo
End Sub

'----------------------------------------------------------
' 並び順
'----------------------------------------------------------
Private Function SortKeyA(info As Variant) As String
    SortKeyA = KindOrder(info(3)) & Format(TodokeOrder(info(1)), "00") & KoteiOrder(info(2)) & info(1) & info(2)
End Function

Private Function SortKeyB(info As Variant, ByVal id As String) As String
    SortKeyB = KindOrder(info(3)) & Format(TodokeOrder(info(1)), "00") & KoteiOrder(info(2)) & id
End Function

Private Function KindOrder(ByVal k As String) As String
    Select Case k
        Case KIND_CASE: KindOrder = "1"
        Case "間接作業": KindOrder = "2"
        Case "社内業務": KindOrder = "3"
        Case "休憩": KindOrder = "4"
        Case Else: KindOrder = "5"
    End Select
End Function

Private Function TodokeOrder(ByVal t As String) As Long
    Select Case t
        Case "給与口座": TodokeOrder = 1
        Case "住居": TodokeOrder = 2
        Case "通勤": TodokeOrder = 3
        Case "扶養": TodokeOrder = 4
        Case "児童": TodokeOrder = 5
        Case "未設定": TodokeOrder = 98
        Case "?": TodokeOrder = 99
        Case Else: TodokeOrder = 50
    End Select
End Function

Private Function KoteiOrder(ByVal t As String) As String
    Select Case t
        Case "1次": KoteiOrder = "1"
        Case "2次": KoteiOrder = "2"
        Case "3次": KoteiOrder = "3"
        Case "未設定": KoteiOrder = "8"
        Case "?": KoteiOrder = "9"
        Case Else: KoteiOrder = "5"
    End Select
End Function

Private Function SortedKeysFor(keys As Object, ByVal prefix As String) As Variant
    Dim arr() As String, sk() As String, k As Variant, n As Long, i As Long, j As Long, t As String, ts As String
    n = -1
    ReDim arr(0 To Application.Max(0, keys.Count - 1)): ReDim sk(0 To Application.Max(0, keys.Count - 1))
    For Each k In keys.keys
        If Left$(k, Len(prefix)) = prefix Then n = n + 1: arr(n) = k: sk(n) = keys(k)
    Next
    If n < 0 Then SortedKeysFor = Array(): Exit Function
    ReDim Preserve arr(0 To n): ReDim Preserve sk(0 To n)
    For i = 1 To n
        t = arr(i): ts = sk(i): j = i - 1
        Do While j >= 0
            If sk(j) <= ts Then Exit Do
            arr(j + 1) = arr(j): sk(j + 1) = sk(j): j = j - 1
        Loop
        arr(j + 1) = t: sk(j + 1) = ts
    Next
    SortedKeysFor = arr
End Function

Private Function SortedPeople(people As Object) As Variant
    Dim arr() As String, k As Variant, n As Long, i As Long, j As Long, t As String
    If people.Count = 0 Then SortedPeople = Array(): Exit Function
    ReDim arr(0 To people.Count - 1)
    For Each k In people.keys: arr(n) = k: n = n + 1: Next
    For i = 1 To n - 1
        t = arr(i): j = i - 1
        Do While j >= 0
            If arr(j) <= t Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = t
    Next
    SortedPeople = arr
End Function

'----------------------------------------------------------
' 共通
'----------------------------------------------------------
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

Private Function NeedCol(v As Variant, ByVal nC As Long, ByVal nm As String) As Long
    Dim c As Long
    For c = 1 To nC
        If Trim$(CStr(v(1, c))) = nm Then NeedCol = c: Exit Function
    Next
    Err.Raise vbObjectError + 1, , WORK_SRC & " に「" & nm & "」列がありません"
End Function

Private Function FindSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nm Then Set FindSheet = ws: Exit Function
    Next
End Function

