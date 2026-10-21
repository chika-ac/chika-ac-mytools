Option Explicit

'==========================================================
' 現在の状況の記録
'   「現在の状況」シートの件数を、今日の日付で「推移」シート（数字で書き込み）と
'   「計算用_状況記録」シート（縦長の控え。同じ日の記録があるかの確認にも使う）に残す。
'   1日1回分。同じ日にもう一度記録すると上書き。
'==========================================================

Private Const SH_NOW As String = "現在の状況"
Private Const SH_REC As String = "計算用_状況記録"   ' 計算用：同じ日の記録があるかの確認に使う縦長の控え
Private Const NOW_GROUP_ROW As Long = 4     ' 現在の状況：グループ見出し（ステータス別／フラグ付き）
Private Const NOW_HEADER_ROW As Long = 5    ' 現在の状況：項目の見出し
Private Const REC_COLS As Long = 7          ' 記録日, 記録日時, 申請種別, 区分, 項目, 件数, 並び順
Private Const SH_TREND As String = "現在の状況（積み上げ）"
Private Const TR_HEADER_ROW As Long = 4     ' 推移：日付の見出し行
Private Const TR_ITEM_COL As Long = 4       ' 推移：項目の列（B=申請種別, C=区分, D=項目）
Private Const TR_FIRST_COL As Long = 5      ' 推移：最初の日付の列
Private Const GRP_OLDEST As String = "最古"   ' 区分：確認中の最古申請日
Private Const INQ_TYPE As String = "問合せ"                 ' 積み上げ：問合せの行の申請種別欄
Private Const INQ_TITLE As String = "■ 前営業日の問合せ"     ' 現在の状況：問合せの表の見出し（前方一致）
Private Const INQ_NONE As String = "前営業日の問合せはありません"
Private Const INQ_TOTAL_GRP As String = "新規受付"
Private Const SH_HOL As String = "休日一覧"

'----------------------------------------------------------
' 操作シートのボタンから呼ぶ
'----------------------------------------------------------
Public Sub 状況を記録()
    Dim lastImp As Variant, ans As VbMsgBoxResult, n As Long, nDel As Long, inqMsg As String

    lastImp = Application.Evaluate("MAX(給与_取込日時)")
    If IsError(lastImp) Then lastImp = 0
    If lastImp = 0 Then
        MsgBox "給与手当集計がまだ取り込まれていません。先に②データ取込をしてください。", vbExclamation, "記録する"
        Exit Sub
    End If
    If Int(lastImp) < Date Then
        ans = MsgBox("給与手当集計の最終取込は " & Format(lastImp, "m/d hh:nn") & " です。" & vbLf & _
                     "今日はまだ取り込んでいませんが、このまま今日（" & Format(Date, "m/d") & "）の記録にしますか？", _
                     vbYesNo + vbExclamation, "記録する")
        If ans <> vbYes Then Exit Sub
    End If
    If CountRecords(Date) > 0 Then
        ans = MsgBox("今日（" & Format(Date, "m/d") & "）の記録はすでにあります。上書きしますか？", vbYesNo + vbQuestion, "記録する")
        If ans <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False
    n = RecordSnapshot(Date, nDel, inqMsg)
    Application.ScreenUpdating = True
    MsgBox Format(Date, "yyyy/mm/dd") & " の状況を記録しました（" & n & "行" & _
           IIf(nDel > 0, "、前の記録を上書き", "") & "）。" & vbLf & "「" & SH_TREND & "」シートの " & Format(Date, "m/d") & " の列に数字で書き込みました。" & _
           IIf(Len(inqMsg) > 0, vbLf & inqMsg, ""), _
           vbInformation, "記録する"
End Sub

' 積み上げシートの見た目だけ整え直す
Public Sub 積み上げの書式を整える()
    Application.ScreenUpdating = False
    FormatTrend ThisWorkbook.Worksheets(SH_TREND)
    Application.ScreenUpdating = True
End Sub

'----------------------------------------------------------
' 記録本体：書き込んだ行数を返す
'----------------------------------------------------------
Public Function RecordSnapshot(ByVal recDate As Date, Optional ByRef nDel As Long, Optional ByRef inqMsg As String) As Long
    Dim wsN As Worksheet, wsR As Worksheet, rTotal As Long, lastC As Long
    Dim r As Long, c As Long, n As Long, item As String, grp As String
    Dim newRows() As Variant, nowTime As Date

    Set wsN = ThisWorkbook.Worksheets(SH_NOW)
    Set wsR = ThisWorkbook.Worksheets(SH_REC)
    Application.Calculate

    rTotal = FindInCol(wsN, 2, "合計")
    If rTotal = 0 Then Err.Raise vbObjectError + 1, , "現在の状況シートに「合計」行が見つかりません"
    lastC = wsN.Cells(NOW_HEADER_ROW, wsN.Columns.Count).End(xlToLeft).Column

    ReDim newRows(1 To REC_COLS, 1 To (rTotal - NOW_HEADER_ROW) * lastC)
    nowTime = Now
    For r = NOW_HEADER_ROW + 1 To rTotal
        For c = 3 To lastC
            item = Trim$(CStr(wsN.Cells(NOW_HEADER_ROW, c).Value))
            If Len(item) > 0 Then
                grp = CStr(wsN.Cells(NOW_GROUP_ROW, c).MergeArea.Cells(1, 1).Value)
                n = n + 1
                newRows(1, n) = recDate
                newRows(2, n) = nowTime
                newRows(3, n) = CStr(wsN.Cells(r, 2).Value)
                newRows(4, n) = GroupOf(grp)
                newRows(5, n) = item
                newRows(6, n) = wsN.Cells(r, c).Value
                newRows(7, n) = (r - NOW_HEADER_ROW) * 1000 + c
            End If
        Next
    Next

    nDel = DeleteRecords(wsR, recDate)
    WriteRows wsR, newRows, n
    WriteTrend newRows, n, recDate
    inqMsg = RecordInquiry(wsN, recDate)
    RecordSnapshot = n
End Function

'----------------------------------------------------------
' 前営業日の問合せ件数（現在の状況の下の表）を、積み上げの受付日の列に書き込む
'   行：問合せ｜業務チーム｜業務分類詳細（先頭に 問合せ｜新規受付｜計）
'   その日に件数がない分類は 0。結果のメッセージを返す（表がなければ空）
'----------------------------------------------------------
Private Function RecordInquiry(wsN As Worksheet, ByVal recDate As Date) As String
    Dim ws As Worksheet, r0 As Long, hr As Long, cCat As Long, cCnt As Long, c As Long
    Dim r As Long, d As Date, col As Long, rowMap As Object, k As String, total As Double
    Dim teams As Object, cats As Object, cnts As Object, i As Long, team As String

    r0 = FindPrefixInCol(wsN, 2, INQ_TITLE)
    If r0 = 0 Then Exit Function
    hr = r0 + 2
    For c = 2 To 20
        Select Case Trim$(CStr(wsN.Cells(hr, c).Value))
            Case "業務分類詳細": cCat = c
            Case "件数": cCnt = c
        End Select
    Next
    If Trim$(CStr(wsN.Cells(hr, 2).Value)) <> "業務チーム" Or cCat = 0 Or cCnt = 0 Then
        RecordInquiry = "※問合せの表の見出しが見つからないため、問合せは記録していません。"
        Exit Function
    End If

    ' 表の中身（業務チーム・業務分類詳細・件数）
    Set teams = New Collection: Set cats = New Collection: Set cnts = New Collection
    r = hr + 1
    Do While Len(Trim$(CStr(wsN.Cells(r, 2).Value))) > 0
        team = Trim$(CStr(wsN.Cells(r, 2).Value))
        If team = INQ_NONE Then Exit Do
        teams.Add team
        cats.Add Trim$(CStr(wsN.Cells(r, cCat).Value))
        cnts.Add Val(CStr(wsN.Cells(r, cCnt).Value))
        total = total + Val(CStr(wsN.Cells(r, cCnt).Value))
        r = r + 1
    Loop

    ' 受付日＝記録日の前営業日（現在の状況の表と同じ計算）
    d = Application.WorksheetFunction.WorkDay(recDate, -1, ThisWorkbook.Worksheets(SH_HOL).Range("A5:A1000"))

    Set ws = ThisWorkbook.Worksheets(SH_TREND)
    col = EnsureDateColumn(ws, d, TR_HEADER_ROW, TR_FIRST_COL)

    ' 行がなければ追加（計の行を先頭に）
    Set rowMap = TrendRowMap(ws)
    If Not rowMap.Exists(TrendKey(INQ_TYPE, INQ_TOTAL_GRP, "計")) Then
        AddTrendRow ws, INQ_TYPE, INQ_TOTAL_GRP, "計"
        Set rowMap = TrendRowMap(ws)
    End If
    For i = 1 To teams.Count
        k = TrendKey(INQ_TYPE, teams(i), cats(i))
        If Not rowMap.Exists(k) Then
            AddTrendRow ws, INQ_TYPE, teams(i), cats(i)
            Set rowMap = TrendRowMap(ws)
        End If
    Next

    ' その日の問合せの行を 0 にしてから書く
    For r = TR_HEADER_ROW + 1 To TrendLastRow(ws)
        If CStr(ws.Cells(r, 2).Value) = INQ_TYPE Then ws.Cells(r, col).Value = 0
    Next
    For i = 1 To teams.Count
        k = TrendKey(INQ_TYPE, teams(i), cats(i))
        ws.Cells(rowMap(k), col).Value = ws.Cells(rowMap(k), col).Value + cnts(i)
    Next
    ws.Cells(rowMap(TrendKey(INQ_TYPE, INQ_TOTAL_GRP, "計")), col).Value = total
    FormatTrend ws

    RecordInquiry = "問合せ：" & Format(d, "m/d") & " 受付分 " & total & "件を、" & Format(d, "m/d") & " の列に書き込みました。"
End Function

Private Function FindPrefixInCol(ws As Worksheet, ByVal col As Long, ByVal text As String) As Long
    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    For r = 1 To lastR
        If Left$(Trim$(CStr(ws.Cells(r, col).Value)), Len(text)) = text Then FindPrefixInCol = r: Exit Function
    Next
End Function

' 積み上げの行を見分けるキー（申請種別｜区分｜項目）
Private Function TrendKey(ByVal typ As String, ByVal grp As String, ByVal item As String) As String
    TrendKey = typ & "|" & grp & "|" & item
End Function

Private Function CountRecords(ByVal recDate As Date) As Long
    Dim wsR As Worksheet, lastR As Long, v As Variant, r As Long
    Set wsR = ThisWorkbook.Worksheets(SH_REC)
    lastR = wsR.Cells(wsR.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Function
    v = wsR.Range("A1").Resize(lastR, 1).Value
    For r = 2 To lastR
        If IsDate(v(r, 1)) Then
            If CDate(v(r, 1)) = recDate Then CountRecords = CountRecords + 1
        End If
    Next
End Function

' 指定日の記録を消して、残りを詰め直す
Private Function DeleteRecords(wsR As Worksheet, ByVal recDate As Date) As Long
    Dim lastR As Long, v As Variant, outArr() As Variant, r As Long, c As Long, o As Long
    lastR = wsR.Cells(wsR.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Function
    v = wsR.Range("A2").Resize(lastR - 1, REC_COLS).Value
    ReDim outArr(1 To lastR - 1, 1 To REC_COLS)
    For r = 1 To lastR - 1
        If IsDate(v(r, 1)) Then
            If CDate(v(r, 1)) = recDate Then
                DeleteRecords = DeleteRecords + 1
                GoTo NextRow
            End If
        End If
        o = o + 1
        For c = 1 To REC_COLS: outArr(o, c) = v(r, c): Next
NextRow:
    Next
    If DeleteRecords = 0 Then Exit Function
    wsR.Range("A2").Resize(lastR - 1, REC_COLS).ClearContents
    If o > 0 Then wsR.Range("A2").Resize(lastR - 1, REC_COLS).Value = outArr
End Function

Private Sub WriteRows(wsR As Worksheet, newRows As Variant, ByVal n As Long)
    Dim startR As Long, outArr() As Variant, i As Long, c As Long
    If n = 0 Then Exit Sub
    startR = wsR.Cells(wsR.Rows.Count, 1).End(xlUp).Row + 1
    If startR < 2 Then startR = 2
    ReDim outArr(1 To n, 1 To REC_COLS)
    For i = 1 To n
        For c = 1 To REC_COLS: outArr(i, c) = newRows(c, i): Next
    Next
    With wsR.Range("A" & startR).Resize(n, REC_COLS)
        .Value = outArr
        .Columns(1).NumberFormat = "yyyy/mm/dd"
        .Columns(2).NumberFormat = "yyyy/mm/dd hh:mm"
        .Columns(6).NumberFormat = "#,##0"
    End With
End Sub

Private Function FindInCol(ws As Worksheet, ByVal col As Long, ByVal text As String) As Long
    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    For r = 1 To lastR
        If Trim$(CStr(ws.Cells(r, col).Value)) = text Then FindInCol = r: Exit Function
    Next
End Function

'----------------------------------------------------------
' 推移シートへ数字で書き込む（数式は使わない）
'   縦：申請種別 × 項目、横：日付（毎日。土日・休日一覧の日はグレー）
'----------------------------------------------------------
Private Sub WriteTrend(newRows As Variant, ByVal n As Long, ByVal recDate As Date)
    Dim ws As Worksheet, col As Long, rowMap As Object, i As Long, k As String, lastR As Long, added As Boolean

    Set ws = ThisWorkbook.Worksheets(SH_TREND)
    col = EnsureDateColumn(ws, recDate, TR_HEADER_ROW, TR_FIRST_COL)

    ' 行がない項目は、その申請種別のまとまりの最後に追加
    Set rowMap = TrendRowMap(ws)
    For i = 1 To n
        k = TrendKey(CStr(newRows(3, i)), CStr(newRows(4, i)), CStr(newRows(5, i)))
        If Not rowMap.Exists(k) Then
            AddTrendRow ws, CStr(newRows(3, i)), CStr(newRows(4, i)), CStr(newRows(5, i))
            Set rowMap = TrendRowMap(ws)
        End If
    Next

    ' その日の列を書き直す（問合せの行は受付日の列に別に書くので消さない）
    lastR = TrendLastRow(ws)
    For i = TR_HEADER_ROW + 1 To lastR
        If CStr(ws.Cells(i, 2).Value) <> INQ_TYPE Then ws.Cells(i, col).ClearContents
    Next
    For i = 1 To n
        k = TrendKey(CStr(newRows(3, i)), CStr(newRows(4, i)), CStr(newRows(5, i)))
        ws.Cells(rowMap(k), col).Value = newRows(6, i)
    Next
    FormatTrend ws
End Sub

Private Function TrendLastRow(ws As Worksheet) As Long
    TrendLastRow = ws.Cells(ws.Rows.Count, TR_ITEM_COL).End(xlUp).Row
    If TrendLastRow < TR_HEADER_ROW Then TrendLastRow = TR_HEADER_ROW
End Function

Private Function TrendRowMap(ws As Worksheet) As Object
    Dim d As Object, r As Long
    Set d = CreateObject("Scripting.Dictionary")
    For r = TR_HEADER_ROW + 1 To TrendLastRow(ws)
        d(TrendKey(CStr(ws.Cells(r, 2).Value), CStr(ws.Cells(r, 3).Value), CStr(ws.Cells(r, TR_ITEM_COL).Value))) = r
    Next
    Set TrendRowMap = d
End Function

' 同じ申請種別・同じ区分の最後に行を追加（最古はステータスの下。区分がまだなければ申請種別のまとまりの最後）
Private Sub AddTrendRow(ws As Worksheet, ByVal typ As String, ByVal grp As String, ByVal item As String)
    Dim r As Long, lastOfType As Long, lastOfGrp As Long, lastOfStatus As Long, lastR As Long, g As String
    lastR = TrendLastRow(ws)
    For r = TR_HEADER_ROW + 1 To lastR
        If CStr(ws.Cells(r, 2).Value) = typ Then
            lastOfType = r
            g = CStr(ws.Cells(r, 3).Value)
            If g = grp Then lastOfGrp = r
            If g = "ステータス" Then lastOfStatus = r
        End If
    Next
    If lastOfGrp = 0 And grp = GRP_OLDEST Then lastOfGrp = lastOfStatus
    If lastOfGrp = 0 Then lastOfGrp = lastOfType
    If lastOfGrp > 0 Then
        r = lastOfGrp + 1
        ws.Rows(r).Insert
    Else
        r = lastR + 1
    End If
    ws.Cells(r, 2).Value = typ
    ws.Cells(r, 3).Value = grp
    ws.Cells(r, TR_ITEM_COL).Value = item
End Sub

Private Sub FormatTrend(ws As Worksheet)
    Dim lastR As Long, lastC As Long, r As Long, it As String, g As String, prevG As String

    lastR = TrendLastRow(ws)
    lastC = ws.Cells(TR_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastR <= TR_HEADER_ROW Or lastC < TR_FIRST_COL Then Exit Sub

    With ws.Range(ws.Cells(TR_HEADER_ROW, TR_FIRST_COL), ws.Cells(TR_HEADER_ROW, lastC))
        .NumberFormat = "m/d(aaa)"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(0, 0, 0)
    End With
    With ws.Range(ws.Cells(TR_HEADER_ROW + 1, 2), ws.Cells(lastR, lastC))
        .Interior.pattern = xlNone
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
    End With
    ws.Range(ws.Cells(TR_HEADER_ROW + 1, TR_FIRST_COL), ws.Cells(lastR, lastC)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(TR_HEADER_ROW, TR_FIRST_COL), ws.Cells(lastR, lastC)).ColumnWidth = 10

    For r = TR_HEADER_ROW + 1 To lastR
        it = CStr(ws.Cells(r, TR_ITEM_COL).Value)
        ' 申請種別のまとまりの先頭行だけ種別名を濃く
        If CStr(ws.Cells(r, 2).Value) <> CStr(ws.Cells(r - 1, 2).Value) Then
            ws.Cells(r, 2).Font.Bold = True
            ws.Range(ws.Cells(r, 2), ws.Cells(r, lastC)).Borders(xlEdgeTop).Weight = xlMedium
            ws.Range(ws.Cells(r, 2), ws.Cells(r, lastC)).Borders(xlEdgeTop).Color = RGB(89, 89, 89)
        Else
            ws.Cells(r, 2).Font.Color = RGB(191, 191, 191)
        End If
        If CStr(ws.Cells(r, 3).Value) = GRP_OLDEST Then
            ws.Range(ws.Cells(r, TR_FIRST_COL), ws.Cells(r, lastC)).NumberFormat = "yy/m/d"
        End If
        ' 大きく二分：ステータス＋最古（緑）／フラグ（オレンジ）
        g = CStr(ws.Cells(r, 3).Value)
        If g = "フラグ" Then
            ws.Range(ws.Cells(r, 3), ws.Cells(r, TR_ITEM_COL)).Interior.Color = RGB(252, 228, 214)
        Else
            ws.Range(ws.Cells(r, 3), ws.Cells(r, TR_ITEM_COL)).Interior.Color = RGB(226, 239, 218)
        End If
        prevG = CStr(ws.Cells(r - 1, 3).Value)
        If CStr(ws.Cells(r, 2).Value) = CStr(ws.Cells(r - 1, 2).Value) Then
            If g = "フラグ" And prevG <> "フラグ" Then
                With ws.Range(ws.Cells(r, 2), ws.Cells(r, lastC)).Borders(xlEdgeTop)
                    .Weight = xlThick: .Color = RGB(197, 89, 30)
                End With
            End If
            If g = prevG Then ws.Cells(r, 3).Font.Color = RGB(150, 150, 150)
        End If
        If g <> prevG Or CStr(ws.Cells(r, 2).Value) <> CStr(ws.Cells(r - 1, 2).Value) Then ws.Cells(r, 3).Font.Bold = True
        If it = "計" Or it = "フラグ計" Then
            ws.Range(ws.Cells(r, TR_ITEM_COL), ws.Cells(r, lastC)).Font.Bold = True
            ws.Range(ws.Cells(r, TR_FIRST_COL), ws.Cells(r, lastC)).Interior.Color = RGB(242, 242, 242)
        End If
    Next

    ' 土日・休日の列はグレー
    ShadeRestDays ws, TR_HEADER_ROW, TR_FIRST_COL, lastR
End Sub

' 現在の状況の上の見出し（グループ）から区分を決める
Private Function GroupOf(ByVal grp As String) As String
    If Left$(grp, 5) = "ステータス" Then
        GroupOf = "ステータス"
    ElseIf InStr(grp, "最古") > 0 Then
        GroupOf = GRP_OLDEST
    Else
        GroupOf = "フラグ"
    End If
End Function
