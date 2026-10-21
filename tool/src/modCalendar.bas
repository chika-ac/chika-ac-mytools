Option Explicit

'==========================================================
' 日付を横に並べるシート（推移・日別処理件数）の共通部品
'   ・日付は毎日1列。土日と休日一覧の日はグレー
'==========================================================

Private Const SH_HOL As String = "休日一覧"
Private Const HOL_HEADER_ROW As Long = 4

' 日付の列を探す（なければ、途中の日も含めて毎日分の列を作る）
Public Function EnsureDateColumn(ws As Worksheet, ByVal d As Date, ByVal hdrRow As Long, ByVal firstCol As Long) As Long
    Dim lastC As Long, firstD As Date, lastD As Date, cnt As Long, i As Long
    d = DateValue(d)
    lastC = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    If lastC < firstCol Then
        ws.Cells(hdrRow, firstCol).Value = d
        EnsureDateColumn = firstCol
        Exit Function
    End If
    firstD = ws.Cells(hdrRow, firstCol).Value
    lastD = ws.Cells(hdrRow, lastC).Value
    If d > lastD Then
        For i = 1 To d - lastD
            ws.Cells(hdrRow, lastC + i).Value = lastD + i
        Next
        EnsureDateColumn = lastC + (d - lastD)
    ElseIf d < firstD Then
        cnt = firstD - d
        ws.Range(ws.Cells(1, firstCol), ws.Cells(1, firstCol + cnt - 1)).EntireColumn.Insert
        For i = 0 To cnt - 1
            ws.Cells(hdrRow, firstCol + i).Value = d + i
        Next
        EnsureDateColumn = firstCol
    Else
        For i = firstCol To lastC
            If CDbl(ws.Cells(hdrRow, i).Value) = CDbl(d) Then EnsureDateColumn = i: Exit Function
        Next
        Err.Raise vbObjectError + 2, , "「" & ws.Name & "」シートに " & Format(d, "m/d") & " の列が見つかりません（日付の列が抜けています）"
    End If
End Function

' 土日・休日の列をグレーにする（見出し行から lastR 行まで）
Public Sub ShadeRestDays(ws As Worksheet, ByVal hdrRow As Long, ByVal firstCol As Long, ByVal lastR As Long)
    Dim hol As Object, lastC As Long, c As Long, d As Date
    Set hol = HolidaySet()
    lastC = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    For c = firstCol To lastC
        d = ws.Cells(hdrRow, c).Value
        If IsRestDay(d, hol) Then
            With ws.Range(ws.Cells(hdrRow, c), ws.Cells(lastR, c))
                .Interior.Color = RGB(217, 217, 217)
                .Font.Color = RGB(128, 128, 128)
            End With
        End If
    Next
End Sub

Public Function IsRestDay(ByVal d As Date, hol As Object) As Boolean
    IsRestDay = (Weekday(d, vbMonday) >= 6) Or hol.Exists(CLng(d))
End Function

' 休日一覧シートのA列の日付
Public Function HolidaySet() As Object
    Dim d As Object, ws As Worksheet, r As Long, lastR As Long, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    Set HolidaySet = d
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_HOL)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = HOL_HEADER_ROW + 1 To lastR
        v = ws.Cells(r, 1).Value
        If IsDate(v) Then d(CLng(CDate(v))) = True
    Next
End Function

' 今日の列をハイライト（条件付き書式。毎日自動で今日の列に移る）
'   offsetYears：1 なら「1年前の今日」（昨年実績のシート用）
Public Sub HighlightToday(ws As Worksheet, ByVal hdrRow As Long, ByVal firstCol As Long, ByVal lastR As Long, _
                          ByVal lastC As Long, Optional ByVal offsetYears As Long = 0)
    Dim rg As Range, hd As Range, ref As String, cond As String, fc As FormatCondition
    ref = ws.Cells(hdrRow, firstCol).Address(True, False)
    If offsetYears = 0 Then
        cond = "AND(ISNUMBER(" & ref & ")," & ref & "=TODAY())"
    Else
        cond = "AND(ISNUMBER(" & ref & ")," & ref & "=DATE(YEAR(TODAY())-" & offsetYears & ",MONTH(TODAY()),DAY(TODAY())))"
    End If
    Set rg = ws.Range(ws.Cells(hdrRow + 1, firstCol), ws.Cells(lastR, lastC))
    Set hd = ws.Range(ws.Cells(hdrRow, firstCol), ws.Cells(hdrRow, lastC))
    Set fc = rg.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & cond)
    fc.Interior.Color = RGB(255, 230, 190)
    fc.SetFirstPriority
    Set fc = hd.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & cond)
    fc.Interior.Color = RGB(237, 125, 49): fc.Font.Color = vbWhite: fc.Font.Bold = True
    fc.SetFirstPriority
End Sub