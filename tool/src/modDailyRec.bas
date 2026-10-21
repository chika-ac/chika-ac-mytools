Option Explicit

'==========================================================
' 日別処理件数の積み上げ
'   「日別処理件数」シート（数式）の値を、「日別処理件数（積み上げ）」シートに数字で残す。
'   記録済みの日で値が変わっているときは、上書きするか確認する。
'==========================================================

Private Const SH_SRC As String = "日別処理件数"
Private Const SH_DST As String = "日別処理件数（積み上げ）"
Private Const SRC_HDR As Long = 4       ' 日別処理件数：日付の見出し行
Private Const SRC_FIRST As Long = 5     ' 日別処理件数：最初の日付の列（D列は期間計）
Private Const DST_HDR As Long = 4       ' 積み上げ：日付の見出し行
Private Const DST_FIRST As Long = 4     ' 積み上げ：最初の日付の列（B=申請種別, C=Status）
Private Const SUB_TOTAL As String = "計"

'----------------------------------------------------------
' 操作シートのボタンから呼ぶ
'----------------------------------------------------------
Public Sub 日別を記録()
    Dim res As String, parts As Variant, nNew As Long, nDiff As Long, mode As Long

    Application.ScreenUpdating = False
    res = RecordDaily(0)
    Application.ScreenUpdating = True
    If Left$(res, 1) = "!" Then MsgBox Mid$(res, 2), vbExclamation, "日別を記録": Exit Sub

    parts = Split(res, "|")
    nNew = CLng(parts(0)): nDiff = CLng(parts(1))
    If nNew = 0 And nDiff = 0 Then
        MsgBox "記録済みの内容と同じです（追加・変更はありません）。", vbInformation, "日別を記録"
        Exit Sub
    End If
    mode = 1
    If nDiff > 0 Then
        Select Case MsgBox("記録済みの日のうち " & nDiff & " 日分は、いまの日別処理件数と値が違います（" & parts(2) & "）。" & vbLf & _
                           "上書きしますか？" & vbLf & vbLf & _
                           "はい：上書きする　／　いいえ：記録済みの値を残す（新しい日だけ追加）", _
                           vbYesNoCancel + vbQuestion, "日別を記録")
            Case vbYes: mode = 2
            Case vbNo: mode = 1
            Case Else: Exit Sub
        End Select
        If mode = 1 And nNew = 0 Then Exit Sub
    End If

    Application.ScreenUpdating = False
    res = RecordDaily(mode)
    Application.ScreenUpdating = True
    MsgBox res, vbInformation, "日別を記録"
End Sub

'----------------------------------------------------------
' mode 0：確認のみ（"新しい日数|違う日数|違う日の一覧" を返す）
' mode 1：新しい日だけ追加　／　mode 2：新しい日を追加し、違う日も上書き
' 問題があるときは "!" で始まるメッセージを返す
'----------------------------------------------------------
Public Function RecordDaily(ByVal mode As Long) As String
    Dim src As Worksheet, dst As Worksheet
    Dim lastSrcR As Long, lastSrcC As Long, r As Long, c As Long, dc As Long
    Dim rowMap As Object, key As String, d As Date, v As Variant
    Dim nNew As Long, nDiff As Long, diffList As String, isNew As Boolean, isDiff As Boolean
    Dim nWritten As Long, nOver As Long

    Set src = ThisWorkbook.Worksheets(SH_SRC)
    Set dst = ThisWorkbook.Worksheets(SH_DST)
    Application.Calculate

    If Not IsDate(src.Cells(SRC_HDR, SRC_FIRST).Value) Then
        RecordDaily = "!日別処理件数にまだデータがありません。先に②データ取込をしてください。"
        Exit Function
    End If
    lastSrcR = SrcLastRow(src)
    lastSrcC = SrcLastCol(src)

    ' 行がない「申請種別×Status」は積み上げ側に追加
    If mode > 0 Then
        Set rowMap = DstRowMap(dst)
        For r = SRC_HDR + 1 To lastSrcR
            key = CStr(src.Cells(r, 2).Value) & "|" & CStr(src.Cells(r, 3).Value)
            If Not rowMap.Exists(key) Then
                AddDstRow dst, CStr(src.Cells(r, 2).Value), CStr(src.Cells(r, 3).Value)
                Set rowMap = DstRowMap(dst)
            End If
        Next
    End If
    Set rowMap = DstRowMap(dst)

    For c = SRC_FIRST To lastSrcC
        d = CDate(src.Cells(SRC_HDR, c).Value)
        dc = FindDstCol(dst, d)
        isNew = (dc = 0)
        isDiff = False
        If Not isNew Then
            For r = SRC_HDR + 1 To lastSrcR
                key = CStr(src.Cells(r, 2).Value) & "|" & CStr(src.Cells(r, 3).Value)
                If rowMap.Exists(key) Then
                    v = dst.Cells(rowMap(key), dc).Value
                    If CDbl(Val(CStr(v))) <> CDbl(Val(CStr(src.Cells(r, c).Value))) Or IsEmpty(v) Then isDiff = True: Exit For
                Else
                    isDiff = True: Exit For
                End If
            Next
        End If
        If isNew Then nNew = nNew + 1
        If isDiff Then
            nDiff = nDiff + 1
            If nDiff <= 5 Then diffList = diffList & IIf(Len(diffList) > 0, "、", "") & Format(d, "m/d")
        End If

        If (mode >= 1 And isNew) Or (mode = 2 And isDiff) Then
            dc = EnsureDateColumn(dst, d, DST_HDR, DST_FIRST)
            For r = SRC_HDR + 1 To lastSrcR
                key = CStr(src.Cells(r, 2).Value) & "|" & CStr(src.Cells(r, 3).Value)
                dst.Cells(rowMap(key), dc).Value = src.Cells(r, c).Value
            Next
            If isNew Then nWritten = nWritten + 1 Else nOver = nOver + 1
        End If
    Next
    If nDiff > 5 Then diffList = diffList & " ほか"

    If mode = 0 Then
        RecordDaily = nNew & "|" & nDiff & "|" & diffList
        Exit Function
    End If
    FormatDst dst
    RecordDaily = "日別処理件数を積み上げシートに記録しました。" & vbLf & _
                  "追加：" & nWritten & " 日分" & IIf(nOver > 0, "　上書き：" & nOver & " 日分", "")
End Function

'----------------------------------------------------------
' 日別処理件数（元）の範囲
'----------------------------------------------------------
Private Function SrcLastRow(src As Worksheet) As Long
    Dim r As Long
    r = SRC_HDR + 1
    Do While Len(CStr(src.Cells(r, 3).Value)) > 0
        r = r + 1
    Loop
    SrcLastRow = r - 1
End Function

Private Function SrcLastCol(src As Worksheet) As Long
    Dim c As Long
    c = SRC_FIRST
    Do While IsDate(src.Cells(SRC_HDR, c + 1).Value)
        c = c + 1
    Loop
    SrcLastCol = c
End Function

'----------------------------------------------------------
' 積み上げシート
'----------------------------------------------------------
Private Function FindDstCol(dst As Worksheet, ByVal d As Date) As Long
    Dim c As Long, lastC As Long
    lastC = dst.Cells(DST_HDR, dst.Columns.Count).End(xlToLeft).Column
    For c = DST_FIRST To lastC
        If IsDate(dst.Cells(DST_HDR, c).Value) Then
            If CLng(CDate(dst.Cells(DST_HDR, c).Value)) = CLng(d) Then
                ' 列はあっても値が1つも入っていなければ「未記録」
                If Application.WorksheetFunction.Count(dst.Range(dst.Cells(DST_HDR + 1, c), dst.Cells(DstLastRow(dst), c))) > 0 Then FindDstCol = c
                Exit Function
            End If
        End If
    Next
End Function

Private Function DstLastRow(dst As Worksheet) As Long
    DstLastRow = dst.Cells(dst.Rows.Count, 3).End(xlUp).Row
    If DstLastRow < DST_HDR Then DstLastRow = DST_HDR
End Function

Private Function DstRowMap(dst As Worksheet) As Object
    Dim d As Object, r As Long
    Set d = CreateObject("Scripting.Dictionary")
    For r = DST_HDR + 1 To DstLastRow(dst)
        d(CStr(dst.Cells(r, 2).Value) & "|" & CStr(dst.Cells(r, 3).Value)) = r
    Next
    Set DstRowMap = d
End Function

' 同じ申請種別のまとまりの最後（「計」があればその上）に行を追加
Private Sub AddDstRow(dst As Worksheet, ByVal typ As String, ByVal stat As String)
    Dim r As Long, lastOfType As Long, subRow As Long, lastR As Long
    lastR = DstLastRow(dst)
    For r = DST_HDR + 1 To lastR
        If CStr(dst.Cells(r, 2).Value) = typ Then
            lastOfType = r
            If CStr(dst.Cells(r, 3).Value) = SUB_TOTAL Then subRow = r
        End If
    Next
    If lastOfType = 0 Then
        r = lastR + 1
    ElseIf subRow > 0 And stat <> SUB_TOTAL Then
        r = subRow
        dst.Rows(r).Insert
    Else
        r = lastOfType + 1
        dst.Rows(r).Insert
    End If
    dst.Cells(r, 2).Value = typ
    dst.Cells(r, 3).Value = stat
End Sub

Private Sub FormatDst(dst As Worksheet)
    Dim lastR As Long, lastC As Long, r As Long

    lastR = DstLastRow(dst)
    lastC = dst.Cells(DST_HDR, dst.Columns.Count).End(xlToLeft).Column
    If lastR <= DST_HDR Or lastC < DST_FIRST Then Exit Sub

    With dst.Range(dst.Cells(DST_HDR, DST_FIRST), dst.Cells(DST_HDR, lastC))
        .NumberFormat = "m/d(aaa)"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(0, 0, 0)
    End With
    With dst.Range(dst.Cells(DST_HDR + 1, 2), dst.Cells(lastR, lastC))
        .Interior.pattern = xlNone
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
    End With
    dst.Range(dst.Cells(DST_HDR + 1, DST_FIRST), dst.Cells(lastR, lastC)).NumberFormat = "#,##0"
    dst.Range(dst.Cells(DST_HDR, DST_FIRST), dst.Cells(lastR, lastC)).ColumnWidth = 10

    For r = DST_HDR + 1 To lastR
        If CStr(dst.Cells(r, 2).Value) <> CStr(dst.Cells(r - 1, 2).Value) Then
            dst.Cells(r, 2).Font.Bold = True
            dst.Range(dst.Cells(r, 2), dst.Cells(r, lastC)).Borders(xlEdgeTop).Weight = xlMedium
            dst.Range(dst.Cells(r, 2), dst.Cells(r, lastC)).Borders(xlEdgeTop).Color = RGB(89, 89, 89)
        Else
            dst.Cells(r, 2).Font.Color = RGB(191, 191, 191)
        End If
        If CStr(dst.Cells(r, 3).Value) = SUB_TOTAL Then
            dst.Range(dst.Cells(r, 3), dst.Cells(r, lastC)).Font.Bold = True
            dst.Range(dst.Cells(r, 3), dst.Cells(r, lastC)).Interior.Color = RGB(242, 242, 242)
        End If
    Next
    ShadeRestDays dst, DST_HDR, DST_FIRST, lastR
End Sub