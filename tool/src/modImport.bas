Option Explicit

'==========================================================
' 元データ取込
'   取込設定シートの内容に従い、元ネタフォルダのファイルを
'   このブックのシートへ取り込む。取込後は取込済フォルダへ移動。
'==========================================================

Private Const SH_MENU As String = "操作"
Private Const SH_SET As String = "取込設定"
Private Const SH_LOG As String = "取込履歴"
Private Const SH_HOL As String = "休日一覧"
Private Const HOL_HEADER_ROW As Long = 4       ' 休日一覧の見出し行
Private Const LAST_COL As Long = 10           ' 操作シートの取込状況の最終列（J）
Private Const COL_FILE As String = "_取込ファイル"
Private Const COL_TIME As String = "_取込日時"
Private Const MENU_STATUS_ROW As Long = 14    ' 操作シート：②の続き（取込状況）の開始行。③④はその下に描く
Private Const CLR_GREEN As Long = 3499348       ' RGB(84,101,53)  ②の緑
Private Const CLR_GREEN_LIGHT As Long = 14348258 ' RGB(226,239,218)
Private Const CLR_ORANGE As Long = 1989061      ' RGB(197,89,30)  ③
Private Const CLR_BLUE As Long = 8014638        ' RGB(46,75,122)  ④
Private Const CLR_PURPLE As Long = 10498160     ' RGB(112,48,160) ⑤
Private Const CLR_TEAL As Long = 7368704        ' RGB(0,112,112)  ⑥
Private Const LOG_HEADER_ROW As Long = 4      ' 取込履歴の見出し行

Private Const MODE_REPLACE As String = "書き換え（毎回全件クリア）"
Private Const MODE_MERGE As String = "蓄積（同じ期間は上書き）"
Private Const MODE_APPEND As String = "蓄積（そのまま追加）"

Private Type ImportSetting
    RowNo As Long
    DataName As String
    keyword As String
    ext As String
    Fmt As String
    Charset As String
    delim As String
    SrcSheet As String
    HeaderRow As Long
    DataRow As Long
    DateCol As String
    PersonCol As String
    Target As String
    mode As String
    MoveTo As String
    Enabled As Boolean
End Type

'----------------------------------------------------------
' 操作シートのボタンから呼ぶ
'----------------------------------------------------------
Public Sub データ取込()
    If MsgBox("元ネタフォルダのファイルを取り込みます。よろしいですか？" & vbLf & _
              "（取り込んだファイルは取込済フォルダへ移動します）", _
              vbOKCancel + vbQuestion, "データ取込") <> vbOK Then Exit Sub
    MsgBox RunImport(), vbInformation, "データ取込"
End Sub

Public Sub 取込状況更新()
    Dim calcMode As XlCalculation
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual      ' 書き換えのたびの再計算を止める（ブック全体の式が多いため）
    RefreshStatus
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
End Sub

'----------------------------------------------------------
' 取込本体（結果メッセージを返す）
'----------------------------------------------------------
Public Function RunImport() As String
    Dim fso As Object, wsSet As Worksheet
    Dim srcDir As String, doneDir As String
    Dim sts() As ImportSetting, n As Long, i As Long
    Dim files As Collection, f As Variant, fileTime As Date
    Dim runTime As Date, summary As String, ln As String
    Dim msg As String, errMsg As String
    Dim nRead As Long, nDel As Long, dMin As Variant, dMax As Variant
    Dim cntOk As Long, cntNg As Long, workTouched As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsSet = ThisWorkbook.Worksheets(SH_SET)
    srcDir = ResolvePath(LabelValue(wsSet, "元ネタフォルダ"))
    doneDir = ResolvePath(LabelValue(wsSet, "取込済フォルダ"))
    If Len(srcDir) = 0 Or Not fso.FolderExists(srcDir) Then
        RunImport = "元ネタフォルダが見つかりません：" & srcDir & vbLf & _
                    "取込設定シートの「元ネタフォルダ」を確認してください。"
        Exit Function
    End If

    On Error GoTo EH
    n = LoadSettings(wsSet, sts)
    runTime = Now
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For i = 1 To n
        If sts(i).Enabled Then
            Set files = FindFiles(fso, srcDir, sts(i))
            If files.Count = 0 Then summary = summary & "・" & sts(i).DataName & "：ファイルなし" & vbLf
            For Each f In files
                msg = "": errMsg = "": nRead = 0: nDel = 0: dMin = Empty: dMax = Empty
                fileTime = fso.GetFile(CStr(f)).DateLastModified

                On Error Resume Next
                ImportFile sts(i), CStr(f), runTime, msg, nRead, nDel, dMin, dMax
                If Err.Number <> 0 Then errMsg = Err.Description
                Err.Clear
                On Error GoTo EH

                If Len(errMsg) = 0 And sts(i).MoveTo = "取込済" Then
                    On Error Resume Next
                    MoveToDone fso, CStr(f), doneDir
                    If Err.Number <> 0 Then msg = AddMsg(msg, "取込済への移動に失敗：" & Err.Description)
                    Err.Clear
                    On Error GoTo EH
                End If

                WriteLog runTime, sts(i), CStr(f), fileTime, nRead, nDel, dMin, dMax, errMsg, msg

                ln = "・" & sts(i).DataName & "：" & fso.GetFileName(CStr(f))
                If Len(errMsg) > 0 Then
                    cntNg = cntNg + 1
                    ln = ln & " → エラー：" & errMsg
                Else
                    cntOk = cntOk + 1
                    If sts(i).Target = WorkSourceSheet() Then workTouched = True
                    ln = ln & " → " & nRead & "件"
                    If Not IsEmpty(dMin) Then ln = ln & "（" & Format(dMin, "yyyy/m/d") & "?" & Format(dMax, "yyyy/m/d") & "）"
                    If nDel > 0 Then ln = ln & "　既存 " & nDel & "件を入替"
                    If Len(msg) > 0 Then ln = ln & vbLf & "　　※" & msg
                End If
                summary = summary & ln & vbLf
            Next
        End If
    Next

    If workTouched Then
        On Error Resume Next
        ln = RefreshWork()
        If Err.Number <> 0 Then ln = "稼働実績の更新に失敗：" & Err.Description
        Err.Clear
        On Error GoTo EH
        summary = summary & "・" & ln & vbLf
    End If

    If cntOk > 0 Then
        On Error Resume Next
        ln = RefreshDetail()
        If Err.Number <> 0 Then ln = "進捗（詳細）の更新に失敗：" & Err.Description
        Err.Clear
        On Error GoTo EH
        summary = summary & "・" & ln & vbLf
    End If

    UpdateSettingStatus wsSet, sts, n
    RefreshStatus
    ThisWorkbook.Worksheets(SH_MENU).Activate
    RunImport = "取込が終わりました。（成功 " & cntOk & "件／エラー " & cntNg & "件）" & vbLf & vbLf & summary

Fin:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Function
EH:
    RunImport = "取込を中断しました：" & Err.Description & vbLf & vbLf & summary
    Resume Fin
End Function

'----------------------------------------------------------
' 1ファイルの取込
'----------------------------------------------------------
Private Sub ImportFile(st As ImportSetting, ByVal path As String, ByVal runTime As Date, _
                       ByRef msg As String, ByRef nRead As Long, ByRef nDel As Long, _
                       ByRef dMin As Variant, ByRef dMax As Variant)
    Dim data As Variant, nR As Long, nC As Long, r As Long, c As Long
    Dim dateIdx As Long, personIdx As Long, persons As Object, v As Variant
    Dim ws As Worksheet, exR As Long, exC As Long, exHdr As Variant, ex As Variant
    Dim keep() As Boolean, nKeep As Long, outArr() As Variant, o As Long, k As Boolean
    Dim fileName As String

    fileName = Mid$(path, InStrRev(path, "\") + 1)
    Select Case st.Fmt
        Case "Excel": data = ReadExcel(path, st)
        Case "HTML表": data = ReadHtml(path, st)
        Case "CSV", "テキスト": data = ReadDelimited(path, st)
        Case Else: Err.Raise vbObjectError + 2, , "未対応の形式です：" & st.Fmt
    End Select
    nR = UBound(data, 1): nC = UBound(data, 2)
    nRead = nR - 1

    dateIdx = HeaderIndex(data, nC, st.DateCol)
    If Len(st.DateCol) > 0 And dateIdx = 0 Then Err.Raise vbObjectError + 3, , "日付列「" & st.DateCol & "」が見つかりません"
    personIdx = HeaderIndex(data, nC, st.PersonCol)
    If Len(st.PersonCol) > 0 And personIdx = 0 Then Err.Raise vbObjectError + 3, , "人別管理列「" & st.PersonCol & "」が見つかりません"
    If st.mode = MODE_MERGE And dateIdx = 0 Then Err.Raise vbObjectError + 3, , "「" & MODE_MERGE & "」には日付列名の設定が必要です"

    ' 日付列を日付型にそろえ、期間と対象者を把握
    Set persons = CreateObject("Scripting.Dictionary")
    For r = 2 To nR
        If dateIdx > 0 Then
            v = data(r, dateIdx)
            If VarType(v) = vbString And IsDate(v) Then data(r, dateIdx) = CDate(v)
            v = data(r, dateIdx)
            If VarType(v) = vbDate Then
                v = CDate(Int(CDbl(v)))
                If IsEmpty(dMin) Then dMin = v: dMax = v
                If v < dMin Then dMin = v
                If v > dMax Then dMax = v
            End If
        End If
        If personIdx > 0 Then persons(SafeStr(data(r, personIdx))) = True
    Next

    ' 取込先シートと列構成の確認
    Set ws = GetOrCreateSheet(st.Target)
    exR = LastRow(ws)
    exC = LastColOfRow1(ws)
    If exR >= 1 And exC > 0 Then
        exHdr = ws.Range("A1").Resize(1, exC).Value
        If Not SameHeader(exHdr, exC, data, nC) Then
            If st.mode = MODE_REPLACE Or exR < 2 Then
                msg = AddMsg(msg, "列構成が前回と変わりました（取込先シートの見出しを確認してください）")
            Else
                Err.Raise vbObjectError + 4, , "列構成が既存シートと違うため取り込めません。元ファイルの見出しを確認し、" & _
                          "新しい列構成で入れ直す場合はシート「" & st.Target & "」を削除してから取り込んでください"
            End If
            exR = 0
        End If
    End If

    ' 既存行のうち残すものを決める
    If st.mode <> MODE_REPLACE And exR >= 2 Then
        ex = ws.Range("A2").Resize(exR - 1, nC + 2).Value
        ReDim keep(1 To exR - 1)
        For r = 1 To exR - 1
            k = True
            If st.mode = MODE_MERGE And Not IsEmpty(dMin) Then
                v = ex(r, dateIdx)
                If VarType(v) = vbDate Then
                    v = CDate(Int(CDbl(v)))
                    If v >= dMin And v <= dMax Then
                        If personIdx = 0 Then
                            k = False
                        ElseIf persons.Exists(SafeStr(ex(r, personIdx))) Then
                            k = False
                        End If
                    End If
                End If
            End If
            keep(r) = k
            If k Then nKeep = nKeep + 1 Else nDel = nDel + 1
        Next
    End If

    ' 書き出し用配列：見出し＋残す既存行＋新規行
    ReDim outArr(1 To 1 + nKeep + nRead, 1 To nC + 2)
    For c = 1 To nC: outArr(1, c) = data(1, c): Next
    outArr(1, nC + 1) = COL_FILE: outArr(1, nC + 2) = COL_TIME
    o = 1
    If nKeep > 0 Then
        For r = 1 To exR - 1
            If keep(r) Then
                o = o + 1
                For c = 1 To nC + 2: outArr(o, c) = ex(r, c): Next
            End If
        Next
    End If
    For r = 2 To nR
        o = o + 1
        For c = 1 To nC: outArr(o, c) = data(r, c): Next
        outArr(o, nC + 1) = fileName
        outArr(o, nC + 2) = runTime
    Next
    WriteSheet ws, outArr, dateIdx

    If nRead = 0 Then msg = AddMsg(msg, "データ行がありません（見出しのみ）")
    If Not IsEmpty(dMax) Then
        If dMax >= Date Then msg = AddMsg(msg, "当日以降のデータを含みます（最新 " & Format(dMax, "yyyy/mm/dd") & "）。確定前のデータの可能性があります")
    End If
End Sub

'----------------------------------------------------------
' 読込：Excel
'----------------------------------------------------------
Private Function ReadExcel(ByVal path As String, st As ImportSetting) As Variant
    Dim wb As Workbook, ws As Worksheet, ur As Range, v As Variant, tmp() As Variant
    Dim lastR As Long, lastC As Long, errDesc As String

    On Error GoTo EH
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Application.DisplayAlerts = True
    If Len(st.SrcSheet) > 0 Then
        On Error Resume Next
        Set ws = wb.Worksheets(st.SrcSheet)
        On Error GoTo EH
        ' シートが1枚だけのブックは、シート名が設定と違っても読む（出力元でシート名が変わることがあるため）
        If ws Is Nothing And wb.Worksheets.Count = 1 Then Set ws = wb.Worksheets(1)
        If ws Is Nothing Then Err.Raise vbObjectError + 10, , "シート「" & st.SrcSheet & "」が見つかりません"
    Else
        Set ws = wb.Worksheets(1)
    End If
    Set ur = ws.UsedRange
    lastR = ur.Row + ur.Rows.Count - 1
    lastC = ur.Column + ur.Columns.Count - 1
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    wb.Close SaveChanges:=False
    If Not IsArray(v) Then
        ReDim tmp(1 To 1, 1 To 1): tmp(1, 1) = v: v = tmp
    End If
    ReadExcel = Normalize(v, st, False)
    Exit Function
EH:
    errDesc = Err.Description
    Application.DisplayAlerts = True
    If wb Is Nothing Then Err.Raise vbObjectError + 10, , "ファイルを開けません：" & errDesc
    If Len(st.SrcSheet) > 0 And ws Is Nothing Then errDesc = "シート「" & st.SrcSheet & "」が見つかりません"
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 10, , errDesc
End Function

'----------------------------------------------------------
' 読込：HTMLの表（拡張子 .xls の中身がHTMLのもの）
'----------------------------------------------------------
Private Function ReadHtml(ByVal path As String, st As ImportSetting) As Variant
    Dim txt As String, reRow As Object, reCell As Object, reTag As Object
    Dim mRows As Object, mCells As Object, raw() As Variant, i As Long, j As Long, maxC As Long

    txt = ReadTextFile(path, CharsetName(st.Charset))
    Set reRow = NewRegExp("<tr(?:\s[^>]*)?>([\s\S]*?)</tr>")
    Set reCell = NewRegExp("<t[hd](?:\s[^>]*)?>([\s\S]*?)</t[hd]>")
    Set reTag = NewRegExp("<[^>]+>")
    Set mRows = reRow.Execute(txt)
    If mRows.Count = 0 Then Err.Raise vbObjectError + 11, , "HTMLの表が見つかりません"
    For i = 0 To mRows.Count - 1
        j = reCell.Execute(mRows(i).SubMatches(0)).Count
        If j > maxC Then maxC = j
    Next
    If maxC = 0 Then Err.Raise vbObjectError + 11, , "HTMLの表にセルがありません"
    ReDim raw(1 To mRows.Count, 1 To maxC)
    For i = 0 To mRows.Count - 1
        Set mCells = reCell.Execute(mRows(i).SubMatches(0))
        For j = 0 To mCells.Count - 1
            raw(i + 1, j + 1) = HtmlText(reTag.Replace(mCells(j).SubMatches(0), ""))
        Next
    Next
    ReadHtml = Normalize(raw, st, True)
End Function

'----------------------------------------------------------
' 読込：CSV／タブ区切りテキスト
'----------------------------------------------------------
Private Function ReadDelimited(ByVal path As String, st As ImportSetting) As Variant
    Dim txt As String, lines As Variant, n As Long, i As Long, j As Long, maxC As Long
    Dim parsed() As Variant, p As Variant, raw() As Variant, d As String

    txt = ReadTextFile(path, CharsetName(st.Charset))
    txt = Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(txt, vbLf)
    n = UBound(lines) + 1
    Do While n > 0
        If Len(lines(n - 1)) > 0 Then Exit Do
        n = n - 1
    Loop
    If n = 0 Then Err.Raise vbObjectError + 12, , "ファイルが空です"

    d = DelimChar(st.delim, st.ext)
    ' 見出し行に設定の区切り文字がなく、もう一方（タブ／カンマ）があればそちらで読む（出力形式の変更に備える）
    p = lines(Application.Min(st.HeaderRow, n) - 1)
    If InStr(p, d) = 0 Then
        If d = vbTab And InStr(p, ",") > 0 Then
            d = ","
        ElseIf d = "," And InStr(p, vbTab) > 0 Then
            d = vbTab
        End If
    End If
    ReDim parsed(1 To n)
    For i = 1 To n
        parsed(i) = ParseLine(CStr(lines(i - 1)), d)
        If UBound(parsed(i)) + 1 > maxC Then maxC = UBound(parsed(i)) + 1
    Next
    ReDim raw(1 To n, 1 To maxC)
    For i = 1 To n
        p = parsed(i)
        For j = 0 To UBound(p): raw(i, j + 1) = p(j): Next
    Next
    ReadDelimited = Normalize(raw, st, True)
End Function

'----------------------------------------------------------
' 読み込んだ表を「見出し1行＋データ行」に整える
'   isText=True のときは文字列を数値に変換（先頭0の番号は文字のまま）
'----------------------------------------------------------
Private Function Normalize(raw As Variant, st As ImportSetting, ByVal isText As Boolean) As Variant
    Dim rr As Long, rC As Long, nC As Long, r As Long, c As Long, first As Long
    Dim rowsOk() As Long, cnt As Long, outArr() As Variant, hasVal As Boolean, v As Variant

    rr = UBound(raw, 1): rC = UBound(raw, 2)
    If st.HeaderRow > rr Then Err.Raise vbObjectError + 13, , "見出し行（" & st.HeaderRow & "行目）がありません"
    For c = rC To 1 Step -1
        If Len(Trim$(SafeStr(raw(st.HeaderRow, c)))) > 0 Then nC = c: Exit For
    Next
    If nC = 0 Then Err.Raise vbObjectError + 13, , "見出し行が空です（見出し行の設定を確認してください）"

    first = st.DataRow
    If first <= st.HeaderRow Then first = st.HeaderRow + 1
    ReDim rowsOk(1 To Application.Max(1, rr))
    For r = first To rr
        hasVal = False
        For c = 1 To nC
            If Len(SafeStr(raw(r, c))) > 0 Then hasVal = True: Exit For
        Next
        If hasVal Then cnt = cnt + 1: rowsOk(cnt) = r
    Next

    ReDim outArr(1 To cnt + 1, 1 To nC)
    For c = 1 To nC: outArr(1, c) = Trim$(SafeStr(raw(st.HeaderRow, c))): Next
    For r = 1 To cnt
        For c = 1 To nC
            v = raw(rowsOk(r), c)
            If isText Then
                outArr(r + 1, c) = ConvText(SafeStr(v))
            Else
                outArr(r + 1, c) = v
            End If
        Next
    Next
    Normalize = outArr
End Function

'----------------------------------------------------------
' シートへ書き出し（文字列の列は文字列書式にして、先頭0や時刻文字列を守る）
'----------------------------------------------------------
Private Sub WriteSheet(ws As Worksheet, outArr As Variant, ByVal dateIdx As Long)
    Dim nR As Long, nC As Long, r As Long, c As Long
    Dim hasStr As Boolean, hasDate As Boolean, hasTime As Boolean, v As Variant

    nR = UBound(outArr, 1): nC = UBound(outArr, 2)
    ws.Cells.Clear
    ws.Rows(1).NumberFormat = "@"
    If nR >= 2 Then
        For c = 1 To nC
            hasStr = False: hasDate = False: hasTime = False
            For r = 2 To nR
                v = outArr(r, c)
                Select Case VarType(v)
                    Case vbString
                        If Len(v) > 0 Then hasStr = True
                    Case vbDate
                        hasDate = True
                        If CDbl(v) <> Int(CDbl(v)) Then hasTime = True
                End Select
            Next
            With ws.Range(ws.Cells(2, c), ws.Cells(nR, c))
                If c = dateIdx Or c = nC Or (hasDate And Not hasStr) Then
                    .NumberFormat = IIf(hasTime, "yyyy/mm/dd hh:mm:ss", "yyyy/mm/dd")
                ElseIf hasStr Then
                    .NumberFormat = "@"
                End If
            End With
        Next
    End If
    ws.Range("A1").Resize(nR, nC).Value = outArr
    With ws.Range("A1").Resize(1, nC)
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    ws.Range(ws.Cells(1, nC - 1), ws.Cells(1, nC)).Interior.Color = RGB(217, 217, 217)
    ws.Range("A1").Resize(Application.Min(nR, 300), nC).Columns.AutoFit
End Sub

'----------------------------------------------------------
' 取込設定シートの読込
'----------------------------------------------------------
Private Function LoadSettings(ws As Worksheet, sts() As ImportSetting) As Long
    Dim hr As Long, r As Long, n As Long
    Dim cName As Long, cKw As Long, cExt As Long, cFmt As Long, cCs As Long, cDl As Long
    Dim cSh As Long, cHr As Long, cDr As Long, cDt As Long, cPs As Long, cTg As Long
    Dim cMd As Long, cMv As Long, cEn As Long

    hr = FindRowInCol(ws, 1, "No", 40)
    If hr = 0 Then Err.Raise vbObjectError + 1, , "取込設定シートの見出し行（No）が見つかりません"
    cName = NeedCol(ws, hr, "データ名")
    cKw = NeedCol(ws, hr, "ファイル名キーワード")
    cExt = HeadCol(ws, hr, "拡張子")
    cFmt = NeedCol(ws, hr, "実際の形式")
    cCs = HeadCol(ws, hr, "文字コード")
    cDl = HeadCol(ws, hr, "区切り文字")
    cSh = HeadCol(ws, hr, "読込シート名")
    cHr = HeadCol(ws, hr, "見出し行")
    cDr = HeadCol(ws, hr, "データ開始行")
    cDt = HeadCol(ws, hr, "日付列名")
    cPs = HeadCol(ws, hr, "人別管理列名")
    cTg = HeadCol(ws, hr, "取込先シート")
    cMd = NeedCol(ws, hr, "取込方式")
    cMv = HeadCol(ws, hr, "取込後の移動先")
    cEn = NeedCol(ws, hr, "有効")

    ReDim sts(1 To 200)
    r = hr + 1
    Do While Len(SafeStr(ws.Cells(r, 1).Value)) > 0 And IsNumeric(ws.Cells(r, 1).Value)
        If Len(CellStr(ws, r, cName)) > 0 Then
            n = n + 1
            With sts(n)
                .RowNo = r
                .DataName = CellStr(ws, r, cName)
                .keyword = CellStr(ws, r, cKw)
                .ext = CellStr(ws, r, cExt)
                .Fmt = CellStr(ws, r, cFmt)
                .Charset = CellStr(ws, r, cCs)
                .delim = CellStr(ws, r, cDl)
                .SrcSheet = CellStr(ws, r, cSh)
                .HeaderRow = Val(CellStr(ws, r, cHr)): If .HeaderRow < 1 Then .HeaderRow = 1
                .DataRow = Val(CellStr(ws, r, cDr)): If .DataRow <= .HeaderRow Then .DataRow = .HeaderRow + 1
                .DateCol = CellStr(ws, r, cDt)
                .PersonCol = CellStr(ws, r, cPs)
                .Target = CellStr(ws, r, cTg): If Len(.Target) = 0 Then .Target = .keyword
                .mode = NormalizeMode(CellStr(ws, r, cMd))
                .MoveTo = CellStr(ws, r, cMv)
                .Enabled = (CellStr(ws, r, cEn) = "○") And Len(.keyword) > 0
            End With
        End If
        r = r + 1
    Loop
    LoadSettings = n
End Function

'----------------------------------------------------------
' 取込設定シート③（取込状況）の更新
'----------------------------------------------------------
Private Sub UpdateSettingStatus(ws As Worksheet, sts() As ImportSetting, ByVal n As Long)
    Dim hr As Long, cTime As Long, cFile As Long, cPer As Long, cCnt As Long, i As Long, a As Variant, d As Object

    hr = FindRowInCol(ws, 1, "No", 40)
    cTime = HeadCol(ws, hr, "最終取込日時")
    cFile = HeadCol(ws, hr, "最終取込ファイル名")
    cPer = HeadCol(ws, hr, "取込済データ期間")
    cCnt = HeadCol(ws, hr, "取込件数")
    For i = 1 To n
        Set d = GetStats(sts(i))
        a = d("")
        If cTime > 0 Then ws.Cells(sts(i).RowNo, cTime).Value = a(3): ws.Cells(sts(i).RowNo, cTime).NumberFormat = "yyyy/mm/dd hh:mm"
        If cFile > 0 Then ws.Cells(sts(i).RowNo, cFile).Value = a(4)
        If cPer > 0 Then
            If IsEmpty(a(0)) Then
                ws.Cells(sts(i).RowNo, cPer).Value = Empty
            Else
                ws.Cells(sts(i).RowNo, cPer).Value = Format(a(0), "yyyy/mm/dd") & "?" & Format(a(1), "yyyy/mm/dd")
            End If
        End If
        If cCnt > 0 Then ws.Cells(sts(i).RowNo, cCnt).Value = a(2)
    Next
End Sub

'----------------------------------------------------------
' 操作シートの取込状況（データ別／人別）を描き直す
'----------------------------------------------------------
Private Sub RefreshStatus()
    Dim ws As Worksheet, sts() As ImportSetting, n As Long, i As Long, r As Long
    Dim d As Object, a As Variant, keys As Variant, j As Long, sh As Worksheet

    Set ws = ThisWorkbook.Worksheets(SH_MENU)
    n = LoadSettings(ThisWorkbook.Worksheets(SH_SET), sts)

    ' ②の続き（取込状況）から下を描き直す。③④もここで描く（表の長さが変わるため）
    With ws.Range(ws.Cells(MENU_STATUS_ROW, 1), ws.Cells(MENU_STATUS_ROW + 2000, 12))
        .Hyperlinks.Delete
        .UnMerge
        .Clear
        .Font.Name = "Meiryo UI"
        .Font.Size = 10
        .EntireRow.RowHeight = ws.StandardHeight
    End With

    r = MENU_STATUS_ROW
    SectionTitle ws, r, "■ データ別 取込状況"
    ws.Cells(r, LAST_COL).Value = "表示更新：" & Format(Now, "yyyy/mm/dd hh:nn")
    ws.Cells(r, LAST_COL).HorizontalAlignment = xlRight: ws.Cells(r, LAST_COL).Font.Size = 9
    r = r + 1
    StatusHeader ws, r, Array("データ名", "取込先シート", "取込方法")
    For i = 1 To n
        r = r + 1
        Set d = GetStats(sts(i))
        a = d("")
        ws.Cells(r, 2).Value = sts(i).DataName
        Set sh = FindSheet(sts(i).Target)
        If sh Is Nothing Then
            ws.Cells(r, 3).Value = sts(i).Target & "（未作成）"
        Else
            ws.Hyperlinks.Add ws.Cells(r, 3), "", "'" & sts(i).Target & "'!A1", , sts(i).Target
        End If
        If sts(i).mode = MODE_REPLACE Then
            ws.Cells(r, 4).Value = "書き換え"
            ws.Cells(r, 4).Interior.Color = RGB(255, 242, 204)
        Else
            ws.Cells(r, 4).Value = "蓄積"
            ws.Cells(r, 4).Interior.Color = CLR_GREEN_LIGHT
        End If
        ws.Cells(r, 4).HorizontalAlignment = xlCenter
        StatusValues ws, r, sts(i), a
    Next
    r = r + 1
    ws.Cells(r, 2).Value = "取込方法　蓄積＝取り込むたびにデータを貯めていく（同じ日付・同じ人のデータは新しいファイルの内容で上書き）" & _
                           "／書き換え＝毎回シートを空にして、最新ファイルの内容だけにする"
    ws.Cells(r, 2).Font.Size = 9: ws.Cells(r, 2).Font.Color = RGB(89, 89, 89)
    r = r + 1
    ws.Cells(r, 2).Value = "状態　　　蓄積＝昨日分まで入っているか（土日・休日一覧の日は数えない）／書き換え＝今日取り込んだか"
    ws.Cells(r, 2).Font.Size = 9: ws.Cells(r, 2).Font.Color = RGB(89, 89, 89)

    For i = 1 To n
        If Len(sts(i).PersonCol) > 0 Then
            Set d = GetStats(sts(i))
            r = r + 2
            SectionTitle ws, r, "■ " & sts(i).DataName & " 人別取込状況（" & sts(i).PersonCol & "ごと）"
            r = r + 1
            StatusHeader ws, r, Array(sts(i).PersonCol, "", "")
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Merge
            keys = SortedKeys(d)
            If UBound(keys) < 0 Then
                r = r + 1
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Merge
                ws.Cells(r, 2).Value = "（未取込）"
                StatusBorder ws, r
            Else
                For j = 0 To UBound(keys)
                    r = r + 1
                    ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Merge
                    ws.Cells(r, 2).NumberFormat = "@"
                    If sts(i).Target = WorkSourceSheet() Then ws.Cells(r, 2).Value = PersonDisplayName(CStr(keys(j))) Else ws.Cells(r, 2).Value = keys(j)
                    StatusValues ws, r, sts(i), d(keys(j))
                Next
            End If
        End If
    Next

    ' ②の終わりの区切り
    r = r + 1
    SectionEnd ws, r, CLR_GREEN

    ' ③ 現在の状況
    r = r + 2
    MenuBand ws, r, "③ 現在の状況", CLR_ORANGE
    r = r + 1
    MenuItem ws, r, "確認", "給与手当の「申請種別 × ステータス・フラグ」の件数を見る（取り込むと自動で最新になります）", 9
    MenuLink ws, r, "現在の状況", "現在の状況を開く"
    r = r + 1
    MenuButtonRow ws, r, "記録", "btnRecord", "記録する", "状況を記録", CLR_ORANGE, _
                  "今日の件数を「現在の状況（積み上げ）」に残します（取込のあと1日1回。同じ日にもう一度押すと上書き）"
    MenuLink ws, r, "現在の状況（積み上げ）", "積み上げを見る"
    r = r + 1
    SectionEnd ws, r, CLR_ORANGE

    ' ④ 処理実績
    r = r + 2
    MenuBand ws, r, "④ 処理実績", CLR_BLUE
    r = r + 1
    MenuItem ws, r, "日別", "案件処理件数一覧を、日ごと × 申請種別ごとに Status 別（登録／ＳＶ承認／確認済締処理登録）で見る（取り込むと自動で最新になります）", 9
    ws.Rows(r).RowHeight = 32
    MenuLink ws, r, "日別処理件数", "日別処理件数を開く"
    r = r + 1
    MenuButtonRow ws, r, "記録", "btnDailyRecord", "日別を記録する", "日別を記録", CLR_BLUE, _
                  "日別処理件数の値を「日別処理件数（積み上げ）」に数字で残します（取込のあとに押す。記録済みの日の値が変わっていたら上書きするか確認）"
    MenuLink ws, r, "日別処理件数（積み上げ）", "積み上げを見る"
    r = r + 1
    SectionEnd ws, r, CLR_BLUE

    ' ⑤ 進捗
    r = r + 2
    MenuBand ws, r, "⑤ 進捗", CLR_PURPLE
    r = r + 1
    MenuItem ws, r, "確認", "届出ごとに、実績（確認中・認定件数（2次）・最古申請日）と予定（見込み申請・確認中・残数）を日別に見る。見込み認定件数（2次）の行に1日あたりの件数を入れると残数を予測します（入れた日から先へ引き継ぎ）。＋で開くと1次・2次の内訳", 9
    MenuLink ws, r, "進捗（簡易版）", "進捗（簡易版）を開く"
    r = r + 1
    MenuItem ws, r, "見込み", "見込み申請件数のもとになる昨年度の日別申請実績を入力する（進捗の明日以降の残数に使います）", 9
    MenuLink ws, r, "見込み申請（昨年実績）", "見込み申請を入力"
    r = r + 1
    MenuButtonRow ws, r, "詳細", "btnDetail", "詳細を更新", "詳細を更新", CLR_PURPLE, _
                  "事由×工程（1次・2次）ごとの進捗を作り直します。人の行に時間を入れると処理見込みが出ます（取込のあとは自動。手で入れた時間・件/時・比率は残ります）"
    MenuLink ws, r, "進捗（詳細）", "進捗（詳細）を開く"
    r = r + 1
    MenuItem ws, r, "スキル", "人ごとの1日の上限時間と、事由×工程ごとの件/時（稼働記録の実績から自動。数字を入れると上書き）", 9
    MenuLink ws, r, "スキルマップ", "スキルマップを開く"
    r = r + 1
    MenuItem ws, r, "比率", "見込み申請・残数を事由に分ける比率（直近4週の実績から自動。上書き可）と、扶養の届出事由→事由の対応表", 9
    MenuLink ws, r, "事由比率", "事由比率を開く"
    r = r + 1
    SectionEnd ws, r, CLR_PURPLE

    ' ⑥ 稼働実績
    r = r + 2
    MenuBand ws, r, "⑥ 稼働実績", CLR_TEAL
    r = r + 1
    MenuButtonRow ws, r, "集計", "btnWork", "稼働実績を更新", "稼働実績を更新", CLR_TEAL, _
                  "稼働記録を、届出×工程・業務別に、全員分と個人別で日ごとに集計します（②で稼働記録を取り込むと自動で更新）"
    MenuLink ws, r, "稼働実績", "稼働実績を開く"
    r = r + 1
    MenuItem ws, r, "マスタ", "業務ID → 届出・工程（1次／2次）・集計区分の対応表。新しい業務IDは自動で追加されるので、届出・工程を確認・入力してください", 9
    MenuLink ws, r, "業務マスタ", "業務マスタを開く"
    r = r + 1
    SectionEnd ws, r, CLR_TEAL
End Sub

'--- 操作シートの部品 ---
Private Sub SectionTitle(ws As Worksheet, ByVal r As Long, ByVal text As String)
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, LAST_COL))
        .Interior.Color = CLR_GREEN_LIGHT
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = CLR_GREEN
    End With
    ws.Cells(r, 2).Value = text
    ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).Font.Size = 11
    ws.Cells(r, 2).Font.Color = CLR_GREEN
    ws.Rows(r).RowHeight = 22
End Sub

Private Sub SectionEnd(ws As Worksheet, ByVal r As Long, ByVal clr As Long)
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, LAST_COL)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThick: .Color = clr
    End With
    ws.Rows(r).RowHeight = 6
End Sub

Private Sub MenuBand(ws As Worksheet, ByVal r As Long, ByVal text As String, ByVal clr As Long)
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, LAST_COL))
        .Merge
        .Value = text
        .Interior.Color = clr
        .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 12
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 24
End Sub

Private Sub MenuItem(ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal text As String, ByVal endCol As Long)
    ws.Cells(r, 2).Value = label: ws.Cells(r, 2).Font.Bold = True
    With ws.Range(ws.Cells(r, 3), ws.Cells(r, endCol))
        .Merge
        .Value = text
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 20
End Sub

Private Sub MenuLink(ws As Worksheet, ByVal r As Long, ByVal sheetName As String, ByVal caption As String)
    ws.Hyperlinks.Add ws.Cells(r, LAST_COL), "", "'" & sheetName & "'!A1", , ChrW(&H25B6) & " " & caption
    ws.Cells(r, LAST_COL).VerticalAlignment = xlCenter
End Sub

' ボタン＋説明の行（ボタンはなければ作り、あれば位置だけ合わせる）
Private Sub MenuButtonRow(ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal shpName As String, _
                          ByVal caption As String, ByVal macro As String, ByVal clr As Long, ByVal text As String)
    Dim shp As Shape
    ws.Rows(r).RowHeight = 40
    ws.Cells(r, 2).Value = label: ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).VerticalAlignment = xlCenter
    With ws.Range(ws.Cells(r, 4), ws.Cells(r, 9))
        .Merge
        .Value = text
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Font.Color = RGB(89, 89, 89)
    End With
    On Error Resume Next
    Set shp = ws.Shapes(shpName)
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 150, 30)
        shp.Name = shpName
        With shp.TextFrame2
            .TextRange.text = caption
            .TextRange.Font.Size = 12
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Name = "Meiryo UI"
            .TextRange.Font.NameFarEast = "Meiryo UI"
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
        shp.Fill.ForeColor.RGB = clr
        shp.Line.Visible = msoFalse
        shp.OnAction = macro
    End If
    shp.Left = ws.Cells(r, 3).Left + 2
    shp.Top = ws.Cells(r, 3).Top + 5
End Sub

Private Sub StatusHeader(ws As Worksheet, ByVal r As Long, ByVal first3 As Variant)
    Dim hs As Variant, c As Long
    hs = Array(first3(0), first3(1), first3(2), "いつのデータから", "いつのデータまで", "件数", "最終取込日時", "最終取込ファイル", "状態")
    For c = 0 To UBound(hs)
        ws.Cells(r, c + 2).Value = hs(c)
    Next
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, LAST_COL))
        .Font.Bold = True: .Font.Color = vbWhite
        .Interior.Color = CLR_GREEN
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(166, 166, 166)
    End With
End Sub

Private Sub StatusValues(ws As Worksheet, ByVal r As Long, st As ImportSetting, a As Variant)
    Dim s As String
    ws.Cells(r, 5).Value = a(0): ws.Cells(r, 5).NumberFormat = "yyyy/mm/dd(aaa)"
    ws.Cells(r, 6).Value = a(1): ws.Cells(r, 6).NumberFormat = "yyyy/mm/dd(aaa)"
    ws.Cells(r, 7).Value = a(2): ws.Cells(r, 7).NumberFormat = "#,##0"
    ws.Cells(r, 8).Value = a(3): ws.Cells(r, 8).NumberFormat = "yyyy/mm/dd hh:mm"
    ws.Cells(r, 9).Value = a(4)
    s = StatusText(st, a)
    ws.Cells(r, LAST_COL).Value = s
    If Left$(s, 1) = "○" Then
        ws.Cells(r, LAST_COL).Font.Color = RGB(84, 130, 53)
    Else
        ws.Cells(r, LAST_COL).Font.Color = RGB(192, 0, 0)
        ws.Cells(r, LAST_COL).Interior.Color = RGB(252, 228, 214)
    End If
    ws.Range(ws.Cells(r, 5), ws.Cells(r, 8)).HorizontalAlignment = xlCenter
    StatusBorder ws, r
End Sub

Private Sub StatusBorder(ws As Worksheet, ByVal r As Long)
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, LAST_COL)).Borders
        .LineStyle = xlContinuous: .Color = RGB(166, 166, 166)
    End With
End Sub

' 状態：書き換えは「今日取り込んだか」、蓄積は「昨日分まで入っているか（土日・休日一覧の日は除く）」で判定
Private Function StatusText(st As ImportSetting, a As Variant) As String
    Dim nd As Long, hol As Variant, firstMiss As Date
    If a(2) = 0 Or IsEmpty(a(3)) Then StatusText = "△ 未取込": Exit Function
    If st.mode = MODE_REPLACE Then
        If Int(CDbl(a(3))) < CDbl(Date) Then
            StatusText = "△ 本日未取込（前回 " & Format(a(3), "m/d") & "）"
        Else
            StatusText = "○ 本日取込済"
        End If
        Exit Function
    End If
    If IsEmpty(a(1)) Then StatusText = "○ 取込済": Exit Function
    If a(1) >= Date Then StatusText = "！今日のデータを含む（確定前）": Exit Function
    If a(1) >= Date - 1 Then StatusText = "○ 昨日分まで取込済": Exit Function
    hol = Holidays()
    If IsEmpty(hol) Then
        nd = Application.WorksheetFunction.NetworkDays(a(1) + 1, Date - 1)
        If nd > 0 Then firstMiss = Application.WorksheetFunction.WorkDay(a(1), 1)
    Else
        nd = Application.WorksheetFunction.NetworkDays(a(1) + 1, Date - 1, hol)
        If nd > 0 Then firstMiss = Application.WorksheetFunction.WorkDay(a(1), 1, hol)
    End If
    If nd <= 0 Then
        StatusText = "○ 昨日分まで取込済"
    Else
        StatusText = "△ " & Format(firstMiss, "m/d") & "以降 未取込（" & nd & "営業日）"
    End If
End Function

' 休日一覧シートのA列（見出し行の下）の日付
Private Function Holidays() As Variant
    Dim ws As Worksheet, lastR As Long
    Set ws = FindSheet(SH_HOL)
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < HOL_HEADER_ROW + 1 Then Exit Function
    Holidays = ws.Range(ws.Cells(HOL_HEADER_ROW + 1, 1), ws.Cells(lastR, 1)).Value
End Function

'----------------------------------------------------------
' 取込先シートの集計：キー ""=全体、それ以外=人別
'   値 = Array(いつのデータから, いつのデータまで, 件数, 最終取込日時, 最終取込ファイル)
'----------------------------------------------------------
Private Function GetStats(st As ImportSetting) As Object
    Dim d As Object, ws As Worksheet, v As Variant, nR As Long, nC As Long, r As Long
    Dim iDate As Long, iPer As Long, iFile As Long, iTime As Long

    Set d = CreateObject("Scripting.Dictionary")
    d("") = Array(Empty, Empty, 0, Empty, "")
    Set GetStats = d
    Set ws = FindSheet(st.Target)
    If ws Is Nothing Then Exit Function
    nR = LastRow(ws): nC = LastColOfRow1(ws)
    If nR < 2 Or nC < 3 Then Exit Function
    v = ws.Range("A1").Resize(nR, nC).Value
    iDate = HeaderIndex(v, nC, st.DateCol)
    iPer = HeaderIndex(v, nC, st.PersonCol)
    iFile = HeaderIndex(v, nC, COL_FILE)
    iTime = HeaderIndex(v, nC, COL_TIME)
    For r = 2 To nR
        AddStat d, "", v, r, iDate, iFile, iTime
        If iPer > 0 Then AddStat d, SafeStr(v(r, iPer)), v, r, iDate, iFile, iTime
    Next
End Function

Private Sub AddStat(d As Object, ByVal key As String, v As Variant, ByVal r As Long, _
                    ByVal iDate As Long, ByVal iFile As Long, ByVal iTime As Long)
    Dim a As Variant, dt As Variant, t As Variant
    If d.Exists(key) Then a = d(key) Else a = Array(Empty, Empty, 0, Empty, "")
    a(2) = a(2) + 1
    If iDate > 0 Then
        dt = v(r, iDate)
        If VarType(dt) = vbDate Then
            dt = CDate(Int(CDbl(dt)))
            If IsEmpty(a(0)) Then a(0) = dt: a(1) = dt
            If dt < a(0) Then a(0) = dt
            If dt > a(1) Then a(1) = dt
        End If
    End If
    If iTime > 0 Then
        t = v(r, iTime)
        If VarType(t) = vbDate Then
            If IsEmpty(a(3)) Then
                a(3) = t
                If iFile > 0 Then a(4) = SafeStr(v(r, iFile))
            ElseIf t > a(3) Then
                a(3) = t
                If iFile > 0 Then a(4) = SafeStr(v(r, iFile))
            End If
        End If
    End If
    d(key) = a
End Sub

'----------------------------------------------------------
' 取込履歴
'----------------------------------------------------------
Private Sub WriteLog(ByVal runTime As Date, st As ImportSetting, ByVal path As String, ByVal fileTime As Date, _
                     ByVal nRead As Long, ByVal nDel As Long, ByVal dMin As Variant, ByVal dMax As Variant, _
                     ByVal errMsg As String, ByVal msg As String)
    Dim ws As Worksheet, r As Long, fn As String, fd As Variant
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r <= LOG_HEADER_ROW Then r = LOG_HEADER_ROW + 1
    fn = Mid$(path, InStrRev(path, "\") + 1)
    fd = FileNameDate(fn, st.keyword)

    ws.Cells(r, 1).Value = runTime: ws.Cells(r, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Cells(r, 2).Value = st.DataName
    ws.Cells(r, 3).NumberFormat = "@": ws.Cells(r, 3).Value = fn
    ws.Cells(r, 4).Value = fd: ws.Cells(r, 4).NumberFormat = "yyyy/mm/dd"
    ws.Cells(r, 5).Value = fileTime: ws.Cells(r, 5).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Cells(r, 6).Value = dMin: ws.Cells(r, 6).NumberFormat = "yyyy/mm/dd"
    ws.Cells(r, 7).Value = dMax: ws.Cells(r, 7).NumberFormat = "yyyy/mm/dd"
    ws.Cells(r, 8).Value = nRead
    ws.Cells(r, 9).Value = IIf(Len(errMsg) > 0, 0, nRead)
    ws.Cells(r, 10).Value = nDel
    If Len(errMsg) > 0 Then
        ws.Cells(r, 11).Value = "エラー": ws.Cells(r, 12).Value = errMsg
        ws.Cells(r, 11).Font.Color = RGB(192, 0, 0)
    ElseIf Len(msg) > 0 Then
        ws.Cells(r, 11).Value = "注意": ws.Cells(r, 12).Value = msg
        ws.Cells(r, 11).Font.Color = RGB(197, 90, 17)
    Else
        ws.Cells(r, 11).Value = "OK"
    End If
    If Len(errMsg) = 0 And nDel > 0 Then
        ws.Cells(r, 12).Value = AddMsg(SafeStr(ws.Cells(r, 12).Value), "同じ期間の既存データ " & nDel & "件を入れ替え")
    End If
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Borders.Color = RGB(166, 166, 166)
End Sub

'----------------------------------------------------------
' ファイル操作
'----------------------------------------------------------
Private Function FindFiles(fso As Object, ByVal dirPath As String, st As ImportSetting) As Collection
    Dim res As New Collection, f As Object, ext As String
    Dim paths() As String, times() As Date, cnt As Long, i As Long, j As Long, tp As String, tt As Date

    ext = LCase$(Replace(st.ext, ".", ""))
    For Each f In fso.GetFolder(dirPath).files
        If Left$(f.Name, 2) <> "~$" And LCase$(Left$(f.Name, Len(st.keyword))) = LCase$(st.keyword) Then
            If Len(ext) = 0 Or LCase$(fso.GetExtensionName(f.Name)) = ext Then
                cnt = cnt + 1
                ReDim Preserve paths(1 To cnt): ReDim Preserve times(1 To cnt)
                paths(cnt) = f.path: times(cnt) = f.DateLastModified
            End If
        End If
    Next
    ' 更新日時が古い順（書き換えでは最後＝最新のファイルの内容が残る）
    For i = 2 To cnt
        tp = paths(i): tt = times(i): j = i - 1
        Do While j >= 1
            If times(j) <= tt Then Exit Do
            paths(j + 1) = paths(j): times(j + 1) = times(j): j = j - 1
        Loop
        paths(j + 1) = tp: times(j + 1) = tt
    Next
    For i = 1 To cnt: res.Add paths(i): Next
    Set FindFiles = res
End Function

Private Sub MoveToDone(fso As Object, ByVal path As String, ByVal doneDir As String)
    Dim dest As String
    If Len(doneDir) = 0 Then Err.Raise vbObjectError + 20, , "取込済フォルダが設定されていません"
    If Not fso.FolderExists(doneDir) Then fso.CreateFolder doneDir
    dest = doneDir & "\" & fso.GetFileName(path)
    If fso.FileExists(dest) Then
        dest = doneDir & "\" & fso.GetBaseName(path) & "_取込" & Format(Now, "yyyymmdd_hhnnss") & "." & fso.GetExtensionName(path)
    End If
    fso.MoveFile path, dest
End Sub

Private Function ReadTextFile(ByVal path As String, ByVal cS As String) As String
    Dim stm As Object, s As String
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = cS
    stm.Open
    stm.LoadFromFile path
    s = stm.ReadText(-1)
    stm.Close
    If Left$(s, 1) = ChrW(&HFEFF) Then s = Mid$(s, 2)
    ReadTextFile = s
End Function

'----------------------------------------------------------
' 共通部品
'----------------------------------------------------------
Private Function ResolvePath(ByVal p As String) As String
    p = Trim$(p)
    If Len(p) = 0 Then Exit Function
    If p = "." Then
        p = ThisWorkbook.path
    ElseIf Left$(p, 2) = ".\" Then
        p = ThisWorkbook.path & Mid$(p, 2)
    ElseIf Not (Mid$(p, 2, 1) = ":" Or Left$(p, 2) = "\\") Then
        p = ThisWorkbook.path & "\" & p
    End If
    Do While Right$(p, 1) = "\" And Len(p) > 3
        p = Left$(p, Len(p) - 1)
    Loop
    ResolvePath = p
End Function

Private Function LabelValue(ws As Worksheet, ByVal label As String) As String
    Dim r As Long
    r = FindRowInCol(ws, 1, label, 20)
    If r > 0 Then LabelValue = Trim$(SafeStr(ws.Cells(r, 2).Value))
End Function

Private Function FindRowInCol(ws As Worksheet, ByVal col As Long, ByVal text As String, ByVal maxRow As Long) As Long
    Dim r As Long
    For r = 1 To maxRow
        If Trim$(SafeStr(ws.Cells(r, col).Value)) = text Then FindRowInCol = r: Exit Function
    Next
End Function

' 見出しの1行目（改行の前）で列を探す
Private Function HeadCol(ws As Worksheet, ByVal hr As Long, ByVal nm As String) As Long
    Dim c As Long, lastC As Long, h As String
    lastC = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        h = Trim$(Split(SafeStr(ws.Cells(hr, c).Value) & vbLf, vbLf)(0))
        If h = nm Then HeadCol = c: Exit Function
    Next
End Function

Private Function NeedCol(ws As Worksheet, ByVal hr As Long, ByVal nm As String) As Long
    NeedCol = HeadCol(ws, hr, nm)
    If NeedCol = 0 Then Err.Raise vbObjectError + 1, , "取込設定シートに「" & nm & "」列がありません"
End Function

Private Function CellStr(ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    If c > 0 Then CellStr = Trim$(SafeStr(ws.Cells(r, c).Value))
End Function

Private Function HeaderIndex(arr As Variant, ByVal nC As Long, ByVal nm As String) As Long
    Dim c As Long
    If Len(nm) = 0 Then Exit Function
    For c = 1 To nC
        If Trim$(SafeStr(arr(1, c))) = nm Then HeaderIndex = c: Exit Function
    Next
End Function

' 既存シートの見出し（管理列2つを除く）と新しい見出しが同じか
Private Function SameHeader(exHdr As Variant, ByVal exC As Long, data As Variant, ByVal nC As Long) As Boolean
    Dim c As Long
    If exC <> nC + 2 Then Exit Function
    For c = 1 To nC
        If SafeStr(exHdr(1, c)) <> SafeStr(data(1, c)) Then Exit Function
    Next
    SameHeader = True
End Function

Private Function FindSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nm Then Set FindSheet = ws: Exit Function
    Next
End Function

Private Function GetOrCreateSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    If nm = SH_MENU Or nm = SH_SET Or nm = SH_LOG Or nm = SH_HOL Then
        Err.Raise vbObjectError + 5, , "取込先シート名「" & nm & "」は使えません"
    End If
    Set ws = FindSheet(nm)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
        ws.Tab.Color = RGB(191, 191, 191)
    End If
    Set GetOrCreateSheet = ws
End Function

Private Function LastRow(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then LastRow = f.Row
End Function

Private Function LastColOfRow1(ws As Worksheet) As Long
    LastColOfRow1 = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If LastColOfRow1 = 1 And Len(SafeStr(ws.Cells(1, 1).Value)) = 0 Then LastColOfRow1 = 0
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or IsObject(v) Then Exit Function
    SafeStr = CStr(v)
End Function

' 文字列→値：数値に見えるものは数値に（先頭0の番号などは文字のまま）
Private Function ConvText(ByVal s As String) As Variant
    Static rE As Object
    If rE Is Nothing Then Set rE = NewRegExp("^-?(0|[1-9][0-9]{0,14})(\.[0-9]+)?$")
    s = Trim$(s)
    If Len(s) = 0 Then
        ConvText = Empty
    ElseIf rE.Test(s) Then
        ConvText = Val(s)
    Else
        ConvText = s
    End If
End Function

Private Function HtmlText(ByVal s As String) As String
    Static reNum As Object
    Dim m As Object, i As Long
    If reNum Is Nothing Then Set reNum = NewRegExp("&#(\d+);")
    s = Replace(Replace(Replace(s, vbCr, ""), vbLf, ""), vbTab, "")
    Set m = reNum.Execute(s)
    For i = m.Count - 1 To 0 Step -1
        s = Replace(s, m(i).Value, ChrW(CLng(m(i).SubMatches(0))))
    Next
    s = Replace(s, "&nbsp;", " ")
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", """")
    s = Replace(s, "&#39;", "'")
    s = Replace(s, "&amp;", "&")
    HtmlText = Trim$(s)
End Function

Private Function ParseLine(ByVal s As String, ByVal d As String) As Variant
    Dim res() As String, n As Long, i As Long, ch As String, cur As String, inQ As Boolean
    ReDim res(0 To 0)
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid$(s, i + 1, 1) = """" Then
                    cur = cur & """": i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        ElseIf ch = """" And Len(cur) = 0 Then
            inQ = True
        ElseIf ch = d Then
            res(n) = cur: cur = "": n = n + 1
            ReDim Preserve res(0 To n)
        Else
            cur = cur & ch
        End If
        i = i + 1
    Loop
    res(n) = cur
    ParseLine = res
End Function

Private Function DelimChar(ByVal delim As String, ByVal ext As String) As String
    Select Case delim
        Case "タブ": DelimChar = vbTab
        Case "カンマ": DelimChar = ","
        Case Else
            If LCase$(Replace(ext, ".", "")) = "csv" Then DelimChar = "," Else DelimChar = vbTab
    End Select
End Function

Private Function CharsetName(ByVal cS As String) As String
    Select Case UCase$(cS)
        Case "UTF-8", "UTF-8(BOM)": CharsetName = "utf-8"
        Case Else: CharsetName = "shift_jis"
    End Select
End Function

' ファイル名のキーワードより後ろにある8桁の日付（yyyymmdd）
Private Function FileNameDate(ByVal fileName As String, ByVal keyword As String) As Variant
    Static rE As Object
    Dim m As Object, s As String
    If rE Is Nothing Then Set rE = NewRegExp("(20\d{2})(\d{2})(\d{2})")
    s = Mid$(fileName, Len(keyword) + 1)
    Set m = rE.Execute(s)
    If m.Count > 0 Then
        On Error Resume Next
        FileNameDate = DateSerial(CInt(m(0).SubMatches(0)), CInt(m(0).SubMatches(1)), CInt(m(0).SubMatches(2)))
        On Error GoTo 0
    End If
End Function

Private Function SortedKeys(d As Object) As Variant
    Dim ks() As String, k As Variant, n As Long, i As Long, j As Long, t As String
    ReDim ks(0 To Application.Max(0, d.Count - 1))
    n = -1
    For Each k In d.keys
        If Len(k) > 0 Then n = n + 1: ks(n) = k
    Next
    If n < 0 Then SortedKeys = Array(): Exit Function
    ReDim Preserve ks(0 To n)
    For i = 1 To n
        t = ks(i): j = i - 1
        Do While j >= 0
            If ks(j) <= t Then Exit Do
            ks(j + 1) = ks(j): j = j - 1
        Loop
        ks(j + 1) = t
    Next
    SortedKeys = ks
End Function

' 取込方式の旧名称も受け付ける
Private Function NormalizeMode(ByVal m As String) As String
    Select Case m
        Case "", "全件置換", MODE_REPLACE: NormalizeMode = MODE_REPLACE
        Case "追記（重複除外）", MODE_MERGE: NormalizeMode = MODE_MERGE
        Case "追記（そのまま）", MODE_APPEND: NormalizeMode = MODE_APPEND
        Case Else: Err.Raise vbObjectError + 6, , "取込方式「" & m & "」が分かりません"
    End Select
End Function

Private Function AddMsg(ByVal msg As String, ByVal s As String) As String
    If Len(msg) > 0 Then AddMsg = msg & "／" & s Else AddMsg = s
End Function

Private Function NewRegExp(ByVal pattern As String) As Object
    Set NewRegExp = CreateObject("VBScript.RegExp")
    NewRegExp.pattern = pattern
    NewRegExp.Global = True
    NewRegExp.IgnoreCase = True
End Function


