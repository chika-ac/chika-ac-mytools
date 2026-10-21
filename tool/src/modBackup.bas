Option Explicit

'==========================================================
' マクロ保存
'   セキュリティ対策などでマクロが消されたときに貼り直せるよう、
'   このブックのマクロのコードを「マクロ保存」シートへ書き出す。
'   ブックを保存するたびに自動で書き出す（ThisWorkbook の BeforeSave から呼ぶ）。
'==========================================================

Private Const SH_BAK As String = "マクロ保存"
Private Const NAME_PREFIX As String = "マクロ_"
Private Const FIRST_ROW As Long = 13          ' コードを書き始める行
Private Const CLR_HEAD As Long = 14348258     ' RGB(226,239,218)

Public Sub マクロ保存を更新()
    Dim msg As String
    msg = BackupMacros()
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "マクロ保存"
    Else
        ThisWorkbook.Worksheets(SH_BAK).Activate
        MsgBox "マクロのコードを「" & SH_BAK & "」シートに書き出しました。", vbInformation, "マクロ保存"
    End If
End Sub

' 保存時に呼ぶ（失敗してもメッセージは出さない）
Public Sub マクロ保存_自動()
    On Error Resume Next
    BackupMacros
End Sub

'----------------------------------------------------------
' 書き出し本体（失敗したときはメッセージを返す）
'----------------------------------------------------------
Public Function BackupMacros() As String
    Dim comps As Object, comp As Object, cm As Object
    Dim ws As Worksheet, pass As Long, i As Long, j As Long, r As Long, nRows As Long
    Dim mNames() As String, mKinds() As String, mLines() As Variant, cnt As Long, total As Long
    Dim arr() As Variant, hdr() As Long, ln As Variant, scr As Boolean

    On Error Resume Next
    Set comps = ThisWorkbook.VBProject.VBComponents
    If Err.Number <> 0 Or comps Is Nothing Then
        BackupMacros = "マクロのコードを読み取れませんでした。" & vbLf & _
            "ファイル → オプション → トラスト センター → トラスト センターの設定 → マクロの設定 で" & vbLf & _
            "「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェックを入れてください。"
        Exit Function
    End If
    On Error GoTo 0

    ' 書き出すモジュール：先に標準・クラス・フォーム、次にコードのあるブック／シートのモジュール
    ReDim mNames(1 To comps.Count): ReDim mKinds(1 To comps.Count): ReDim mLines(1 To comps.Count)
    For pass = 1 To 2
        For Each comp In comps
            Set cm = comp.CodeModule
            If (pass = 1) = (comp.Type <> 100) And cm.CountOfLines > 0 Then
                If HasCode(cm) Then
                    cnt = cnt + 1
                    mNames(cnt) = comp.Name
                    mKinds(cnt) = KindName(comp)
                    mLines(cnt) = Split(cm.lines(1, cm.CountOfLines), vbCrLf)
                    total = total + UBound(mLines(cnt)) + 1
                End If
            End If
        Next
    Next

    scr = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = GetBackupSheet()
    WriteGuide ws, cnt, total

    ' 見出し＋コード＋空行
    nRows = total + cnt * 2
    ReDim arr(1 To Application.Max(1, nRows), 1 To 2)
    ReDim hdr(1 To Application.Max(1, cnt))
    r = 0
    For i = 1 To cnt
        r = r + 1: hdr(i) = r
        arr(r, 1) = "■"
        arr(r, 2) = mNames(i) & "（" & mKinds(i) & "）　" & (UBound(mLines(i)) + 1) & "行"
        j = 0
        For Each ln In mLines(i)
            j = j + 1: r = r + 1
            arr(r, 1) = j
            arr(r, 2) = "'" & ln      ' 先頭の ' や = をそのまま文字として残す
        Next
        r = r + 1
    Next

    With ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(Application.Max(FIRST_ROW, ws.UsedRange.Row + ws.UsedRange.Rows.Count), 2))
        .Clear
    End With
    If nRows > 0 Then
        With ws.Cells(FIRST_ROW, 1).Resize(nRows, 2)
            .Value = arr
            .Font.Name = "ＭＳ ゴシック"
            .Font.Size = 10
            .VerticalAlignment = xlTop
        End With
        With ws.Cells(FIRST_ROW, 1).Resize(nRows, 1)
            .Font.Color = RGB(150, 150, 150)
            .HorizontalAlignment = xlRight
        End With
    End If

    ' モジュールの見出しと、名前ボックスから選べる範囲名
    DeleteNames
    For i = 1 To cnt
        r = FIRST_ROW + hdr(i) - 1
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 2))
            .Font.Name = "Meiryo UI"
            .Font.Bold = True
            .Font.Size = 11
            .Font.Color = RGB(0, 0, 0)
            .Interior.Color = CLR_HEAD
            .HorizontalAlignment = xlLeft
        End With
        ThisWorkbook.Names.Add Name:=NAME_PREFIX & mNames(i), _
            RefersTo:="='" & SH_BAK & "'!$B$" & (r + 1) & ":$B$" & (r + UBound(mLines(i)) + 1)
    Next
    Application.ScreenUpdating = scr
End Function

'----------------------------------------------------------
' 部品
'----------------------------------------------------------
Private Function GetBackupSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SH_BAK Then Set GetBackupSheet = ws: Exit Function
    Next
    Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SH_BAK
    ws.Tab.Color = RGB(191, 191, 191)
    ws.Columns(1).ColumnWidth = 6
    ws.Columns(2).ColumnWidth = 130
    ActiveWindow.DisplayGridlines = False
    Set GetBackupSheet = ws
End Function

Private Sub WriteGuide(ws As Worksheet, ByVal cnt As Long, ByVal total As Long)
    Dim t As Variant, i As Long
    t = Array( _
        "マクロ保存", _
        "このブックのマクロのコードの控えです。ブックを保存するたびに自動で最新になります（手動で更新するときは Alt+F8 →「マクロ保存を更新」）。", _
        "最終書き出し：" & Format(Now, "yyyy/mm/dd hh:nn") & "（" & cnt & "モジュール／" & total & "行）", _
        "", _
        "【マクロが消されたときの戻し方】", _
        "1. ブックを「Excel マクロ有効ブック（.xlsm）」で保存し直し、Alt+F11 で VBE を開く。", _
        "2. 標準モジュール：メニュー「挿入」→「標準モジュール」。F4 のプロパティで (オブジェクト名) を下の見出しのモジュール名（例：modImport）に変える。", _
        "3. 左上の名前ボックスで「" & NAME_PREFIX & "モジュール名」を選ぶと、そのコードの範囲（B列）が選ばれる。Ctrl+C でコピーし、VBE のモジュールに Ctrl+V で貼る。", _
        "4. ThisWorkbook：VBE 左の「ThisWorkbook」をダブルクリックし、「" & NAME_PREFIX & "ThisWorkbook」を同じようにコピーして貼る。", _
        "5. すべて貼ったら保存し、操作シートのボタンでマクロが動くか確認する。", _
        "※ 貼った先の先頭に Option Explicit が2行並んだら1行消してください。このシートは書き換えないでください（保存のたびに上書きされます）。")
    ws.Range("A1:B" & FIRST_ROW - 1).Clear
    For i = 0 To UBound(t)
        ws.Cells(i + 1, 1).Value = t(i)
    Next
    With ws.Range("A1:B" & FIRST_ROW - 1)
        .Font.Name = "Meiryo UI"
        .Font.Size = 10
    End With
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A5").Font.Bold = True
    ws.Range("A3").Font.Color = RGB(89, 89, 89)
    ws.Range("A11").Font.Color = RGB(192, 0, 0)
End Sub

Private Function HasCode(cm As Object) As Boolean
    Dim i As Long, s As String
    For i = 1 To cm.CountOfLines
        s = Trim$(cm.lines(i, 1))
        If Len(s) > 0 And Left$(s, 7) <> "Option " Then HasCode = True: Exit Function
    Next
End Function

Private Function KindName(comp As Object) As String
    Select Case comp.Type
        Case 1: KindName = "標準モジュール"
        Case 2: KindName = "クラスモジュール"
        Case 3: KindName = "ユーザーフォーム"
        Case 100
            If comp.Name = "ThisWorkbook" Then KindName = "ThisWorkbook に貼る" Else KindName = "シートのモジュールに貼る"
        Case Else: KindName = "その他"
    End Select
End Function

Private Sub DeleteNames()
    Dim i As Long
    For i = ThisWorkbook.Names.Count To 1 Step -1
        If Left$(ThisWorkbook.Names(i).Name, Len(NAME_PREFIX)) = NAME_PREFIX Then ThisWorkbook.Names(i).Delete
    Next
End Sub

