Option Explicit

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' マクロが消されたときのために、保存のたびにコードの控えを「マクロ保存」シートへ書き出す
    マクロ保存_自動
End Sub