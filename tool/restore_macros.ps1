# 処理件数集計 のマクロを「マクロ保存」シートから組み直す
#   使い方（PowerShell）：
#     powershell -NoProfile -ExecutionPolicy Bypass -File restore_macros.ps1 [-Path ブックのパス] [-FromSrc]
#   -Path    省略時は、このフォルダの1つ上の 処理件数集計.xlsm（なければ .xlsx）
#   -FromSrc シートではなく、このフォルダの src\*.bas（テキストの控え）から組み直す
#   前提：Excel のトラストセンター →「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」がオン
param([string]$Path = '', [switch]$FromSrc)
$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
if ($Path -eq '') {
    $Path = Join-Path (Split-Path $here) '処理件数集計.xlsm'
    if (-not (Test-Path $Path)) { $Path = Join-Path (Split-Path $here) '処理件数集計.xlsx' }
}
$Path = (Resolve-Path $Path).Path
$dir = Split-Path $Path
$leaf = Split-Path $Path -Leaf

# 1. ブックが開いていないか
if (Test-Path -LiteralPath (Join-Path $dir ('~$' + $leaf))) { throw "ブックが開いています。保存して閉じてから実行してください：$Path" }

# 2. バックアップ（00.old）
$old = Join-Path $dir '00.old'
if (-not (Test-Path $old)) { New-Item -ItemType Directory $old | Out-Null }
$bak = Join-Path $old ('処理件数集計_マクロ復元前_' + (Get-Date -Format 'yyyyMMdd_HHmm') + [IO.Path]::GetExtension($Path))
Copy-Item -LiteralPath $Path $bak
"バックアップ：$bak"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$xl.AutomationSecurity = 3      # 開くときはマクロを動かさない
$wb = $xl.Workbooks.Open($Path)
try {
    try { $vbp = $wb.VBProject; $null = $vbp.VBComponents.Count }
    catch { throw 'VBA プロジェクトにアクセスできません。Excel のファイル → オプション → トラストセンター → マクロの設定 →「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をオンにしてください。' }

    # 3. モジュール名 → コード
    $mods = [ordered]@{}
    if ($FromSrc) {
        foreach ($f in Get-ChildItem (Join-Path $here 'src') -Filter *.bas) {
            $mods[$f.BaseName] = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)
        }
    } else {
        foreach ($n in @($wb.Names)) {
            if ($n.Name -like 'マクロ_*') {
                $rg = $n.RefersToRange; $v = $rg.Value2
                $lines = for ($i = 1; $i -le $rg.Rows.Count; $i++) { "$($v[$i, 1])" }
                $mods[$n.Name.Substring(4)] = ($lines -join "`r`n")
            }
        }
    }
    if ($mods.Count -eq 0) { throw '組み直すコードが見つかりません（マクロ保存シートの名前「マクロ_○○」がない）。-FromSrc を付けて src の控えから組み直してください。' }

    # 4. 組み直し（標準モジュールはなければ作る。ThisWorkbook はブックのモジュールへ）
    foreach ($m in $mods.Keys) {
        if ($m -eq 'ThisWorkbook') {
            $comp = $vbp.VBComponents.Item($wb.CodeName)
        } else {
            $comp = $null
            foreach ($c in $vbp.VBComponents) { if ($c.Name -eq $m) { $comp = $c } }
            if ($comp -eq $null) { $comp = $vbp.VBComponents.Add(1); $comp.Name = $m }
        }
        $cm = $comp.CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString(($mods[$m] -replace '(?m)^Attribute VB_.*\r?\n', ''))
        "$m ： $($cm.CountOfLines) 行"
    }

    # 5. マクロ有効ブック（.xlsm）で保存
    $target = [IO.Path]::ChangeExtension($Path, '.xlsm')
    if ($Path -ne $target) { $wb.SaveAs($target, 52); "保存：$target（.xlsm に変換）" } else { $wb.Save(); "保存：$target" }
} finally {
    $wb.Close($false); $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
"完了。Excel で開き、Alt+F11 →「デバッグ」→「VBAProject のコンパイル」でエラーがないか、操作シートのボタンが動くか確認してください。"
