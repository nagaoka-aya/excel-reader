# Copyright 2025 nagaoka.aya

# Windows PowerShell
# Excelファイルを読み込む

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"
$VerbosePreference = "Continue"
$DebugPreference = "Continue"


# Excelを起動する
$excel = New-Object -ComObject Excel.Application

$excel.ScreenUpdating = $false
$excel.DisplayStatusBar = $false
$excel.EnableEvents = $false
$excel.Visible = $false


# ブックを開く (Start)
$excel.Workbooks.Open($Args[0]) | %{
    # シートを読み込み (Start)
    $_.Worksheets | %{
        #$_ | Out-Default
        Write-Output "Sheet_Start=$($_.Name)"

        # 行を読み込む (Start)
        $_.UsedRange.Rows | %{
            # 列を読み込む (Start)
            $_.Columns | %{
                if($_.Text){
                   Write-Output "$($_.Row) , $($_.Column) , $($_.Text)"
                }
            }
            # 列を読み込む (End)
        }
        # 行を読み込む (End)

        Write-Output "Sheet_End=$($_.Name)"
    }
    # シートを読み込み (End)

}
# ブックを開く (End)

# Excelを終了する
$excel.Quit()


[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null

