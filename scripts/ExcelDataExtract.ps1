# Copyright 2025 nagaoka.aya

# Windows PowerShell
# Excel�t�@�C����ǂݍ���

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"
$VerbosePreference = "Continue"
$DebugPreference = "Continue"


# Excel���N������
$excel = New-Object -ComObject Excel.Application

$excel.ScreenUpdating = $false
$excel.DisplayStatusBar = $false
$excel.EnableEvents = $false
$excel.Visible = $false


# �u�b�N���J�� (Start)
$excel.Workbooks.Open($Args[0]) | %{
    # �V�[�g��ǂݍ��� (Start)
    $_.Worksheets | %{
        #$_ | Out-Default
        Write-Output "Sheet_Start=$($_.Name)"

        # �s��ǂݍ��� (Start)
        $_.UsedRange.Rows | %{
            # ���ǂݍ��� (Start)
            $_.Columns | %{
                if($_.Text){
                   Write-Output "$($_.Row) , $($_.Column) , $($_.Text)"
                }
            }
            # ���ǂݍ��� (End)
        }
        # �s��ǂݍ��� (End)

        Write-Output "Sheet_End=$($_.Name)"
    }
    # �V�[�g��ǂݍ��� (End)

}
# �u�b�N���J�� (End)

# Excel���I������
$excel.Quit()


[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null

