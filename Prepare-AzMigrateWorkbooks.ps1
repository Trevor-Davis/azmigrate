<#
.SYNOPSIS
  Validate and repair Azure Migrate workbooks so azuremigrate.py can read them.

.DESCRIPTION
  For each *.xlsx in the input directory:
    1. Checks the first two bytes; a real .xlsx starts with "PK" (zip signature).
    2. If it's not a real zip, opens it in Excel and re-saves as a proper .xlsx
       (handles the common "legacy .xls saved with .xlsx extension" case).
    3. If after the re-save it's still not a real zip, the file is almost certainly
       MIP / sensitivity-labeled (encrypted). The script reports that and asks you
       to open it once in Excel, change Sensitivity to Public (or remove the label),
       Save As .xlsx, then re-run.

.PARAMETER Path
  Directory containing the workbooks (defaults to current location).

.EXAMPLE
  .\Prepare-AzMigrateWorkbooks.ps1 -Path "C:\...\detailed-report"
#>
[CmdletBinding()]
param(
    [string]$Path
)

# ----------------------------------------------------------------------
# Default workbook folder when -Path is not supplied on the command line.
# Edit this to point at the directory containing Discovery.xlsx etc.
# ----------------------------------------------------------------------
$DefaultPath = "C:\Users\tredavis\OneDrive - Microsoft\M+M\AzMigrate\PPTBuild\Sandbox\report\azure-migrate\detailed-report"

if (-not $Path) { $Path = $DefaultPath }

function Test-IsRealXlsx {
    param([string]$File)
    try {
        $fs = [System.IO.File]::OpenRead($File)
        try {
            $b = New-Object byte[] 2
            [void]$fs.Read($b, 0, 2)
            return ($b[0] -eq 0x50 -and $b[1] -eq 0x4B)
        } finally { $fs.Close() }
    } catch { return $false }
}

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Host "Path not found: $Path" -ForegroundColor Red
    exit 1
}

$TargetFiles = @(
    'Discovery.xlsx',
    'Strategy_Lift_and_shift.xlsx',
    'Strategy_PaaS_Preferred.xlsx'
)

$files = foreach ($name in $TargetFiles) {
    $full = Join-Path $Path $name
    if (Test-Path -LiteralPath $full) {
        Get-Item -LiteralPath $full
    } else {
        Write-Host ("[MISS] {0} not found in path" -f $name) -ForegroundColor Red
    }
}
$files = @($files | Where-Object { $_ -is [System.IO.FileInfo] })

if (-not $files) {
    Write-Host "None of the expected workbooks were found in $Path" -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "Checking $($files.Count) workbook(s) in:" -ForegroundColor Cyan
Write-Host "  $Path"
Write-Host ""

$excel = $null
$needsManual = @()

foreach ($f in $files) {
    if (Test-IsRealXlsx -File $f.FullName) {
        Write-Host ("[OK ] {0}" -f $f.Name) -ForegroundColor Green
        continue
    }

    Write-Host ("[FIX] {0} - re-saving via Excel..." -f $f.Name) -ForegroundColor Yellow

    if ($null -eq $excel) {
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
        } catch {
            Write-Host "Could not start Excel. Is Microsoft Excel installed? $($_.Exception.Message)" -ForegroundColor Red
            exit 1
        }
    }

    $wb = $null
    $tmp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(),
            [System.Guid]::NewGuid().ToString() + '.xlsx')
    try {
        # UpdateLinks=0, ReadOnly=$true so we don't trigger save prompts on original
        $wb = $excel.Workbooks.Open($f.FullName, 0, $true)
        $wb.SaveAs($tmp, 51)   # 51 = xlOpenXMLWorkbook
        $wb.Close($false)
        $wb = $null

        if (Test-IsRealXlsx -File $tmp) {
            Copy-Item -LiteralPath $tmp -Destination $f.FullName -Force
            Write-Host ("       -> repaired") -ForegroundColor Green
        } else {
            Write-Host ("       -> Excel re-save still not a valid .xlsx (sensitivity label still applied)") -ForegroundColor Red
            Write-Host ("       -> opening in Excel for manual fix...") -ForegroundColor Yellow

            # Open the file in a visible Excel window so user can clear the label
            $visibleExcel = New-Object -ComObject Excel.Application
            $visibleExcel.Visible = $true
            try {
                [void]$visibleExcel.Workbooks.Open($f.FullName)
            } catch {
                Write-Host ("       -> could not open: {0}" -f $_.Exception.Message) -ForegroundColor Red
            }

            Write-Host ""
            Write-Host "       In the Excel window that just opened:" -ForegroundColor Cyan
            Write-Host "         1. Sensitivity -> Public  (or remove the label)"
            Write-Host "         2. File -> Save  (Ctrl+S),  then close the workbook"
            Write-Host ""
            Read-Host "       Press ENTER here after you have saved and closed the file"

            try { $visibleExcel.Quit() } catch {}
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($visibleExcel)
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()

            if (Test-IsRealXlsx -File $f.FullName) {
                Write-Host ("       -> repaired") -ForegroundColor Green
            } else {
                Write-Host ("       -> still not a valid .xlsx; label was not removed") -ForegroundColor Red
                $needsManual += $f.Name
            }
        }
    } catch {
        Write-Host ("       -> failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
        $needsManual += $f.Name
        if ($wb) { try { $wb.Close($false) } catch {} }
    } finally {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
}

if ($excel) {
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ""
if ($needsManual.Count -gt 0) {
    Write-Host "Manual action required for these files (sensitivity label present):" -ForegroundColor Yellow
    foreach ($n in $needsManual) { Write-Host "  - $n" }
    Write-Host ""
    Write-Host "For each one:" -ForegroundColor Yellow
    Write-Host "  1. Open the file in Excel"
    Write-Host "  2. Sensitivity -> Public  (or remove the label)"
    Write-Host "  3. File -> Save As -> Excel Workbook (*.xlsx), overwrite"
    Write-Host "  4. Close Excel, then re-run this script to verify"
    exit 2
}

Write-Host "All workbooks are valid .xlsx. Safe to run azuremigrate.py." -ForegroundColor Green
exit 0
