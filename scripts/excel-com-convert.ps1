param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $InputPath)) {
    Write-Error "Input file not found: $InputPath"
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    # Read-only open; Excel descifra libros a3ERP sin pedir contraseña.
    $workbook = $excel.Workbooks.Open($InputPath, 0, $true)
    $workbook.SaveAs($OutputPath, 51) # xlOpenXMLWorkbook (.xlsx)
    $workbook.Close($false)
    $workbook = $null
}
finally {
    if ($workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (-not (Test-Path -LiteralPath $OutputPath)) {
    Write-Error "Excel did not produce output file"
}
