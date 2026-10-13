[CmdletBinding()]
param([string]$Directory = (Split-Path -Parent $PSScriptRoot))

$ErrorActionPreference = 'Stop'
$checks = 0
. (Join-Path $PSScriptRoot 'Compile-Vba.ps1')
foreach ($name in @('SnowflakeExcelAddin.xlam', 'SnowflakeExcelAddinReadOnly.xlam')) {
    $excel = $null
    $addin = $null
    $book = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 1
        $addin = $excel.Workbooks.Open((Join-Path ([IO.Path]::GetFullPath($Directory)) $name), 0, $true)
        $project = $addin.GetType().InvokeMember('VBProject', [Reflection.BindingFlags]::GetProperty, $null, $addin, $null)
        Assert-VbaCompilation $excel $project
        foreach ($developmentName in @('TestHarness', 'Testing', 'FormulaTests', 'FakeConnection')) {
            $developmentComponent = $null
            try { $developmentComponent = $project.VBComponents.Item($developmentName) } catch {}
            if ($null -ne $developmentComponent) { throw ($name + ': contains developer test code ' + $developmentName) }
        }
        $checks++
        $prefix = "'" + $addin.Name.Replace("'", "''") + "'!"
        $excel.Run($prefix + 'ThisWorkbook.Workbook_Open')
        $book = $excel.Workbooks.Add()
        $sheet = $book.Worksheets.Item(1)
        $initialSheets = $book.Worksheets.Count
        $initialNames = $book.Names.Count
        $sheet.Range('A1').Formula2 = '=SFVALUE("SELECT 1")'
        $sheet.Range('B1').Formula2 = '=SFQUERY("SELECT 1")'
        $sheet.Range('D1').Formula2 = '=SFSTATUS(A1)'
        $sheet.Range('D13').Formula2 = '=SFSTATUS()'
        $sheet.Calculate()
        foreach ($address in @('A1', 'B1')) {
            if ($sheet.Range($address).Text -ne '#N/A') { throw ($name + ': expected a connection error at ' + $address) }
            $checks++
        }
        if ($sheet.Range('E1').Value2 -ne 'Not connected') { throw ($name + ': missing connection status') }
        $checks++
        if ($sheet.Range('E13').Value2 -ne 'Not connected') { throw ($name + ': missing standalone connection status') }
        $checks++
        if ($book.Worksheets.Count -ne $initialSheets -or $book.Names.Count -ne $initialNames) { throw ($name + ': formulas changed workbook structure') }
        $checks++
        $expectedReadOnly = if ($name.Contains('ReadOnly')) { 'True' } else { 'False' }
        if ([string]$addin.Names.Item('sfReadOnly').RefersToRange.Value2 -ne $expectedReadOnly) { throw ($name + ': incorrect read-only mode') }
        $checks++
        Write-Output ($name + ': formula startup and disconnected-session smoke checks passed.')
    } finally {
        if ($null -ne $book) { $book.Close($false) }
        if ($null -ne $addin) { $addin.Close($false) }
        if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    }
}
Write-Output ($checks.ToString() + ' package smoke checks passed; no Snowflake connection attempted.')
