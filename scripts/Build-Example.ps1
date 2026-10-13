[CmdletBinding()]
param([string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'examples\RecurringSchedule.xlsx'))

$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = 'Monthly schedule'
    $sheet.Range('A1:H1').Merge()
    $sheet.Range('A1').Value2 = 'Monthly account schedule'
    $sheet.Range('A1').Font.Size = 20
    $sheet.Range('A1').Font.Bold = $true
    $sheet.Range('A2:H2').Merge()
    $sheet.Range('A2').Value2 = 'Connect with Snowflake, choose the inputs below, then change Refresh to fetch fresh results.'
    $sheet.Range('A4').Value2 = 'Entity'
    $sheet.Range('B4').Value2 = 'EXAMPLE_ENTITY'
    $sheet.Range('A5').Value2 = 'Period start'
    $sheet.Range('B5').Value2 = ([datetime]'2026-09-01').ToOADate()
    $sheet.Range('B5').NumberFormat = 'yyyy-mm-dd'
    $sheet.Range('A6').Value2 = 'Source table'
    $sheet.Range('B6').Value2 = 'FINANCE.PUBLIC.GL_ENTRIES'
    $sheet.Range('A7').Value2 = 'Refresh'
    $sheet.Range('B7').Value2 = 0
    $sheet.Range('A8').Value2 = 'Maximum detail rows'
    $sheet.Range('B8').Value2 = 10000
    $sheet.Range('A10').Value2 = 'SQL filter'
    $sheet.Range('B10').Formula2 = '="FROM "&B6&" WHERE entity = ''"&SUBSTITUTE(SUBSTITUTE(B4,"\","\\"),"''","''''")&"'' AND posting_date >= TO_DATE(''"&TEXT(B5,"yyyy-mm-dd")&"'') AND posting_date < DATEADD(month,1,TO_DATE(''"&TEXT(B5,"yyyy-mm-dd")&"''))"'
    $sheet.Range('A11').Value2 = 'Total SQL'
    $sheet.Range('B11').Formula2 = '="SELECT SUM(amount) AS total_amount "&B10'
    $sheet.Range('A12').Value2 = 'Detail SQL'
    $sheet.Range('B12').Formula2 = '="SELECT CAST(account_code AS VARCHAR) AS account_code, account_name, SUM(amount) AS amount "&B10&" GROUP BY account_code, account_name ORDER BY account_code"'
    $sheet.Range('A14').Value2 = 'Total amount'
    $sheet.Range('B14').Formula2 = '=SFVALUE(B11,$B$7)'
    $sheet.Range('A17:C17').Merge()
    $sheet.Range('A17').Value2 = 'Account detail — leave the output area below clear'
    $sheet.Range('A18').Formula2 = '=SFQUERY($B$12,TRUE,$B$7,$B$8)'
    $sheet.Range('E17:F17').Merge()
    $sheet.Range('E17').Value2 = 'Detail query status'
    $sheet.Range('E18').Formula2 = '=SFSTATUS(A18)'
    $sheet.Range('E4:H4').Merge()
    $sheet.Range('E4').Value2 = 'Getting started'
    $sheet.Range('E4').Font.Bold = $true
    $notes = @(
        '1. Install the rebuilt Excelerator add-in and use Connect in this workbook.',
        '2. Replace the example entity, period and source table. The SQL assumes entity, posting_date, account_code, account_name and amount columns.',
        '3. Increase Refresh by 1 to fetch again. Ordinary recalculation reuses results.',
        '4. High-precision amounts are returned as exact text. Choose an explicit numeric CAST in SQL only when its precision and scale fit your reporting requirements.',
        '5. Keep the detail spill area clear. Place formulas and notes beside the output.',
        'Agents: use SFDATABASES, SFSCHEMAS, SFTABLES and SFCOLUMNS to discover the actual source before adapting this template.'
    )
    for ($i = 0; $i -lt $notes.Count; $i++) {
        $row = 5 + $i * 2
        $sheet.Range("E${row}:H$($row+1)").Merge()
        $noteCell = $sheet.Range("E$row")
        $noteCell.Value2 = [string]$notes[$i]
        $sheet.Range("E${row}:H$($row+1)").WrapText = $true
        $sheet.Range("E${row}:H$($row+1)").VerticalAlignment = -4160
    }
    $sheet.Range('A:A').ColumnWidth = 25
    $sheet.Range('B:B').ColumnWidth = 45
    $sheet.Range('C:C').ColumnWidth = 24
    $sheet.Range('D:D').ColumnWidth = 3
    $sheet.Range('E:H').ColumnWidth = 19
    $sheet.Range('B4:B8').Interior.Color = 15790320
    $sheet.Range('A10:B12').WrapText = $true
    $sheet.Range('A10:B12').RowHeight = 60
    $sheet.Range('A17:C17').Font.Bold = $true
    $sheet.Range('E17:F17').Font.Bold = $true
    $resolved = [IO.Path]::GetFullPath($OutputPath)
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $resolved))
    $book.SaveAs($resolved, 51)
    Write-Output ('Created ' + $resolved)
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
