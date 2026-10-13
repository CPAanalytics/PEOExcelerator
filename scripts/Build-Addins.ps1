[CmdletBinding()]
param(
    [string]$OutputDirectory = (Split-Path -Parent $PSScriptRoot),
    [switch]$TestOnly
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$excel = $null
$workbook = $null
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('Excelerator-build-' + [guid]::NewGuid())
$stage = Join-Path $temporary 'SnowflakeExcelAddin.xlam'
. (Join-Path $PSScriptRoot 'Compile-Vba.ps1')

function Get-VbaProject($book) {
    try {
        return $book.GetType().InvokeMember('VBProject', [Reflection.BindingFlags]::GetProperty, $null, $book, $null)
    } catch {
        throw ('Unable to access the VBA project: ' + $_.Exception.Message + ' Enable Excel Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model if access is not trusted. This script does not change that setting.')
    }
}

try {
    [void](New-Item -ItemType Directory -Path $temporary)
    Copy-Item -LiteralPath (Join-Path $repo 'SnowflakeExcelAddin.xlam') -Destination $stage
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    # Startup events stay disabled; allow explicit test macros in the staging file.
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($stage, 0, $false)
    Write-Verbose 'Opened staging add-in; importing source.'
    $project = Get-VbaProject $workbook
    if ($project.Protection -ne 0) { throw 'The add-in VBA project is locked.' }

    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $repo 'src') -File) {
        Write-Verbose ('Importing ' + $file.Name)
        if ($file.BaseName -in @('TestHarness', 'Testing')) {
            # These old, live-database developer harnesses use obsolete APIs.
            # Keep their source for development; never ship them in releases.
            $developmentComponent = $null
            try { $developmentComponent = $project.VBComponents.Item($file.BaseName) } catch {}
            if ($null -ne $developmentComponent) { $project.VBComponents.Remove($developmentComponent) }
            continue
        }
        if ($file.Name.EndsWith('.sheet.cls')) {
            $componentName = $file.Name.Substring(0, $file.Name.Length - '.sheet.cls'.Length)
            $code = $project.VBComponents.Item($componentName).CodeModule
            if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
            $code.AddFromString([IO.File]::ReadAllText($file.FullName))
        } elseif ($file.Extension -in '.bas', '.frm') {
            $existing = $null
            try { $existing = $project.VBComponents.Item($file.BaseName) } catch {}
            if ($null -ne $existing) { $project.VBComponents.Remove($existing) }
            [void]$project.VBComponents.Import($file.FullName)
        }
    }
    # Excel can import exported .cls files as standard modules. Create the test
    # class explicitly and add only its code, excluding the export header.
    $fake = $project.VBComponents.Add(2)
    $fake.Name = 'FakeConnection'
    $fakeSource = [IO.File]::ReadAllText((Join-Path $repo 'tests\FakeConnection.cls'))
    $fake.CodeModule.AddFromString($fakeSource.Substring($fakeSource.IndexOf('Option Explicit')))
    [void]$project.VBComponents.Import((Join-Path $repo 'tests\FormulaTests.bas'))
    Write-Verbose 'Imported source; running offline Excel tests before packaging.'
    # Run only the reviewed test entrypoint; startup/close events remain disabled.
    $prefix = "'" + $workbook.Name.Replace("'", "''") + "'!"
    foreach ($reference in $project.References) {
        if ($reference.IsBroken) { throw ('Broken VBA reference: ' + $reference.Name) }
    }
    Write-Verbose ('Imported VBA project ' + $project.Name + ' with ' + $project.VBComponents.Count + ' components.')
    Assert-VbaCompilation $excel $project
    $excel.Run($prefix + 'WorksheetFunctions.RegisterFunctions')
    $result = $excel.Run($prefix + 'FormulaTests.Run')
    Write-Output $result
    $project = Get-VbaProject $workbook
    foreach ($name in @('FakeConnection', 'FormulaTests')) {
        $project.VBComponents.Remove($project.VBComponents.Item($name))
    }
    Assert-VbaCompilation $excel $project

    if (-not $TestOnly) {
        $resolvedOutput = [IO.Path]::GetFullPath($OutputDirectory)
        [void][IO.Directory]::CreateDirectory($resolvedOutput)
        $workbook.Names.Item('sfReadOnly').RefersToRange.Value2 = 'False'
        $workbook.Save()
        $fullVersion = Join-Path $temporary 'full.xlam'
        $workbook.SaveCopyAs($fullVersion)
        $workbook.Names.Item('sfReadOnly').RefersToRange.Value2 = 'True'
        $workbook.Save()
        $readOnlyVersion = Join-Path $temporary 'readonly.xlam'
        $workbook.SaveCopyAs($readOnlyVersion)
        $workbook.Close($false)
        $workbook = $null

        foreach ($file in @($fullVersion, $readOnlyVersion)) {
            $excel.AutomationSecurity = 3
            $check = $excel.Workbooks.Open($file, 0, $true)
            try {
                $checkProject = Get-VbaProject $check
                foreach ($name in @('WorksheetFunctions', 'FormulaEngine', 'FormulaResults', 'SQLValidation', 'PublicMacros', 'Load', 'Query')) {
                    [void]$checkProject.VBComponents.Item($name)
                }
                if ($checkProject.VBComponents.Count -lt 7) { throw 'Packaged add-in is missing VBA modules.' }
            } finally { $check.Close($false) }
        }
        Copy-Item -LiteralPath $fullVersion -Destination (Join-Path $resolvedOutput 'SnowflakeExcelAddin.xlam') -Force
        Copy-Item -LiteralPath $readOnlyVersion -Destination (Join-Path $resolvedOutput 'SnowflakeExcelAddinReadOnly.xlam') -Force
        Write-Output ('Built and verified both add-ins in ' + $resolvedOutput)
    }
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    # Delete only this script's verified unique temporary directory.
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    $resolvedTemp = [IO.Path]::GetFullPath($temporary)
    if ($resolvedTemp.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path -Leaf $resolvedTemp).StartsWith('Excelerator-build-')) {
        Remove-Item -LiteralPath $resolvedTemp -Recurse -Force -ErrorAction SilentlyContinue
    }
}
