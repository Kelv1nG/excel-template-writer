# Regeneration only: requires installed desktop Excel; never invoke from default CI.
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([object] $Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Set-ReferenceCell {
    param([object] $Sheet, [string] $Address, [string] $Text, [switch] $Formula)
    $range = $null
    try {
        $range = $Sheet.Range($Address)
        if ($Formula) {
            # Formula (not FormulaLocal) uses invariant English syntax and decimal points.
            $range.Formula = $Text
            $range.NumberFormat = '0.000000000000000'
        } else {
            $range.Value2 = $Text
        }
    } finally {
        Release-ComObject $range
    }
}

function Publish-TemporaryFile {
    param([string] $Temporary, [string] $Destination)
    if ([IO.File]::Exists($Destination)) {
        [IO.File]::Replace($Temporary, $Destination, $null)
    } else {
        [IO.File]::Move($Temporary, $Destination)
    }
}

$fixturePath = Join-Path $PSScriptRoot 'excel_numeric_reference.xlsx'
$manifestPath = Join-Path $PSScriptRoot 'excel_numeric_reference.json'
$temporaryId = [Guid]::NewGuid().ToString('N')
$temporaryWorkbook = Join-Path $PSScriptRoot "excel_numeric_reference.$temporaryId.tmp.xlsx"
$temporaryManifest = Join-Path $PSScriptRoot "excel_numeric_reference.$temporaryId.tmp.json"
$excel = $null
$workbooks = $null
$workbook = $null
$worksheets = $null
$sheet = $null
$names = $null
$columns = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
    $excel.EnableEvents = $false
    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Add(-4167) # xlWBATWorksheet: exactly one worksheet
    $excel.Calculation = -4105 # xlCalculationAutomatic
    $excel.Iteration = $false
    $workbook.PrecisionAsDisplayed = $false
    $worksheets = $workbook.Worksheets
    $sheet = $worksheets.Item(1)
    $sheet.Name = 'NumericReference'
    $names = $workbook.Names
    Set-ReferenceCell $sheet 'A1' 'Case / significant digits'
    Set-ReferenceCell $sheet 'B1' 'Excel-calculated result'
    Set-ReferenceCell $sheet 'C1' 'Comparison mode'
    Set-ReferenceCell $sheet 'A17' 'Literal amounts above prove decimal transport only.'
    Set-ReferenceCell $sheet 'A18' 'Exact binary-fraction operation'
    Set-ReferenceCell $sheet 'C18' 'Same binary inputs; one explicit operation'

    $cells = [Collections.Generic.List[object]]::new()
    $digits = '314159265358979'
    for ($precision = 1; $precision -le 15; $precision++) {
        $amount = '0.' + $digits.Substring(0, $precision)
        $cells.Add([ordered]@{
            name = 'precision_{0:00}' -f $precision
            sheet = 'NumericReference'
            cell = 'B{0}' -f ($precision + 1)
            formula = '=' + $amount
            decimal = $amount
            precision = $precision
            comparison_mode = 'decimal_amount'
        })
    }
    $operations = @(
        @('exact_add', '=0.5+0.25', '0.75', 'add', '0.5', '0.25'),
        @('exact_subtract', '=1.5-0.25', '1.25', 'subtract', '1.5', '0.25'),
        @('exact_multiply', '=1.125*2', '2.25', 'multiply', '1.125', '2'),
        @('exact_divide', '=5/2', '2.5', 'divide', '5', '2')
    )
    for ($index = 0; $index -lt $operations.Count; $index++) {
        $operation = $operations[$index]
        $cells.Add([ordered]@{
            name = $operation[0]
            sheet = 'NumericReference'
            cell = 'B{0}' -f ($index + 19)
            formula = $operation[1]
            decimal = $operation[2]
            operation = $operation[3]
            inputs = @($operation[4], $operation[5])
            evaluation_order = 'single binary operation on literal inputs'
            comparison_mode = 'binary64_exact'
        })
    }
    foreach ($record in $cells) {
        $row = $record.cell.Substring(1)
        Set-ReferenceCell $sheet "A$row" $record.name
        Set-ReferenceCell $sheet $record.cell $record.formula -Formula
        Set-ReferenceCell $sheet "C$row" $record.comparison_mode
        $definedName = $null
        try {
            $refersTo = "='NumericReference'!`$B`$$row"
            $definedName = $names.Add($record.name, $refersTo)
        } finally {
            Release-ComObject $definedName
        }
    }
    $columns = $sheet.Columns
    [void]$columns.AutoFit()
    $excel.CalculateFullRebuild()
    if ($excel.CalculationState -ne 0) {
        throw 'Excel calculation did not finish after CalculateFullRebuild.'
    }
    foreach ($record in $cells) {
        $range = $null
        try {
            $range = $sheet.Range($record.cell)
            $record.cached_decimal = ([double]$range.Value2).ToString(
                'R', [Globalization.CultureInfo]::InvariantCulture
            )
        } finally {
            Release-ComObject $range
        }
    }
    $executable = Join-Path $excel.Path 'EXCEL.EXE'
    $manifest = [ordered]@{
        schema_version = 1
        producer = 'Microsoft Excel desktop via PowerShell COM'
        workbook = 'excel_numeric_reference.xlsx'
        created_utc = [DateTime]::UtcNow.ToString('o').Replace('+00:00', 'Z')
        excel = [ordered]@{
            version = [string]$excel.Version
            build = [string]$excel.Build
            file_version = [Diagnostics.FileVersionInfo]::GetVersionInfo($executable).FileVersion
            executable = $executable
        }
        locale = [ordered]@{
            culture = [Globalization.CultureInfo]::CurrentCulture.Name
            country_setting = [int]$excel.International(2) # xlCountrySetting
            decimal_separator = [string]$excel.International(3) # xlDecimalSeparator
            use_system_separators = [bool]$excel.UseSystemSeparators
        }
        calculation = [ordered]@{
            mode = 'automatic'
            mode_value = [int]$excel.Calculation
            precision_as_displayed = [bool]$workbook.PrecisionAsDisplayed
            iteration = [bool]$excel.Iteration
            full_rebuild = $true
            state_after_rebuild = [int]$excel.CalculationState
            automation_security = [int]$excel.AutomationSecurity
            save_format = 51 # xlOpenXMLWorkbook; macro-free .xlsx
        }
        cells = $cells.ToArray()
        regeneration = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/fixtures/numeric_fidelity/build_excel_reference.ps1'
    }
    $workbook.SaveAs($temporaryWorkbook, 51)
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $null
    Publish-TemporaryFile $temporaryWorkbook $fixturePath
    $manifest.sha256 = (Get-FileHash -LiteralPath $fixturePath -Algorithm SHA256).Hash.ToLowerInvariant()
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $temporaryManifest -Encoding UTF8
    Publish-TemporaryFile $temporaryManifest $manifestPath
    Write-Output "Generated $fixturePath with Excel $($manifest.excel.version) build $($manifest.excel.build)"
    Write-Output "SHA-256: $($manifest.sha256)"
} finally {
    # Close our workbook on every failure path; never attach to a user's workbook.
    try {
        if ($null -ne $workbook) { $workbook.Close($false) }
    } finally {
        try {
            if ($null -ne $excel) { $excel.Quit() }
        } finally {
            Release-ComObject $columns
            Release-ComObject $names
            Release-ComObject $sheet
            Release-ComObject $worksheets
            Release-ComObject $workbook
            Release-ComObject $workbooks
            Release-ComObject $excel
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            foreach ($temporaryPath in @($temporaryWorkbook, $temporaryManifest)) {
                if (Test-Path -LiteralPath $temporaryPath) {
                    Remove-Item -LiteralPath $temporaryPath
                }
            }
        }
    }
}
