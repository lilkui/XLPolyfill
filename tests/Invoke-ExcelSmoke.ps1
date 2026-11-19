[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet(32, 64)]
    [int]$ExcelBitness
)

$ErrorActionPreference = 'Stop'

function Assert-Cells($Sheet, [hashtable]$Expected) {
    foreach ($cell in $Expected.Keys) {
        $actual = $Sheet.Range($cell).Value2
        if (-not $Expected[$cell].Equals($actual)) {
            throw "Cell ${cell}: expected '$($Expected[$cell])', got '$actual'."
        }
    }
}

function Set-FixtureCells($Sheet, [hashtable]$Values) {
    foreach ($cell in $Values.Keys) {
        $value = $Values[$cell]
        if ($value -is [string]) {
            $Sheet.Range($cell).Value2 = [string]$value
        } else {
            $Sheet.Range($cell).Value2 = [double]$value
        }
    }
}

$suffix = if ($ExcelBitness -eq 64) { 'AddIn64-packed.xll' } else { 'AddIn-packed.xll' }
$xll = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../src/bin/Release/net11.0-windows/publish/XLPolyfill-$suffix")).Path
$xllHash = (Get-FileHash -LiteralPath $xll -Algorithm SHA256).Hash
$excelType = [type]::GetTypeFromProgID('Excel.Application', $false)
$excelServer = if ($null -ne $excelType) {
    $serverKey = "Registry::HKEY_CLASSES_ROOT\CLSID\$($excelType.GUID.ToString('B'))\LocalServer32"
    if (Test-Path -LiteralPath $serverKey) {
        (Get-Item -LiteralPath $serverKey).GetValue('')
    }
}
if ($excelServer -notmatch '\\EXCEL\.EXE(?:"|\s|$)') {
    throw 'The smoke check requires Microsoft Excel desktop registered as Excel.Application.'
}
$probeXll = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "XLPolyfill.ExcelSmoke/bin/Release/net11.0-windows/publish/XLPolyfill.ExcelSmoke-$suffix")).Path

$excel = New-Object -ComObject Excel.Application
$workbook = $null
$roundTripPath = Join-Path ([IO.Path]::GetTempPath()) ("XLPolyfill-Smoke-{0}.xlsx" -f [guid]::NewGuid())
try {
    $excel.DisplayAlerts = $false
    if (-not $excel.RegisterXLL($xll)) {
        throw "Excel could not load '$xll'. Check Excel bitness, the .NET 11 Desktop Runtime, Trust Center policy, and whether another add-in already loaded a different modern .NET major version."
    }

    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'SmokeData'
    $referenceSheet = $workbook.Worksheets.Add()
    $referenceSheet.Name = 'SmokeCriteria'
    $excel.Calculation = -4105 # xlCalculationAutomatic

    $data = @{
        A2 = 1.0; A3 = 2.0; A4 = 3.0
        B2 = 10.0; B3 = 20.0; B4 = 30.0
        C2 = 'a'; C3 = 'b'; C4 = 'c'
        L5 = 3.0; M18 = 1.0; M19 = -1.0
        A514 = 513.0; B514 = 5130.0
        P2 = 'A1'; Q2 = 20.0; P3 = 'A!'; Q3 = 10.0
        R1 = 9.0; R2 = 4.0
    }
    Set-FixtureCells $sheet $data
    $sheet.Range('H9:I9').FormulaArray = '=XLOOKUP({9,8},A2:A4,B2:B4,{"missing 9","missing 8"})'
    $sheet.Range('E10').FormulaArray = '=ISNA(FILTER(A2:C4,B2:B4>100))'
    if (-not ('missing 8'.Equals($sheet.Range('I9').Value2) -and $true.Equals($sheet.Range('E10').Value2))) {
        throw 'The smoke formulas did not produce XLPolyfill results. Run the smoke check in Excel 2016 or 2019, where built-in XLOOKUP and FILTER cannot take precedence.'
    }

    $criteria = @{ J2 = 0.0; J4 = 0.0; J6 = 'none'; K6 = 0.0; J7 = 1.0; K7 = 2.0 }
    Set-FixtureCells $referenceSheet $criteria
    $referenceSheet.Range('J3').Formula = '=J7'
    $sheet.Range('D2').Formula = '="TRUE"'
    $sheet.Range('D3').Formula = '="false"'
    $sheet.Range('D4').Formula = '="TrUe"'
    $sheet.Range('K5').Formula = '=SmokeCriteria!K7'
    $sheet.Range('N2:N514').FormulaR1C1 = '=ROW()-1'
    $sheet.Range('O2:O514').FormulaR1C1 = '=RC[-1]*10'

    $sort = $sheet.Sort
    $sort.SortFields.Clear()
    $null = $sort.SortFields.Add($sheet.Range('P2:P3'), 0, 1)
    $sort.SetRange($sheet.Range('P2:Q3'))
    $sort.Header = 2 # xlNo
    $sort.Apply()

    $invalidMode = if ($excel.International(4) -eq ',') { '0,1' } else { 'invalid' }
    $formulas = @{
        H2 = '=XLOOKUP(2,A2:A4,B2:B4)'
        H3 = '=XLOOKUP(4,A2:A4,B2:B4,"missing")'
        H4 = '=XLOOKUP(2.5,A2:A4,B2:B4,,-0.5,2.5)'
        H5 = '=XLOOKUP("B*",C2:C4,B2:B4,,2)'
        H6 = '=XLOOKUP(3,{1;3;3;3;4},{11;21;22;23;33},,0,2)'
        H11 = '=XLOOKUP(K5,A2:A4,B2:B4)'
        H12 = '=XLOOKUP(1,{0;2},{10;20},"miss","100%")'
        H13 = '=XLOOKUP(2,{0;2},{10;20},"miss",0,UNICHAR(65297))'
        H14 = '=ERROR.TYPE(XLOOKUP(2,{{0;2}},{{10;20}},"miss",0,"{0}"))' -f $invalidMode
        H15 = '=IFNA(XLOOKUP(1,(A2,A4),B2:B3),"disjoint")'
        H16 = '=XLOOKUP("A1",P2:P3,Q2:Q3,"miss",0,2)'
        H27 = '=XLOOKUP(4,A2:A5,{10;20;30;40},"miss",1)'
        H28 = '=XLOOKUP(CHAR(1),C2:C4,B2:B4,"miss",2)'
        H29 = '=XLOOKUP(4,A2:A4,B2:B4,C2,M18,M18)'
        H30 = '=XLOOKUP(,R1:R3,{10;20;30},"miss",0,-2)'
        H31 = '=XLOOKUP(0,R1:R3,{10;20;30},"miss",1,-2)'
        H47 = '=ERROR.TYPE(XLOOKUP(9,A2:A4,B2:B4,,"bad",1/0))'
    }
    foreach ($cell in $formulas.Keys) {
        $sheet.Range($cell).Formula = $formulas[$cell]
    }
    $referenceSheet.Range('H6').Formula = '=SUM(XLOOKUP(1,SmokeData!A2:A4,SmokeData!B2:B4):XLOOKUP(3,SmokeData!A2:A4,SmokeData!B2:B4))'
    $referenceSheet.Range('H13').Formula = '=XLOOKUP(513,SmokeData!A2:A514,SmokeData!B2:B514,-1)'
    $referenceSheet.Range('N2:V2').FormulaArray = '=XLOOKUP({1,2,3,4,5,6,7,8,513},SmokeData!N2:N514,SmokeData!O2:O514,,0,2)'

    $arrayFormulas = @{
        'H7:I7' = '=XLOOKUP(2,A2:A4,B2:C4)'
        'H8:I8' = '=XLOOKUP({2,2},A2:A4,B2:B4)'
        'H17:I17' = '=XLOOKUP({"A1","A!"},P2:P3,Q2:Q3,"miss",0,2)'
        'H18:H19' = '=XLOOKUP(1,{1,1},{10,11},0,0,M18:M19)'
        'H20:I20' = '=XLOOKUP(2,{1,3},{10,30},0,{-1,1})'
        H21 = '=XLOOKUP(UNICHAR(13217),CHOOSE({1,2},"m2",UNICHAR(13217)),{10,20},0)'
        H22 = '=XLOOKUP(UNICHAR(8486)&"*",CHOOSE({1,2},UNICHAR(937)&"1",UNICHAR(8486)&"1"),{10,20},0,2)'
        H23 = '=XLOOKUP(,A2:A4,B2:B4,"miss",-1,-1)'
        H24 = '=XLOOKUP("a.b",CHOOSE({1,2},"a"&UNICHAR(8228)&"b","a.b"),{10,20})'
        H25 = '=XLOOKUP("a.*",CHOOSE({1,2},"a"&UNICHAR(8228)&"b","a.b"),{10,20},,2)'
        'H26:I26' = '=IFERROR(XLOOKUP({9,2},A2:A4,B2:B4,C2,0,{1,-1}),"bad fallback")'
        'H32:I32' = '=XLOOKUP({"",";"},CHOOSE({1,2,3,4},UNICHAR(173),"",UNICHAR(894),";"),{10,20,30,40},"miss")'
        H33 = '=XLOOKUP(";",CHOOSE({1,2},UNICHAR(894),";"),{30,40},"miss",2)'
        'H34:I34' = '=XLOOKUP(CHOOSE({1,2},UNICHAR(953)&"*",UNICHAR(837)&"*"),CHOOSE({1,2},UNICHAR(837)&"1",UNICHAR(953)&"1"),{10,20},,2)'
        'H35:I35' = '=XLOOKUP(CHOOSE({1,2},UNICHAR(66560)&"*",UNICHAR(66600)&"*"),CHOOSE({1,2},UNICHAR(66600)&"1",UNICHAR(66560)&"1"),{10,20},,2)'
        'H36:I36' = '=XLOOKUP(CHOOSE({1,2},UNICHAR(65515),UNICHAR(8594)),CHOOSE({1,2},UNICHAR(8594),UNICHAR(65515)),{10,20})'
        H37 = '=COUNT(FILTER(A5,TRUE))'
        H38 = '=COUNT(XLOOKUP({1;2},A2:A4,A5:A7,"miss"))'
        'H39:H40' = '=ROWS(IFERROR(XLOOKUP({1;2},A2:A4,B2:B4,0,0,(R1,R3)),0))'
        'H41:H42' = '=XLOOKUP({4;2},A2:A4,B2:B4,{"a","b"})'
        'H43:H45' = '=ERROR.TYPE(XLOOKUP(CHOOSE({1;2;3},1,1/0,2),{1,2;3,4},B2:C3,"miss"))'
        H46 = '=COLUMNS(XLOOKUP({2},A2:A4,B2:C4))'
        'H48:I48' = '=ERROR.TYPE(XLOOKUP(9,A2:A4,B2:B4,,{"bad",4},1/0))'
        'K8:L8' = '=XLOOKUP(K5:L5,A2:A4,B2:B4)'
        'E2:G3' = '=FILTER(A2:C4,B2:B4>15,"none")'
        'E5:F7' = '=FILTER(A2:C4,{TRUE,FALSE,TRUE})'
        E9 = '=FILTER(A2:C4,B2:B4>100,"none")'
        'E11:E12' = '=FILTER(A2:A4,D2:D4,"none")'
        'K12:L12' = '=FILTER(A2:B4,SmokeCriteria!J2:J4,SmokeCriteria!J6:K6)'
    }
    foreach ($range in $arrayFormulas.Keys) {
        $sheet.Range($range).FormulaArray = $arrayFormulas[$range]
    }

    Assert-Cells $sheet @{
        H2 = 20.0; H3 = 'missing'; H4 = 20.0; H5 = 20.0; H6 = 21.0
        H7 = 20.0; I7 = 'b'; H8 = 20.0; I8 = 20.0
        H9 = 'missing 9'; I9 = 'missing 8'
        H11 = 20.0; K8 = 20.0; L8 = 30.0; K12 = 2.0; L12 = 20.0
        H12 = 20.0; H13 = 20.0; H14 = 3.0; H15 = 'disjoint'; H16 = 20.0; H17 = 20.0; I17 = 10.0
        H18 = 10.0; H19 = 11.0; H20 = 10.0; I20 = 30.0; H21 = 10.0; H22 = 10.0
        H23 = 30.0; H24 = 10.0; H25 = 10.0; H26 = 'bad fallback'; I26 = 20.0
        H27 = 40.0; H28 = 10.0; H29 = 'a'
        H30 = 'miss'; H31 = 30.0; H32 = 20.0; I32 = 30.0; H33 = 30.0
        H34 = 20.0; I34 = 10.0; H35 = 20.0; I35 = 10.0; H36 = 20.0; I36 = 10.0
        H37 = 0.0; H38 = 0.0; H39 = 2.0; H40 = 2.0; H41 = 'a'; H42 = 20.0
        H43 = 3.0; H44 = 2.0; H45 = 3.0; H46 = 1.0
        H47 = 3.0; H48 = 3.0; I48 = 2.0
        E2 = 2.0; F2 = 20.0; G2 = 'b'; E3 = 3.0; F3 = 30.0; G3 = 'c'
        E5 = 1.0; F5 = 'a'; E6 = 2.0; F6 = 'b'; E7 = 3.0; F7 = 'c'
        E9 = 'none'; E10 = $true; E11 = 1.0; E12 = 3.0
    }
    Assert-Cells $referenceSheet @{ H6 = 60.0; H13 = 5130.0 }
    Assert-Cells $referenceSheet @{
        N2 = 10.0; O2 = 20.0; P2 = 30.0; Q2 = 40.0; R2 = 50.0; S2 = 60.0; T2 = 70.0; U2 = 80.0; V2 = 5130.0
    }

    $sheet.Range('M19').Value2 = 1.0
    Assert-Cells $sheet @{ H18 = 10.0; H19 = 10.0 }

    $sheet.Range('O514').Value2 = 7777.0
    Assert-Cells $referenceSheet @{ V2 = 7777.0 }

    $sheet.Range('A514').Value2 = 514.0
    Assert-Cells $referenceSheet @{ H13 = -1.0 }

    $sheet.Range('B3').Value2 = 25.0
    Assert-Cells $sheet @{ H2 = 25.0; H7 = 25.0; H8 = 25.0; I8 = 25.0; I9 = 'missing 8'; F2 = 25.0; H11 = 25.0; K8 = 25.0; L12 = 25.0 }
    Assert-Cells $referenceSheet @{ H6 = 65.0 }

    $sheet.Range('B2').Value2 = 17.0
    Assert-Cells $sheet @{ E2 = 1.0; F2 = 17.0; G2 = 'a'; E3 = 2.0; F3 = 25.0; G3 = 'b' }

    $referenceSheet.Range('K7').Value2 = 1.0
    Assert-Cells $sheet @{ H11 = 17.0; K8 = 17.0; L8 = 30.0 }

    $referenceSheet.Range('J7').Value2 = 0.0
    Assert-Cells $sheet @{ K12 = 'none'; L12 = 0.0 }
    $referenceSheet.Range('K6').Value2 = 99.0
    Assert-Cells $sheet @{ K12 = 'none'; L12 = 99.0 }
    $referenceSheet.Range('J2').Value2 = 1.0
    Assert-Cells $sheet @{ K12 = 1.0; L12 = 17.0 }

    $workbook.SaveAs($roundTripPath, 51) # xlOpenXMLWorkbook
    $workbook.Close($false)
    $workbook = $null
    $workbook = $excel.Workbooks.Open($roundTripPath, 0)
    $sheet = $workbook.Worksheets.Item('SmokeData')
    $sheet.Range('B3').Value2 = 26.0
    $excel.CalculateFullRebuild()
    Assert-Cells $sheet @{ H2 = 26.0; I9 = 'missing 8'; F3 = 26.0; E10 = $true }

    $sheet.Range('AA2:AA8193').FormulaR1C1 = '=ROW()-1'
    $sheet.Range('AB1').Value2 = 0.0
    if (-not $excel.RegisterXLL($probeXll)) {
        throw "Excel could not load the cleanup probe '$probeXll'."
    }
    $multithreading = $excel.MultiThreadedCalculation.Enabled
    try {
        # The return buffer belongs to the thread that ran FILTER.
        $excel.MultiThreadedCalculation.Enabled = $false
        $sheet.Range('AE1').Value2 = $xll
        $probeFormula = '=XLPOLYFILL_SMOKE_BUFFER_ALLOCATED(AE1)'
        $sheet.Range('AC2:AC3').FormulaArray = '=FILTER(AA2:AA3,AA2:AA3>AB1,-1)'
        $sheet.Range('AD1').Formula = $probeFormula
        Assert-Cells $sheet @{ AC2 = 1.0; AC3 = 2.0; AD1 = $true }

        $sheet.Range('AC2:AC3').ClearContents()
        $sheet.Range('AC2:AC8193').FormulaArray = '=FILTER(AA2:AA8193,AA2:AA8193>AB1,-1)'
        $sheet.Range('AD2').Formula = $probeFormula
        Assert-Cells $sheet @{ AC2 = 1.0; AC8193 = 8192.0; AD2 = $false }
    }
    finally {
        $excel.MultiThreadedCalculation.Enabled = $multithreading
    }
    $sheet.Range('AB1').Value2 = 8192.0
    Assert-Cells $sheet @{ AC2 = -1.0 }
    $sheet.Range('AB1').Value2 = 1.0
    Assert-Cells $sheet @{ AC2 = 2.0; AC8192 = 8192.0 }

    $calculation = $excel.Calculation
    try {
        $excel.MultiThreadedCalculation.Enabled = $true
        $excel.Calculation = -4135 # xlCalculationManual
        $sheet.Range('AF2:AF8001').Value2 = ('x' * 16000)
        $expected = @{}
        for ($row = 2; $row -le 10; $row++) {
            $cell = "AG$row"
            $sheet.Range($cell).FormulaArray = '=LEN(INDEX(FILTER($AF$2:$AF$8001,TRUE),8000))'
            $expected[$cell] = 16000.0
        }
        $excel.CalculateFull()
        Assert-Cells $sheet $expected
    }
    finally {
        $excel.Calculation = $calculation
        $excel.MultiThreadedCalculation.Enabled = $multithreading
    }

    $hostInfo = "Excel $($excel.Version) build $($excel.Build), product $($excel.ProductCode)"
}
finally {
    try {
        if ($null -ne $workbook) {
            $workbook.Close($false)
        }
    }
    finally {
        try {
            $excel.Quit()
        }
        finally {
            if (Test-Path -LiteralPath $roundTripPath) {
                Remove-Item -LiteralPath $roundTripPath
            }
        }
    }
}

Write-Output "Smoke check passed in $hostInfo with $([IO.Path]::GetFileName($xll)) (${ExcelBitness}-bit), SHA-256 $xllHash."
