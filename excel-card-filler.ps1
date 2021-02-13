$global:INPUT_FILE_NAME = "input.xlsx"
$global:OUTPUT_FILE_NAME = "output.xlsx"

$global:MEDIC_NAME_CODES = @{
    @"
Барский
"@ = "4000";
    @"
Маслов
"@ = "4001";
}

$global:POLICY_PREFIX = @"
ЕНП - 
"@

$global:POLICY_CODES = @{
    @"
АВМ
"@ = "1";
    @"
МАКС-М
"@ = "2";
    @"
АСКОМЕД
"@ = "3";
    @"
АЛЬЯНСМЕД
"@ = "4";
}

$global:COMMON_CLINIC_PREFIX = @"
ГБУЗ
"@

function GetCommonProcessor
{
    #    [OutputType([Function])]
    param(
        [Parameter(Mandatory)]
        [int]$rowIndex,
        [Parameter(Mandatory)]
        [int]$columnIndex
    )
    function CommonProcessor
    {
        [OutputType([String])]
        param(
            [Parameter(Mandatory)]
            [String]$inputCellValue,
            [Parameter(Mandatory)]
            [String]$previousResult,
            [Parameter(Mandatory)]
            [Sheet]$outputSheet
        )
        $outputSheet.Cells.Item($rowIndex, $columnIndex).Value2 = $inputCellValue
        return $previousResult
    }
    return ${function:CommonProcessor}
}

function ProcessSecondNameOrName
{
    [OutputType([String])]
    param(
        [Parameter(Mandatory)]
        [String]$inputCellValue,
        [Parameter(Mandatory)]
        [String]$previousResult,
        [Parameter(Mandatory)]
        [Sheet]$outputSheet
    )
    return $previousResult + $inputCellValue
}

function ProcessSurname
{
    [OutputType([String])]
    param(
        [Parameter(Mandatory)]
        [String]$inputCellValue,
        [Parameter(Mandatory)]
        [String]$previousResult,
        [Parameter(Mandatory)]
        [Sheet]$outputSheet
    )
    $outputSheet.Cells.Item(5, 2).Value2 = $previousResult + $inputCellValue
    return $previousResult
}

function ProcessPolicy
{
    [OutputType([String])]
    param(
        [Parameter(Mandatory)]
        [String]$inputCellValue,
        [Parameter(Mandatory)]
        [String]$previousResult,
        [Parameter(Mandatory)]
        [Sheet]$outputSheet
    )
    $outputSheet.Cells.Item(9, 2).Value2 = $POLICY_PREFIX + $inputCellValue
    return $previousResult
}

function ProcessPolicyCode
{
    [OutputType([String])]
    param(
        [Parameter(Mandatory)]
        [String]$inputCellValue,
        [Parameter(Mandatory)]
        [String]$previousResult,
        [Parameter(Mandatory)]
        [Sheet]$outputSheet
    )
    $POLICY_CODES.GetEnumerator() | foreach {
        if ($inputCellValue -Match $_.key)
        {
            $outputSheet.Cells.Item(10, 2).Value2 = $_.value
        }
    }
    return $previousResult
}

function ProcessClinicName
{
    [OutputType([String])]
    param(
        [Parameter(Mandatory)]
        [String]$inputCellValue,
        [Parameter(Mandatory)]
        [String]$previousResult,
        [Parameter(Mandatory)]
        [Sheet]$outputSheet
    )
    $startIndex = $inputCellValue.IndexOf($COMMON_CLINIC_PREFIX)
    $clinicName = $inputCellValue
    if ($startIndex -ne -1)
    {
        $clinicName = $inputCellValue.Substring($startIndex)
    }
    $outputSheet.Cells.Item(4, 2).Value2 = $clinicName
    return $previousResult
}

$global:PROCESSORS = @{
    2 = GetCommonProcessor 12 2;
    3 = ${function:ProcessSecondNameOrName};
    4 = ${function:ProcessSecondNameOrName};
    5 = ${function:ProcessSurname};
    6 = GetCommonProcessor 7 2;
    7 = GetCommonProcessor 6 2;
    8 = ${function:ProcessPolicy};
    9 = ${function:ProcessPolicyCode};
    10 = GetCommonProcessor 8 2;
    11 = ${function:ProcessClinicName};
}

function ProcessDate
{
    [OutputType([Void])]
    param(
        [Parameter(Mandatory)]
        [Sheet]$inputSheet,
        [Parameter(Mandatory)]
        [Sheet]$outputSheet
    )
    #    todo should we locate cell dynamically?
    $fullDate = $sheet.Cells.Item(3, 3).Value2
    $dateParts = $fullDate -split " "
    $outputSheet.Cells.Item(1, 2).Value2 = $dateParts[0]
}

function ProcessMedicName
{
    [OutputType([Void])]
    param(
        [Parameter(Mandatory)]
        [Sheet]$inputSheet,
        [Parameter(Mandatory)]
        [Sheet]$outputSheet
    )
    #    todo should we locate cell dynamically?
    $medicName = $sheet.Cells.Item(4, 4).Value2
    $outputSheet.Cells.Item(1, 2).Value2 = MEDIC_NAME_CODES[$medicName]
}

function ProcessLine
{
    [OutputType([String])]
    param(
        [Parameter(Mandatory)]
        [Range]$row,
        [Parameter(Mandatory)]
        [Sheet]$outputSheet
    )
    $previousResult = ""
    $sortedProcessors = $PROCESSORS.GetEnumerator() | Sort-Object -Property key | foreach {
        $processor = $_.value
        $cellValue = $row.Item($_.key).Value2
        $previousResult = $processor.Invoke($cellValue, $previousResult, $outputSheet)
    }
}

function FillCards
{
    $excel = New-Object -Com Excel.Application
    $inputWB = $excel.Workbooks.Open($INPUT_FILE_NAME)
    $inputSheet = $inputWB.Sheets.Item(1)

    #    filling each card
    $xlCellTypeLastCell = 11
    $endColumn = $inputSheet.UsedRange.SpecialCells($xlCellTypeLastCell).Column
    $rowIndex = 1
    $firstRowCell = $inputSheet.Cells.Item($rowIndex, 1)
    #    finding first non-null cell
    while ($firstRowCell.Value2.length -eq 0)
    {
        $rowIndex++
        $firstRowCell = $inputSheet.Cells.Item($rowIndex, 1)
    }
    #    for each valuable row, process it
    $result = ""
    while ($firstRowCell.Value2.length -ne 0)
    {
        $rangeAddress = $inputSheet.Cells.Item($rowIndex, 1).Address() + ":" + $inputSheet.Cells.Item($rowIndex, $endColumn).Address()
        $processedRow = $inputSheet.Range($rangeAddress)

        $newOutputTmpFileName = "tmp" + $OUTPUT_FILE_NAME
        Copy-Item $OUTPUT_FILE_NAME -Destination $newOutputTmpFileName
        $outputWB = $excel.Workbooks.Open($newOutputFileName)
        $outputSheet = $outputWB.Sheets.Item(1)

        ProcessDate -inputSheet $inputSheet -outputSheet $outputSheet
        ProcessMedicName -inputSheet $inputSheet -outputSheet $outputSheet

        $result = ProcessLine -row $processedRow -outputSheet $outputSheet

        if ($result.length -ne 0)
        {
            Rename-Item -Path $outputWB.Path -NewName $result
        }

        $outputWB.Save()
        $outputWB.Close()

        $rowIndex++
        $firstRowCell = $inputSheet.Cells.Item($rowIndex, 1)
    }
    $inputWB.Close()
    $excel.Quit()
}