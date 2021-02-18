[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$global:EXCEL_EXTENSION = ".xlsx"
$global:INPUT_FILE_NAME = "input" + $EXCEL_EXTENSION
$global:OUTPUT_FILE_NAME = "output" + $EXCEL_EXTENSION
$global:OUTPUT_SHEET_INDEX = 3

$global:PATH_DELIMITER = "\"

$global:MEDIC_NAME_CODES = @{
    "Барский" = 4000;
    "Маслов" = 4001;
}

$global:POLICY_PREFIX = "ЕНП - "

$global:POLICY_CODES = @{
    "АВМ" = 1;
    "МАКС-М" = 2;
    "АСКОМЕД" = 3;
    "АЛЬЯНСМЕД" = 4;
}

$global:COMMON_CLINIC_PREFIX = "ГБУЗ"

# todo correct typings

function CommonProcessor
{
    [OutputType([String])]
    param(
        [Parameter(Mandatory)]
        [int]$rowIndex,
        [Parameter(Mandatory)]
        [int]$columnIndex,
        [Parameter(Mandatory)]
        [String]$inputCellValue,
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [String]$previousResult,
        [Parameter(Mandatory)] #[Sheet]
        $outputSheet
    )
    $outputSheet.Cells.Item($rowIndex, $columnIndex) = $inputCellValue
    return $previousResult
}

function ProcessSecondNameOrName
{
    [OutputType([String])]
    param(
        [Parameter(Mandatory)]
        [String]$inputCellValue,
        [Parameter(Mandatory)]
        [String]$previousResult,
        [Parameter(Mandatory)] #[Sheet]
        $outputSheet
    )
    return $previousResult + " " + $inputCellValue
}

function ProcessSurname
{
    [OutputType([String])]
    param(
        [Parameter(Mandatory)]
        [String]$inputCellValue,
        [Parameter(Mandatory)]
        [String]$previousResult,
        [Parameter(Mandatory)] #[Sheet]
        $outputSheet
    )
    $previousResult = $previousResult + " " + $inputCellValue
    $outputSheet.Cells.Item(5, 2) = $previousResult
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
        [Parameter(Mandatory)] #[Sheet]
        $outputSheet
    )
    $outputSheet.Cells.Item(9, 2) = $POLICY_PREFIX + $inputCellValue
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
        [Parameter(Mandatory)] #[Sheet]
        $outputSheet
    )
    $POLICY_CODES.GetEnumerator() | foreach {
        if ($inputCellValue -Match $_.key)
        {
            $outputSheet.Cells.Item(10, 2) = $_.value
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
        [Parameter(Mandatory)] #[Sheet]
        $outputSheet
    )
    $startIndex = $inputCellValue.IndexOf($COMMON_CLINIC_PREFIX)
    $clinicName = $inputCellValue
    if ($startIndex -ne -1)
    {
        $clinicName = $inputCellValue.Substring($startIndex)
    }
    $outputSheet.Cells.Item(4, 2) = $clinicName
    return $previousResult
}

$global:PROCESSORS = @{
    2 = { param($1, $2, $3) CommonProcessor 12 2 $1 $2 $3 };
    3 = ${function:ProcessSecondNameOrName};
    4 = ${function:ProcessSecondNameOrName};
    5 = ${function:ProcessSurname};
    6 = { param($1, $2, $3) CommonProcessor 7 2 $1 $2 $3 };
    7 = { param($1, $2, $3) CommonProcessor 6 2 $1 $2 $3 };
    8 = ${function:ProcessPolicy};
    9 = ${function:ProcessPolicyCode};
    10 = { param($1, $2, $3) CommonProcessor 8 2 $1 $2 $3 };
    11 = ${function:ProcessClinicName};
}

function ProcessDate
{
    [OutputType([Void])]
    param(
        [Parameter(Mandatory)] #[Sheet]
        $inputSheet,
        [Parameter(Mandatory)] #[Sheet]
        $outputSheet
    )
    #    todo should we locate cell dynamically?
    $inputDate = $inputSheet.Cells.Item(3, 4).Value2
    $outputDate = [math]::Floor($inputDate)
    $outputSheet.Cells.Item(1, 2) = $outputDate
}

function ProcessMedicName
{
    [OutputType([Void])]
    param(
        [Parameter(Mandatory)] #[Sheet]
        $inputSheet,
        [Parameter(Mandatory)] #[Sheet]
        $outputSheet
    )
    #    todo should we locate cell dynamically?
    $medicName = $inputSheet.Cells.Item(4, 4).Value2.Trim()
    $medicCode = $MEDIC_NAME_CODES[$medicName]
    $outputSheet.Cells.Item(3, 2) = $medicCode
}

function ProcessLine
{
    [OutputType([String])]
    param(
        [Parameter(Mandatory)] #[Range]
        $row,
        [Parameter(Mandatory)] #[Sheet]
        $outputSheet
    )
    $previousResult = ""
    $sortedProcessors = $PROCESSORS.GetEnumerator() | Sort-Object -Property key | foreach {
        $processor = $_.value
        $cellValue = $row.Item($_.key).Value2
        $previousResult = $processor.Invoke($cellValue, $previousResult, $outputSheet)
    }
    return $previousResult
}

# todo error handling
function FillCards
{
    $inputExcel = New-Object -Com Excel.Application
    $inputExcel.Visible = $true

    $outputExcel = New-Object -Com Excel.Application
    $outputExcel.Visible = $true

    $inputWBPath = $PSScriptRoot + $PATH_DELIMITER + $INPUT_FILE_NAME
    $outputWBPath = $PSScriptRoot + $PATH_DELIMITER + $OUTPUT_FILE_NAME

    $inputWB = $inputExcel.Workbooks.Open($inputWBPath)
    $inputSheet = $inputWB.Sheets.Item(1)

    #    filling each card
    $xlCellTypeLastCell = 11
    $endColumn = $inputSheet.UsedRange.SpecialCells($xlCellTypeLastCell).Column
    $rowIndex = 1
    $firstRowCell = $inputSheet.Cells.Item($rowIndex, 1)
    #    finding first non-null cell
    while ($firstRowCell.Value2.length -eq 0)
    {
        $firstRowCell = $inputSheet.Cells.Item($rowIndex, 1)
        $rowIndex++
    }
    $firstRowCell = $inputSheet.Cells.Item($rowIndex, 1)
    #    for each valuable row, process it
    $result = ""
    while ($firstRowCell.Value2.length -ne 0)
    {
        $rangeAddress = $inputSheet.Cells.Item($rowIndex, 1).Address() + ":" + $inputSheet.Cells.Item($rowIndex, $endColumn).Address()
        $processedRow = $inputSheet.Range($rangeAddress)

        $newOutputTmpFileName = "tmp" + $OUTPUT_FILE_NAME
        $newOutputTmpFilePath = $PSScriptRoot + $PATH_DELIMITER + $newOutputTmpFileName

        Copy-Item $outputWBPath -Destination $newOutputTmpFilePath
        $outputWB = $outputExcel.Workbooks.Open($newOutputTmpFilePath)

        $outputSheet = $outputWB.Sheets.Item($OUTPUT_SHEET_INDEX)

        ProcessDate -inputSheet $inputSheet -outputSheet $outputSheet
        ProcessMedicName -inputSheet $inputSheet -outputSheet $outputSheet

        $result = ProcessLine -row $processedRow -outputSheet $outputSheet
        # todo SaveAs
        $outputWB.Save()
        $outputWB.Close()

        if ($result.length -ne 0)
        {
            $newName = $result + $EXCEL_EXTENSION
            Rename-Item -Path $newOutputTmpFilePath -NewName $newName
        }

        $rowIndex++
        $firstRowCell = $inputSheet.Cells.Item($rowIndex, 1)
    }
    $inputWB.Close()
    $inputExcel.Quit()
    $outputExcel.Quit()
}

FillCards