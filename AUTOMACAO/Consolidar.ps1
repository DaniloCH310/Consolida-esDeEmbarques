[CmdletBinding()]
param(
    [string] $RaizConsolidacoes,
    [switch] $SemInteracao
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$automationDirectory = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($RaizConsolidacoes)) {
    $RaizConsolidacoes = Split-Path -Parent $automationDirectory
}
$RaizConsolidacoes = [IO.Path]::GetFullPath($RaizConsolidacoes)

$modulePath = Join-Path $automationDirectory "Consolidacao.Core.psm1"
Import-Module -Name $modulePath -Force

$runtimeAutomationDirectory = Join-Path $RaizConsolidacoes "AUTOMACAO"
$logsDirectory = Join-Path $runtimeAutomationDirectory "LOGS"
$errorsDirectory = Join-Path $runtimeAutomationDirectory "ERROS"
$stateDirectory = Join-Path $runtimeAutomationDirectory "ESTADO"
$statePath = Join-Path $stateDirectory "processados.jsonl"
$lockPath = Join-Path $runtimeAutomationDirectory "processamento.lock"
$shipmentGroupsPath = Join-Path $automationDirectory "GRUPOS_EMBARQUES.csv"
$assetsDirectory = Join-Path $automationDirectory "ASSETS"
$logoPath = Join-Path $assetsDirectory "CTRES_perfilNOME.png"
$brandingWarnings = New-Object "System.Collections.Generic.List[string]"
if (-not (Test-Path -LiteralPath $logoPath -PathType Leaf)) {
    $brandingWarnings.Add("Logo C·TRÊS não encontrado; os arquivos foram gerados somente com cores, títulos e tipografia institucionais.")
}
foreach ($directory in @($logsDirectory, $errorsDirectory, $stateDirectory)) {
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
}
$logPath = Join-Path $logsDirectory ((Get-Date).ToString("yyyy-MM-dd") + ".log")

$processes = @(
    [pscustomobject]@{
        Type = "EMBARQUES"
        Input = Join-Path $RaizConsolidacoes "01_CONSOLIDAÇÃO_DE_EMBARQUES\INPUT"
        Output = Join-Path $RaizConsolidacoes "01_CONSOLIDAÇÃO_DE_EMBARQUES\OUTPUT"
        Processed = Join-Path $RaizConsolidacoes "01_CONSOLIDAÇÃO_DE_EMBARQUES\PROCESSADOS"
    },
    [pscustomobject]@{
        Type = "DUIMP"
        Input = Join-Path $RaizConsolidacoes "02_CONSOLIDAÇÃO_DE_DUIMP\INPUT"
        Output = Join-Path $RaizConsolidacoes "02_CONSOLIDAÇÃO_DE_DUIMP\OUTPUT"
        Processed = Join-Path $RaizConsolidacoes "02_CONSOLIDAÇÃO_DE_DUIMP\PROCESSADOS"
    }
)

function Write-OperationalLog {
    param(
        [string] $Level,
        [string] $Message
    )

    $line = "{0} [{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}

function Write-Phase {
    param(
        [Parameter(Mandatory = $true)] [datetime] $StartedAt,
        [Parameter(Mandatory = $true)] [string] $Message
    )

    $elapsedSeconds = [int] [Math]::Floor(((Get-Date) - $StartedAt).TotalSeconds)
    $minutes = [int] [Math]::Floor($elapsedSeconds / 60)
    $seconds = $elapsedSeconds % 60
    Write-Host ("[{0:D2}:{1:D2}] {2}" -f $minutes, $seconds, $Message) -ForegroundColor DarkCyan
}

function Get-ElapsedText {
    param([Parameter(Mandatory = $true)] [datetime] $StartedAt)

    return ("{0:N1}s" -f ((Get-Date) - $StartedAt).TotalSeconds)
}

function Set-ExcelManualCalculation {
    param($Excel)

    try { $Excel.Calculation = -4135 } catch {}
    try { $Excel.CalculateBeforeSave = $false } catch {}
}

function Get-RelativePath {
    param(
        [string] $Root,
        [string] $Path
    )

    $rootWithSlash = $Root.TrimEnd("\") + "\"
    $rootUri = New-Object Uri($rootWithSlash)
    $pathUri = New-Object Uri([IO.Path]::GetFullPath($Path))
    return [Uri]::UnescapeDataString($rootUri.MakeRelativeUri($pathUri).ToString()).Replace("/", "\")
}

function Read-StateEntries {
    if (-not (Test-Path -LiteralPath $statePath)) {
        return @()
    }

    $entries = New-Object "System.Collections.Generic.List[object]"
    foreach ($line in Get-Content -LiteralPath $statePath -Encoding UTF8) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        try {
            $entries.Add(($line | ConvertFrom-Json))
        }
        catch {
            Write-OperationalLog "AVISO" "Uma linha inválida do histórico foi ignorada."
        }
    }
    return $entries.ToArray()
}

function Add-StateEntry {
    param(
        [string] $Hash,
        [string] $ProcessType,
        [string] $InputName,
        [string] $OutputPath,
        [string] $ArchivedPath,
        [string] $BatchId = "",
        [int] $BatchSize = 1,
        [switch] $Migrated
    )

    $entry = [ordered]@{
        Hash = $Hash
        ProcessType = $ProcessType
        InputName = $InputName
        OutputRelativePath = Get-RelativePath -Root $RaizConsolidacoes -Path $OutputPath
        ArchivedRelativePath = Get-RelativePath -Root $RaizConsolidacoes -Path $ArchivedPath
        Migrated = [bool] $Migrated
        ProcessedAt = (Get-Date).ToString("o")
        Computer = $env:COMPUTERNAME
        User = $env:USERNAME
    }
    if (-not [string]::IsNullOrWhiteSpace($BatchId)) {
        $entry.BatchId = $BatchId
        $entry.BatchSize = $BatchSize
    }
    Add-Content -LiteralPath $statePath -Value ($entry | ConvertTo-Json -Compress) -Encoding UTF8
}

function Add-StateEntriesBatch {
    param(
        [Parameter(Mandatory = $true)] [object[]] $Entries,
        [Parameter(Mandatory = $true)] [string] $BatchId,
        [Parameter(Mandatory = $true)] [int] $BatchSize
    )

    $lines = New-Object "System.Collections.Generic.List[string]"
    foreach ($item in $Entries) {
        $entry = [ordered]@{
            Hash = [string] $item.Hash
            ProcessType = "EMBARQUES"
            InputName = [string] $item.InputName
            OutputRelativePath = Get-RelativePath -Root $RaizConsolidacoes -Path $item.OutputPath
            ArchivedRelativePath = Get-RelativePath -Root $RaizConsolidacoes -Path $item.ArchivedPath
            Migrated = $false
            BatchId = $BatchId
            BatchSize = $BatchSize
            ProcessedAt = (Get-Date).ToString("o")
            Computer = $env:COMPUTERNAME
            User = $env:USERNAME
        }
        $lines.Add(($entry | ConvertTo-Json -Compress))
    }
    Add-Content -LiteralPath $statePath -Value $lines.ToArray() -Encoding UTF8
}

function Write-ErrorReport {
    param(
        [IO.FileInfo] $InputFile,
        [string] $ProcessType,
        [Exception] $Exception
    )

    $safeName = ConvertTo-SafeFilePart $InputFile.BaseName
    $reportPath = Join-Path $errorsDirectory (
        "{0}_{1}_{2}_ERRO.txt" -f (Get-Date).ToString("yyyyMMdd_HHmmssfff"), $ProcessType, $safeName
    )
    $guidance = @"
ARQUIVO: $($InputFile.Name)
PROCESSO: $ProcessType
DATA/HORA: $((Get-Date).ToString("dd/MM/yyyy HH:mm:ss"))

MOTIVO:
$($Exception.Message)

ORIENTAÇÃO:
- Confirme que o arquivo é .xlsx e está no INPUT correto.
- Feche o arquivo no Excel e aguarde a sincronização do OneDrive.
- Não altere a estrutura original da relação da seguradora.
- Depois, execute novamente PROCESSAR_CONSOLIDACOES.cmd.
"@
    Set-Content -LiteralPath $reportPath -Value $guidance -Encoding UTF8
    return $reportPath
}

function Write-ShipmentBatchErrorReport {
    param(
        [string] $BatchKey,
        [IO.FileInfo[]] $InputFiles,
        [Exception] $Exception
    )

    $safeKey = ConvertTo-SafeFilePart ($BatchKey -replace "\|", "_")
    if ([string]::IsNullOrWhiteSpace($safeKey)) { $safeKey = "LOTE_NAO_IDENTIFICADO" }
    $reportPath = Join-Path $errorsDirectory (
        "{0}_EMBARQUES_LOTE_{1}_ERRO.txt" -f (Get-Date).ToString("yyyyMMdd_HHmmssfff"), $safeKey
    )
    $fileLines = @($InputFiles | Sort-Object Name | ForEach-Object { "- $($_.Name)" }) -join [Environment]::NewLine
    $guidance = @"
PROCESSO: EMBARQUES — LOTE UNIFICADO
LOTE: $BatchKey
DATA/HORA: $((Get-Date).ToString("dd/MM/yyyy HH:mm:ss"))

ARQUIVOS DO LOTE:
$fileLines

MOTIVO:
$($Exception.Message)

ORIENTAÇÃO:
- Corrija ou substitua a relação indicada pela mensagem.
- Mantenha no INPUT todas as relações que devem compor o lote completo.
- Feche os arquivos no Excel e aguarde a sincronização do OneDrive.
- Depois, execute novamente PROCESSAR_CONSOLIDACOES.cmd.
"@
    Set-Content -LiteralPath $reportPath -Value $guidance -Encoding UTF8
    return $reportPath
}

function Test-FileReady {
    param([IO.FileInfo] $File)

    $before = Get-Item -LiteralPath $File.FullName
    Start-Sleep -Milliseconds 400
    $after = Get-Item -LiteralPath $File.FullName
    if ($before.Length -ne $after.Length -or $before.LastWriteTimeUtc -ne $after.LastWriteTimeUtc) {
        throw "O arquivo ainda está sendo sincronizado ou gravado."
    }

    $stream = $null
    try {
        $stream = [IO.File]::Open($File.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
    }
    catch {
        throw "O arquivo está aberto, bloqueado ou ainda sincronizando. Feche-o e tente novamente."
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Convert-ComValuesToRows {
    param($Values)

    $rows = New-Object "System.Collections.Generic.List[object[]]"
    if ($Values -isnot [Array] -or $Values.Rank -ne 2) {
        $rows.Add(@($Values))
        return $rows.ToArray()
    }

    $rowLower = $Values.GetLowerBound(0)
    $rowUpper = $Values.GetUpperBound(0)
    $columnLower = $Values.GetLowerBound(1)
    $columnUpper = $Values.GetUpperBound(1)
    for ($row = $rowLower; $row -le $rowUpper; $row++) {
        $current = New-Object object[] ($columnUpper - $columnLower + 1)
        for ($column = $columnLower; $column -le $columnUpper; $column++) {
            $current[$column - $columnLower] = $Values.GetValue($row, $column)
        }
        $rows.Add($current)
    }
    return $rows.ToArray()
}

function Release-ComObject {
    param($Object)

    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void] [Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Read-SourceData {
    param(
        $Excel,
        [string] $Path
    )

    $workbook = $null
    $sheet = $null
    $used = $null
    try {
        $workbook = $Excel.Workbooks.Open($Path, 0, $true)
        $sheet = $workbook.Worksheets.Item(1)
        $used = $sheet.UsedRange
        $rows = Convert-ComValuesToRows $used.Value2
        return [pscustomobject]@{
            Rows = $rows
            UsedRows = $used.Rows.Count
            UsedColumns = $used.Columns.Count
            SheetName = [string] $sheet.Name
        }
    }
    finally {
        if ($null -ne $workbook) {
            [void] $workbook.Close($false)
        }
        Release-ComObject $used
        Release-ComObject $sheet
        Release-ComObject $workbook
    }
}

function Get-ExcelColor {
    param([string] $Hex)

    $red = [Convert]::ToInt32($Hex.Substring(1, 2), 16)
    $green = [Convert]::ToInt32($Hex.Substring(3, 2), 16)
    $blue = [Convert]::ToInt32($Hex.Substring(5, 2), 16)
    return $red + ($green * 256) + ($blue * 65536)
}

function Set-RangeBorders {
    param(
        $Range,
        [int] $Color = 14277081
    )

    $Range.Borders.LineStyle = 1
    $Range.Borders.Weight = 2
    $Range.Borders.Color = $Color
}

function Assert-NoFormulaErrors {
    param($Range)

    $errorCells = $null
    try {
        $errorCells = $Range.SpecialCells(-4123, 16)
    }
    catch {
        $errorCells = $null
    }
    if ($null -ne $errorCells) {
        $count = $errorCells.Count
        Release-ComObject $errorCells
        throw "A planilha gerada contém $count fórmula(s) com erro."
    }
}

function Get-CtresBranding {
    return [pscustomobject]@{
        Navy = Get-ExcelColor "#003060"
        SecondaryBlue = Get-ExcelColor "#004090"
        Orange = Get-ExcelColor "#F05010"
        LightBlue = Get-ExcelColor "#EAF2F8"
        SoftGray = Get-ExcelColor "#F4F6F8"
        Border = Get-ExcelColor "#D5DFE8"
        DarkText = Get-ExcelColor "#243447"
        White = Get-ExcelColor "#FFFFFF"
        Font = "Aptos"
        LogoPath = $logoPath
        LogoCrop = [pscustomobject]@{
            Left = 138.0 / 1024.0
            Top = 369.0 / 1024.0
            Right = 146.0 / 1024.0
            Bottom = 368.0 / 1024.0
        }
    }
}

function Add-BrandingWarning {
    param([string] $Message)

    if (-not $brandingWarnings.Contains($Message)) {
        $brandingWarnings.Add($Message)
    }
}

function Remove-WorksheetShapes {
    param($Sheet)

    for ($index = $Sheet.Shapes.Count; $index -ge 1; $index--) {
        $shape = $null
        try {
            $shape = $Sheet.Shapes.Item($index)
            [void] $shape.Delete()
        }
        finally {
            Release-ComObject $shape
        }
    }
}

function Add-CtresLogo {
    param(
        $Sheet,
        [string] $TargetRangeAddress,
        $Branding
    )

    if (-not (Test-Path -LiteralPath $Branding.LogoPath -PathType Leaf)) {
        return $false
    }

    $target = $null
    $shape = $null
    try {
        $target = $Sheet.Range($TargetRangeAddress)
        $shape = $Sheet.Shapes.AddPicture(
            $Branding.LogoPath,
            0,
            -1,
            [single] ($target.Left + 5),
            [single] ($target.Top + 4),
            -1,
            -1
        )
        $shape.Name = "CTRES_LOGO"
        $shape.LockAspectRatio = -1
        $naturalWidth = [double] $shape.Width
        $naturalHeight = [double] $shape.Height
        $shape.PictureFormat.CropLeft = [single] ($naturalWidth * [double] $Branding.LogoCrop.Left)
        $shape.PictureFormat.CropTop = [single] ($naturalHeight * [double] $Branding.LogoCrop.Top)
        $shape.PictureFormat.CropRight = [single] ($naturalWidth * [double] $Branding.LogoCrop.Right)
        $shape.PictureFormat.CropBottom = [single] ($naturalHeight * [double] $Branding.LogoCrop.Bottom)
        $shape.AlternativeText = "Grupo C·TRÊS - Corretora de Seguros"
        $maximumWidth = [double] $target.Width - 10
        $maximumHeight = [double] $target.Height - 8
        $croppedWidth = [double] $shape.Width
        $croppedHeight = [double] $shape.Height
        $scale = [Math]::Min($maximumWidth / $croppedWidth, $maximumHeight / $croppedHeight)
        $shape.Width = [single] ($croppedWidth * $scale)
        $shape.Left = [single] ($target.Left + (($target.Width - $shape.Width) / 2))
        $shape.Top = [single] ($target.Top + (($target.Height - $shape.Height) / 2))
        return $true
    }
    catch {
        Add-BrandingWarning "O logo C·TRÊS não pôde ser aplicado; os arquivos foram gerados somente com cores, títulos e tipografia institucionais."
        return $false
    }
    finally {
        Release-ComObject $shape
        Release-ComObject $target
    }
}

function Set-BrandFallback {
    param(
        $Sheet,
        [string] $TargetCell,
        $Branding
    )

    $cell = $Sheet.Range($TargetCell)
    $cell.Value2 = "C·TRÊS`nCORRETORA DE SEGUROS"
    $cell.Font.Name = $Branding.Font
    $cell.Font.Bold = $true
    $cell.Font.Size = 12
    $cell.Font.Color = $Branding.Navy
    $cell.HorizontalAlignment = -4108
    $cell.VerticalAlignment = -4108
    $cell.WrapText = $true
    Release-ComObject $cell
}

function Set-CtresPageLayout {
    param(
        $Sheet,
        [string] $PrintArea,
        [string] $PrintTitleRows = ""
    )

    $Sheet.PageSetup.Orientation = 2
    $Sheet.PageSetup.Zoom = $false
    $Sheet.PageSetup.FitToPagesWide = 1
    $Sheet.PageSetup.FitToPagesTall = $false
    $Sheet.PageSetup.CenterHorizontally = $true
    $Sheet.PageSetup.LeftMargin = $Sheet.Application.CentimetersToPoints(0.5)
    $Sheet.PageSetup.RightMargin = $Sheet.Application.CentimetersToPoints(0.5)
    $Sheet.PageSetup.TopMargin = $Sheet.Application.CentimetersToPoints(0.7)
    $Sheet.PageSetup.BottomMargin = $Sheet.Application.CentimetersToPoints(0.7)
    $Sheet.PageSetup.PrintArea = $PrintArea
    if (-not [string]::IsNullOrWhiteSpace($PrintTitleRows)) {
        $Sheet.PageSetup.PrintTitleRows = $PrintTitleRows
    }
}

function Apply-CtresShipmentBranding {
    param(
        $Sheet,
        $Metadata,
        [int] $FirstDataRow,
        [int] $LastDataRow,
        [int] $TotalRow,
        $Branding
    )

    [void] $Sheet.Range("A1", "G4").Merge()
    [void] $Sheet.Range("H1", "V2").Merge()
    foreach ($address in @("H3:M3", "H4:M4", "N3:Q3", "N4:Q4", "R3:V3", "R4:V4")) {
        [void] $Sheet.Range($address).Merge()
    }

    $Sheet.Range("A1", "G4").Interior.Color = $Branding.White
    $Sheet.Range("H1", "V2").Interior.Color = $Branding.Navy
    $Sheet.Range("H1").Value2 = "RELAÇÃO CONSOLIDADA DE EMBARQUES"
    $Sheet.Range("H1").Font.Name = $Branding.Font
    $Sheet.Range("H1").Font.Bold = $true
    $Sheet.Range("H1").Font.Size = 17
    $Sheet.Range("H1").Font.Color = $Branding.White
    $Sheet.Range("H1").HorizontalAlignment = -4131
    $Sheet.Range("H1").VerticalAlignment = -4108

    foreach ($address in @("H3:M4", "N3:Q4", "R3:V4")) {
        $Sheet.Range($address).Interior.Color = $Branding.LightBlue
    }
    $Sheet.Range("H3").Value2 = "CLIENTE"
    $Sheet.Range("H4").NumberFormat = "@"
    $Sheet.Range("H4").Value2 = [string] $Metadata.Client
    $Sheet.Range("N3").Value2 = "COMPETÊNCIA"
    $Sheet.Range("N4").NumberFormat = "@"
    $Sheet.Range("N4").Value2 = [string] $Metadata.Reference
    $Sheet.Range("R3").Value2 = "PROCESSO"
    $Sheet.Range("R4").NumberFormat = "@"
    $Sheet.Range("R4").Value2 = "EMBARQUES"
    foreach ($address in @("H3", "N3", "R3")) {
        $Sheet.Range($address).Font.Name = $Branding.Font
        $Sheet.Range($address).Font.Bold = $true
        $Sheet.Range($address).Font.Size = 8
        $Sheet.Range($address).Font.Color = $Branding.SecondaryBlue
        $Sheet.Range($address).VerticalAlignment = -4108
    }
    foreach ($address in @("H4", "N4", "R4")) {
        $Sheet.Range($address).Font.Name = $Branding.Font
        $Sheet.Range($address).Font.Bold = $true
        $Sheet.Range($address).Font.Size = 10
        $Sheet.Range($address).Font.Color = $Branding.DarkText
        $Sheet.Range($address).VerticalAlignment = -4108
    }
    $Sheet.Range("A5", "V5").Interior.Color = $Branding.Orange
    $Sheet.Rows.Item(1).RowHeight = 25
    $Sheet.Rows.Item(2).RowHeight = 25
    $Sheet.Rows.Item(3).RowHeight = 17
    $Sheet.Rows.Item(4).RowHeight = 25
    $Sheet.Rows.Item(5).RowHeight = 4

    if (-not (Add-CtresLogo -Sheet $Sheet -TargetRangeAddress "A1:G4" -Branding $Branding)) {
        Set-BrandFallback -Sheet $Sheet -TargetCell "A1" -Branding $Branding
    }

    $header = $Sheet.Range("A6", "V6")
    $header.Interior.Color = $Branding.Navy
    $header.Font.Bold = $true
    $header.Font.Color = $Branding.White
    $header.Font.Name = $Branding.Font
    $header.Font.Size = 10
    $header.HorizontalAlignment = -4108
    $header.VerticalAlignment = -4108
    $header.WrapText = $true
    $header.RowHeight = 42

    $details = $Sheet.Range("A$FirstDataRow", "V$LastDataRow")
    $details.Font.Name = $Branding.Font
    $details.Font.Size = 9
    $details.Font.Color = $Branding.DarkText
    $details.VerticalAlignment = -4108
    $details.RowHeight = 19
    $Sheet.Range("A$FirstDataRow", "G$LastDataRow").HorizontalAlignment = -4131
    $Sheet.Range("H$FirstDataRow", "V$LastDataRow").HorizontalAlignment = -4152
    # Uma única regra condicional substitui chamadas COM por linha e mantém o
    # zebrado para qualquer quantidade de documentos.
    $details.FormatConditions.Delete()
    $alternateRows = $details.FormatConditions.Add(2, $null, "=MOD(LIN()-$FirstDataRow;2)=1")
    $alternateRows.Interior.Color = $Branding.LightBlue

    $total = $Sheet.Range("A$TotalRow", "V$TotalRow")
    $total.Interior.Color = $Branding.Navy
    $total.Font.Bold = $true
    $total.Font.Color = $Branding.White
    $total.Font.Name = $Branding.Font
    $total.Font.Size = 10
    $total.RowHeight = 23
    $total.Borders.Item(8).Color = $Branding.Orange
    $total.Borders.Item(8).Weight = 3

    Set-RangeBorders -Range $Sheet.Range("A6", "V$TotalRow") -Color $Branding.Border
    [void] $Sheet.Range("A6", "V$LastDataRow").AutoFilter()
    $Sheet.Tab.Color = $Branding.Orange
    Set-CtresPageLayout -Sheet $Sheet -PrintArea "`$A`$1:`$V`$$TotalRow" -PrintTitleRows "`$6:`$6"
}

function Apply-CtresDuimpBranding {
    param(
        $Sheet,
        $Metadata,
        [int] $HeaderRows,
        [int] $DocumentCount,
        [int] $BlockSize,
        [int] $FinalRows,
        $Branding
    )

    Remove-WorksheetShapes -Sheet $Sheet
    [void] $Sheet.Range("A1", "L5").UnMerge()
    [void] $Sheet.Range("A1", "L5").ClearContents()
    [void] $Sheet.Range("A1", "G2").Merge()
    [void] $Sheet.Range("A3", "G4").Merge()
    [void] $Sheet.Range("H1", "L4").Merge()

    $Sheet.Range("A1", "G2").Interior.Color = $Branding.Navy
    $Sheet.Range("A1").Value2 = "RELAÇÃO CONSOLIDADA DE DUIMP"
    $Sheet.Range("A1").Font.Name = $Branding.Font
    $Sheet.Range("A1").Font.Bold = $true
    $Sheet.Range("A1").Font.Size = 15
    $Sheet.Range("A1").Font.Color = $Branding.White
    $Sheet.Range("A1").HorizontalAlignment = -4131
    $Sheet.Range("A1").VerticalAlignment = -4108
    $Sheet.Range("A3", "G4").Interior.Color = $Branding.LightBlue
    $Sheet.Range("A3").Value2 = "CONSOLIDAÇÃO DE DOCUMENTOS DE IMPORTAÇÃO"
    $Sheet.Range("A3").Font.Name = $Branding.Font
    $Sheet.Range("A3").Font.Bold = $true
    $Sheet.Range("A3").Font.Size = 9
    $Sheet.Range("A3").Font.Color = $Branding.SecondaryBlue
    $Sheet.Range("A3").HorizontalAlignment = -4131
    $Sheet.Range("A3").VerticalAlignment = -4108
    $Sheet.Range("H1", "L4").Interior.Color = $Branding.White
    $Sheet.Range("A5", "L5").Interior.Color = $Branding.Orange
    $Sheet.Rows.Item(1).RowHeight = 24
    $Sheet.Rows.Item(2).RowHeight = 24
    $Sheet.Rows.Item(3).RowHeight = 18
    $Sheet.Rows.Item(4).RowHeight = 24
    $Sheet.Rows.Item(5).RowHeight = 4

    if (-not (Add-CtresLogo -Sheet $Sheet -TargetRangeAddress "H1:L4" -Branding $Branding)) {
        Set-BrandFallback -Sheet $Sheet -TargetCell "H1" -Branding $Branding
    }

    foreach ($address in @("B6", "E6", "G6", "B8")) {
        $Sheet.Range($address).NumberFormat = "@"
    }
    $Sheet.Range("B6").Value2 = [string] $Metadata.Policy
    $Sheet.Range("E6").Value2 = [string] $Metadata.Subgroup
    $Sheet.Range("G6").Value2 = [string] $Metadata.Reference
    $Sheet.Range("B8").Value2 = [string] $Metadata.Client

    $used = $Sheet.Range("A1", "L$FinalRows")
    $used.Font.Name = $Branding.Font
    $used.Font.Color = $Branding.DarkText
    $Sheet.Range("A1", "G2").Font.Color = $Branding.White
    $Sheet.Range("A6", "L8").Interior.Color = $Branding.SoftGray
    Set-RangeBorders -Range $Sheet.Range("A6", "L8") -Color $Branding.Border
    foreach ($address in @("A6", "D6", "F6", "A7", "D7", "A8")) {
        $Sheet.Range($address).Font.Bold = $true
        $Sheet.Range($address).Font.Color = $Branding.SecondaryBlue
    }

    # Os blocos têm tamanho fixo. Formatação condicional aplica as quatro
    # faixas visuais em lote, evitando centenas de chamadas COM por arquivo.
    $firstDocumentRow = $HeaderRows + 1
    $documentRange = $Sheet.Range("A$firstDocumentRow", "L$FinalRows")
    $documentRange.FormatConditions.Delete()

    $blockHeader = $documentRange.FormatConditions.Add(2, $null, "=MOD(LIN()-$firstDocumentRow;$BlockSize)=0")
    $blockHeader.Interior.Color = $Branding.SecondaryBlue
    $blockHeader.Font.Bold = $true
    $blockHeader.Font.Color = $Branding.White

    $blockSecondRow = $documentRange.FormatConditions.Add(2, $null, "=MOD(LIN()-$firstDocumentRow;$BlockSize)=1")
    $blockSecondRow.Interior.Color = $Branding.LightBlue

    $financialHeader = $documentRange.FormatConditions.Add(2, $null, "=MOD(LIN()-$firstDocumentRow;$BlockSize)=$($BlockSize - 9)")
    $financialHeader.Interior.Color = $Branding.SoftGray
    $financialHeader.Font.Bold = $true
    $financialHeader.Font.Color = $Branding.Navy

    $blockTotal = $documentRange.FormatConditions.Add(2, $null, "=MOD(LIN()-$firstDocumentRow;$BlockSize)=$($BlockSize - 2)")
    $blockTotal.Interior.Color = $Branding.LightBlue
    $blockTotal.Font.Bold = $true

    # As alturas já vêm do arquivo da seguradora. Mantê-las evita chamadas COM
    # por bloco e conserva a separação visual original entre documentos.

    $Sheet.Tab.Color = $Branding.Navy
    Set-CtresPageLayout -Sheet $Sheet -PrintArea "`$A`$1:`$L`$$FinalRows"
}

function Add-ShipmentWorksheet {
    param(
        $Excel,
        $Sheet,
        $Metadata,
        [object[]] $Records,
        [string] $SheetName
    )

    $used = $null
    try {
        $branding = Get-CtresBranding
        $Sheet.Name = $SheetName

        $headers = @(
            "APÓLICE", "SUBGRUPO", "DIVISÃO", "TAXA RISCO", "REFERÊNCIA", "DI / DUIMP", "PO",
            "Valor FOB USD", "MOEDA", "CÂMBIO", "Valor FOB BRL", "Prêmio FOB",
            "Frete USD", "Frete BRL", "Prêmio Frete", "Despesas", "Prêmio Despesas",
            "Lucro Esperado", "Prêmio Lucro Esperado", "Impostos", "Prêmio Impostos", "Prêmio Total"
        )
        $headerMatrix = New-Object "object[,]" 1, 22
        for ($column = 0; $column -lt 22; $column++) {
            $headerMatrix[0, $column] = $headers[$column]
        }
        $Sheet.Range("A6", "V6").Value2 = $headerMatrix

        $dataMatrix = New-Object "object[,]" $Records.Count, 22
        for ($index = 0; $index -lt $Records.Count; $index++) {
            $record = $Records[$index]
            $currencyValue = if ($record.Currency -match "^\d+$") {
                [int] $record.Currency
            }
            else {
                [string] $record.Currency
            }
            $values = @(
                [string] $Metadata.Policy,
                [string] $Metadata.Subgroup,
                [string] $Metadata.Client,
                [double] $record.RiskRate,
                [string] $Metadata.Reference,
                [string] $record.Document,
                "",
                [double] $record.FobUsd,
                $currencyValue,
                [double] $record.Exchange,
                $null, $null,
                [double] $record.FreightUsd,
                $null, $null, $null, $null, $null, $null,
                [double] $record.TaxesBrl,
                $null, $null
            )
            for ($column = 0; $column -lt 22; $column++) {
                $dataMatrix[$index, $column] = $values[$column]
            }
        }

        $firstDataRow = 7
        $lastDataRow = $Records.Count + 6
        $totalRow = $lastDataRow + 1
        $Sheet.Range("A$firstDataRow", "C$lastDataRow").NumberFormat = "@"
        $Sheet.Range("E$firstDataRow", "G$lastDataRow").NumberFormat = "@"
        $Sheet.Range("A$firstDataRow", "V$lastDataRow").Value2 = $dataMatrix

        $Sheet.Range("K$firstDataRow", "K$lastDataRow").FormulaR1C1 = "=RC[-3]*RC[-1]"
        $Sheet.Range("L$firstDataRow", "L$lastDataRow").FormulaR1C1 = "=RC[-1]*RC[-8]"
        $Sheet.Range("N$firstDataRow", "N$lastDataRow").FormulaR1C1 = "=RC[-1]*RC[-4]"
        $Sheet.Range("O$firstDataRow", "O$lastDataRow").FormulaR1C1 = "=RC[-1]*RC[-11]"
        $Sheet.Range("P$firstDataRow", "P$lastDataRow").FormulaR1C1 = "=(RC[-5]+RC[-2])*10%"
        $Sheet.Range("Q$firstDataRow", "Q$lastDataRow").FormulaR1C1 = "=RC[-1]*RC[-13]"
        $Sheet.Range("R$firstDataRow", "R$lastDataRow").FormulaR1C1 = "=(RC[-7]+RC[-4]+RC[-2])*10%"
        $Sheet.Range("S$firstDataRow", "S$lastDataRow").FormulaR1C1 = "=RC[-1]*RC[-15]"
        $Sheet.Range("U$firstDataRow", "U$lastDataRow").FormulaR1C1 = "=RC[-1]*RC[-17]"
        $Sheet.Range("V$firstDataRow", "V$lastDataRow").FormulaR1C1 = "=SUM(RC[-10],RC[-7],RC[-5],RC[-3],RC[-1])"

        $Sheet.Cells.Item($totalRow, 1).Value2 = "SUBTOTAL"
        foreach ($column in @(8, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22)) {
            $Sheet.Cells.Item($totalRow, $column).FormulaR1C1 = "=SUBTOTAL(9,R${firstDataRow}C:R[-1]C)"
        }

        $Sheet.Range("D$firstDataRow", "D$lastDataRow").NumberFormatLocal = "0,0000%"
        foreach ($column in @("H", "M")) {
            $Sheet.Range("$column$firstDataRow", "$column$totalRow").NumberFormatLocal = '"US$" #.##0,00'
        }
        $Sheet.Range("I$firstDataRow", "I$lastDataRow").NumberFormatLocal = "000"
        $Sheet.Range("J$firstDataRow", "J$lastDataRow").NumberFormatLocal = "0,0000"
        foreach ($column in @("K", "L", "N", "O", "P", "Q", "R", "S", "T", "U", "V")) {
            $Sheet.Range("$column$firstDataRow", "$column$totalRow").NumberFormatLocal = '"R$" #.##0,00'
        }

        $widths = @(17, 11, 34, 13, 13, 46, 14, 19, 10, 12, 20, 16, 18, 19, 16, 19, 17, 19, 20, 19, 17, 17)
        for ($column = 1; $column -le 22; $column++) {
            $Sheet.Columns.Item($column).ColumnWidth = $widths[$column - 1]
        }
        Apply-CtresShipmentBranding `
            -Sheet $Sheet `
            -Metadata $Metadata `
            -FirstDataRow $firstDataRow `
            -LastDataRow $lastDataRow `
            -TotalRow $totalRow `
            -Branding $branding
        [void] $Sheet.Activate()
        $Excel.ActiveWindow.SplitRow = 6
        $Excel.ActiveWindow.FreezePanes = $true
        $Excel.ActiveWindow.DisplayGridlines = $false

        # O cálculo manual evita recálculos globais a cada aba. Esta chamada
        # calcula somente a aba criada antes da validação financeira abaixo.
        [void] $Sheet.Calculate()
        $used = $Sheet.Range("A1", "V$totalRow")
        Assert-NoFormulaErrors $used

        [decimal] $expectedPremium = 0
        foreach ($record in $Records) {
            [decimal] $fobBrl = [decimal] $record.FobUsd * [decimal] $record.Exchange
            [decimal] $freightBrl = [decimal] $record.FreightUsd * [decimal] $record.Exchange
            [decimal] $expenses = ($fobBrl + $freightBrl) * [decimal] 0.10
            [decimal] $profit = ($fobBrl + $freightBrl + $expenses) * [decimal] 0.10
            $expectedPremium += ($fobBrl + $freightBrl + $expenses + $profit + [decimal] $record.TaxesBrl) * [decimal] $record.RiskRate
        }
        $actualPremium = [decimal] $Sheet.Cells.Item($totalRow, 22).Value2
        if ([Math]::Abs([double] ($actualPremium - $expectedPremium)) -gt 0.01) {
            throw "O prêmio total não conciliou com os registros extraídos."
        }

        return [pscustomobject]@{
            SheetName = $SheetName
            Documents = $Records.Count
            Premium = [math]::Round([double] $actualPremium, 2)
        }
    }
    finally {
        Release-ComObject $used
    }
}

function Write-ShipmentBatchWorkbook {
    param(
        $Excel,
        [object[]] $Items,
        [string] $TemporaryPath
    )

    $workbook = $null
    $sheet = $null
    try {
        if ($Items.Count -eq 0) {
            throw "O lote de embarques não contém relações válidas."
        }
        $workbook = $Excel.Workbooks.Add()
        Set-ExcelManualCalculation -Excel $Excel
        while ($workbook.Worksheets.Count -gt 1) {
            [void] $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
        }

        $usedNames = New-Object "System.Collections.Generic.HashSet[string]" ([StringComparer]::OrdinalIgnoreCase)
        $results = New-Object "System.Collections.Generic.List[object]"
        for ($index = 0; $index -lt $Items.Count; $index++) {
            if ($index -eq 0) {
                $sheet = $workbook.Worksheets.Item(1)
            }
            else {
                $lastSheet = $workbook.Worksheets.Item($workbook.Worksheets.Count)
                try {
                    $sheet = $workbook.Worksheets.Add([Type]::Missing, $lastSheet)
                }
                finally {
                    Release-ComObject $lastSheet
                }
            }
            $item = $Items[$index]
            $sheetName = Get-UniqueShipmentWorksheetName `
                -Client $item.Metadata.Client `
                -Subgroup $item.Metadata.Subgroup `
                -UsedNames $usedNames
            $results.Add((Add-ShipmentWorksheet `
                -Excel $Excel `
                -Sheet $sheet `
                -Metadata $item.Metadata `
                -Records $item.Records `
                -SheetName $sheetName))
            Release-ComObject $sheet
            $sheet = $null
        }

        [void] $workbook.Worksheets.Item(1).Activate()
        $Excel.ActiveWindow.SplitRow = 6
        $Excel.ActiveWindow.FreezePanes = $true
        $Excel.ActiveWindow.DisplayGridlines = $false
        [void] $workbook.SaveAs($TemporaryPath, 51)
        return [pscustomobject]@{
            Sheets = $results.Count
            Documents = [int] (($results | Measure-Object -Property Documents -Sum).Sum)
            Results = $results.ToArray()
        }
    }
    finally {
        if ($null -ne $workbook) {
            [void] $workbook.Close($false)
        }
        Release-ComObject $sheet
        Release-ComObject $workbook
    }
}

function Write-DuimpWorkbook {
    param(
        $Excel,
        $Metadata,
        [string] $SourcePath,
        [object[]] $Rows,
        [int] $OriginalUsedRows,
        [string] $TemporaryPath
    )

    $consolidationStartedAt = Get-Date
    $sourceBlocks = @(Get-DocumentBlocks -Rows $Rows)
    $layout = Get-DuimpBlockLayout -Blocks $sourceBlocks
    $blockSize = [int] $layout.BlockSize
    $fixedBlocks = New-Object "System.Collections.Generic.List[object]"
    foreach ($block in $sourceBlocks) {
        $fixedBlocks.Add((New-DocumentBlock -Rows @($block.Rows[0..($blockSize - 1)]) -StartIndex $block.StartIndex))
    }
    $sourceBlocks = $fixedBlocks.ToArray()
    Initialize-DuimpBlockLayoutCache -Blocks $sourceBlocks
    Write-Phase -StartedAt $runStartedAt -Message "Blocos DUIMP identificados: $($sourceBlocks.Count); agrupando valores..."

    $headerRows = [int] $layout.HeaderRows

    $firstDuimpIndex = -1
    for ($index = 0; $index -lt $sourceBlocks.Count; $index++) {
        if ($sourceBlocks[$index].Type -eq "DUIMP") {
            $firstDuimpIndex = $index
            break
        }
    }
    if ($firstDuimpIndex -lt 0) {
        throw "Nenhuma DUIMP fracionada foi encontrada."
    }

    $consolidated = @(Merge-DuimpDocumentBlocks -Blocks $sourceBlocks)
    Initialize-DuimpBlockLayoutCache -Blocks $consolidated
    $differences = @(Compare-DuimpFinancialTotals -OriginalBlocks $sourceBlocks -ConsolidatedBlocks $consolidated)
    if ($differences.Count -gt 0) {
        throw "A consolidação de DUIMP apresentou divergências financeiras."
    }
    Write-Phase -StartedAt $runStartedAt -Message "Consolidação financeira validada em $(Get-ElapsedText $consolidationStartedAt); montando workbook..."

    Copy-Item -LiteralPath $SourcePath -Destination $TemporaryPath
    [void] (Remove-XlsxSheetProtection -Path $TemporaryPath)
    $workbook = $null
    $sheet = $null
    $targetRange = $null
    $rowsToDelete = $null
    $focusCell = $null
    $viewWindow = $null
    try {
        $branding = Get-CtresBranding
        $workbook = $Excel.Workbooks.Open($TemporaryPath, 0, $false)
        Set-ExcelManualCalculation -Excel $Excel
        $sheet = $workbook.Worksheets.Item(1)

        $remainingBlocks = @($consolidated[$firstDuimpIndex..($consolidated.Count - 1)])
        $flatRows = New-Object "System.Collections.Generic.List[object[]]"
        foreach ($block in $remainingBlocks) {
            foreach ($row in @($block.Rows[0..($blockSize - 1)])) {
                $flatRows.Add(@($row))
            }
        }
        $matrix = ConvertTo-TwoDimensionalArray -Rows $flatRows.ToArray() -Columns 12
        $firstExcelRow = $headerRows + ($firstDuimpIndex * $blockSize) + 1
        $lastWriteRow = $firstExcelRow + $flatRows.Count - 1
        $targetRange = $sheet.Range("A$firstExcelRow", "L$lastWriteRow")
        $targetRange.Value2 = $matrix

        $finalRows = $headerRows + ($consolidated.Count * $blockSize)
        if ($finalRows -lt $OriginalUsedRows) {
            $rowsToDelete = $sheet.Rows.Item("$($finalRows + 1):$OriginalUsedRows")
            [void] $rowsToDelete.Delete()
        }

        $documents = @($consolidated | ForEach-Object { $_.Document })
        $documentSummary = Test-DuimpResultDocuments -Documents $documents
        Apply-CtresDuimpBranding `
            -Sheet $sheet `
            -Metadata $Metadata `
            -HeaderRows $headerRows `
            -DocumentCount $consolidated.Count `
            -BlockSize $blockSize `
            -FinalRows $finalRows `
            -Branding $branding
        for ($worksheetIndex = 2; $worksheetIndex -le $workbook.Worksheets.Count; $worksheetIndex++) {
            $auxiliarySheet = $null
            try {
                $auxiliarySheet = $workbook.Worksheets.Item($worksheetIndex)
                $auxiliarySheet.Visible = 0
            }
            finally {
                Release-ComObject $auxiliarySheet
            }
        }
        [void] $sheet.Activate()
        $viewWindow = $workbook.Windows.Item(1)
        if ($viewWindow.FreezePanes) {
            $viewWindow.FreezePanes = $false
        }
        $viewWindow.SplitRow = 0
        $viewWindow.SplitColumn = 0
        $viewWindow.ScrollRow = 1
        $viewWindow.ScrollColumn = 1
        $focusCell = $sheet.Cells.Item($headerRows + 1, 1)
        [void] $focusCell.Select()
        $viewWindow.SplitRow = $headerRows
        $viewWindow.FreezePanes = $true
        $viewWindow.DisplayGridlines = $false

        [void] $workbook.Save()
        return [pscustomobject]@{
            SourceBlocks = $sourceBlocks.Count
            FinalBlocks = $consolidated.Count
            DiDocuments = @($sourceBlocks | Where-Object { $_.Type -eq "DI" }).Count
            DuimpDocuments = @($documents | Where-Object { $_ -match "^\d{2}/BR\d+-\d$" }).Count
            RepeatedOccurrences = $documentSummary.RepeatedOccurrences
        }
    }
    finally {
        if ($null -ne $workbook) {
            [void] $workbook.Close($true)
        }
        Release-ComObject $rowsToDelete
        Release-ComObject $targetRange
        Release-ComObject $focusCell
        Release-ComObject $viewWindow
        Release-ComObject $sheet
        Release-ComObject $workbook
    }
}

function Get-ProcessForFile {
    param([string] $Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    foreach ($process in $processes) {
        $inputWithSlash = [IO.Path]::GetFullPath($process.Input).TrimEnd("\") + "\"
        if ($fullPath.StartsWith($inputWithSlash, [StringComparison]::OrdinalIgnoreCase)) {
            return $process
        }
    }
    throw "O arquivo selecionado não está em um dos INPUTs permitidos."
}

function Get-InputFiles {
    $files = New-Object "System.Collections.Generic.List[IO.FileInfo]"
    foreach ($process in $processes) {
        New-Item -ItemType Directory -Path $process.Input -Force | Out-Null
        New-Item -ItemType Directory -Path $process.Output -Force | Out-Null
        New-Item -ItemType Directory -Path $process.Processed -Force | Out-Null
        foreach ($file in Get-ChildItem -LiteralPath $process.Input -File -Filter "*.xlsx") {
            if (-not $file.Name.StartsWith("~$")) {
                $files.Add($file)
            }
        }
    }
    return $files.ToArray()
}

$lockStream = $null
$excel = $null
$excelSettings = $null
$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) "CtresConsolidacoes"
$runStartedAt = Get-Date
$completed = 0
$completedShipmentBatches = 0
$completedShipmentFiles = 0
$completedShipmentSheets = 0
$completedDuimpFiles = 0
$rejectedShipmentBatches = 0
$archivedPrevious = 0
$rejected = 0

try {
    try {
        $lockStream = [IO.File]::Open($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    }
    catch {
        throw "Já existe um processamento em andamento nesta máquina."
    }

    $shipmentAliases = Read-ShipmentGroupAliases -Path $shipmentGroupsPath

    $inputFiles = @(Get-InputFiles)
    if ($inputFiles.Count -eq 0) {
        Write-Host "Nenhum arquivo .xlsx foi encontrado para processamento." -ForegroundColor Yellow
        Write-OperationalLog "INFO" "Nenhum arquivo encontrado."
        exit 0
    }

    $stateEntries = @(Read-StateEntries)
    $pendingJobs = New-Object "System.Collections.Generic.List[object]"
    $legacyJobs = New-Object "System.Collections.Generic.List[object]"
    $preflightRejections = New-Object "System.Collections.Generic.List[object]"

    foreach ($inputFile in $inputFiles) {
        $process = $null
        try {
            $process = Get-ProcessForFile -Path $inputFile.FullName
            if ($inputFile.Extension -ne ".xlsx" -or $inputFile.Name.StartsWith("~$")) {
                throw "Somente arquivos .xlsx não temporários são aceitos."
            }

            Test-FileReady $inputFile
            $hash = (Get-FileHash -LiteralPath $inputFile.FullName -Algorithm SHA256).Hash
            $legacyEntry = Get-LegacyProcessedEntry `
                -Entries $stateEntries `
                -Hash $hash `
                -ProcessType $process.Type `
                -InputName $inputFile.Name `
                -RootDirectory $RaizConsolidacoes
            if ($null -ne $legacyEntry) {
                $legacyJobs.Add([pscustomobject]@{
                    InputFile = $inputFile
                    Process = $process
                    Hash = $hash
                    StateEntry = $legacyEntry
                })
                continue
            }

            $pendingJobs.Add([pscustomobject]@{
                InputFile = $inputFile
                Process = $process
                Hash = $hash
            })
        }
        catch {
            $rejected++
            $processType = if ($null -eq $process) { "DESCONHECIDO" } else { $process.Type }
            $report = if ($processType -eq "EMBARQUES") {
                $null
            }
            else {
                Write-ErrorReport -InputFile $inputFile -ProcessType $processType -Exception $_.Exception
            }
            $preflightRejections.Add([pscustomobject]@{
                InputFile = $inputFile
                ProcessType = $processType
                Message = $_.Exception.Message
                Report = $report
            })
            Write-OperationalLog "REJEITADO NA PRÉ-ANÁLISE" "$($inputFile.Name) | $processType | $($_.Exception.Message)"
        }
    }

    Write-Host ""
    Write-Host "Fila encontrada: $($inputFiles.Count) arquivo(s)" -ForegroundColor Cyan
    Write-Host "  Novos para processar: $($pendingJobs.Count)"
    Write-Host "  Já processados para arquivar: $($legacyJobs.Count)"
    Write-Host "  Rejeitados na pré-análise: $($preflightRejections.Count)"
    Write-Host ""

    foreach ($preflightRejection in $preflightRejections) {
        Write-Host "[REJEITADO NA PRÉ-ANÁLISE] $($preflightRejection.InputFile.Name): $($preflightRejection.Message)" -ForegroundColor Red
        if (-not [string]::IsNullOrWhiteSpace([string] $preflightRejection.Report)) {
            Write-Host "                            Relatório: $($preflightRejection.Report)" -ForegroundColor DarkYellow
        }
    }

    foreach ($legacyJob in $legacyJobs) {
        $inputFile = $legacyJob.InputFile
        $process = $legacyJob.Process
        $archivePath = $null
        $sourceArchived = $false
        try {
            Test-FileReady $inputFile
            $currentHash = (Get-FileHash -LiteralPath $inputFile.FullName -Algorithm SHA256).Hash
            if ($currentHash -ne $legacyJob.Hash) {
                throw "O arquivo foi alterado após a pré-análise. Execute novamente para arquivar a versão atual."
            }

            $archivePath = Get-NextOutputPath -OutputDirectory $process.Processed -FileName $inputFile.Name
            $legacyOutputPath = Join-Path $RaizConsolidacoes ([string] $legacyJob.StateEntry.OutputRelativePath)
            Move-Item -LiteralPath $inputFile.FullName -Destination $archivePath
            $sourceArchived = $true
            try {
                Add-StateEntry `
                    -Hash $legacyJob.Hash `
                    -ProcessType $process.Type `
                    -InputName $inputFile.Name `
                    -OutputPath $legacyOutputPath `
                    -ArchivedPath $archivePath `
                    -Migrated
            }
            catch {
                if ($sourceArchived -and
                    (Test-Path -LiteralPath $archivePath) -and
                    -not (Test-Path -LiteralPath $inputFile.FullName)) {
                    Move-Item -LiteralPath $archivePath -Destination $inputFile.FullName
                    $sourceArchived = $false
                }
                throw
            }

            $archivedPrevious++
            Write-Host "[JÁ PROCESSADO — ARQUIVADO] $($inputFile.Name) -> $([IO.Path]::GetFileName($archivePath))" -ForegroundColor DarkGray
            Write-OperationalLog "JÁ PROCESSADO — ARQUIVADO" "$($inputFile.Name) | $($process.Type) | original movido para PROCESSADOS."
        }
        catch {
            $rejected++
            $report = Write-ErrorReport -InputFile $inputFile -ProcessType $process.Type -Exception $_.Exception
            Write-Host "[REJEITADO AO ARQUIVAR] $($inputFile.Name): $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "                         Relatório: $report" -ForegroundColor DarkYellow
            Write-OperationalLog "REJEITADO AO ARQUIVAR" "$($inputFile.Name) | $($process.Type) | $($_.Exception.Message)"
        }
    }

    if ($pendingJobs.Count -eq 0) {
        Write-Host "Nenhuma demanda nova para processar." -ForegroundColor Green
    }
    else {
        Write-Phase -StartedAt $runStartedAt -Message "Preparando Excel para a fila..."
        $excel = New-Object -ComObject Excel.Application
        $excelSettings = [pscustomobject]@{
            ScreenUpdating = $excel.ScreenUpdating
            EnableEvents = $excel.EnableEvents
            DisplayAlerts = $excel.DisplayAlerts
            AskToUpdateLinks = $excel.AskToUpdateLinks
            Calculation = $excel.Calculation
            CalculateBeforeSave = $excel.CalculateBeforeSave
        }
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.AutomationSecurity = 3
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $false
        # Algumas instalações corporativas só permitem alterar o modo de
        # cálculo depois que há um workbook aberto. Não interromper a fila
        # nesses casos; os geradores também tentam aplicar o modo manual.
        try { $excel.Calculation = -4135 } catch {
            Write-OperationalLog "AVISO" "O Excel não aceitou o cálculo manual antes da abertura do workbook."
        }
        try { $excel.CalculateBeforeSave = $false } catch {}

        $shipmentJobs = @($pendingJobs | Where-Object { $_.Process.Type -eq "EMBARQUES" })
        $duimpJobs = @($pendingJobs | Where-Object { $_.Process.Type -eq "DUIMP" })
        $shipmentGroups = @{}
        $shipmentFailures = New-Object "System.Collections.Generic.List[object]"

        foreach ($preflightRejection in @($preflightRejections | Where-Object { $_.ProcessType -eq "EMBARQUES" })) {
            $shipmentFailures.Add([pscustomobject]@{
                Key = Get-ShipmentBatchKeyFromFileName -FileName $preflightRejection.InputFile.Name
                InputFile = $preflightRejection.InputFile
                Exception = [Exception]::new($preflightRejection.Message)
                AlreadyCounted = $true
            })
        }

        foreach ($job in $shipmentJobs) {
            $metadata = $null
            try {
                $readStartedAt = Get-Date
                Write-Phase -StartedAt $runStartedAt -Message "Lendo relação de Embarques: $($job.InputFile.Name)"
                Test-FileReady $job.InputFile
                $currentHash = (Get-FileHash -LiteralPath $job.InputFile.FullName -Algorithm SHA256).Hash
                if ($currentHash -ne $job.Hash) {
                    throw "O arquivo foi alterado após a pré-análise. Execute novamente para processar a versão atual."
                }
                $source = Read-SourceData -Excel $excel -Path $job.InputFile.FullName
                Write-Phase -StartedAt $runStartedAt -Message "Relação lida em $(Get-ElapsedText $readStartedAt); validando estrutura..."
                $metadata = Get-WorkbookMetadata -Rows $source.Rows
                $key = Get-ShipmentBatchKey -Metadata $metadata
                $records = @(ConvertTo-ShipmentRecords -Rows $source.Rows)
                if (-not $shipmentGroups.ContainsKey($key)) {
                    $shipmentGroups[$key] = New-Object "System.Collections.Generic.List[object]"
                }
                $shipmentGroups[$key].Add([pscustomobject]@{
                    Job = $job
                    InputFile = $job.InputFile
                    Hash = $job.Hash
                    Metadata = $metadata
                    Records = $records
                })
            }
            catch {
                $failureKey = if ($null -ne $metadata) {
                    try { Get-ShipmentBatchKey -Metadata $metadata } catch { $null }
                }
                else {
                    Get-ShipmentBatchKeyFromFileName -FileName $job.InputFile.Name
                }
                $shipmentFailures.Add([pscustomobject]@{
                    Key = $failureKey
                    InputFile = $job.InputFile
                    Exception = $_.Exception
                    AlreadyCounted = $false
                })
                $rejected++
            }
        }

        $unknownShipmentFailure = @($shipmentFailures | Where-Object { [string]::IsNullOrWhiteSpace([string] $_.Key) }).Count -gt 0
        $blockedKeys = @($shipmentFailures | Where-Object { -not [string]::IsNullOrWhiteSpace([string] $_.Key) } | ForEach-Object { $_.Key } | Select-Object -Unique)
        $groupKeys = @($shipmentGroups.Keys | Sort-Object)
        Write-Host "Lotes de Embarques identificados: $($groupKeys.Count) lote(s), $($shipmentJobs.Count) arquivo(s)." -ForegroundColor Cyan

        if ($unknownShipmentFailure -and ($shipmentJobs.Count -gt 0 -or $shipmentFailures.Count -gt 0)) {
            $allShipmentFiles = @(
                @($shipmentJobs | ForEach-Object { $_.InputFile }) +
                @($shipmentFailures | ForEach-Object { $_.InputFile })
            ) | Sort-Object FullName -Unique
            $exception = [Exception]::new("Um arquivo bloqueado ou inválido não permitiu identificar apólice e competência; todos os lotes de Embarques desta execução foram suspensos.")
            $report = Write-ShipmentBatchErrorReport -BatchKey "NÃO IDENTIFICADO" -InputFiles $allShipmentFiles -Exception $exception
            $validNotCounted = @($shipmentJobs | Where-Object {
                $name = $_.InputFile.FullName
                @($shipmentFailures | Where-Object { $_.InputFile.FullName -eq $name }).Count -eq 0
            }).Count
            $rejected += $validNotCounted
            $rejectedShipmentBatches += [Math]::Max(1, $groupKeys.Count)
            Write-Host "[LOTES DE EMBARQUES SUSPENSOS] $($exception.Message)" -ForegroundColor Red
            Write-Host "                                 Relatório: $report" -ForegroundColor DarkYellow
            Write-OperationalLog "LOTES DE EMBARQUES SUSPENSOS" $exception.Message
        }
        else {
            $batchPosition = 0
            foreach ($key in $groupKeys) {
                $batchPosition++
                $items = @($shipmentGroups[$key].ToArray() | Sort-Object `
                    @{ Expression = { [int] $_.Metadata.Subgroup } }, `
                    @{ Expression = { [string] $_.Metadata.Client } })
                $failuresForKey = @($shipmentFailures | Where-Object { $_.Key -eq $key })
                if ($failuresForKey.Count -gt 0) {
                    $files = @(@($items | ForEach-Object { $_.InputFile }) + @($failuresForKey | ForEach-Object { $_.InputFile })) | Sort-Object FullName -Unique
                    $report = Write-ShipmentBatchErrorReport -BatchKey $key -InputFiles $files -Exception $failuresForKey[0].Exception
                    $rejected += $items.Count
                    $rejectedShipmentBatches++
                    Write-Host "[LOTE REJEITADO $batchPosition/$($groupKeys.Count)] ${key}: $($failuresForKey[0].Exception.Message)" -ForegroundColor Red
                    Write-Host "                                             Relatório: $report" -ForegroundColor DarkYellow
                    Write-OperationalLog "LOTE REJEITADO" "$key | $($files.Count) arquivo(s) | $($failuresForKey[0].Exception.Message)"
                    continue
                }

                $jobDirectory = $null
                Write-Host "[PROCESSANDO LOTE $batchPosition/$($groupKeys.Count)] $key | $($items.Count) arquivo(s)..." -ForegroundColor Cyan
                try {
                    foreach ($item in $items) {
                        Test-FileReady $item.InputFile
                        $currentHash = (Get-FileHash -LiteralPath $item.InputFile.FullName -Algorithm SHA256).Hash
                        if ($currentHash -ne $item.Hash) {
                            throw "O arquivo $($item.InputFile.Name) foi alterado após a leitura do lote."
                        }
                    }

                    $firstMetadata = $items[0].Metadata
                    $process = $items[0].Job.Process
                    $fileName = Get-ShipmentBatchOutputFileName `
                        -Policy $firstMetadata.Policy `
                        -Reference $firstMetadata.Reference `
                        -Aliases $shipmentAliases
                    $outputPath = Get-NextOutputPath -OutputDirectory $process.Output -FileName $fileName
                    $jobDirectory = Join-Path $temporaryRoot ([guid]::NewGuid().ToString("N"))
                    New-Item -ItemType Directory -Path $jobDirectory -Force | Out-Null
                    $temporaryPath = Join-Path $jobDirectory "resultado.xlsx"
                    $generationStartedAt = Get-Date
                    Write-Phase -StartedAt $runStartedAt -Message "Gerando e formatando lote de Embarques $batchPosition/$($groupKeys.Count)..."
                    $result = Write-ShipmentBatchWorkbook -Excel $excel -Items $items -TemporaryPath $temporaryPath
                    Write-Phase -StartedAt $runStartedAt -Message "Lote pronto em $(Get-ElapsedText $generationStartedAt); publicando..."
                    if (-not (Test-Path -LiteralPath $temporaryPath -PathType Leaf)) {
                        throw "O workbook unificado temporário não foi criado."
                    }
                    if (Test-Path -LiteralPath $outputPath) {
                        $outputPath = Get-NextOutputPath -OutputDirectory $process.Output -FileName $fileName
                    }

                    $publishSources = New-Object "System.Collections.Generic.List[object]"
                    $stateItems = New-Object "System.Collections.Generic.List[object]"
                    foreach ($item in $items) {
                        $archivePath = Get-NextOutputPath -OutputDirectory $process.Processed -FileName $item.InputFile.Name
                        $publishSources.Add([pscustomobject]@{
                            SourcePath = $item.InputFile.FullName
                            ArchivePath = $archivePath
                        })
                        $stateItems.Add([pscustomobject]@{
                            Hash = $item.Hash
                            InputName = $item.InputFile.Name
                            OutputPath = $outputPath
                            ArchivedPath = $archivePath
                        })
                    }
                    $batchId = [guid]::NewGuid().ToString("N")
                    $batchSize = $items.Count
                    $stateArray = $stateItems.ToArray()
                    $commitState = {
                        Add-StateEntriesBatch -Entries $stateArray -BatchId $batchId -BatchSize $batchSize
                    }
                    Publish-ConsolidationBatchFiles `
                        -Sources $publishSources.ToArray() `
                        -TemporaryResultPath $temporaryPath `
                        -OutputPath $outputPath `
                        -CommitState $commitState

                    $stateEntries = @(Read-StateEntries)
                    $completed++
                    $completedShipmentBatches++
                    $completedShipmentFiles += $items.Count
                    $completedShipmentSheets += $result.Sheets
                    Write-Host "[LOTE CONCLUÍDO $batchPosition/$($groupKeys.Count)] $([IO.Path]::GetFileName($outputPath)) | $($items.Count) arquivo(s), $($result.Sheets) aba(s), $($result.Documents) documento(s)." -ForegroundColor Green
                    Write-OperationalLog "LOTE CONCLUÍDO" "$key | $($items.Count) arquivo(s) | $($result.Sheets) aba(s) | $([IO.Path]::GetFileName($outputPath)) | $(Get-ElapsedText $generationStartedAt)."
                }
                catch {
                    $rejected += $items.Count
                    $rejectedShipmentBatches++
                    $files = @($items | ForEach-Object { $_.InputFile })
                    $report = Write-ShipmentBatchErrorReport -BatchKey $key -InputFiles $files -Exception $_.Exception
                    Write-Host "[LOTE REJEITADO $batchPosition/$($groupKeys.Count)] ${key}: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "                                             Relatório: $report" -ForegroundColor DarkYellow
                    Write-OperationalLog "LOTE REJEITADO" "$key | $($items.Count) arquivo(s) | $($_.Exception.Message)"
                }
                finally {
                    if ($null -ne $jobDirectory -and (Test-Path -LiteralPath $jobDirectory)) {
                        Remove-Item -LiteralPath $jobDirectory -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
            }

            foreach ($blockedKey in @($blockedKeys | Where-Object { -not $shipmentGroups.ContainsKey($_) })) {
                $failuresForKey = @($shipmentFailures | Where-Object { $_.Key -eq $blockedKey })
                $files = @($failuresForKey | ForEach-Object { $_.InputFile })
                $report = Write-ShipmentBatchErrorReport -BatchKey $blockedKey -InputFiles $files -Exception $failuresForKey[0].Exception
                $rejectedShipmentBatches++
                Write-Host "[LOTE REJEITADO] ${blockedKey}: $($failuresForKey[0].Exception.Message)" -ForegroundColor Red
                Write-Host "                  Relatório: $report" -ForegroundColor DarkYellow
            }
        }

        $jobPosition = 0
        foreach ($job in $duimpJobs) {
            $jobPosition++
            $inputFile = $job.InputFile
            $process = $job.Process
            $hash = $job.Hash
            $jobDirectory = $null
            Write-Host "[PROCESSANDO DUIMP $jobPosition/$($duimpJobs.Count)] $($inputFile.Name)..." -ForegroundColor Cyan
            try {
                $readStartedAt = Get-Date
                Write-Phase -StartedAt $runStartedAt -Message "Lendo relação de DUIMP $jobPosition/$($duimpJobs.Count)..."
                Test-FileReady $inputFile
                $currentHash = (Get-FileHash -LiteralPath $inputFile.FullName -Algorithm SHA256).Hash
                if ($currentHash -ne $hash) {
                    throw "O arquivo foi alterado após a pré-análise. Execute novamente para processar a versão atual."
                }

                $source = Read-SourceData -Excel $excel -Path $inputFile.FullName
                Write-Phase -StartedAt $runStartedAt -Message "Relação lida em $(Get-ElapsedText $readStartedAt); consolidando blocos..."
                $metadata = Get-WorkbookMetadata -Rows $source.Rows
                $fileName = Get-OutputFileName -Client $metadata.Client -Reference $metadata.Reference -ProcessType "DUIMP"
                $outputPath = Get-NextOutputPath -OutputDirectory $process.Output -FileName $fileName
                $jobDirectory = Join-Path $temporaryRoot ([guid]::NewGuid().ToString("N"))
                New-Item -ItemType Directory -Path $jobDirectory -Force | Out-Null
                $temporaryPath = Join-Path $jobDirectory "resultado.xlsx"
                $generationStartedAt = Get-Date
                Write-Phase -StartedAt $runStartedAt -Message "Gerando, formatando e validando DUIMP..."
                $result = Write-DuimpWorkbook -Excel $excel -Metadata $metadata -SourcePath $inputFile.FullName -Rows $source.Rows -OriginalUsedRows $source.UsedRows -TemporaryPath $temporaryPath
                Write-Phase -StartedAt $runStartedAt -Message "DUIMP pronta em $(Get-ElapsedText $generationStartedAt); publicando..."
                $detail = "$($result.FinalBlocks) documentos finais"
                if ($result.RepeatedOccurrences -gt 0) {
                    $detail += "; $($result.RepeatedOccurrences) ocorrência(s) repetida(s)"
                }
                if (-not (Test-Path -LiteralPath $temporaryPath -PathType Leaf)) {
                    throw "O arquivo temporário validado não foi criado."
                }
                if (Test-Path -LiteralPath $outputPath) {
                    $outputPath = Get-NextOutputPath -OutputDirectory $process.Output -FileName $fileName
                }
                $archivePath = Get-NextOutputPath -OutputDirectory $process.Processed -FileName $inputFile.Name
                $inputName = $inputFile.Name
                $commitState = {
                    Add-StateEntry `
                        -Hash $hash `
                        -ProcessType "DUIMP" `
                        -InputName $inputName `
                        -OutputPath $outputPath `
                        -ArchivedPath $archivePath
                }
                Publish-ConsolidationFiles `
                    -SourcePath $inputFile.FullName `
                    -TemporaryResultPath $temporaryPath `
                    -OutputPath $outputPath `
                    -ArchivePath $archivePath `
                    -CommitState $commitState

                $stateEntries = @(Read-StateEntries)
                $completed++
                $completedDuimpFiles++
                Write-Host "[DUIMP CONCLUÍDA $jobPosition/$($duimpJobs.Count)] $($inputFile.Name) -> $([IO.Path]::GetFileName($outputPath)) ($detail)" -ForegroundColor Green
                Write-OperationalLog "DUIMP CONCLUÍDA" "$($inputFile.Name) | $([IO.Path]::GetFileName($outputPath)) | $detail | $(Get-ElapsedText $generationStartedAt)."
            }
            catch {
                $rejected++
                $report = Write-ErrorReport -InputFile $inputFile -ProcessType "DUIMP" -Exception $_.Exception
                Write-Host "[DUIMP REJEITADA $jobPosition/$($duimpJobs.Count)] $($inputFile.Name): $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "                                           Relatório: $report" -ForegroundColor DarkYellow
                Write-OperationalLog "DUIMP REJEITADA" "$($inputFile.Name) | $($_.Exception.Message)"
            }
            finally {
                if ($null -ne $jobDirectory -and (Test-Path -LiteralPath $jobDirectory)) {
                    Remove-Item -LiteralPath $jobDirectory -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }
}
catch {
    $rejected++
    Write-Host "[ERRO GERAL] $($_.Exception.Message)" -ForegroundColor Red
    Write-OperationalLog "ERRO GERAL" $_.Exception.Message
}
finally {
    if ($null -ne $excel) {
        if ($null -ne $excelSettings) {
            try { $excel.ScreenUpdating = $excelSettings.ScreenUpdating } catch {}
            try { $excel.EnableEvents = $excelSettings.EnableEvents } catch {}
            try { $excel.DisplayAlerts = $excelSettings.DisplayAlerts } catch {}
            try { $excel.AskToUpdateLinks = $excelSettings.AskToUpdateLinks } catch {}
            try { $excel.Calculation = $excelSettings.Calculation } catch {}
            try { $excel.CalculateBeforeSave = $excelSettings.CalculateBeforeSave } catch {}
        }
        try { [void] $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    if ($null -ne $lockStream) {
        $lockStream.Dispose()
        Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Resumo final:" -ForegroundColor Cyan
Write-Host "  Lotes de Embarques concluídos: $completedShipmentBatches"
Write-Host "  Arquivos/abas de Embarques concluídos: $completedShipmentFiles/$completedShipmentSheets"
Write-Host "  Arquivos de DUIMP concluídos: $completedDuimpFiles"
Write-Host "  Já processados arquivados: $archivedPrevious"
Write-Host "  Lotes de Embarques rejeitados: $rejectedShipmentBatches"
Write-Host "  Arquivos rejeitados: $rejected"
Write-Host "  Tempo total: $(Get-ElapsedText $runStartedAt)"
foreach ($warning in $brandingWarnings) {
    Write-Host "  AVISO: $warning" -ForegroundColor Yellow
    Write-OperationalLog "AVISO" $warning
}
if (-not $SemInteracao) {
    Write-Host "Pressione ENTER para fechar."
    [void] (Read-Host)
}

if ($rejected -gt 0) {
    exit 1
}
exit 0
