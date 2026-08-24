Set-StrictMode -Version Latest

function ConvertTo-SafeFilePart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string] $Value
    )

    $normalized = $Value.Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object Text.StringBuilder

    foreach ($character in $normalized.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($character)
        if ($category -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void] $builder.Append($character)
        }
    }

    $safe = $builder.ToString().Normalize([Text.NormalizationForm]::FormC).ToUpperInvariant()
    $safe = $safe -replace "[^A-Z0-9]+", "_"
    $safe = $safe.Trim("_")

    if ([string]::IsNullOrWhiteSpace($safe)) {
        throw "Não foi possível formar um nome de arquivo válido."
    }

    return $safe
}

function ConvertTo-NormalizedKey {
    param([AllowNull()] $Value)

    if ($null -eq $Value) {
        return ""
    }

    $text = ([string] $Value).Trim().Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object Text.StringBuilder
    foreach ($character in $text.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($character)
        if ($category -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void] $builder.Append($character)
        }
    }
    return ($builder.ToString().ToLowerInvariant() -replace "[^a-z0-9]", "")
}

function Copy-Rows {
    param([Parameter(Mandatory = $true)] [object[]] $Rows)

    $copy = New-Object "System.Collections.Generic.List[object[]]"
    foreach ($row in $Rows) {
        $copy.Add(@($row))
    }
    return $copy.ToArray()
}

function ConvertTo-DecimalValue {
    param([AllowNull()] $Value)

    if ($null -eq $Value) {
        return $null
    }
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or
        $Value -is [int64] -or $Value -is [single] -or $Value -is [double] -or
        $Value -is [decimal]) {
        return [decimal] $Value
    }

    $raw = ([string] $Value).Trim() -replace "\*", "" -replace "\s", ""
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $null
    }

    $styles = [Globalization.NumberStyles]::AllowLeadingSign -bor
        [Globalization.NumberStyles]::AllowDecimalPoint -bor
        [Globalization.NumberStyles]::AllowThousands
    [decimal] $parsed = 0
    if ($raw.Contains(",")) {
        $culture = [Globalization.CultureInfo]::GetCultureInfo("pt-BR")
        if ([decimal]::TryParse($raw, $styles, $culture, [ref] $parsed)) {
            return $parsed
        }
    }
    else {
        $culture = [Globalization.CultureInfo]::InvariantCulture
        if ([decimal]::TryParse($raw, $styles, $culture, [ref] $parsed)) {
            return $parsed
        }
    }
    return $null
}

function Get-DecimalPlaces {
    param([AllowNull()] $Value)

    if ($Value -is [string]) {
        $text = $Value.Trim()
        if ($text -match ",(\d+)$") {
            return $Matches[1].Length
        }
        if ($text -match "\.(\d+)$" -and $text -notmatch "\.\d{3}(?:\.|$)") {
            return $Matches[1].Length
        }
        return 0
    }
    return 2
}

function Format-DecimalLike {
    param(
        [Parameter(Mandatory = $true)] [decimal] $Value,
        [Parameter(Mandatory = $true)] $Example
    )

    if ($Example -is [string]) {
        $places = Get-DecimalPlaces $Example
        $culture = [Globalization.CultureInfo]::GetCultureInfo("pt-BR")
        return $Value.ToString("N$places", $culture)
    }
    return [double] $Value
}

function Get-DuimpParts {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [string] $Document)

    $pattern = "^DUIMP/Adi(?:ção|cao)/Item:\s*(\d{2}/BR\d+-\d)/(\d{4})/(\d{5})$"
    $match = [regex]::Match($Document.Trim(), $pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {
        return $null
    }
    return [pscustomobject]@{
        Base = $match.Groups[1].Value
        Addition = $match.Groups[2].Value
        Item = $match.Groups[3].Value
    }
}

function Get-OutputFileName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $Client,
        [Parameter(Mandatory = $true)] [string] $Reference,
        [Parameter(Mandatory = $true)]
        [ValidateSet("EMBARQUES", "DUIMP")]
        [string] $ProcessType
    )

    if ($Reference -notmatch "^\s*(\d{1,2})\s*/\s*(\d{4})\s*$") {
        throw "Competência inválida: $Reference"
    }
    $month = ([int] $Matches[1]).ToString("00")
    $year = $Matches[2]
    $clientPart = ConvertTo-SafeFilePart $Client

    if ($ProcessType -eq "EMBARQUES") {
        return "${clientPart}_${year}_${month}_EMBARQUES_CONSOLIDADOS.xlsx"
    }
    return "${clientPart}_${year}_${month}_DUIMP_CONSOLIDADA.xlsx"
}

function Get-ShipmentBatchKey {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] $Metadata)

    $policy = ([string] $Metadata.Policy).Trim()
    $reference = ([string] $Metadata.Reference).Trim()
    if ($policy -notmatch "^\d{13}$") {
        throw "Apólice inválida para o lote de embarques: $policy"
    }
    if ($reference -notmatch "^\s*(\d{1,2})\s*/\s*(\d{4})\s*$") {
        throw "Competência inválida para o lote de embarques: $reference"
    }
    $month = [int] $Matches[1]
    $year = [int] $Matches[2]
    if ($month -lt 1 -or $month -gt 12) {
        throw "Competência inválida para o lote de embarques: $reference"
    }
    return "{0}|{1:0000}-{2:00}" -f $policy, $year, $month
}

function Read-ShipmentGroupAliases {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [string] $Path)

    $aliases = @{}
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $aliases
    }

    foreach ($row in @(Import-Csv -LiteralPath $Path -Delimiter ";" -Encoding UTF8)) {
        $policy = ([string] $row.Apolice).Trim()
        $name = ([string] $row.NomeLote).Trim()
        if ([string]::IsNullOrWhiteSpace($policy) -and [string]::IsNullOrWhiteSpace($name)) {
            continue
        }
        if ($policy -notmatch "^\d{13}$") {
            throw "Apólice inválida em GRUPOS_EMBARQUES.csv: $policy"
        }
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw "NomeLote não informado para a apólice $policy."
        }
        if ($aliases.ContainsKey($policy)) {
            throw "A apólice $policy está duplicada em GRUPOS_EMBARQUES.csv."
        }
        $aliases[$policy] = $name
    }
    return $aliases
}

function Get-ShipmentBatchOutputFileName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $Policy,
        [Parameter(Mandatory = $true)] [string] $Reference,
        [Parameter(Mandatory = $true)] [hashtable] $Aliases
    )

    $metadata = [pscustomobject]@{ Policy = $Policy; Reference = $Reference }
    $key = Get-ShipmentBatchKey -Metadata $metadata
    $parts = $key.Split("|")
    $period = $parts[1].Split("-")
    $year = $period[0]
    $month = [int] $period[1]
    if ($Aliases.ContainsKey($Policy)) {
        $months = @(
            "", "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
            "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
        )
        $lotName = ([string] $Aliases[$Policy]).Trim()
        return "{0} - PRÉVIA - COMPETÊNCIA {1} - {2}.xlsx" -f $lotName, $months[$month], $year
    }
    return "APOLICE_{0}_{1}_{2:00}_EMBARQUES_CONSOLIDADOS.xlsx" -f $Policy, $year, $month
}

function Get-UniqueShipmentWorksheetName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $Client,
        [Parameter(Mandatory = $true)] [string] $Subgroup,
        [Parameter(Mandatory = $true)] $UsedNames
    )

    $base = ConvertTo-SafeFilePart $Client
    if ([string]::IsNullOrWhiteSpace($base)) {
        $base = "EMBARQUES"
    }
    if ($base.Length -gt 31) {
        $base = $base.Substring(0, 31)
    }
    if ($UsedNames.Add($base)) {
        return $base
    }

    $subgroupPart = ConvertTo-SafeFilePart $Subgroup
    if ([string]::IsNullOrWhiteSpace($subgroupPart)) {
        $subgroupPart = "SG"
    }
    $suffix = "_$subgroupPart"
    $prefixLength = [Math]::Max(1, 31 - $suffix.Length)
    $candidate = $base.Substring(0, [Math]::Min($base.Length, $prefixLength)) + $suffix
    if ($UsedNames.Add($candidate)) {
        return $candidate
    }

    $version = 2
    do {
        $numberSuffix = "_${subgroupPart}_$version"
        $prefixLength = [Math]::Max(1, 31 - $numberSuffix.Length)
        $candidate = $base.Substring(0, [Math]::Min($base.Length, $prefixLength)) + $numberSuffix
        $version++
    } while (-not $UsedNames.Add($candidate))
    return $candidate
}

function Get-ShipmentBatchKeyFromFileName {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [string] $FileName)

    $policyMatch = [regex]::Match($FileName, "(?<!\d)(\d{13})(?!\d)")
    if (-not $policyMatch.Success) {
        return $null
    }
    $policy = $policyMatch.Groups[1].Value
    $tail = $FileName.Substring($policyMatch.Index + $policyMatch.Length)

    $monthYear = [regex]::Match($tail, "(?:^|[_-])(0[1-9]|1[0-2])[_-](20\d{2})(?:\D|$)")
    if ($monthYear.Success) {
        return "{0}|{1}-{2}" -f $policy, $monthYear.Groups[2].Value, $monthYear.Groups[1].Value
    }
    $yearMonth = [regex]::Match($tail, "(?:^|[_-])(20\d{2})[_-](0[1-9]|1[0-2])(?:\D|$)")
    if ($yearMonth.Success) {
        return "{0}|{1}-{2}" -f $policy, $yearMonth.Groups[1].Value, $yearMonth.Groups[2].Value
    }
    return $null
}

function Get-LabelValue {
    param(
        [Parameter(Mandatory = $true)] [object[]] $Rows,
        [Parameter(Mandatory = $true)] [string] $Label,
        [int] $MaxRows = 20
    )

    $wanted = ConvertTo-NormalizedKey $Label
    $limit = [Math]::Min($Rows.Count, $MaxRows)
    for ($rowIndex = 0; $rowIndex -lt $limit; $rowIndex++) {
        $row = @($Rows[$rowIndex])
        for ($column = 0; $column -lt $row.Count; $column++) {
            if ((ConvertTo-NormalizedKey $row[$column]) -ne $wanted) {
                continue
            }
            for ($next = $column + 1; $next -lt $row.Count; $next++) {
                $value = ([string] $row[$next]).Trim()
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    return $value
                }
            }
        }
    }
    return ""
}

function Get-WorkbookMetadata {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [object[]] $Rows)

    $policy = Get-LabelValue -Rows $Rows -Label "Apolice"
    $subgroup = Get-LabelValue -Rows $Rows -Label "SubGrupo"
    $reference = (Get-LabelValue -Rows $Rows -Label "Referencia") -replace "\s", ""
    $client = Get-LabelValue -Rows $Rows -Label "Segurado"

    if ($subgroup -match "^\d+$") {
        $subgroup = $subgroup.PadLeft(3, "0")
    }
    if ([string]::IsNullOrWhiteSpace($policy) -or
        [string]::IsNullOrWhiteSpace($subgroup) -or
        [string]::IsNullOrWhiteSpace($reference) -or
        [string]::IsNullOrWhiteSpace($client)) {
        throw "Metadados obrigatórios não encontrados na relação."
    }

    return [pscustomobject]@{
        Policy = $policy
        Subgroup = $subgroup
        Reference = $reference
        Client = $client
    }
}

function New-DocumentBlock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [object[]] $Rows,
        [Parameter(Mandatory = $true)] [int] $StartIndex
    )

    $copied = Copy-Rows $Rows
    $document = ""
    foreach ($row in $copied) {
        if ((ConvertTo-NormalizedKey $row[0]).StartsWith("observa")) {
            $document = ([string] $row[1]).Trim()
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace($document)) {
        throw "Documento não encontrado no bloco iniciado na linha $($StartIndex + 1)."
    }

    $duimp = Get-DuimpParts $document
    return [pscustomobject]@{
        StartIndex = $StartIndex
        Rows = $copied
        Document = $document
        Type = if ($null -eq $duimp) { "DI" } else { "DUIMP" }
        Base = if ($null -eq $duimp) { $document } else { $duimp.Base }
        Addition = if ($null -eq $duimp) { $null } else { $duimp.Addition }
        Item = if ($null -eq $duimp) { $null } else { $duimp.Item }
    }
}

function Get-DocumentBlocks {
    param([Parameter(Mandatory = $true)] [object[]] $Rows)

    $starts = New-Object "System.Collections.Generic.List[int]"
    for ($index = 0; $index -lt $Rows.Count; $index++) {
        if ((ConvertTo-NormalizedKey $Rows[$index][0]) -eq "definitiva") {
            $starts.Add($index)
        }
    }
    if ($starts.Count -eq 0) {
        throw "Nenhum bloco 'Definitiva' foi encontrado."
    }

    $blocks = New-Object "System.Collections.Generic.List[object]"
    for ($index = 0; $index -lt $starts.Count; $index++) {
        $start = $starts[$index]
        $end = if ($index + 1 -lt $starts.Count) { $starts[$index + 1] } else { $Rows.Count }
        $length = $end - $start
        if ($length -lt 17) {
            throw "Bloco incompleto iniciado na linha $($start + 1)."
        }
        $slice = @($Rows[$start..($end - 1)])
        $blocks.Add((New-DocumentBlock -Rows $slice -StartIndex $start))
    }
    return $blocks.ToArray()
}

function Get-BlockRow {
    param(
        [Parameter(Mandatory = $true)] $Block,
        [Parameter(Mandatory = $true)] [string] $Label
    )

    $wanted = ConvertTo-NormalizedKey $Label
    foreach ($row in $Block.Rows) {
        if ((ConvertTo-NormalizedKey $row[0]) -eq $wanted) {
            return $row
        }
    }
    return $null
}

function Get-BlockLabelPositions {
    param(
        [Parameter(Mandatory = $true)] $Block,
        [Parameter(Mandatory = $true)] [string] $Label
    )

    $wanted = ConvertTo-NormalizedKey $Label
    $positions = New-Object "System.Collections.Generic.List[object]"
    for ($rowIndex = 0; $rowIndex -lt $Block.Rows.Count; $rowIndex++) {
        $row = @($Block.Rows[$rowIndex])
        for ($column = 0; $column -lt $row.Count; $column++) {
            if ((ConvertTo-NormalizedKey $row[$column]) -eq $wanted) {
                $positions.Add([pscustomobject]@{ Row = $rowIndex; Column = $column })
            }
        }
    }
    return $positions.ToArray()
}

function Get-DuimpFinancialSlots {
    param(
        [Parameter(Mandatory = $true)] $Block,
        [Parameter(Mandatory = $true)] [string] $Label
    )

    $wanted = ConvertTo-NormalizedKey $Label
    $cacheProperty = $Block.PSObject.Properties["DuimpFinancialSlotMap"]
    if ($null -eq $cacheProperty) {
        $financialKeys = @(
            (ConvertTo-NormalizedKey "FOB"),
            (ConvertTo-NormalizedKey "Frete"),
            (ConvertTo-NormalizedKey "Despesas"),
            (ConvertTo-NormalizedKey "Lucro Esp"),
            (ConvertTo-NormalizedKey "Imposto"),
            (ConvertTo-NormalizedKey "Seguro"),
            (ConvertTo-NormalizedKey "Total")
        )
        $positionsByKey = @{}
        for ($rowIndex = 0; $rowIndex -lt $Block.Rows.Count; $rowIndex++) {
            $row = @($Block.Rows[$rowIndex])
            for ($column = 0; $column -lt $row.Count; $column++) {
                $key = ConvertTo-NormalizedKey $row[$column]
                if ($key -notin $financialKeys) {
                    continue
                }
                if (-not $positionsByKey.ContainsKey($key)) {
                    $positionsByKey[$key] = New-Object "System.Collections.Generic.List[object]"
                }
                $positionsByKey[$key].Add([pscustomobject]@{ Row = $rowIndex; Column = $column })
            }
        }

        $slotMap = @{}
        foreach ($key in $positionsByKey.Keys) {
            $positions = @($positionsByKey[$key].ToArray())
            $rowIndex = $positions[0].Row
            $row = @($Block.Rows[$rowIndex])
            $slots = New-Object "System.Collections.Generic.List[object]"
            $usedColumns = @{}
            for ($index = 0; $index -lt $positions.Count; $index++) {
                $valueColumn = $positions[$index].Column + 1
                if ($valueColumn -ge $row.Count -or $usedColumns.ContainsKey($valueColumn)) {
                    continue
                }
                $slots.Add([pscustomobject]@{
                    Key = "${key}:VALUE:$index"
                    Row = $rowIndex
                    Column = $valueColumn
                })
                $usedColumns[$valueColumn] = $true
            }
            foreach ($column in @(9, 10, 11)) {
                if ($column -ge $row.Count -or $usedColumns.ContainsKey($column)) {
                    continue
                }
                $slots.Add([pscustomobject]@{
                    Key = "${key}:EXTRA:$column"
                    Row = $rowIndex
                    Column = $column
                })
                $usedColumns[$column] = $true
            }
            $slotMap[$key] = $slots.ToArray()
        }
        $Block | Add-Member -MemberType NoteProperty -Name "DuimpFinancialSlotMap" -Value $slotMap
        $cacheProperty = $Block.PSObject.Properties["DuimpFinancialSlotMap"]
    }

    if (-not $cacheProperty.Value.ContainsKey($wanted)) {
        return @()
    }
    return @($cacheProperty.Value[$wanted])
}

function Get-DuimpConditionSlot {
    param([Parameter(Mandatory = $true)] $Block)

    $positions = @(Get-BlockLabelPositions -Block $Block -Label "Cond./Franquia")
    if ($positions.Count -gt 0) {
        $position = $positions[0]
        if ($position.Column + 1 -lt $Block.Rows[$position.Row].Count) {
            return [pscustomobject]@{ Row = $position.Row; Column = $position.Column + 1 }
        }
    }

    # Compatibilidade com o layout legado e com as matrizes de teste antigas,
    # nas quais o valor existe em I9, mas o rotulo nao e repetido na linha.
    if ($Block.Rows.Count -gt 8 -and @($Block.Rows[8]).Count -gt 7) {
        return [pscustomobject]@{ Row = 8; Column = 7 }
    }
    return $null
}

function Get-DuimpBlockLayout {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [object[]] $Blocks)

    if ($Blocks.Count -eq 0) {
        throw "A relação não contém blocos para identificar o padrão DUIMP."
    }

    $headerRows = [int] $Blocks[0].StartIndex
    if ($Blocks.Count -gt 1) {
        $blockSize = [int] $Blocks[1].StartIndex - [int] $Blocks[0].StartIndex
    }
    else {
        $hasIncoterm = @(Get-BlockLabelPositions -Block $Blocks[0] -Label "Incoterm").Count -gt 0
        $blockSize = if ($hasIncoterm) { 19 } else { 18 }
    }

    if ($blockSize -notin @(18, 19)) {
        throw "O padrão DUIMP possui $blockSize linhas por bloco; são aceitos os padrões de 18 e 19 linhas."
    }

    for ($index = 0; $index -lt $Blocks.Count; $index++) {
        $expectedStart = $headerRows + ($index * $blockSize)
        if ([int] $Blocks[$index].StartIndex -ne $expectedStart) {
            throw "A sequência de blocos não segue um padrão uniforme de $blockSize linhas."
        }
        if ($Blocks[$index].Rows.Count -lt $blockSize) {
            throw "O bloco iniciado na linha $([int] $Blocks[$index].StartIndex + 1) possui menos de $blockSize linhas."
        }
    }

    return [pscustomobject]@{
        HeaderRows = $headerRows
        BlockSize = $blockSize
    }
}

function ConvertTo-ShipmentRecords {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [object[]] $Rows)

    $records = New-Object "System.Collections.Generic.List[object]"
    $seen = @{}
    foreach ($block in @(Get-DocumentBlocks -Rows $Rows)) {
        if ($seen.ContainsKey($block.Document)) {
            throw "Documento duplicado encontrado: $($block.Document)"
        }
        $seen[$block.Document] = $true

        $fob = Get-BlockRow -Block $block -Label "FOB"
        $freight = Get-BlockRow -Block $block -Label "Frete"
        $taxes = Get-BlockRow -Block $block -Label "Imposto"
        if ($null -eq $fob -or $null -eq $freight -or $null -eq $taxes) {
            throw "Bloco incompleto para o documento $($block.Document)."
        }

        $exchange = ConvertTo-DecimalValue $block.Rows[1][11]
        if ($null -eq $exchange) {
            $exchange = ConvertTo-DecimalValue $fob[3]
        }
        $displayedRate = ConvertTo-DecimalValue $fob[7]
        $fobBrl = ConvertTo-DecimalValue $fob[5]
        if ($null -eq $fobBrl) { $fobBrl = ConvertTo-DecimalValue $fob[6] }
        $freightBrl = ConvertTo-DecimalValue $freight[5]
        if ($null -eq $freightBrl) { $freightBrl = ConvertTo-DecimalValue $freight[6] }
        $taxesBrl = ConvertTo-DecimalValue $taxes[5]
        if ($null -eq $taxesBrl) { $taxesBrl = ConvertTo-DecimalValue $taxes[6] }
        if ($null -eq $taxesBrl) { $taxesBrl = ConvertTo-DecimalValue $taxes[1] }

        if ($null -eq $exchange -or $exchange -eq 0 -or $null -eq $displayedRate -or
            $null -eq $fobBrl -or $null -eq $freightBrl -or $null -eq $taxesBrl) {
            throw "Valores obrigatórios ausentes para o documento $($block.Document)."
        }

        $records.Add([pscustomobject]@{
            Document = $block.Document
            RiskRate = ([decimal] $displayedRate) / ([decimal] 100)
            Exchange = [decimal] $exchange
            FobUsd = ([decimal] $fobBrl) / ([decimal] $exchange)
            FreightUsd = ([decimal] $freightBrl) / ([decimal] $exchange)
            Currency = ([string] $fob[2]).Trim()
            TaxesBrl = [decimal] $taxesBrl
        })
    }
    return $records.ToArray()
}

function Get-UniqueNonBlank {
    param(
        [Parameter(Mandatory = $true)] [object[]] $Blocks,
        [Parameter(Mandatory = $true)] [int] $Row,
        [Parameter(Mandatory = $true)] [int] $Column
    )

    return @($Blocks |
        ForEach-Object { ([string] $_.Rows[$Row][$Column]).Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique)
}

function Merge-DuimpGroup {
    param(
        [Parameter(Mandatory = $true)] [string] $Base,
        [Parameter(Mandatory = $true)] [object[]] $Blocks
    )

    $rows = Copy-Rows $Blocks[0].Rows
    $observation = @(Get-BlockLabelPositions -Block $Blocks[0] -Label "Observacoes")
    if ($observation.Count -eq 0 -or $observation[0].Column + 1 -ge $rows[$observation[0].Row].Count) {
        throw "A linha de observações não foi encontrada para a DUIMP $Base."
    }
    $rows[$observation[0].Row][$observation[0].Column + 1] = $Base

    if ((@(Get-UniqueNonBlank -Blocks $Blocks -Row 5 -Column 1)).Count -gt 1) {
        $rows[5][1] = "DIVERSAS MERCADORIAS"
    }
    if ((@(Get-UniqueNonBlank -Blocks $Blocks -Row 5 -Column 7)).Count -gt 1) {
        $rows[5][7] = "DIVERSAS"
    }
    $conditionValues = New-Object "System.Collections.Generic.List[string]"
    foreach ($block in $Blocks) {
        $condition = Get-DuimpConditionSlot -Block $block
        if ($null -eq $condition) {
            continue
        }
        $value = ([string] $block.Rows[$condition.Row][$condition.Column]).Trim()
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $conditionValues.Add($value)
        }
    }
    if (@($conditionValues | Select-Object -Unique).Count -gt 1) {
        $targetCondition = Get-DuimpConditionSlot -Block $Blocks[0]
        if ($null -ne $targetCondition) {
            $rows[$targetCondition.Row][$targetCondition.Column] = "CONFORME ITENS"
        }
    }

    foreach ($label in @("FOB", "Frete", "Despesas", "Lucro Esp", "Imposto", "Seguro", "Total")) {
        $targetSlots = @(Get-DuimpFinancialSlots -Block $Blocks[0] -Label $label)
        foreach ($targetSlot in $targetSlots) {
            $values = New-Object "System.Collections.Generic.List[object]"
            foreach ($block in $Blocks) {
                $sourceSlot = @(Get-DuimpFinancialSlots -Block $block -Label $label |
                    Where-Object { $_.Key -eq $targetSlot.Key })
                if ($sourceSlot.Count -eq 0) {
                    continue
                }
                $rawValue = $block.Rows[$sourceSlot[0].Row][$sourceSlot[0].Column]
                if ($null -ne (ConvertTo-DecimalValue $rawValue)) {
                    $values.Add($rawValue)
                }
            }
            if ($values.Count -eq 0) {
                continue
            }
            [decimal] $sum = 0
            foreach ($value in $values) {
                $sum += ConvertTo-DecimalValue $value
            }
            $rows[$targetSlot.Row][$targetSlot.Column] = Format-DecimalLike -Value $sum -Example $values[0]
        }
    }

    $result = New-DocumentBlock -Rows $rows -StartIndex $Blocks[0].StartIndex
    $result.Document = $Base
    $result.Type = "DUIMP"
    $result.Base = $Base
    $result.Addition = $null
    $result.Item = $null
    return $result
}

function Merge-DuimpDocumentBlocks {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [object[]] $Blocks)

    $groups = @{}
    foreach ($block in $Blocks) {
        if ($block.Type -eq "DUIMP") {
            if (-not $groups.ContainsKey($block.Base)) {
                $groups[$block.Base] = New-Object "System.Collections.Generic.List[object]"
            }
            $groups[$block.Base].Add($block)
        }
    }

    $emitted = @{}
    $results = New-Object "System.Collections.Generic.List[object]"
    foreach ($block in $Blocks) {
        if ($block.Type -ne "DUIMP") {
            $results.Add($block)
            continue
        }
        if ($emitted.ContainsKey($block.Base)) {
            continue
        }
        $results.Add((Merge-DuimpGroup -Base $block.Base -Blocks $groups[$block.Base].ToArray()))
        $emitted[$block.Base] = $true
    }
    return $results.ToArray()
}

function Test-DuimpResultDocuments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]] $Documents
    )

    if (@($Documents | Where-Object { $_ -match "/\d{4}/\d{5}$" }).Count -gt 0) {
        throw "Ainda existem itens de DUIMP fracionados no resultado."
    }

    $uniqueCount = @($Documents | Select-Object -Unique).Count
    return [pscustomobject]@{
        Documents = $Documents.Count
        UniqueDocuments = $uniqueCount
        RepeatedOccurrences = $Documents.Count - $uniqueCount
    }
}

function Get-NextOutputPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $OutputDirectory,
        [Parameter(Mandatory = $true)] [string] $FileName
    )

    $candidate = Join-Path $OutputDirectory $FileName
    if (-not (Test-Path -LiteralPath $candidate)) {
        return $candidate
    }

    $stem = [IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [IO.Path]::GetExtension($FileName)
    $version = 2
    do {
        $candidate = Join-Path $OutputDirectory ("{0}_V{1}{2}" -f $stem, $version, $extension)
        $version++
    } while (Test-Path -LiteralPath $candidate)

    return $candidate
}

function Publish-ConsolidationFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $SourcePath,
        [Parameter(Mandatory = $true)] [string] $TemporaryResultPath,
        [Parameter(Mandatory = $true)] [string] $OutputPath,
        [Parameter(Mandatory = $true)] [string] $ArchivePath,
        [Parameter(Mandatory = $true)] [scriptblock] $CommitState
    )

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        throw "O arquivo original não está mais no INPUT."
    }
    if (-not (Test-Path -LiteralPath $TemporaryResultPath -PathType Leaf)) {
        throw "O arquivo temporário validado não foi encontrado."
    }
    if (Test-Path -LiteralPath $OutputPath) {
        throw "O caminho reservado no OUTPUT já existe."
    }
    if (Test-Path -LiteralPath $ArchivePath) {
        throw "O caminho reservado em PROCESSADOS já existe."
    }

    $sourceArchived = $false
    try {
        Move-Item -LiteralPath $SourcePath -Destination $ArchivePath
        $sourceArchived = $true
        Copy-Item -LiteralPath $TemporaryResultPath -Destination $OutputPath
        & $CommitState
    }
    catch {
        $originalError = $_
        if (Test-Path -LiteralPath $OutputPath) {
            Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue
        }
        if ($sourceArchived -and (Test-Path -LiteralPath $ArchivePath)) {
            if (Test-Path -LiteralPath $SourcePath) {
                throw "Não foi possível desfazer a publicação porque o nome original voltou a existir no INPUT."
            }
            Move-Item -LiteralPath $ArchivePath -Destination $SourcePath
        }
        throw $originalError
    }
}

function Publish-ConsolidationBatchFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [object[]] $Sources,
        [Parameter(Mandatory = $true)] [string] $TemporaryResultPath,
        [Parameter(Mandatory = $true)] [string] $OutputPath,
        [Parameter(Mandatory = $true)] [scriptblock] $CommitState
    )

    if ($Sources.Count -eq 0) {
        throw "O lote não contém arquivos de origem."
    }
    if (-not (Test-Path -LiteralPath $TemporaryResultPath -PathType Leaf)) {
        throw "O arquivo temporário validado não foi encontrado."
    }
    if (Test-Path -LiteralPath $OutputPath) {
        throw "O caminho reservado no OUTPUT já existe."
    }

    $seenSources = @{}
    $seenArchives = @{}
    foreach ($source in $Sources) {
        $sourcePath = [IO.Path]::GetFullPath([string] $source.SourcePath)
        $archivePath = [IO.Path]::GetFullPath([string] $source.ArchivePath)
        if ($seenSources.ContainsKey($sourcePath)) {
            throw "Um arquivo de origem foi informado mais de uma vez no lote."
        }
        if ($seenArchives.ContainsKey($archivePath)) {
            throw "Um caminho de arquivamento foi reservado mais de uma vez no lote."
        }
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            throw "Um arquivo original não está mais no INPUT: $([IO.Path]::GetFileName($sourcePath))"
        }
        if (Test-Path -LiteralPath $archivePath) {
            throw "Um caminho reservado em PROCESSADOS já existe: $([IO.Path]::GetFileName($archivePath))"
        }
        $seenSources[$sourcePath] = $true
        $seenArchives[$archivePath] = $true
    }

    $archived = New-Object "System.Collections.Generic.List[object]"
    try {
        foreach ($source in $Sources) {
            Move-Item -LiteralPath $source.SourcePath -Destination $source.ArchivePath
            $archived.Add($source)
        }
        Copy-Item -LiteralPath $TemporaryResultPath -Destination $OutputPath
        & $CommitState
    }
    catch {
        $originalError = $_
        if (Test-Path -LiteralPath $OutputPath) {
            Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue
        }
        for ($index = $archived.Count - 1; $index -ge 0; $index--) {
            $source = $archived[$index]
            if (-not (Test-Path -LiteralPath $source.ArchivePath)) {
                continue
            }
            if (Test-Path -LiteralPath $source.SourcePath) {
                throw "Não foi possível desfazer o lote porque o nome original voltou a existir no INPUT: $([IO.Path]::GetFileName($source.SourcePath))"
            }
            Move-Item -LiteralPath $source.ArchivePath -Destination $source.SourcePath
        }
        throw $originalError
    }
}

function Test-PreviousSuccess {
    [CmdletBinding()]
    param(
        [AllowNull()] [object[]] $Entries,
        [Parameter(Mandatory = $true)] [string] $Hash,
        [Parameter(Mandatory = $true)] [string] $ProcessType,
        [Parameter(Mandatory = $true)] [string] $RootDirectory
    )

    foreach ($entry in @($Entries)) {
        if ($null -eq $entry) {
            continue
        }
        if (([string] $entry.Hash) -ne $Hash -or
            ([string] $entry.ProcessType) -ne $ProcessType) {
            continue
        }
        $outputPath = Join-Path $RootDirectory ([string] $entry.OutputRelativePath)
        if (Test-Path -LiteralPath $outputPath) {
            return $true
        }
    }
    return $false
}

function Get-LegacyProcessedEntry {
    [CmdletBinding()]
    param(
        [AllowNull()] [object[]] $Entries,
        [Parameter(Mandatory = $true)] [string] $Hash,
        [Parameter(Mandatory = $true)] [string] $ProcessType,
        [Parameter(Mandatory = $true)] [string] $InputName,
        [Parameter(Mandatory = $true)] [string] $RootDirectory
    )

    $matches = @($Entries | Where-Object {
        $null -ne $_ -and
        ([string] $_.Hash) -eq $Hash -and
        ([string] $_.ProcessType) -eq $ProcessType -and
        ([string] $_.InputName) -eq $InputName
    })
    if ($matches.Count -eq 0) {
        return $null
    }

    foreach ($entry in $matches) {
        $archiveProperty = $entry.PSObject.Properties["ArchivedRelativePath"]
        if ($null -ne $archiveProperty -and
            -not [string]::IsNullOrWhiteSpace([string] $archiveProperty.Value)) {
            return $null
        }
    }

    foreach ($entry in $matches) {
        $outputPath = Join-Path $RootDirectory ([string] $entry.OutputRelativePath)
        if (Test-Path -LiteralPath $outputPath -PathType Leaf) {
            return $entry
        }
    }
    return $null
}

function Get-DuimpFinancialTotals {
    param([Parameter(Mandatory = $true)] [object[]] $Blocks)

    $totals = @{}
    foreach ($block in $Blocks) {
        foreach ($label in @("FOB", "Frete", "Despesas", "Lucro Esp", "Imposto", "Seguro", "Total")) {
            foreach ($slot in @(Get-DuimpFinancialSlots -Block $block -Label $label)) {
                $value = ConvertTo-DecimalValue $block.Rows[$slot.Row][$slot.Column]
                if ($null -eq $value) {
                    continue
                }
                $key = $slot.Key
                if (-not $totals.ContainsKey($key)) {
                    [decimal] $totals[$key] = 0
                }
                $totals[$key] = [decimal] $totals[$key] + [decimal] $value
            }
        }
    }
    return $totals
}

function Compare-DuimpFinancialTotals {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [object[]] $OriginalBlocks,
        [Parameter(Mandatory = $true)] [object[]] $ConsolidatedBlocks
    )

    $original = Get-DuimpFinancialTotals -Blocks $OriginalBlocks
    $consolidated = Get-DuimpFinancialTotals -Blocks $ConsolidatedBlocks
    $keys = @($original.Keys + $consolidated.Keys | Select-Object -Unique)
    $differences = New-Object "System.Collections.Generic.List[object]"
    foreach ($key in $keys) {
        [decimal] $before = if ($original.ContainsKey($key)) { $original[$key] } else { 0 }
        [decimal] $after = if ($consolidated.ContainsKey($key)) { $consolidated[$key] } else { 0 }
        if ($before -ne $after) {
            $differences.Add([pscustomobject]@{
                Slot = $key
                Original = $before
                Consolidated = $after
                Difference = $after - $before
            })
        }
    }
    return $differences.ToArray()
}

function Remove-XlsxSheetProtection {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [string] $Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Arquivo XLSX não encontrado: $Path"
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $updates = New-Object "System.Collections.Generic.List[object]"
        foreach ($entry in @($archive.Entries | Where-Object { $_.FullName -match "^xl/worksheets/sheet\d+\.xml$" })) {
            $reader = New-Object IO.StreamReader($entry.Open())
            try {
                $content = $reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
            }

            $pattern = "<sheetProtection\b[^>]*/>|<sheetProtection\b[^>]*>.*?</sheetProtection>"
            $updated = [regex]::Replace(
                $content,
                $pattern,
                "",
                [Text.RegularExpressions.RegexOptions]::Singleline
            )
            if ($updated -ne $content) {
                $updates.Add([pscustomobject]@{
                    Entry = $entry
                    Name = $entry.FullName
                    Content = $updated
                })
            }
        }

        foreach ($update in $updates) {
            $update.Entry.Delete()
            $newEntry = $archive.CreateEntry($update.Name, [IO.Compression.CompressionLevel]::Optimal)
            $writer = New-Object IO.StreamWriter(
                $newEntry.Open(),
                (New-Object Text.UTF8Encoding($false))
            )
            try {
                $writer.Write($update.Content)
            }
            finally {
                $writer.Dispose()
            }
        }
        return $updates.Count
    }
    finally {
        $archive.Dispose()
    }
}

function ConvertTo-TwoDimensionalArray {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [object[]] $Rows,
        [Parameter(Mandatory = $true)] [int] $Columns
    )

    $matrix = New-Object "object[,]" $Rows.Count, $Columns
    for ($row = 0; $row -lt $Rows.Count; $row++) {
        for ($column = 0; $column -lt $Columns; $column++) {
            if ($column -lt $Rows[$row].Count) {
                $matrix[$row, $column] = $Rows[$row][$column]
            }
        }
    }
    return ,$matrix
}

Export-ModuleMember -Function @(
    "ConvertTo-SafeFilePart",
    "Get-DuimpParts",
    "Get-OutputFileName",
    "Get-ShipmentBatchKey",
    "Read-ShipmentGroupAliases",
    "Get-ShipmentBatchOutputFileName",
    "Get-UniqueShipmentWorksheetName",
    "Get-ShipmentBatchKeyFromFileName",
    "Get-WorkbookMetadata",
    "New-DocumentBlock",
    "Get-DocumentBlocks",
    "Get-DuimpBlockLayout",
    "ConvertTo-ShipmentRecords",
    "Merge-DuimpDocumentBlocks",
    "Test-DuimpResultDocuments",
    "Get-NextOutputPath",
    "Publish-ConsolidationFiles",
    "Publish-ConsolidationBatchFiles",
    "Test-PreviousSuccess",
    "Get-LegacyProcessedEntry",
    "Compare-DuimpFinancialTotals",
    "Remove-XlsxSheetProtection",
    "ConvertTo-TwoDimensionalArray"
)
