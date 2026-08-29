Set-StrictMode -Version Latest

# ChangeLog:
# 2026.08.29.1 - Set word/settings.xml UpdateFieldsOnOpen to false in Convert-AsBuiltMarkdownToDocx (via reflection, for Windows PowerShell 5.1 compatibility) so DOCPROPERTY/TOC fields no longer silently recalculate on open.
# 2026.08.27.1 - Added vertical cell-merge (w:vMerge) support to the markdown-to-docx table renderer, driven by an invisible-marker convention (Get-AsBuiltTableMergeMarker), for the StorageGRID MAV Configuration table.
$script:AsBuiltCommonModuleVersion = "2026.08.29.1"

# Invisible-separator marker (U+2063): placed in a table cell to signal that the docx renderer
# should vertically merge that cell with the one above it, instead of repeating the same value.
$script:AsBuiltTableMergeMarker = [string][char]0x2063

function Get-AsBuiltTableMergeMarker {
    [CmdletBinding()]
    param()
    return $script:AsBuiltTableMergeMarker
}

function Initialize-AsBuiltPath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }

    return (Resolve-Path -LiteralPath $Path).ProviderPath
}

function Resolve-AsBuiltPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][string]$BasePath = $PSScriptRoot
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return (Join-Path -Path $BasePath -ChildPath ($Path -replace '^[.]\\', ''))
}

function Resolve-AsBuiltExistingFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][string]$BasePath = $PSScriptRoot
    )

    $candidate = Resolve-AsBuiltPath -Path $Path -BasePath $BasePath
    if ($null -eq $candidate -or -not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        return $null
    }

    return (Resolve-Path -LiteralPath $candidate).ProviderPath
}

function ConvertTo-AsBuiltSafeName {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Value)

    $safe = $Value.ToLowerInvariant()
    $safe = [regex]::Replace($safe, '[^a-z0-9-]+', '_').Trim('_')

    if ([string]::IsNullOrWhiteSpace($safe)) {
        return 'unknown'
    }

    return $safe
}

function Get-AsBuiltPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$Object,
        [Parameter(Mandatory = $false)][string]$PropertyName
    )

    if ($null -eq $Object -or [string]::IsNullOrWhiteSpace($PropertyName)) {
        return $null
    }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($PropertyName)) {
            return $Object[$PropertyName]
        }
        return $null
    }

    $property = $Object.PSObject.Properties[$PropertyName]
    if ($null -ne $property) {
        return $property.Value
    }

    return $null
}

function ConvertTo-AsBuiltArray {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)]$Value)

    if ($null -eq $Value) {
        return ,@()
    }

    if ($Value -is [array]) {
        return ,$Value
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string]) -and -not ($Value -is [hashtable]) -and -not ($Value -is [pscustomobject])) {
        return ,@($Value)
    }

    return ,@($Value)
}

function New-AsBuiltSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('StorageGRID', 'ESeries')][string]$Platform,
        [Parameter(Mandatory = $false)][string]$Target,
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false)][string]$OutputDir,
        [Parameter(Mandatory = $false)][bool]$ValidateCerts = $false,
        [Parameter(Mandatory = $false)][bool]$UseSystemProxy = $false,
        [Parameter(Mandatory = $false)][bool]$CollectionOnly = $false,
        [Parameter(Mandatory = $false)][bool]$KeepIntermediateOutputs = $false,
        [Parameter(Mandatory = $false)][hashtable]$PlatformOptions = @{}
    )

    return [pscustomobject][ordered]@{
        Platform                = $Platform
        Target                  = $Target
        Credential              = $Credential
        OutputDir               = $OutputDir
        ValidateCerts           = $ValidateCerts
        UseSystemProxy          = $UseSystemProxy
        CollectionOnly          = $CollectionOnly
        KeepIntermediateOutputs = $KeepIntermediateOutputs
        PlatformOptions         = $PlatformOptions
    }
}

$script:AsBuiltOpenXmlAssembliesLoaded = $false

function Initialize-AsBuiltOpenXmlAssemblies {
    [CmdletBinding()]
    param()

    if ($script:AsBuiltOpenXmlAssembliesLoaded) {
        return
    }

    $libRoot = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'Lib\OpenXml'
    $editionFolder = if ($PSVersionTable.PSEdition -eq 'Core') { 'Core' } else { 'Desktop' }
    $libDir = Join-Path -Path $libRoot -ChildPath $editionFolder

    if (-not (Test-Path -LiteralPath $libDir -PathType Container)) {
        throw "OpenXML SDK assemblies not found at '$libDir'. Re-extract the distribution package so the Lib\OpenXml folder sits alongside collect_asbuilt.ps1."
    }

    if ($editionFolder -eq 'Desktop') {
        Add-Type -AssemblyName WindowsBase
    }
    else {
        Add-Type -Path (Join-Path -Path $libDir -ChildPath 'System.IO.Packaging.dll')
    }

    Add-Type -Path (Join-Path -Path $libDir -ChildPath 'DocumentFormat.OpenXml.Framework.dll')
    Add-Type -Path (Join-Path -Path $libDir -ChildPath 'DocumentFormat.OpenXml.dll')

    $script:AsBuiltOpenXmlAssembliesLoaded = $true
}

$script:AsBuiltHyperlinkPlaceholderPrefix = 'ASBUILT_HYPERLINK_PLACEHOLDER::'

function ConvertTo-AsBuiltDocxRunsXml {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text)

    if ([string]::IsNullOrEmpty($Text)) {
        return ''
    }

    # Markdown link syntax [text](url) becomes a real hyperlink run; a placeholder relationship id is
    # used here because the actual r:id can only be minted once a MainDocumentPart exists (see
    # Resolve-AsBuiltDocxHyperlinkPlaceholders, called after the body is assembled). Plain URLs with no
    # [text](url) wrapper are left as ordinary text and never become hyperlinks.
    $tokens = [System.Text.RegularExpressions.Regex]::Split($Text, '(\*\*.+?\*\*|<br\s*/?>|\[[^\]]+\]\(https?://[^\s\)]+\))')
    $runsXml = New-Object System.Text.StringBuilder

    foreach ($token in $tokens) {
        if ([string]::IsNullOrEmpty($token)) { continue }

        if ($token -match '(?i)^<br\s*/?>$') {
            [void]$runsXml.Append('<w:r><w:br/></w:r>')
        }
        elseif ($token -match '^\[([^\]]+)\]\((https?://[^\s\)]+)\)$') {
            $linkText = [System.Security.SecurityElement]::Escape($Matches[1])
            $linkUrl = [System.Security.SecurityElement]::Escape($Matches[2])
            [void]$runsXml.Append('<w:hyperlink r:id="' + $script:AsBuiltHyperlinkPlaceholderPrefix + $linkUrl + '"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">' + $linkText + '</w:t></w:r></w:hyperlink>')
        }
        elseif ($token -match '^\*\*(.+)\*\*$') {
            $boldText = [System.Security.SecurityElement]::Escape($Matches[1])
            [void]$runsXml.Append('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">' + $boldText + '</w:t></w:r>')
        }
        else {
            $plainText = [System.Security.SecurityElement]::Escape($token)
            [void]$runsXml.Append('<w:r><w:t xml:space="preserve">' + $plainText + '</w:t></w:r>')
        }
    }

    return $runsXml.ToString()
}

function ConvertTo-AsBuiltDocxParagraphXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$Text = '',
        [Parameter(Mandatory = $false)][string]$StyleId
    )

    $pPrXml = if (-not [string]::IsNullOrWhiteSpace($StyleId)) { '<w:pPr><w:pStyle w:val="' + $StyleId + '"/></w:pPr>' } else { '' }
    $runsXml = ConvertTo-AsBuiltDocxRunsXml -Text $Text
    return '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + $pPrXml + $runsXml + '</w:p>'
}

function Split-AsBuiltMarkdownTableRow {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Line)

    $trimmed = $Line.Trim() -replace '^\|', '' -replace '\|$', ''
    $rawCells = [System.Text.RegularExpressions.Regex]::Split($trimmed, '(?<!\\)\|')

    # Built with an explicit List(Of String) rather than @(... | ForEach-Object {...}): PowerShell's
    # pipeline output flattens a single-element array result back to a bare scalar, which silently
    # broke single-column tables (the caller's .Count access would then fail).
    $cells = New-Object System.Collections.Generic.List[string]
    foreach ($rawCell in $rawCells) {
        $cells.Add(($rawCell -replace '\\\|', '|').Trim())
    }
    return , $cells.ToArray()
}

function ConvertTo-AsBuiltDocxTableXml {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string[]]$TableLines)

    # Row 2 of a GFM table (e.g. "| --- | --- |") is the header separator and carries no content.
    $dataLines = New-Object System.Collections.Generic.List[string]
    foreach ($tableLine in $TableLines) {
        if ($tableLine.Trim() -notmatch '^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$') {
            $dataLines.Add($tableLine)
        }
    }
    if ($dataLines.Count -eq 0) {
        return $null
    }

    $rows = New-Object System.Collections.Generic.List[string[]]
    foreach ($dataLine in $dataLines) {
        $rows.Add((Split-AsBuiltMarkdownTableRow -Line $dataLine))
    }

    $columnCount = $rows[0].Count

    # Pre-pass: cells containing the merge marker (data rows only, never the header row) are
    # "continue" merges; the nearest preceding non-marker cell in that column becomes "restart".
    $mergeMarker = $script:AsBuiltTableMergeMarker
    $mergeTypes = New-Object 'string[,]' $rows.Count, $columnCount
    $lastNonMergedRowByColumn = New-Object 'int[]' $columnCount
    for ($col = 0; $col -lt $columnCount; $col++) { $lastNonMergedRowByColumn[$col] = -1 }

    for ($rowIndex = 1; $rowIndex -lt $rows.Count; $rowIndex++) {
        for ($col = 0; $col -lt $columnCount; $col++) {
            if ($col -lt $rows[$rowIndex].Count -and $rows[$rowIndex][$col] -eq $mergeMarker) {
                $mergeTypes[$rowIndex, $col] = 'continue'
                $anchorRow = $lastNonMergedRowByColumn[$col]
                $anchorMergeType = if ($anchorRow -ge 0) { $mergeTypes[$anchorRow, $col] } else { 'n/a' }
                if ($anchorRow -ge 0 -and [string]::IsNullOrEmpty($anchorMergeType)) {
                    $mergeTypes[$anchorRow, $col] = 'restart'
                }
            }
            else {
                $lastNonMergedRowByColumn[$col] = $rowIndex
            }
        }
    }

    $gridColumnsXml = ('<w:gridCol/>' * $columnCount) -join ''

    $tableXml = New-Object System.Text.StringBuilder
    [void]$tableXml.Append('<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid>' + $gridColumnsXml + '</w:tblGrid>')

    for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
        [void]$tableXml.Append('<w:tr>')
        $row = $rows[$rowIndex]
        for ($col = 0; $col -lt $columnCount; $col++) {
            $cell = if ($col -lt $row.Count) { $row[$col] } else { '' }
            $mergeType = $mergeTypes[$rowIndex, $col]
            $vMergeXml = switch ($mergeType) {
                'restart'  { '<w:vMerge w:val="restart"/>' }
                'continue' { '<w:vMerge/>' }
                default    { '' }
            }
            $cellText = if ($mergeType -eq 'continue') { '' } else { $cell }
            $runsXml = ConvertTo-AsBuiltDocxRunsXml -Text $cellText
            [void]$tableXml.Append('<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>' + $vMergeXml + '</w:tcPr><w:p>' + $runsXml + '</w:p></w:tc>')
        }
        [void]$tableXml.Append('</w:tr>')
    }
    [void]$tableXml.Append('</w:tbl>')

    return $tableXml.ToString()
}


function ConvertTo-AsBuiltDocxBodyFragments {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowEmptyCollection()][string[]]$MarkdownLines = @())

    $fragments = New-Object System.Collections.Generic.List[string]
    $index = 0

    while ($index -lt $MarkdownLines.Count) {
        $line = $MarkdownLines[$index]
        $trimmed = $line.Trim()

        if ($trimmed -eq '```{=openxml}') {
            $index++
            $rawLines = New-Object System.Collections.Generic.List[string]
            while ($index -lt $MarkdownLines.Count -and $MarkdownLines[$index].Trim() -ne '```') {
                $rawLines.Add($MarkdownLines[$index])
                $index++
            }
            $index++ # skip closing fence

            $rawXml = ($rawLines -join "`n").Trim()
            if (-not [string]::IsNullOrWhiteSpace($rawXml)) {
                $wrapped = '<asbuilt-fragment-root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + $rawXml + '</asbuilt-fragment-root>'
                [xml]$rawDocument = $wrapped
                foreach ($child in $rawDocument.DocumentElement.ChildNodes) {
                    $fragments.Add($child.OuterXml)
                }
            }
            continue
        }

        if ($trimmed.StartsWith('|')) {
            $tableLines = New-Object System.Collections.Generic.List[string]
            while ($index -lt $MarkdownLines.Count -and $MarkdownLines[$index].Trim().StartsWith('|')) {
                $tableLines.Add($MarkdownLines[$index])
                $index++
            }
            $tableXml = ConvertTo-AsBuiltDocxTableXml -TableLines $tableLines
            if ($null -ne $tableXml) {
                $fragments.Add($tableXml)
            }
            continue
        }

        if ($trimmed -match '^(#{1,6})\s+(.*)$') {
            $fragments.Add((ConvertTo-AsBuiltDocxParagraphXml -Text $Matches[2] -StyleId ("Heading{0}" -f $Matches[1].Length)))
            $index++
            continue
        }

        if ($trimmed -match '^-\s+(.*)$') {
            $bulletText = [string][char]0x2022 + ' ' + $Matches[1]
            $fragments.Add((ConvertTo-AsBuiltDocxParagraphXml -Text $bulletText -StyleId 'ListParagraph'))
            $index++
            continue
        }

        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            $index++
            continue
        }

        $fragments.Add((ConvertTo-AsBuiltDocxParagraphXml -Text $line))
        $index++
    }

    return $fragments
}

function Get-AsBuiltOpenXmlDescendants {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Element,
        [Parameter(Mandatory = $true)][type]$Type,
        [Parameter(Mandatory = $true)]$ResultList
    )

    # Written as plain recursion over .ChildElements (not the SDK's generic Descendants<T>()/
    # GetFirstChild<T>() methods) to avoid PowerShell's enumeration/generic-method quirks with
    # OpenXmlElement seen elsewhere in this file.
    foreach ($child in @($Element.ChildElements)) {
        if ($Type.IsInstanceOfType($child)) {
            $ResultList.Add($child)
        }
        Get-AsBuiltOpenXmlDescendants -Element $child -Type $Type -ResultList $ResultList
    }
}

function Resolve-AsBuiltDocxHyperlinkPlaceholders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Body,
        [Parameter(Mandatory = $true)]$MainDocumentPart
    )

    $hyperlinks = New-Object System.Collections.Generic.List[object]
    Get-AsBuiltOpenXmlDescendants -Element $Body -Type ([DocumentFormat.OpenXml.Wordprocessing.Hyperlink]) -ResultList $hyperlinks

    # One external relationship per unique URL; AddHyperlinkRelationship mints the real r:id that
    # replaces the placeholder written by ConvertTo-AsBuiltDocxRunsXml.
    $resolvedRelationshipIds = @{}
    foreach ($hyperlink in $hyperlinks) {
        $currentId = [string]$hyperlink.Id.Value
        if (-not $currentId.StartsWith($script:AsBuiltHyperlinkPlaceholderPrefix, [System.StringComparison]::Ordinal)) {
            continue
        }

        $targetUrl = $currentId.Substring($script:AsBuiltHyperlinkPlaceholderPrefix.Length)
        if (-not $resolvedRelationshipIds.ContainsKey($targetUrl)) {
            $relationship = $MainDocumentPart.AddHyperlinkRelationship([Uri]$targetUrl, $true)
            $resolvedRelationshipIds[$targetUrl] = $relationship.Id
        }

        $hyperlink.Id = $resolvedRelationshipIds[$targetUrl]
    }
}

function Convert-AsBuiltMarkdownToDocx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$MarkdownPath,
        [Parameter(Mandatory = $true)][string]$DocxPath,
        [Parameter(Mandatory = $true)][string]$ReferenceDocumentPath,
        [Parameter(Mandatory = $false)][string]$TableStyleName,
        [Parameter(Mandatory = $false)][bool]$NumberSections = $false
    )

    if (-not (Test-Path -LiteralPath $MarkdownPath -PathType Leaf)) {
        throw "Markdown input not found: $MarkdownPath"
    }
    if (-not (Test-Path -LiteralPath $ReferenceDocumentPath -PathType Leaf)) {
        throw "Reference template not found: $ReferenceDocumentPath"
    }

    Initialize-AsBuiltOpenXmlAssemblies

    # Heading auto-numbering (NumberSections) is inherited from the reference template's own
    # Heading1-3 style/numbering definitions; this function does not synthesize numbering.xml.
    Copy-Item -LiteralPath $ReferenceDocumentPath -Destination $DocxPath -Force

    $markdownLines = [System.IO.File]::ReadAllLines($MarkdownPath)
    $bodyFragmentsXml = ConvertTo-AsBuiltDocxBodyFragments -MarkdownLines $markdownLines

    $wordDocument = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($DocxPath, $true)
    try {
        $mainPart = $wordDocument.MainDocumentPart
        $body = $mainPart.Document.Body

        # The reference template defines two sections: a title-page section (its sectPr is embedded
        # in a paragraph mid-body, carrying the first-page header/footer/titlePg) and a bare final
        # body section. Both must be preserved - the title-page one is discarded by RemoveAllChildren()
        # below unless captured first, and the platform modules' own title-page content ends with a
        # placeholder section-break paragraph (bare sectPr, no header/footer) that needs the real
        # title-page section spliced back into it.
        $templateSectionProperties = New-Object System.Collections.Generic.List[object]
        Get-AsBuiltOpenXmlDescendants -Element $body -Type ([DocumentFormat.OpenXml.Wordprocessing.SectionProperties]) -ResultList $templateSectionProperties

        $finalSectionPropertiesTemplate = $null
        $titlePageSectionPropertiesTemplate = $null
        if ($templateSectionProperties.Count -gt 0) {
            $finalSectionPropertiesTemplate = $templateSectionProperties[$templateSectionProperties.Count - 1].CloneNode($true)
        }
        if ($templateSectionProperties.Count -gt 1) {
            $titlePageSectionPropertiesTemplate = $templateSectionProperties[0].CloneNode($true)
        }

        $body.RemoveAllChildren()

        foreach ($fragmentXml in $bodyFragmentsXml) {
            [xml]$fragmentDocument = $fragmentXml
            $fragmentLocalName = $fragmentDocument.DocumentElement.LocalName

            # Assign directly rather than through a switch expression: OpenXmlCompositeElement
            # implements IEnumerable over its children, so a switch/pipeline output silently
            # unwraps a Paragraph/Table into its first child element instead of returning it whole.
            $element = $null
            if ($fragmentLocalName -eq 'p') {
                $element = [DocumentFormat.OpenXml.Wordprocessing.Paragraph]::new($fragmentXml)
            }
            elseif ($fragmentLocalName -eq 'tbl') {
                $element = [DocumentFormat.OpenXml.Wordprocessing.Table]::new($fragmentXml)
            }

            if ($null -ne $element) {
                $body.AppendChild($element) | Out-Null
            }
        }

        if ($null -ne $titlePageSectionPropertiesTemplate) {
            $placeholderSectionBreaks = New-Object System.Collections.Generic.List[object]
            Get-AsBuiltOpenXmlDescendants -Element $body -Type ([DocumentFormat.OpenXml.Wordprocessing.SectionProperties]) -ResultList $placeholderSectionBreaks

            foreach ($placeholderSectionProperties in $placeholderSectionBreaks) {
                $hasHeaderOrFooterReference = $false
                foreach ($sectionChild in @($placeholderSectionProperties.ChildElements)) {
                    if ($sectionChild -is [DocumentFormat.OpenXml.Wordprocessing.HeaderReference] -or $sectionChild -is [DocumentFormat.OpenXml.Wordprocessing.FooterReference]) {
                        $hasHeaderOrFooterReference = $true
                    }
                }

                if (-not $hasHeaderOrFooterReference) {
                    $paragraphProperties = $placeholderSectionProperties.Parent
                    if ($null -ne $paragraphProperties) {
                        $replacementSectionProperties = $titlePageSectionPropertiesTemplate.CloneNode($true)
                        $paragraphProperties.ReplaceChild($replacementSectionProperties, $placeholderSectionProperties) | Out-Null
                    }
                }
            }
        }

        if ($null -ne $finalSectionPropertiesTemplate) {
            $body.AppendChild($finalSectionPropertiesTemplate) | Out-Null
        }
        else {
            $body.AppendChild([DocumentFormat.OpenXml.Wordprocessing.SectionProperties]::new()) | Out-Null
        }

        Resolve-AsBuiltDocxHyperlinkPlaceholders -Body $body -MainDocumentPart $mainPart

        # Trial: force fields (DOCPROPERTY/TOC) to NOT auto-recalculate on open, matching the
        # OOXML-spec default behavior for an omitted <w:updateFields> element (see
        # DocumentFormat.OpenXml.Wordprocessing.UpdateFieldsOnOpen), made explicit here rather
        # than relying on the reference template to omit it.
        # Reflection is used for the generic OpenXml SDK calls below because the
        # `$obj.Method[Type]()` explicit-generic-argument syntax is not valid in Windows
        # PowerShell 5.1 (it parses as an array indexer and throws a parser error).
        $settingsPart = $mainPart.DocumentSettingsPart
        if ($null -eq $settingsPart) {
            $addNewPartMethod = $mainPart.GetType().GetMethods() |
                Where-Object { $_.Name -eq 'AddNewPart' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 0 } |
                Select-Object -First 1
            $settingsPart = $addNewPartMethod.MakeGenericMethod([DocumentFormat.OpenXml.Packaging.DocumentSettingsPart]).Invoke($mainPart, $null)
            $settingsPart.Settings = [DocumentFormat.OpenXml.Wordprocessing.Settings]::new()
        }
        $settings = $settingsPart.Settings
        $getFirstChildMethod = $settings.GetType().GetMethod('GetFirstChild', [Type[]]@())
        $updateFieldsOnOpen = $getFirstChildMethod.MakeGenericMethod([DocumentFormat.OpenXml.Wordprocessing.UpdateFieldsOnOpen]).Invoke($settings, $null)
        if ($null -eq $updateFieldsOnOpen) {
            $updateFieldsOnOpen = [DocumentFormat.OpenXml.Wordprocessing.UpdateFieldsOnOpen]::new()
            $settings.PrependChild($updateFieldsOnOpen) | Out-Null
        }
        $updateFieldsOnOpen.Val = $false
        $settings.Save()

        $mainPart.Document.Save()
    }
    finally {
        $wordDocument.Dispose()
    }

    if (-not (Test-Path -LiteralPath $DocxPath -PathType Leaf)) {
        throw "DOCX generation failed: $DocxPath"
    }

    return (Resolve-Path -LiteralPath $DocxPath).ProviderPath
}

Export-ModuleMember -Function @(
    'Initialize-AsBuiltPath',
    'Resolve-AsBuiltPath',
    'Resolve-AsBuiltExistingFile',
    'ConvertTo-AsBuiltSafeName',
    'Get-AsBuiltPropertyValue',
    'ConvertTo-AsBuiltArray',
    'New-AsBuiltSettings',
    'Initialize-AsBuiltOpenXmlAssemblies',
    'Convert-AsBuiltMarkdownToDocx',
    'Get-AsBuiltTableMergeMarker'
)