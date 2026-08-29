Set-StrictMode -Version Latest
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'AsBuilt.Common.psm1') -Force

function Invoke-ESeriesAsBuilt {
[CmdletBinding()]
param(
    # Preferred form: https://santricity.example.com or https://santricity.example.com:8443
    [Parameter(Mandatory = $false)]
    [string]$Target,

    # Used only with legacy host-only Target input; full URLs use their explicit port or 8443.
    [Parameter(Mandatory = $false)]
    [int]$Port = 8443,

    [Parameter(Mandatory = $false)]
    [string]$ApiUsername,

    # Backward-compatible secure password input; -Credential is preferred.
    [Parameter(Mandatory = $false)]
    [System.Security.SecureString]$ApiPasswordSecure,

    # Backward-compatible alias for legacy plain-text password input; -Credential is preferred.
    [Parameter(Mandatory = $false)]
    [Alias("ApiPassword")]
    [string]$LegacyApiSecret,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$Ssid,

    [Parameter(Mandatory = $false)]
    [switch]$ValidateCerts,

    [Parameter(Mandatory = $false)]
    [switch]$UseSystemProxy,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir = ".\\asbuilt_output_eseries",

    [Parameter(Mandatory = $false)]
    [switch]$CollectionOnly,

    # Load previously collected JSON instead of running a live collection.
    [Parameter(Mandatory = $false)]
    [string]$JsonInputPath,

    [Parameter(Mandatory = $false)]
    [string]$ReportPrefix = "",

    [Parameter(Mandatory = $false)]
    [switch]$CleanupIntermediateOutputs,

    [Parameter(Mandatory = $false)]
    [switch]$KeepIntermediateOutputs = $false,

    [Parameter(Mandatory = $false)]
    [switch]$EnableDocxToc,

    [Parameter(Mandatory = $false)]
    [int]$DocxTocDepth = 3,

    [Parameter(Mandatory = $false)]
    [switch]$EnableDocxNumberSections,

    [Parameter(Mandatory = $false)]
    [string]$DocxTableStyleName = "NetAppTable1",

    [Parameter(Mandatory = $false)]
    [string]$DocxTableHeaderParagraphStyle = "TableHeading",

    [Parameter(Mandatory = $false)]
    [string]$DocxTableBodyParagraphStyle = "TableText",

    [Parameter(Mandatory = $false)]
    [switch]$DocxTableAutofitToWindow,

    [Parameter(Mandatory = $false)]
    [switch]$DocxEnableTitlePage,

    [Parameter(Mandatory = $false)]
    [Alias("TitleCustomerName")]
    [string]$CustomerName,

    [Parameter(Mandatory = $false)]
    [Alias("TitleCustomerLocation")]
    [string]$CustomerLocation,

    [Parameter(Mandatory = $false)]
    [Alias("TitleProjectName")]
    [string]$ProjectName,

    [Parameter(Mandatory = $false)]
    [hashtable]$TitlePageFields,

    [Parameter(Mandatory = $false)]
    [string]$TitlePageFieldsJson,

    [Parameter(Mandatory = $false)]
    [switch]$PromptForTitlePageFields,

    [Parameter(Mandatory = $true)]
    [string]$WorkspaceRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Import-Module (Join-Path -Path $WorkspaceRoot -ChildPath 'Modules\AsBuilt.Common.psm1') -Force
# ChangeLog:
# 2026.08.29.1 - Removed the stale w:dirty flag and added cached placeholder text on the TOC field so Word no longer prompts to update fields on open; relies on AsBuilt.Common.psm1 2026.08.29.1 for the accompanying UpdateFieldsOnOpen=false setting.
# 2026.08.26.1 - Replaced Pandoc-based DOCX generation with a direct DocumentFormat.OpenXml SDK pipeline (Lib\OpenXml); no external DOCX conversion binary is required. Fixed title-page section header/footer/titlePg preservation and single-column table parsing, added selective hyperlink support for markdown [text](url) content, and removed all Pandoc-related code and dependencies.
$ScriptVersion = "2026.08.29.1"

if (-not $PSBoundParameters.ContainsKey('CleanupIntermediateOutputs')) {
    $CleanupIntermediateOutputs = $true
}
if (-not $PSBoundParameters.ContainsKey('EnableDocxToc')) {
    $EnableDocxToc = $true
}
if (-not $PSBoundParameters.ContainsKey('DocxTableAutofitToWindow')) {
    $DocxTableAutofitToWindow = $true
}
if (-not $PSBoundParameters.ContainsKey('DocxEnableTitlePage')) {
    $DocxEnableTitlePage = $true
}

Write-Host "PowerShell Script Version: $ScriptVersion"

function Initialize-Path {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Resolve-OptionalPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return (Join-Path -Path $WorkspaceRoot -ChildPath ($Path -replace "^[.]\\", ""))
}

function Resolve-ExistingFilePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $candidate = Resolve-OptionalPath -Path $Path
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        return $null
    }

    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
        return (Resolve-Path -LiteralPath $candidate).ProviderPath
    }

    return $null
}

function Set-DocxTableStyle {
    param(
        [Parameter(Mandatory = $true)][string]$DocxPath,
        [Parameter(Mandatory = $true)][string]$TableStyleId,
        [Parameter(Mandatory = $false)][string]$HeaderParagraphStyle = "",
        [Parameter(Mandatory = $false)][string]$BodyParagraphStyle = "",
        [Parameter(Mandatory = $false)][bool]$AutofitToWindow = $true
    )

    if (-not (Test-Path -LiteralPath $DocxPath -PathType Leaf)) {
        throw "DOCX not found: $DocxPath"
    }

    if ([string]::IsNullOrWhiteSpace($TableStyleId)) {
        throw "TableStyleId must not be empty."
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::Open($DocxPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        if ($null -eq $entry) {
            throw "word/document.xml not found in DOCX: $DocxPath"
        }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        $documentXmlText = $reader.ReadToEnd()
        $reader.Dispose()

        [xml]$documentXml = $documentXmlText
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

        function Get-OrCreate-XmlChildNode {
            param(
                [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Parent,
                [Parameter(Mandatory = $true)][string]$ChildName,
                [Parameter(Mandatory = $false)][bool]$InsertAtStart = $false
            )

            $existing = $Parent.SelectSingleNode('./w:' + $ChildName, $namespaceManager)
            if ($null -ne $existing) {
                return $existing
            }

            $newNode = $documentXml.CreateElement('w', $ChildName, 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            if ($InsertAtStart -and $Parent.HasChildNodes) {
                $null = $Parent.PrependChild($newNode)
            }
            else {
                $null = $Parent.AppendChild($newNode)
            }

            return $newNode
        }

        function Set-ParagraphStyleInCell {
            param(
                [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Cell,
                [Parameter(Mandatory = $true)][string]$ParagraphStyle
            )

            if ([string]::IsNullOrWhiteSpace($ParagraphStyle)) {
                return 0
            }

            $updated = 0
            $paragraphs = @($Cell.SelectNodes('.//w:p', $namespaceManager))
            foreach ($paragraph in $paragraphs) {
                $pPr = Get-OrCreate-XmlChildNode -Parent $paragraph -ChildName 'pPr' -InsertAtStart $true
                $pStyle = $pPr.SelectSingleNode('./w:pStyle', $namespaceManager)
                if ($null -eq $pStyle) {
                    $pStyle = $documentXml.CreateElement('w', 'pStyle', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                    $null = $pPr.PrependChild($pStyle)
                }

                $valAttr = $pStyle.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                if ($valAttr -ne $ParagraphStyle) {
                    [void]$pStyle.SetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', $ParagraphStyle)
                    $updated += 1
                }
            }

            return $updated
        }

        $tables = @($documentXml.SelectNodes('//w:tbl', $namespaceManager))
        if ($tables.Count -eq 0) {
            $entry.Delete()
            $replacementEntry = $zip.CreateEntry('word/document.xml')
            $writer = New-Object System.IO.StreamWriter($replacementEntry.Open())
            $writer.Write($documentXml.OuterXml)
            $writer.Dispose()
            Write-Host "INFO: No tables found in document"
            return
        }

        $updatedTableStyle = 0
        $updatedHeaderParagraphs = 0
        $updatedBodyParagraphs = 0
        $updatedAutofit = 0
        $updatedCantSplit = 0

        foreach ($table in $tables) {
            $tblPr = Get-OrCreate-XmlChildNode -Parent $table -ChildName 'tblPr' -InsertAtStart $true
            $tblStyle = $tblPr.SelectSingleNode('./w:tblStyle', $namespaceManager)
            if ($null -eq $tblStyle) {
                $tblStyle = $documentXml.CreateElement('w', 'tblStyle', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                $null = $tblPr.PrependChild($tblStyle)
            }

            if ($tblStyle.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') -ne $TableStyleId) {
                [void]$tblStyle.SetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', $TableStyleId)
                $updatedTableStyle += 1
            }

            if ($AutofitToWindow) {
                $tblW = Get-OrCreate-XmlChildNode -Parent $tblPr -ChildName 'tblW'
                $changed = $false
                if ($tblW.GetAttribute('type', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') -ne 'pct') {
                    [void]$tblW.SetAttribute('type', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'pct')
                    $changed = $true
                }
                if ($tblW.GetAttribute('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') -ne '5000') {
                    [void]$tblW.SetAttribute('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', '5000')
                    $changed = $true
                }

                $tblLayout = Get-OrCreate-XmlChildNode -Parent $tblPr -ChildName 'tblLayout'
                if ($tblLayout.GetAttribute('type', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') -ne 'autofit') {
                    [void]$tblLayout.SetAttribute('type', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'autofit')
                    $changed = $true
                }

                $tblLook = Get-OrCreate-XmlChildNode -Parent $tblPr -ChildName 'tblLook'
                if ($tblLook.GetAttribute('firstRow', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') -ne '1') {
                    [void]$tblLook.SetAttribute('firstRow', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', '1')
                    $changed = $true
                }

                if ($changed) {
                    $updatedAutofit += 1
                }
            }

            $rows = @($table.SelectNodes('./w:tr', $namespaceManager))
            if ($rows.Count -gt 0) {
                foreach ($row in $rows) {
                    $trPr = Get-OrCreate-XmlChildNode -Parent $row -ChildName 'trPr' -InsertAtStart $true
                    $cantSplit = Get-OrCreate-XmlChildNode -Parent $trPr -ChildName 'cantSplit'
                    if ($null -ne $cantSplit) {
                        $updatedCantSplit += 1
                    }
                }

                foreach ($cell in @($rows[0].SelectNodes('./w:tc', $namespaceManager))) {
                    $updatedHeaderParagraphs += Set-ParagraphStyleInCell -Cell $cell -ParagraphStyle $HeaderParagraphStyle
                }

                for ($rowIndex = 1; $rowIndex -lt $rows.Count; $rowIndex++) {
                    foreach ($cell in @($rows[$rowIndex].SelectNodes('./w:tc', $namespaceManager))) {
                        $updatedBodyParagraphs += Set-ParagraphStyleInCell -Cell $cell -ParagraphStyle $BodyParagraphStyle
                    }
                }
            }
        }

        $entry.Delete()
        $replacementEntry = $zip.CreateEntry('word/document.xml')
        $writer = New-Object System.IO.StreamWriter($replacementEntry.Open())
        $writer.Write($documentXml.OuterXml)
        $writer.Dispose()

        Write-Verbose "INFO: Applied table style '$TableStyleId' to $updatedTableStyle table(s)"
        Write-Verbose "INFO: Updated header paragraphs with style '$HeaderParagraphStyle': $updatedHeaderParagraphs"
        Write-Verbose "INFO: Updated body paragraphs with style '$BodyParagraphStyle': $updatedBodyParagraphs"
        Write-Verbose "INFO: Applied autofit-to-window settings on $updatedAutofit table(s)"
        Write-Verbose "INFO: Applied non-splitting row setting on $updatedCantSplit table row(s)"
    }
    finally {
        $zip.Dispose()
    }
}

function Convert-ToSafeName {
    param([Parameter(Mandatory = $true)][string]$Value)

    $safe = $Value.ToLowerInvariant()
    $safe = [regex]::Replace($safe, "[^a-z0-9-]+", "_")
    $safe = $safe.Trim("_")

    if ([string]::IsNullOrWhiteSpace($safe)) {
        return "unknown"
    }

    return $safe
}

function Read-ReportMetadata {
    param([Parameter(Mandatory = $false)][string]$MetadataPath = ".\docx\report-metadata-e-series.yml")

    $reportMetadata = @{
        title = "As-Built Documentation"
        subject = "NetApp E-Series"
        author = "NetApp Professional Services"
        "toc-title" = "Table of Contents"
    }

    $reportMetadataPath = Resolve-ExistingFilePath -Path $MetadataPath
    if ($null -eq $reportMetadataPath) {
        return $reportMetadata
    }

    foreach ($line in Get-Content -Path $reportMetadataPath -ErrorAction SilentlyContinue) {
        if ($line -match '^\s*([A-Za-z0-9_-]+):\s*(.*)$') {
            $metadataKey = $matches[1]
            $metadataValue = $matches[2].Trim()

            if ($metadataValue.StartsWith('"') -and $metadataValue.EndsWith('"') -and $metadataValue.Length -ge 2) {
                $metadataValue = $metadataValue.Substring(1, $metadataValue.Length - 2)
            }
            elseif ($metadataValue.StartsWith("'") -and $metadataValue.EndsWith("'") -and $metadataValue.Length -ge 2) {
                $metadataValue = $metadataValue.Substring(1, $metadataValue.Length - 2)
            }

            if (-not [string]::IsNullOrWhiteSpace($metadataValue)) {
                $reportMetadata[$metadataKey] = $metadataValue
            }
        }
    }

    return $reportMetadata
}

function Set-DocxMetadataProperties {
    param(
        [Parameter(Mandatory = $true)][string]$DocxPath,
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][string]$Author,
        [Parameter(Mandatory = $false)][string]$CustomerName,
        [Parameter(Mandatory = $false)][string]$CustomerLocation,
        [Parameter(Mandatory = $false)][string]$ProjectName,
        [Parameter(Mandatory = $false)][string]$SystemName
    )

    if (-not (Test-Path -LiteralPath $DocxPath -PathType Leaf)) {
        throw "DOCX not found: $DocxPath"
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $openXmlNs = "http://schemas.openxmlformats.org/package/2006/relationships"
    $coreNs = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    $dcNs = "http://purl.org/dc/elements/1.1/"
    $customNs = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    $vtNs = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    $customFmtId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

    function Read-ZipEntryText {
        param(
            [Parameter(Mandatory = $true)]$Zip,
            [Parameter(Mandatory = $true)][string]$EntryPath
        )

        $entry = $Zip.GetEntry($EntryPath)
        if ($null -eq $entry) {
            return $null
        }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }

    function Write-ZipEntryText {
        param(
            [Parameter(Mandatory = $true)]$Zip,
            [Parameter(Mandatory = $true)][string]$EntryPath,
            [Parameter(Mandatory = $true)][string]$Text
        )

        $entry = $Zip.GetEntry($EntryPath)
        if ($null -ne $entry) {
            $entry.Delete()
        }

        $newEntry = $Zip.CreateEntry($EntryPath)
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $writer = New-Object System.IO.StreamWriter($newEntry.Open(), $utf8NoBom)
        try {
            $writer.Write($Text)
        }
        finally {
            $writer.Dispose()
        }
    }

    function Get-OrCreate-CoreElement {
        param(
            [Parameter(Mandatory = $true)][xml]$Xml,
            [Parameter(Mandatory = $true)][System.Xml.XmlElement]$Root,
            [Parameter(Mandatory = $true)][string]$Prefix,
            [Parameter(Mandatory = $true)][string]$LocalName,
            [Parameter(Mandatory = $true)][string]$NamespaceUri
        )

        $nsmgr = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
        $nsmgr.AddNamespace('cp', $coreNs)
        $nsmgr.AddNamespace('dc', $dcNs)
        $existing = $Root.SelectSingleNode('./' + $Prefix + ':' + $LocalName, $nsmgr)
        if ($null -ne $existing) {
            return $existing
        }

        $newNode = $Xml.CreateElement($Prefix, $LocalName, $NamespaceUri)
        $null = $Root.AppendChild($newNode)
        return $newNode
    }

    function Set-CustomStringProperty {
        param(
            [Parameter(Mandatory = $true)][xml]$Xml,
            [Parameter(Mandatory = $true)][System.Xml.XmlElement]$Root,
            [Parameter(Mandatory = $true)][string]$Name,
            [Parameter(Mandatory = $true)][string]$Value
        )

        $nsmgr = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
        $nsmgr.AddNamespace('cp', $customNs)
        $nsmgr.AddNamespace('vt', $vtNs)

        $propertyNode = $Root.SelectSingleNode('./cp:property[@name="' + $Name + '"]', $nsmgr)
        if ($null -eq $propertyNode) {
            $nextPid = 2
            foreach ($existingProp in @($Root.SelectNodes('./cp:property', $nsmgr))) {
                $candidatePid = 0
                if ([int]::TryParse([string]$existingProp.GetAttribute('pid'), [ref]$candidatePid)) {
                    if ($candidatePid -ge $nextPid) {
                        $nextPid = $candidatePid + 1
                    }
                }
            }

            $propertyNode = $Xml.CreateElement('property', $customNs)
            [void]$propertyNode.SetAttribute('fmtid', $customFmtId)
            [void]$propertyNode.SetAttribute('pid', [string]$nextPid)
            [void]$propertyNode.SetAttribute('name', $Name)
            $null = $Root.AppendChild($propertyNode)
        }

        foreach ($child in @($propertyNode.ChildNodes)) {
            $null = $propertyNode.RemoveChild($child)
        }

        $valueNode = $Xml.CreateElement('vt', 'lpwstr', $vtNs)
        $valueNode.InnerText = $Value
        $null = $propertyNode.AppendChild($valueNode)
    }

    $zip = [System.IO.Compression.ZipFile]::Open($DocxPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $coreText = Read-ZipEntryText -Zip $zip -EntryPath 'docProps/core.xml'
        if ([string]::IsNullOrWhiteSpace($coreText)) {
            $coreText = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="' + $coreNs + '" xmlns:dc="' + $dcNs + '" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>'
        }

        [xml]$coreXml = $coreText
        $coreRoot = $coreXml.DocumentElement

        $titleNode = Get-OrCreate-CoreElement -Xml $coreXml -Root $coreRoot -Prefix 'dc' -LocalName 'title' -NamespaceUri $dcNs
        $titleNode.InnerText = $Title

        $subjectNode = Get-OrCreate-CoreElement -Xml $coreXml -Root $coreRoot -Prefix 'dc' -LocalName 'subject' -NamespaceUri $dcNs
        $subjectNode.InnerText = $Subject

        $creatorNode = Get-OrCreate-CoreElement -Xml $coreXml -Root $coreRoot -Prefix 'dc' -LocalName 'creator' -NamespaceUri $dcNs
        $creatorNode.InnerText = $Author

        $lastModifiedByNode = Get-OrCreate-CoreElement -Xml $coreXml -Root $coreRoot -Prefix 'cp' -LocalName 'lastModifiedBy' -NamespaceUri $coreNs
        $lastModifiedByNode.InnerText = $Author

        Write-ZipEntryText -Zip $zip -EntryPath 'docProps/core.xml' -Text $coreXml.OuterXml

        $contentTypesText = Read-ZipEntryText -Zip $zip -EntryPath '[Content_Types].xml'
        if (-not [string]::IsNullOrWhiteSpace($contentTypesText)) {
            [xml]$contentTypesXml = $contentTypesText
            $ctNsMgr = New-Object System.Xml.XmlNamespaceManager($contentTypesXml.NameTable)
            $ctNsMgr.AddNamespace('ct', 'http://schemas.openxmlformats.org/package/2006/content-types')
            $customOverride = $contentTypesXml.SelectSingleNode('/ct:Types/ct:Override[@PartName="/docProps/custom.xml"]', $ctNsMgr)
            if ($null -eq $customOverride) {
                $overrideNode = $contentTypesXml.CreateElement('Override', 'http://schemas.openxmlformats.org/package/2006/content-types')
                [void]$overrideNode.SetAttribute('PartName', '/docProps/custom.xml')
                [void]$overrideNode.SetAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.custom-properties+xml')
                $null = $contentTypesXml.DocumentElement.AppendChild($overrideNode)
                Write-ZipEntryText -Zip $zip -EntryPath '[Content_Types].xml' -Text $contentTypesXml.OuterXml
            }
        }

        $relsText = Read-ZipEntryText -Zip $zip -EntryPath '_rels/.rels'
        if (-not [string]::IsNullOrWhiteSpace($relsText)) {
            [xml]$relsXml = $relsText
            $relsNsMgr = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
            $relsNsMgr.AddNamespace('r', $openXmlNs)

            $customRel = $relsXml.SelectSingleNode('/r:Relationships/r:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"]', $relsNsMgr)
            if ($null -eq $customRel) {
                $highestRid = 0
                foreach ($relNode in @($relsXml.SelectNodes('/r:Relationships/r:Relationship', $relsNsMgr))) {
                    if ($relNode.Id -match '^rId(\d+)$') {
                        $idNum = [int]$matches[1]
                        if ($idNum -gt $highestRid) {
                            $highestRid = $idNum
                        }
                    }
                }

                $newRel = $relsXml.CreateElement('Relationship', $openXmlNs)
                [void]$newRel.SetAttribute('Id', 'rId' + ($highestRid + 1))
                [void]$newRel.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties')
                [void]$newRel.SetAttribute('Target', 'docProps/custom.xml')
                $null = $relsXml.DocumentElement.AppendChild($newRel)
                Write-ZipEntryText -Zip $zip -EntryPath '_rels/.rels' -Text $relsXml.OuterXml
            }
        }

        $customText = Read-ZipEntryText -Zip $zip -EntryPath 'docProps/custom.xml'
        if ([string]::IsNullOrWhiteSpace($customText)) {
            $customText = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="' + $customNs + '" xmlns:vt="' + $vtNs + '"></Properties>'
        }

        [xml]$customXml = $customText
        $customRoot = $customXml.DocumentElement

        Set-CustomStringProperty -Xml $customXml -Root $customRoot -Name 'CustomerName' -Value ([string]$CustomerName)
        Set-CustomStringProperty -Xml $customXml -Root $customRoot -Name 'CustomerLocation' -Value ([string]$CustomerLocation)
        Set-CustomStringProperty -Xml $customXml -Root $customRoot -Name 'ProjectName' -Value ([string]$ProjectName)
        Set-CustomStringProperty -Xml $customXml -Root $customRoot -Name 'SystemName' -Value ([string]$SystemName)

        Write-ZipEntryText -Zip $zip -EntryPath 'docProps/custom.xml' -Text $customXml.OuterXml
    }
    finally {
        $zip.Dispose()
    }
}

function Set-TlsPolicy {
    param([Parameter(Mandatory = $true)][bool]$ValidateCertificates)

    $protocols = [Net.SecurityProtocolType]::Tls12
    try {
        $protocols = $protocols -bor [Net.SecurityProtocolType]::Tls11
    }
    catch {
    }
    try {
        $protocols = $protocols -bor [Net.SecurityProtocolType]::Tls
    }
    catch {
    }

    [Net.ServicePointManager]::SecurityProtocol = $protocols
    [System.Net.ServicePointManager]::Expect100Continue = $false
    $script:UsePowerShell7SkipCertificateCheck = (
        -not $ValidateCertificates -and
        $PSVersionTable.PSVersion.Major -ge 6 -and
        (Get-Command Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck')
    )

    if (-not ("TrustAllCertsCallback" -as [type])) {
        Add-Type -TypeDefinition @"
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

public static class TrustAllCertsCallback
{
    public static readonly RemoteCertificateValidationCallback Callback =
        delegate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        };
}
"@
    }

    if (-not $ValidateCertificates) {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = [TrustAllCertsCallback]::Callback
    }
    else {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
    }
}

function Set-ProxyPolicy {
    param([Parameter(Mandatory = $true)][bool]$UseProxy)

    if (-not $UseProxy) {
        [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy
    }
}

function Get-ExceptionMessageChain {
    param([Parameter(Mandatory = $true)][System.Exception]$Exception)

    $messages = New-Object System.Collections.Generic.List[string]
    $cursor = $Exception
    while ($null -ne $cursor) {
        if (-not [string]::IsNullOrWhiteSpace($cursor.Message)) {
            $messages.Add($cursor.Message)
        }
        $cursor = $cursor.InnerException
    }

    if ($messages.Count -eq 0) {
        return "Unknown request failure"
    }

    return ($messages -join " | ")
}

function Get-PropertyValue {
    param(
        [Parameter(Mandatory = $false)]$Object,
        [Parameter(Mandatory = $false)][string]$PropertyName
    )

    if ($null -eq $Object) {
        return $null
    }

    # Defensive: avoid interactive prompts if a caller forgets -PropertyName.
    if ([string]::IsNullOrWhiteSpace($PropertyName)) {
        return $null
    }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($PropertyName)) {
            return $Object[$PropertyName]
        }

        if ($Object -is [hashtable] -and $Object.ContainsKey($PropertyName)) {
            return $Object[$PropertyName]
        }

        return $null
    }

    $prop = $Object.PSObject.Properties[$PropertyName]
    if ($null -ne $prop) {
        return $prop.Value
    }

    return $null
}

function ConvertTo-NormalizedApiPayload {
    param([Parameter(Mandatory = $false)]$Payload)

    if ($null -eq $Payload) {
        return @{}
    }

    $valueProperty = Get-PropertyValue -Object $Payload -PropertyName "value"
    $countProperty = Get-PropertyValue -Object $Payload -PropertyName "Count"

    if ($null -ne $valueProperty -and $null -ne $countProperty -and -not ($Payload -is [string])) {
        return $valueProperty
    }

    return $Payload
}

function Convert-ToArrayPayload {
    param([Parameter(Mandatory = $false)]$Payload)

    if ($null -eq $Payload) {
        return @()
    }

    if ($Payload -is [array]) {
        return [object[]]$Payload
    }

    if (
        $Payload -is [System.Collections.IEnumerable] -and
        -not ($Payload -is [string]) -and
        -not ($Payload -is [hashtable]) -and
        -not ($Payload -is [pscustomobject])
    ) {
        $items = @($Payload)
        if ($items.Count -eq 0) {
            return @()
        }
        return [object[]]$items
    }

    if ($Payload -is [hashtable]) {
        if ((@($Payload.GetEnumerator())).Count -eq 0) {
            return @()
        }
        return [object[]](,$Payload)
    }

    $propNames = @($Payload.PSObject.Properties.Name)
    if ($propNames.Count -gt 0) {
        return [object[]](,$Payload)
    }

    return [object[]](,$Payload)
}

function Get-ResponseData {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Responses,
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $false)][switch]$AsArray
    )

    if (-not $Responses.ContainsKey($Key) -or $null -eq $Responses[$Key]) {
        if ($AsArray) {
            return @()
        }
        return @{}
    }

    $data = $Responses[$Key].Data
    if ($AsArray) {
        $normalized = Convert-ToArrayPayload -Payload $data
        return [object[]]@($normalized)
    }

    if ($null -eq $data) {
        return @{}
    }

    return $data
}

function New-BasicAuthHeader {
    param([Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential)

    $plainPassword = [System.Net.NetworkCredential]::new("", $Credential.Password).Password
    $pair = "{0}:{1}" -f $Credential.UserName, $plainPassword
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $encoded = [System.Convert]::ToBase64String($bytes)
    return "Basic $encoded"
}

function Format-MarkdownCell {
    param([Parameter(Mandatory = $false)]$Value)

    if ($null -eq $Value) {
        return "N/A"
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return "N/A"
    }

    $text = $text -replace "\|", "\\|"
    $text = $text -replace "`r?`n", "<br />"
    return $text
}

function Convert-BytesToTiB {
    param([Parameter(Mandatory = $false)]$Bytes)

    if ($null -eq $Bytes) {
        return "0"
    }

    $n = 0
    if (-not [double]::TryParse([string]$Bytes, [ref]$n)) {
        return "0"
    }

    return ([math]::Round(($n / 1099511627776), 2)).ToString("0.##")
}

function Format-ShelfDisplayValue {
    param(
        [Parameter(Mandatory = $false)]$ShelfValue,
        [Parameter(Mandatory = $false)][int]$Width = 2
    )

    if ($null -eq $ShelfValue) {
        return "N/A"
    }

    $text = [string]$ShelfValue
    if ([string]::IsNullOrWhiteSpace($text)) {
        return "N/A"
    }

    $parsed = 0
    if ([int]::TryParse($text, [ref]$parsed)) {
        return $parsed.ToString(("D{0}" -f $Width))
    }

    return $text
}

function Get-FirstItem {
    param([Parameter(Mandatory = $false)]$Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [array]) {
        if ($Value.Count -gt 0) {
            return $Value[0]
        }
        return $null
    }

    return $Value
}

function Add-MarkdownTable {
    param(
        [Parameter(Mandatory = $true)]$LineBuffer,
        [Parameter(Mandatory = $true)][string[]]$Headers,
        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [object[]]$Rows = @()
    )

    if ($null -eq $LineBuffer -or ($LineBuffer -is [string])) {
        throw "LineBuffer was not initialized correctly for markdown rendering."
    }

    $LineBuffer.Add("| $($Headers -join ' | ') |")
    $separatorCells = @()
    foreach ($h in $Headers) {
        $separatorCells += "---"
    }
    $LineBuffer.Add("| $($separatorCells -join ' | ') |")

    if ($Rows.Count -eq 0) {
        $naRow = @()
        foreach ($h in $Headers) {
            $naRow += "N/A"
        }
        $LineBuffer.Add("| $($naRow -join ' | ') |")
        return
    }

    foreach ($row in $Rows) {
        $cells = @()
        foreach ($cell in $row) {
            $cells += (Format-MarkdownCell -Value $cell)
        }
        $LineBuffer.Add("| $($cells -join ' | ') |")
    }
}

function Write-EseriesMarkdown {
    param(
        [Parameter(Mandatory = $true)]$Export,
        [Parameter(Mandatory = $true)][string]$ApiUrl,
        [Parameter(Mandatory = $true)][string]$Ssid,
        [Parameter(Mandatory = $true)][string]$Iso8601,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)]$RestQueryFailures,
        [Parameter(Mandatory = $true)][hashtable]$TitleFields,
        [Parameter(Mandatory = $true)][bool]$IncludeTitlePage
    )

    $facts = $Export.storage_array_facts
    $storageSystem = Get-FirstItem -Value (Get-PropertyValue -Object $facts -PropertyName "netapp_storage_systems_details")
    $storagePools = @((Get-PropertyValue -Object $facts -PropertyName "netapp_storage_pools_details"))
    $volumes = @((Get-PropertyValue -Object $facts -PropertyName "netapp_volumes_details"))
    $hosts = @((Get-PropertyValue -Object $facts -PropertyName "netapp_hosts_details"))
    $hostTypes = @((Get-PropertyValue -Object $facts -PropertyName "netapp_host_types_details"))
    $hostGroups = @((Get-PropertyValue -Object $facts -PropertyName "netapp_host_groups_details"))
    $snapshotGroups = @((Get-PropertyValue -Object $facts -PropertyName "netapp_snapshot_groups_details"))
    $snapshotImages = @((Get-PropertyValue -Object $facts -PropertyName "netapp_snapshot_images_details"))
    $snapshotSchedules = @((Get-PropertyValue -Object $facts -PropertyName "netapp_snapshot_schedules_details"))
    $snapshotConsistencyGroups = @((Get-PropertyValue -Object $facts -PropertyName "netapp_snapshot_consistency_groups_details"))
    $snapshotConsistencyGroupMemberVolumes = @((Get-PropertyValue -Object $facts -PropertyName "netapp_snapshot_consistency_group_member_volumes_details"))
    $hw = Get-PropertyValue -Object $facts -PropertyName "netapp_hardware_inventory"
    $embeddedFwRoot = Get-PropertyValue -Object $facts -PropertyName "netapp_embedded_firmware_versions"
    $embeddedFwVersions = @((Get-PropertyValue -Object $embeddedFwRoot -PropertyName "codeVersions"))
    if ($embeddedFwVersions.Count -eq 0 -and $embeddedFwRoot -is [array]) {
        $embeddedFwVersions = @($embeddedFwRoot)
    }
    $ldapSettings = Get-PropertyValue -Object $facts -PropertyName "netapp_ldap_settings"
    $deviceAlerts = Get-PropertyValue -Object $facts -PropertyName "netapp_device_alerts_settings"
    $snmpAlertSettings = Get-PropertyValue -Object $facts -PropertyName "netapp_snmp_alert_settings"
    $deviceAlertSyslogSettings = Get-PropertyValue -Object $facts -PropertyName "netapp_device_alert_syslog_settings"
    $auditLogConfig = Get-PropertyValue -Object $facts -PropertyName "netapp_audit_log_config"
    $auditLogSyslogSettings = @((Get-PropertyValue -Object $facts -PropertyName "netapp_audit_log_syslog_settings"))
    $deviceAsup = Get-PropertyValue -Object $facts -PropertyName "netapp_device_asup_settings"
    $loginBanner = Get-PropertyValue -Object $facts -PropertyName "netapp_login_banner_settings"

    $controllers = @()
    $trays = @()
    $drawers = @()
    $drives = @()
    $esms = @()
    if ($null -ne $hw) {
        $controllers = @((Get-PropertyValue -Object $hw -PropertyName "controllers"))
        $trays = @((Get-PropertyValue -Object $hw -PropertyName "trays"))
        $drawers = @((Get-PropertyValue -Object $hw -PropertyName "drawers"))
        $drives = @((Get-PropertyValue -Object $hw -PropertyName "drives"))
        $esms = @((Get-PropertyValue -Object $hw -PropertyName "esms"))
    }

    $epochToUtc = {
        param($Epoch)
        if ($null -eq $Epoch) { return "N/A" }
        $e = 0L
        if (-not [long]::TryParse([string]$Epoch, [ref]$e) -or $e -eq 0) { return "N/A" }
        return [DateTimeOffset]::FromUnixTimeSeconds($e).UtcDateTime.ToString("dd/MM/yyyy HH:mm:ss")
    }

    $boolText = {
        param($Value)
        if ($null -eq $Value) { return "N/A" }
        $s = [string]$Value
        if ([string]::IsNullOrWhiteSpace($s)) { return "N/A" }
        return $s
    }

    $nameOf = {
        param($Obj)
        $n = Get-PropertyValue -Object $Obj -PropertyName "name"
        if ([string]::IsNullOrWhiteSpace([string]$n)) {
            $n = Get-PropertyValue -Object $Obj -PropertyName "label"
        }
        if ([string]::IsNullOrWhiteSpace([string]$n)) {
            return "N/A"
        }
        return [string]$n
    }

    $titleField = {
        param(
            [string]$Key,
            [string]$Fallback = "N/A"
        )

        if ($null -ne $TitleFields -and $TitleFields.ContainsKey($Key) -and -not [string]::IsNullOrWhiteSpace([string]$TitleFields[$Key])) {
            return [string]$TitleFields[$Key]
        }

        return $Fallback
    }

    $reportMetadata = Read-ReportMetadata -MetadataPath ".\docx\report-metadata-e-series.yml"

    $emitOpenXmlParagraph = {
        param(
            [Parameter(Mandatory = $true)][string]$StyleId,
            [Parameter(Mandatory = $false)][string]$Text,
            [Parameter(Mandatory = $false)][switch]$Blank,
            [Parameter(Mandatory = $false)][int]$LeftIndentTwips = 0
        )

        $paragraphPropertiesXml = '<w:pPr><w:pStyle w:val="' + $StyleId + '"/>'
        if ($LeftIndentTwips -gt 0) {
            $paragraphPropertiesXml += '<w:ind w:left="' + [string]$LeftIndentTwips + '"/>'
        }
        $paragraphPropertiesXml += '</w:pPr>'

        if ($Blank.IsPresent) {
            $markdownLines.Add('<w:p>' + $paragraphPropertiesXml + '</w:p>')
            return
        }

        $escapedText = [System.Security.SecurityElement]::Escape([string]$Text)
        $markdownLines.Add('<w:p>' + $paragraphPropertiesXml + '<w:r><w:t xml:space="preserve">' + $escapedText + '</w:t></w:r></w:p>')
    }

    $emitOpenXmlFieldParagraph = {
        param(
            [Parameter(Mandatory = $true)][string]$StyleId,
            [Parameter(Mandatory = $false)][AllowEmptyString()][string]$Label = "",
            [Parameter(Mandatory = $true)][string]$Instruction,
            [Parameter(Mandatory = $true)][string]$FallbackText,
            [Parameter(Mandatory = $false)][int]$LeftIndentTwips = 0
        )

        $escapedLabel = [System.Security.SecurityElement]::Escape([string]$Label)
        $escapedInstruction = [System.Security.SecurityElement]::Escape([string]$Instruction)
        $escapedFallback = [System.Security.SecurityElement]::Escape([string]$FallbackText)

        $fieldXml = '<w:p><w:pPr><w:pStyle w:val="' + $StyleId + '"/>'
        if ($LeftIndentTwips -gt 0) {
            $fieldXml += '<w:ind w:left="' + [string]$LeftIndentTwips + '"/>'
        }
        $fieldXml += '</w:pPr>'
        $fieldXml += '<w:r><w:t xml:space="preserve">' + $escapedLabel + '</w:t></w:r>'
        $fieldXml += '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        $fieldXml += '<w:r><w:instrText xml:space="preserve"> ' + $escapedInstruction + ' </w:instrText></w:r>'
        $fieldXml += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        $fieldXml += '<w:r><w:t xml:space="preserve">' + $escapedFallback + '</w:t></w:r>'
        $fieldXml += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        $fieldXml += '</w:p>'
        $markdownLines.Add($fieldXml)
    }

    $emitOpenXmlFieldBlockParagraph = {
        param(
            [Parameter(Mandatory = $true)][string]$StyleId,
            [Parameter(Mandatory = $true)][object[]]$Lines,
            [Parameter(Mandatory = $false)][int]$LeftIndentTwips = 0
        )

        $fieldXml = '<w:p><w:pPr><w:pStyle w:val="' + $StyleId + '"/>'
        if ($LeftIndentTwips -gt 0) {
            $fieldXml += '<w:ind w:left="' + [string]$LeftIndentTwips + '"/>'
        }
        $fieldXml += '</w:pPr>'

        for ($lineIndex = 0; $lineIndex -lt $Lines.Count; $lineIndex++) {
            $line = $Lines[$lineIndex]
            $lineLabel = [System.Security.SecurityElement]::Escape([string]$line.Label)
            $lineInstruction = [System.Security.SecurityElement]::Escape([string]$line.Instruction)
            $lineFallback = [System.Security.SecurityElement]::Escape([string]$line.FallbackText)

            $fieldXml += '<w:r><w:t xml:space="preserve">' + $lineLabel + '</w:t></w:r>'
            $fieldXml += '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            $fieldXml += '<w:r><w:instrText xml:space="preserve"> ' + $lineInstruction + ' </w:instrText></w:r>'
            $fieldXml += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            $fieldXml += '<w:r><w:t xml:space="preserve">' + $lineFallback + '</w:t></w:r>'
            $fieldXml += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

            if ($lineIndex -lt ($Lines.Count - 1)) {
                $fieldXml += '<w:r><w:br/></w:r>'
            }
        }

        $fieldXml += '</w:p>'
        $markdownLines.Add($fieldXml)
    }

    $emitOpenXmlMergedFirstColumnTable = {
        param(
            [Parameter(Mandatory = $true)][string[]]$Headers,
            [Parameter(Mandatory = $false)][object[]]$Rows = @(),
            [Parameter(Mandatory = $false)][int[]]$GridWidths = @()
        )

        $columnCount = $Headers.Count
        if ($columnCount -lt 2) {
            throw "Merged table emitter requires at least two columns."
        }

        if ($GridWidths.Count -ne $columnCount) {
            $GridWidths = @()
            for ($widthIndex = 0; $widthIndex -lt $columnCount; $widthIndex++) {
                $GridWidths += 2400
            }
        }

        $tableGridXml = ""
        foreach ($gridWidth in $GridWidths) {
            $tableGridXml += '<w:gridCol w:w="' + [string]$gridWidth + '"/>'
        }

        $markdownLines.Add('```{=openxml}')
        $markdownLines.Add('<w:tbl>')
        $markdownLines.Add('<w:tblPr><w:tblStyle w:val="NetAppTable1"/><w:tblW w:w="5000" w:type="pct"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>')
        $markdownLines.Add('<w:tblGrid>' + $tableGridXml + '</w:tblGrid>')

        $headerXml = '<w:tr><w:trPr><w:tblHeader/></w:trPr>'
        foreach ($header in $Headers) {
            $escapedHeader = [System.Security.SecurityElement]::Escape([string]$header)
            $headerXml += '<w:tc><w:p><w:pPr><w:pStyle w:val="TableHeading"/></w:pPr><w:r><w:t xml:space="preserve">' + $escapedHeader + '</w:t></w:r></w:p></w:tc>'
        }
        $headerXml += '</w:tr>'
        $markdownLines.Add($headerXml)

        $rowsToRender = @($Rows)
        if ($rowsToRender.Count -eq 0) {
            $naRow = @("N/A")
            for ($naIndex = 1; $naIndex -lt $columnCount; $naIndex++) {
                $naRow += "N/A"
            }
            $rowsToRender = ,$naRow
        }

        foreach ($rowObject in $rowsToRender) {
            $cells = @($rowObject)
            if ($cells.Count -lt $columnCount) {
                for ($padIndex = $cells.Count; $padIndex -lt $columnCount; $padIndex++) {
                    $cells += $null
                }
            }

            $firstCellRaw = [string]$cells[0]
            $isMergeRestart = -not [string]::IsNullOrWhiteSpace($firstCellRaw)

            $rowXml = '<w:tr>'
            if ($isMergeRestart) {
                $escapedFirst = [System.Security.SecurityElement]::Escape($firstCellRaw)
                $rowXml += '<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableText"/></w:pPr><w:r><w:t xml:space="preserve">' + $escapedFirst + '</w:t></w:r></w:p></w:tc>'
            }
            else {
                $rowXml += '<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableText"/></w:pPr></w:p></w:tc>'
            }

            for ($colIndex = 1; $colIndex -lt $columnCount; $colIndex++) {
                $cellText = [string]$cells[$colIndex]
                if ([string]::IsNullOrWhiteSpace($cellText)) {
                    $cellText = "N/A"
                }
                $escapedCellText = [System.Security.SecurityElement]::Escape($cellText)
                $rowXml += '<w:tc><w:p><w:pPr><w:pStyle w:val="TableText"/></w:pPr><w:r><w:t xml:space="preserve">' + $escapedCellText + '</w:t></w:r></w:p></w:tc>'
            }

            $rowXml += '</w:tr>'
            $markdownLines.Add($rowXml)
        }

        $markdownLines.Add('</w:tbl>')
        $markdownLines.Add('```')
    }

    $hostTypeByIndex = @{}
    foreach ($ht in $hostTypes) {
        $idx = Get-PropertyValue -Object $ht -PropertyName "index"
        if ($null -ne $idx) {
            $hostTypeByIndex[[string]$idx] = (& $nameOf $ht)
        }
    }

    $hostGroupById = @{}
    foreach ($hg in $hostGroups) {
        $id = Get-PropertyValue -Object $hg -PropertyName "id"
        if ($null -ne $id) {
            $hostGroupById[[string]$id] = $hg
        }
    }

    $hostById = @{}
    foreach ($h in $hosts) {
        $id = Get-PropertyValue -Object $h -PropertyName "id"
        if ($null -ne $id) {
            $hostById[[string]$id] = $h
        }
    }

    $volumeById = @{}
    foreach ($v in $volumes) {
        $id = Get-PropertyValue -Object $v -PropertyName "id"
        if ($null -ne $id) {
            $volumeById[[string]$id] = $v
        }
    }

    $poolById = @{}
    foreach ($p in $storagePools) {
        $id = Get-PropertyValue -Object $p -PropertyName "id"
        if ($null -ne $id) {
            $poolById[[string]$id] = $p
        }
    }

    $snapshotGroupById = @{}
    foreach ($sg in $snapshotGroups) {
        $id = Get-PropertyValue -Object $sg -PropertyName "id"
        if ($null -ne $id) {
            $snapshotGroupById[[string]$id] = $sg
        }
    }

    $snapshotConsistencyGroupById = @{}
    foreach ($cg in $snapshotConsistencyGroups) {
        $id = Get-PropertyValue -Object $cg -PropertyName "id"
        if ($null -ne $id) {
            $snapshotConsistencyGroupById[[string]$id] = $cg
        }
    }

    $trayByRef = @{}
    foreach ($t in $trays) {
        $ref = Get-PropertyValue -Object $t -PropertyName "trayRef"
        if ($null -ne $ref) {
            $trayByRef[[string]$ref] = $t
        }
    }

    $drawerByRef = @{}
    foreach ($d in $drawers) {
        $ref = Get-PropertyValue -Object $d -PropertyName "drawerRef"
        if ($null -ne $ref) {
            $drawerByRef[[string]$ref] = $d
        }
    }

    $santricityOs = "N/A"
    foreach ($code in $embeddedFwVersions) {
        if (([string](Get-PropertyValue -Object $code -PropertyName "codeModule")) -eq "bundleDisplay") {
            $candidate = Get-PropertyValue -Object $code -PropertyName "versionString"
            if (-not [string]::IsNullOrWhiteSpace([string]$candidate)) {
                $santricityOs = [string]$candidate
                break
            }
        }
    }

    $markdownLines = New-Object System.Collections.Generic.List[string]

    if ($IncludeTitlePage) {
        $reportDateDisplay = (Get-Date $Iso8601).ToString("dd/MM/yyyy")
        $systemNameForTitle = (& $nameOf $storageSystem)
        $titleBlockIndentTwips = 720

        $markdownLines.Add('```{=openxml}')
        foreach ($blankIndex in 1..11) {
            & $emitOpenXmlParagraph -StyleId "Normal" -Blank
        }
        & $emitOpenXmlFieldParagraph -StyleId "Eyebrow" -Label "" -Instruction 'DOCPROPERTY  Title  \* MERGEFORMAT' -FallbackText ([string]$reportMetadata.title) -LeftIndentTwips $titleBlockIndentTwips
        & $emitOpenXmlFieldParagraph -StyleId "Normal20Pt" -Label "" -Instruction 'DOCPROPERTY  Subject  \* MERGEFORMAT' -FallbackText ([string]$reportMetadata.subject) -LeftIndentTwips $titleBlockIndentTwips
        & $emitOpenXmlFieldParagraph -StyleId "Normal20PtBold" -Label "" -Instruction 'DOCPROPERTY  SystemName  \* MERGEFORMAT' -FallbackText $systemNameForTitle -LeftIndentTwips $titleBlockIndentTwips
        & $emitOpenXmlParagraph -StyleId "Normal" -Blank
        & $emitOpenXmlFieldBlockParagraph -StyleId "Normal" -LeftIndentTwips $titleBlockIndentTwips -Lines @(
            @{ Label = 'Prepared for: '; Instruction = 'DOCPROPERTY  CustomerName  \* MERGEFORMAT'; FallbackText = (& $titleField '<<CUSTOMER_NAME>>') },
            @{ Label = 'Location: '; Instruction = 'DOCPROPERTY  CustomerLocation  \* MERGEFORMAT'; FallbackText = (& $titleField '<<CUSTOMER_LOCATION>>') },
            @{ Label = 'Project: '; Instruction = 'DOCPROPERTY  ProjectName  \* MERGEFORMAT'; FallbackText = (& $titleField '<<PROJECT_NAME>>') },
            @{ Label = 'Author: '; Instruction = 'AUTHOR  \* MERGEFORMAT'; FallbackText = ([string]$reportMetadata.author) },
            @{ Label = 'Date: '; Instruction = 'DATE  \@ "dd/MM/yyyy"  \* MERGEFORMAT'; FallbackText = $reportDateDisplay }
        )
        foreach ($blankIndex in 1..9) {
            & $emitOpenXmlParagraph -StyleId "Normal" -Blank
        }
        & $emitOpenXmlParagraph -StyleId "Abstract" -Text "Abstract"
        & $emitOpenXmlParagraph -StyleId "BodyText" -Text "This document provides a comprehensive configuration summary of all the NetApp storage systems for which this report was generated."
        $markdownLines.Add('<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>')
        $markdownLines.Add('```')
        $markdownLines.Add("")

        if ($EnableDocxToc) {
            $markdownLines.Add('```{=openxml}')
            & $emitOpenXmlParagraph -StyleId "TOCHeading" -Text $reportMetadata["toc-title"]
            $markdownLines.Add('<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve">TOC \o "1-' + $DocxTocDepth + '" \h \z \u</w:instrText><w:fldChar w:fldCharType="separate"/><w:t xml:space="preserve">Right-click and select Update Field to generate the table of contents.</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>')
            $markdownLines.Add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
            $markdownLines.Add('```')
            $markdownLines.Add("")
        }
    }

    $markdownLines.Add("# Document Version Information")
    $markdownLines.Add("")
    $markdownLines.Add("The following table lists details about the data collection performed to facilitate creation of this As-Built document.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
        @("PowerShell Script Version", $ScriptVersion),
        @("Date Data Collected", $Iso8601),
        @("E-Series System API URL", $ApiUrl),
        @("Array SSID", $Ssid)
    )

    $markdownLines.Add("")
    $markdownLines.Add('```{=openxml}')
    $markdownLines.Add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    $markdownLines.Add('```')

    $markdownLines.Add("")
    $markdownLines.Add("# Overview")
    $markdownLines.Add("This report will provide the following information:")
    $markdownLines.Add("")
    $markdownLines.Add("- Array Hardware Inventory")
    $markdownLines.Add("  - Controllers and their Management and Host-Side Interfaces")
    $markdownLines.Add("  - Disk Shelves")
    $markdownLines.Add("  - IOM Modules")
    $markdownLines.Add("  - Drives")
    $markdownLines.Add("- Storage Pools")
    $markdownLines.Add("- Volumes")
    $markdownLines.Add("- Hosts, Host Groups, and Volume Mappings")
    $markdownLines.Add("- Volume Snapshots")
    $markdownLines.Add("- System Configuration Settings")
    $markdownLines.Add("  - DNS and NTP")
    $markdownLines.Add("  - General Array Settings")
    $markdownLines.Add("  - Alerts (SMTP / SNMP / Syslog)")
    $markdownLines.Add("  - LDAP Authentication")
    $markdownLines.Add("  - Audit Logging")
    $markdownLines.Add("  - AutoSupport")

    $markdownLines.Add("")
    $markdownLines.Add("# System Summary")
    $markdownLines.Add("")
    $markdownLines.Add("NetApp E-Series Storage Arrays include shelves, controllers, drives, software, and firmware. An array can be installed in a rack or cabinet, with customizable hardware for one or two controllers, in a 12-, 24-, or 60-drive shelf. You can connect the storage array to a SAN from multiple interface types and present block storage devices to a variety of host operating systems.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides a summary of the E-Series Storage System reviewed in this report. Further detailed information is provided in the subsequent sections.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("System Summary Item", "System Summary Detail") -Rows @(
        @("Name", (& $nameOf $storageSystem)),
        @("Chassis Serial", (Get-PropertyValue -Object $storageSystem -PropertyName "chassisSerialNumber")),
        @("SANtricity OS", $santricityOs),
        @("Firmware", (Get-PropertyValue -Object $storageSystem -PropertyName "fwVersion")),
        @("WWN/WWID", (Get-PropertyValue -Object $storageSystem -PropertyName "wwn")),
        @("Platform Model", (Get-PropertyValue -Object $storageSystem -PropertyName "model")),
        @("NVSRAM Version", (Get-PropertyValue -Object $storageSystem -PropertyName "nvsramVersion")),
        @("Simplex Mode", (& $boolText (Get-PropertyValue -Object $storageSystem -PropertyName "simplexModeEnabled"))),
        @("Current Status", (Get-PropertyValue -Object $storageSystem -PropertyName "status")),
        @("Recovery Mode Enabled", (& $boolText (Get-PropertyValue -Object $storageSystem -PropertyName "recoveryModeEnabled"))),
        @("Controllers Detected", $controllers.Count),
        @("Storage Pools", $storagePools.Count),
        @("Volumes", $volumes.Count),
        @("Disks", $drives.Count),
        @("Hosts", $hosts.Count),
        @("Host Groups", $hostGroups.Count)
    )

    $markdownLines.Add("")
    $markdownLines.Add("# Hardware Inventory")
    $markdownLines.Add("")
    $markdownLines.Add("## Controllers")
    $markdownLines.Add("An E-Series controller is a hardware compute module that runs the SANtricity operating system. It controls drives, enclosures, and storage access, and also implements management software functions. Controllers include network interfaces for management and host-side connectivity.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides a summary of the controllers in the E-Series Storage System.")
    $markdownLines.Add("")
    $controllerRows = @()
    foreach ($c in $controllers) {
        $loc = Get-PropertyValue -Object (Get-PropertyValue -Object $c -PropertyName "physicalLocation") -PropertyName "label"
        $controllerRows += [pscustomobject]@{
            Controller = [string]$loc
            Serial     = (Get-PropertyValue -Object $c -PropertyName "serialNumber")
            Model      = (Get-PropertyValue -Object $c -PropertyName "modelName")
            Status     = (Get-PropertyValue -Object $c -PropertyName "status")
            Active     = (& $boolText (Get-PropertyValue -Object $c -PropertyName "active"))
        }
    }
    $controllerRows = @($controllerRows | Sort-Object Controller)
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Controller", "Serial Number", "Model", "Status", "Active") -Rows @(
        $controllerRows | ForEach-Object { ,@($_.Controller, $_.Serial, $_.Model, $_.Status, $_.Active) }
    )

    $markdownLines.Add("")
    $markdownLines.Add("## Management Interfaces")
    $markdownLines.Add("The management interface is the primary interface for managing the E-Series Storage System. SANtricity System Manager is web-based management software embedded on each controller and is accessible via HTTPS to the controller IP address.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the management interfaces in the E-Series Storage System.")
    $markdownLines.Add("")
    $mgmtRows = @()
    foreach ($c in $controllers) {
        $ctrlLabel = Get-PropertyValue -Object (Get-PropertyValue -Object $c -PropertyName "physicalLocation") -PropertyName "label"
        $sshEnabled = & $boolText (Get-PropertyValue -Object (Get-PropertyValue -Object $c -PropertyName "networkSettings") -PropertyName "remoteAccessEnabled")
        foreach ($iface in @((Get-PropertyValue -Object $c -PropertyName "netInterfaces"))) {
            if (([string](Get-PropertyValue -Object $iface -PropertyName "interfaceType")).ToLowerInvariant() -ne "ethernet") { continue }
            $eth = Get-PropertyValue -Object $iface -PropertyName "ethernet"
            if ($null -eq $eth) { continue }
            $mgmtRows += [pscustomobject]@{
                Controller = [string]$ctrlLabel
                Name       = [string](Get-PropertyValue -Object $eth -PropertyName "interfaceName")
                Port       = (Get-PropertyValue -Object (Get-PropertyValue -Object $eth -PropertyName "physicalLocation") -PropertyName "label")
                LinkStatus = (Get-PropertyValue -Object $eth -PropertyName "linkStatus")
                IPv4       = (Get-PropertyValue -Object $eth -PropertyName "ipv4Address")
                Speed      = (Get-PropertyValue -Object $eth -PropertyName "speed")
                MAC        = (Get-PropertyValue -Object $eth -PropertyName "macAddr")
                SSH        = $sshEnabled
            }
        }
    }
    $mgmtRows = @($mgmtRows | Sort-Object Controller, Name)
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Controller", "Port", "Link Status", "IPv4", "Speed", "MAC", "SSH") -Rows @(
        $mgmtRows | ForEach-Object { ,@($_.Controller, $_.Port, $_.LinkStatus, $_.IPv4, $_.Speed, $_.MAC, $_.SSH) }
    )

    $markdownLines.Add("")
    $markdownLines.Add("## Host-Side Interfaces")
    $markdownLines.Add("Host-side interfaces connect the E-Series Storage System to host systems. These interfaces can be FC, iSCSI, NVMe, SAS, or InfiniBand and provide host connectivity to storage volumes.")
    $markdownLines.Add("")
    $markdownLines.Add("The following tables provide details of the host-side interfaces in the E-Series Storage System.")
    $markdownLines.Add("")
    $hostIfRows = @()
    $fcRows = @()
    $iscsiRows = @()
    $iscsiIqns = New-Object System.Collections.Generic.HashSet[string]
    foreach ($c in $controllers) {
        $ctrlLabel = Get-PropertyValue -Object (Get-PropertyValue -Object $c -PropertyName "physicalLocation") -PropertyName "label"
        foreach ($hi in @((Get-PropertyValue -Object $c -PropertyName "hostInterfaces"))) {
            $ifType = ([string](Get-PropertyValue -Object $hi -PropertyName "interfaceType")).ToLowerInvariant()
            if ($ifType -eq "fc") {
                $fibre = Get-PropertyValue -Object $hi -PropertyName "fibre"
                $channel = Get-PropertyValue -Object $fibre -PropertyName "channel"
                $hostIfRows += [pscustomobject]@{ Controller=$ctrlLabel; Type="FC"; Channel=$channel; LinkStatus=(Get-PropertyValue -Object $fibre -PropertyName "linkStatus"); CurrentSpeed=(Get-PropertyValue -Object $fibre -PropertyName "currentInterfaceSpeed"); MaxSpeed=(Get-PropertyValue -Object $fibre -PropertyName "maximumInterfaceSpeed") }
                $fcRows += [pscustomobject]@{ Controller=$ctrlLabel; Channel=$channel; WWPN=(Get-PropertyValue -Object $fibre -PropertyName "niceAddressId"); Topology=(Get-PropertyValue -Object $fibre -PropertyName "topology") }
            }
            elseif ($ifType -eq "iscsi") {
                $iscsi = Get-PropertyValue -Object $hi -PropertyName "iscsi"
                $ethData = Get-PropertyValue -Object (Get-PropertyValue -Object $iscsi -PropertyName "interfaceData") -PropertyName "ethernetData"
                $ipv4Data = Get-PropertyValue -Object $iscsi -PropertyName "ipv4Data"
                $channel = Get-PropertyValue -Object $iscsi -PropertyName "channel"
                $hostIfRows += [pscustomobject]@{ Controller=$ctrlLabel; Type="iSCSI"; Channel=$channel; LinkStatus=(Get-PropertyValue -Object $ethData -PropertyName "linkStatus"); CurrentSpeed=(Get-PropertyValue -Object $ethData -PropertyName "currentInterfaceSpeed"); MaxSpeed=(Get-PropertyValue -Object $ethData -PropertyName "maximumInterfaceSpeed") }
                $iscsiRows += [pscustomobject]@{ Controller=$ctrlLabel; Channel=$channel; IPv4=(Get-PropertyValue -Object $ipv4Data -PropertyName "ipv4Address"); MTU=(Get-PropertyValue -Object $ethData -PropertyName "maximumFramePayloadSize"); TCPPort=(Get-PropertyValue -Object $iscsi -PropertyName "tcpListenPort"); FEC=(Get-PropertyValue -Object $ethData -PropertyName "configuredFECMode"); VlanId=(Get-PropertyValue -Object (Get-PropertyValue -Object $ipv4Data -PropertyName "ipv4VlanId") -PropertyName "value") }
                $iqn = Get-PropertyValue -Object $iscsi -PropertyName "iqn"
                if (-not [string]::IsNullOrWhiteSpace([string]$iqn)) {
                    $iscsiIqns.Add([string]$iqn) | Out-Null
                }
            }
        }
    }
    $hostIfRows = @($hostIfRows | Sort-Object Controller, Channel)
    $fcRows = @($fcRows | Sort-Object Controller, Channel)
    $iscsiRows = @($iscsiRows | Sort-Object Controller, Channel)
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Controller", "Type", "Channel", "Link Status", "Current Speed", "Max Speed") -Rows @(
        $hostIfRows | ForEach-Object { ,@($_.Controller, $_.Type, $_.Channel, $_.LinkStatus, $_.CurrentSpeed, $_.MaxSpeed) }
    )

    $markdownLines.Add("")
    $markdownLines.Add("### FC Interface Details")
    $markdownLines.Add("The following table provides details of the Fibre Channel (FC) host-side interfaces in the E-Series Storage System.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Controller", "Channel", "WWPN", "Topology") -Rows @(
        $fcRows | ForEach-Object { ,@($_.Controller, $_.Channel, $_.WWPN, $_.Topology) }
    )

    $markdownLines.Add("")
    $markdownLines.Add("### iSCSI Interface Details")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the iSCSI host-side interfaces in the E-Series Storage System.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Controller", "Channel", "IPv4", "MTU", "TCP Port", "FEC", "VLAN ID") -Rows @(
        $iscsiRows | ForEach-Object { ,@($_.Controller, $_.Channel, $_.IPv4, $_.MTU, $_.TCPPort, $_.FEC, $_.VlanId) }
    )

    $markdownLines.Add("")
    $markdownLines.Add("### iSCSI Target IQNs")
    $markdownLines.Add("")
    $markdownLines.Add("The iSCSI Target IQN (Internet Qualified Name) is a unique iSCSI identifier for the E-Series Storage System. It is used by iSCSI initiators to connect to the storage system and access the storage volumes.")
    $markdownLines.Add("")
    $iqnRows = @()
    foreach ($iqnItem in @($iscsiIqns | Sort-Object)) {
        $iqnRows += ,@($iqnItem)
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Target IQN") -Rows $iqnRows

    $markdownLines.Add("")
    $markdownLines.Add("## Shelves")
    $markdownLines.Add("The shelves in the E-Series Storage System are the physical enclosures that house the controllers and the drives. A shelf may support a single row of drives, or can contain multiple trays, and each tray can contain multiple drives. Shelf drive capacities can range from as few as 12 drives up to 60 drives in a single shelf.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the shelves in the E-Series Storage System.")
    $markdownLines.Add("")
    $shelfRows = @()
    foreach ($tray in $trays) {
        $trayRef = Get-PropertyValue -Object $tray -PropertyName "trayRef"
        $populated = 0
        foreach ($drive in $drives) {
            $dTrayRef = Get-PropertyValue -Object (Get-PropertyValue -Object $drive -PropertyName "physicalLocation") -PropertyName "trayRef"
            if ([string]$dTrayRef -eq [string]$trayRef) {
                $populated += 1
            }
        }
        $trayId = Get-PropertyValue -Object $tray -PropertyName "trayId"
        $trayIdNum = 2147483647
        if (-not [string]::IsNullOrWhiteSpace([string]$trayId)) {
            $parsedTray = 0
            if ([int]::TryParse([string]$trayId, [ref]$parsedTray)) {
                $trayIdNum = $parsedTray
            }
        }
        $shelfRows += [pscustomobject]@{
            ShelfId       = (Format-ShelfDisplayValue -ShelfValue $trayId)
            Model         = (Get-PropertyValue -Object $tray -PropertyName "partNumber")
            NumDrawers    = (Get-PropertyValue -Object $tray -PropertyName "numDrawers")
            NumDriveSlots = (Get-PropertyValue -Object $tray -PropertyName "numDriveSlots")
            Populated     = $populated
            SortGroup     = $(if ($trayIdNum -eq 99) { 0 } else { 1 })
            SortNum       = $trayIdNum
        }
    }
    $shelfRows = @($shelfRows | Sort-Object SortGroup, SortNum)
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Shelf", "Model Number", "Number of Drawers", "Total Drive Slots", "Populated Slots") -Rows @(
        $shelfRows | ForEach-Object { ,@($_.ShelfId, $_.Model, $_.NumDrawers, $_.NumDriveSlots, $_.Populated) }
    )

    $markdownLines.Add("")
    $markdownLines.Add("## IOM / ESM Modules")
    $markdownLines.Add("The IOM (Input/Output Module) or ESM (Enclosure Services Module) is a hardware component that provides connectivity and management for the drives in the E-Series Storage System. It is responsible for managing the communication between the drives and the controllers, as well as providing power and cooling to the drives.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the IOM / ESM modules in the E-Series Storage System.")
    $markdownLines.Add("")
    $esmRows = @()
    foreach ($esm in $esms) {
        $esmPhysical = Get-PropertyValue -Object $esm -PropertyName "physicalLocation"
        $esmTrayRef = Get-PropertyValue -Object $esmPhysical -PropertyName "trayRef"
        $esmTray = $trayByRef[[string]$esmTrayRef]
        $trayId = Get-PropertyValue -Object $esmTray -PropertyName "trayId"
        $trayNum = 2147483647
        if (-not [string]::IsNullOrWhiteSpace([string]$trayId)) {
            $parsed = 0
            if ([int]::TryParse([string]$trayId, [ref]$parsed)) {
                $trayNum = $parsed
            }
        }
        $label = [string](Get-PropertyValue -Object $esmPhysical -PropertyName "label")
        $esmRows += [pscustomobject]@{
            Shelf    = (Format-ShelfDisplayValue -ShelfValue $trayId)
            Label    = $label
            FType    = (Get-PropertyValue -Object $esm -PropertyName "fruType")
            Part     = (Get-PropertyValue -Object $esm -PropertyName "partNumber")
            Serial   = (Get-PropertyValue -Object $esm -PropertyName "serialNumber")
            Firmware = (Get-PropertyValue -Object $esm -PropertyName "softwareVersion")
            Status   = (Get-PropertyValue -Object $esm -PropertyName "status")
            SortGroup = $(if ($trayNum -eq 99) { 0 } else { 1 })
            SortNum  = $trayNum
        }
    }
    $esmRows = @($esmRows | Sort-Object SortGroup, SortNum, Label)
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Shelf", "Module", "FRU Type", "Part Number", "Serial Number", "Firmware", "Status") -Rows @(
        $esmRows | ForEach-Object { ,@($_.Shelf, $_.Label, $_.FType, $_.Part, $_.Serial, $_.Firmware, $_.Status) }
    )

    $markdownLines.Add("")
    $markdownLines.Add("## Drives")
    $markdownLines.Add("NetApp E-Series Arrays support a variety of drive types, including SATA, SSD, and NVMe drives. The drives are installed in the shelves and are managed by the controllers. Each drive has a unique serial number and can be identified by its location in the shelf.")
    $markdownLines.Add("")
    $markdownLines.Add("Drives are hot-swappable, meaning they can be replaced without shutting down the system, and drive firmware can be updated to improve performance and reliability. The drives are also monitored for health and status, and any issues are reported to the management software.")
    $markdownLines.Add("")
    $markdownLines.Add("### Drive Assignments")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provide a summary of the drive assignments in the E-Series Storage System. Unassigned drives are available for allocation to Volume Groups or Dynamic Disk Pools.")
    $markdownLines.Add("")
    $markdownLines.Add("Global Hot Spares are drives manually assigned as a standby for additional data protection in RAID 1, RAID 5, or RAID 6 volume groups. If a drive fails in one of these volume groups, the controller reconstructs data from the failed drive to the hot spare.")
    $markdownLines.Add("")
    $totalDriveCount = $drives.Count
    $hotSpare = 0
    $assigned = 0
    $unassigned = 0
    foreach ($d in $drives) {
        $isHotSpare = [string](Get-PropertyValue -Object $d -PropertyName "hotSpare")
        $poolRef = [string](Get-PropertyValue -Object $d -PropertyName "currentVolumeGroupRef")
        $isAssigned = (-not [string]::IsNullOrWhiteSpace($poolRef)) -and ($poolRef -notmatch "^0+$")
        if ($isHotSpare -eq "True" -and -not $isAssigned) { $hotSpare += 1 }
        elseif ($isAssigned) { $assigned += 1 }
        else { $unassigned += 1 }
    }
    $driveAssignmentRows = ,@($totalDriveCount, $hotSpare, $assigned, $unassigned)
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Total Drive Count", "Global Hot Spares", "Assigned Drives", "Unassigned Drives") -Rows $driveAssignmentRows

    $markdownLines.Add("")
    $markdownLines.Add("### Drive Details")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the drives in the E-Series Storage System.")
    $markdownLines.Add("")
    $driveRows = @()
    foreach ($d in $drives) {
        $physical = Get-PropertyValue -Object $d -PropertyName "physicalLocation"
        $trayRef = Get-PropertyValue -Object $physical -PropertyName "trayRef"
        $drawerRef = Get-PropertyValue -Object $physical -PropertyName "drawerRef"
        if ([string]::IsNullOrWhiteSpace([string]$drawerRef)) {
            $locationParent = Get-PropertyValue -Object $physical -PropertyName "locationParent"
            $typedRef = Get-PropertyValue -Object $locationParent -PropertyName "typedReference"
            $componentType = [string](Get-PropertyValue -Object $typedRef -PropertyName "componentType")
            if ($componentType.ToLowerInvariant() -eq "drawer") {
                $drawerRef = Get-PropertyValue -Object $typedRef -PropertyName "symbolRef"
            }
        }
        $tray = $trayByRef[[string]$trayRef]
        $drawer = $drawerByRef[[string]$drawerRef]
        $drawerLabel = Get-PropertyValue -Object $drawer -PropertyName "label"
        if ([string]::IsNullOrWhiteSpace([string]$drawerLabel)) {
            $drawerLabel = Get-PropertyValue -Object (Get-PropertyValue -Object $drawer -PropertyName "physicalLocation") -PropertyName "label"
        }
        $shelfId = Get-PropertyValue -Object $tray -PropertyName "trayId"
        $bayLabel = Get-PropertyValue -Object $physical -PropertyName "label"
        $trayNum = 2147483647
        $drawerNum = 2147483647
        $bayNum = 2147483647
        $tmp = 0
        if ([int]::TryParse([string]$shelfId, [ref]$tmp)) { $trayNum = $tmp }
        if ([int]::TryParse([string]$drawerLabel, [ref]$tmp)) { $drawerNum = $tmp }
        if ([int]::TryParse([string]$bayLabel, [ref]$tmp)) { $bayNum = $tmp }
        $driveRows += [pscustomobject]@{
            Shelf    = (Format-ShelfDisplayValue -ShelfValue $shelfId)
            Drawer   = $drawerLabel
            Bay      = $bayLabel
            Capacity = (Convert-BytesToTiB -Bytes (Get-PropertyValue -Object $d -PropertyName "usableCapacity"))
            Type     = (Get-PropertyValue -Object $d -PropertyName "driveMediaType")
            FW       = (Get-PropertyValue -Object $d -PropertyName "firmwareVersion")
            Product  = (Get-PropertyValue -Object $d -PropertyName "productID")
            Serial   = (Get-PropertyValue -Object $d -PropertyName "serialNumber")
            Status   = (Get-PropertyValue -Object $d -PropertyName "status")
            SortGroup = $(if ($trayNum -eq 99) { 0 } else { 1 })
            SortTray = $trayNum
            SortDrawer = $drawerNum
            SortBay  = $bayNum
        }
    }
    $driveRows = @($driveRows | Sort-Object SortGroup, SortTray, SortDrawer, SortBay)
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Shelf", "Drawer", "Bay", "Capacity (TiB)", "Type", "FW", "Product ID", "Serial Number", "Status") -Rows @(
        $driveRows | ForEach-Object { ,@($_.Shelf, $_.Drawer, $_.Bay, $_.Capacity, $_.Type, $_.FW, $_.Product, $_.Serial, $_.Status) }
    )

    $markdownLines.Add("")
    $markdownLines.Add("# Storage Configuration and Presentation")
    $markdownLines.Add("")
    $markdownLines.Add("")
    $markdownLines.Add("## Storage Pools")
    $markdownLines.Add("")
    $markdownLines.Add("Storage Pools are collections of physical drives that are grouped together to provide storage capacity for volumes. Storage pools can be configured as Dynamic Disk Pools (DDP) or as Volume Groups.")
    $markdownLines.Add("")
    $markdownLines.Add("DDP uses RAID6 and distributes data across a pool of drives to provide automated data layout, fast background reconstruction, and simplified capacity management.")
    $markdownLines.Add("")
    $markdownLines.Add("Volume Groups use traditional RAID levels (RAID 0, 1, 5, 6, 10) to provide data protection and performance. Volume Groups can be configured with different RAID levels to meet specific performance and capacity requirements.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the storage pools in the E-Series Storage System.")
    $markdownLines.Add("")
    $poolRows = @()
    foreach ($pool in $storagePools) {
        $poolType = if (([string](Get-PropertyValue -Object $pool -PropertyName "raidLevel")) -eq "raidDiskPool") { "DDP" } else { "VolGroup" }
        $poolRows += ,@(
            (& $nameOf $pool),
            $poolType,
            (Convert-BytesToTiB -Bytes (Get-PropertyValue -Object $pool -PropertyName "totalRaidedSpace")),
            (Convert-BytesToTiB -Bytes (Get-PropertyValue -Object $pool -PropertyName "usedSpace")),
            (Convert-BytesToTiB -Bytes (Get-PropertyValue -Object $pool -PropertyName "freeSpace"))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Name", "Type", "Total Capacity (TiB)", "Used Capacity (TiB)", "Free Capacity (TiB)") -Rows $poolRows

    $markdownLines.Add("")
    $markdownLines.Add("### Dynamic Disk Pools")
    $markdownLines.Add("")
    $markdownLines.Add("This section provides details of each Dynamic Disk Pool (DDP) in the storage system. DDP provides simplified capacity management, fast background reconstruction, and automated data layout across a pool of drives, with built-in data protection and preservation capacity.")
    foreach ($pool in $storagePools) {
        if (([string](Get-PropertyValue -Object $pool -PropertyName "raidLevel")) -ne "raidDiskPool") { continue }
        $markdownLines.Add("")
        $markdownLines.Add("#### DDP Details - $(& $nameOf $pool)")
        $markdownLines.Add("")
        $dppData = Get-PropertyValue -Object (Get-PropertyValue -Object $pool -PropertyName "volumeGroupData") -PropertyName "diskPoolData"
        $extents = @((Get-PropertyValue -Object $pool -PropertyName "extents"))
        $dppRaidCaps = @((Get-PropertyValue -Object ($extents | Select-Object -First 1) -PropertyName "ddpRAIDCapacities"))
        $dppRaidLevel = Get-PropertyValue -Object ($dppRaidCaps | Select-Object -First 1) -PropertyName "ddpVolRAIDLevel"
        $poolId = [string](Get-PropertyValue -Object $pool -PropertyName "id")
        $dppDriveCount = 0
        foreach ($dppDrive in $drives) {
            if ([string](Get-PropertyValue -Object $dppDrive -PropertyName "currentVolumeGroupRef") -eq $poolId) {
                $dppDriveCount += 1
            }
        }
        $preservationCount = Get-PropertyValue -Object $dppData -PropertyName "reconstructionReservedDriveCount"
        $preservationCapacity = if ($null -ne $preservationCount -and -not [string]::IsNullOrWhiteSpace([string]$preservationCount)) { "$preservationCount Drives" } else { "N/A" }
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
            @("Name", (& $nameOf $pool)),
            @("RAID Status", (Get-PropertyValue -Object $pool -PropertyName "raidStatus")),
            @("DDP RAID Level", $dppRaidLevel),
            @("Drive Count", $dppDriveCount),
            @("Total Capacity (TiB)", (Convert-BytesToTiB -Bytes (Get-PropertyValue -Object $pool -PropertyName "totalRaidedSpace"))),
            @("Used Capacity (TiB)", (Convert-BytesToTiB -Bytes (Get-PropertyValue -Object $pool -PropertyName "usedSpace"))),
            @("Free Capacity (TiB)", (Convert-BytesToTiB -Bytes (Get-PropertyValue -Object $pool -PropertyName "freeSpace"))),
            @("Utilization Critical Threshold (%)", (Get-PropertyValue -Object $dppData -PropertyName "poolUtilizationCriticalThreshold")),
            @("Critical Reconstruction Priority", (Get-PropertyValue -Object $dppData -PropertyName "criticalReconstructPriority")),
            @("Degraded Reconstruction Priority", (Get-PropertyValue -Object $dppData -PropertyName "degradedReconstructPriority")),
            @("Background Operation Priority", (Get-PropertyValue -Object $dppData -PropertyName "backgroundOperationPriority")),
            @("Preservation Capacity", $preservationCapacity),
            @("Drawer Loss Protection", (& $boolText (Get-PropertyValue -Object $pool -PropertyName "drawerLossProtection"))),
            @("Tray Loss Protection", (& $boolText (Get-PropertyValue -Object $pool -PropertyName "trayLossProtection"))),
            @("Security Type", (Get-PropertyValue -Object $pool -PropertyName "securityType"))
        )
    }

    $markdownLines.Add("")
    $markdownLines.Add("### Volume Groups")
    $markdownLines.Add("")
    $markdownLines.Add("This section provides details of each Volume Group in the storage system. Volume Groups are logical collections of physical drives that are configured with a specific RAID level to provide data protection and performance. Global Hot Spares can be assigned to Volume Groups to provide automatic drive replacement in the event of a drive failure.")
    foreach ($pool in $storagePools) {
        if (([string](Get-PropertyValue -Object $pool -PropertyName "raidLevel")) -eq "raidDiskPool") { continue }
        $markdownLines.Add("")
        $markdownLines.Add("#### Volume Group Details - $(& $nameOf $pool)")
        $markdownLines.Add("")
        $poolId = [string](Get-PropertyValue -Object $pool -PropertyName "id")
        $vgDriveCount = 0
        foreach ($vgDrive in $drives) {
            if ([string](Get-PropertyValue -Object $vgDrive -PropertyName "currentVolumeGroupRef") -eq $poolId) {
                $vgDriveCount += 1
            }
        }
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
            @("Name", (& $nameOf $pool)),
            @("RAID Level", (Get-PropertyValue -Object $pool -PropertyName "raidLevel")),
            @("RAID Status", (Get-PropertyValue -Object $pool -PropertyName "raidStatus")),
            @("Drive Count", $vgDriveCount),
            @("Total Capacity (TiB)", (Convert-BytesToTiB -Bytes (Get-PropertyValue -Object $pool -PropertyName "totalRaidedSpace"))),
            @("Used Capacity (TiB)", (Convert-BytesToTiB -Bytes (Get-PropertyValue -Object $pool -PropertyName "usedSpace"))),
            @("Free Capacity (TiB)", (Convert-BytesToTiB -Bytes (Get-PropertyValue -Object $pool -PropertyName "freeSpace"))),
            @("Drive Media Type", (Get-PropertyValue -Object $pool -PropertyName "driveMediaType")),
            @("Drive Physical Type", (Get-PropertyValue -Object $pool -PropertyName "drivePhysicalType")),
            @("Spindle Speed", (Get-PropertyValue -Object $pool -PropertyName "spindleSpeed")),
            @("Drawer Loss Protection", (& $boolText (Get-PropertyValue -Object $pool -PropertyName "drawerLossProtection"))),
            @("Tray Loss Protection", (& $boolText (Get-PropertyValue -Object $pool -PropertyName "trayLossProtection"))),
            @("Security Type", (Get-PropertyValue -Object $pool -PropertyName "securityType"))
        )
    }

    $markdownLines.Add("")
    $markdownLines.Add("## Volumes")
    $markdownLines.Add("E-Series volumes are logical storage units created from storage pools or volume groups. Volumes are mapped to LUNs or NSIDs and are visible to host systems over one or more supported host access protocols.")
    $markdownLines.Add("")
    $volumeRows = @()
    foreach ($v in $volumes) {
        $parentRef = Get-PropertyValue -Object $v -PropertyName "volumeGroupRef"
        if ($null -eq $parentRef) { $parentRef = Get-PropertyValue -Object $v -PropertyName "parentPoolId" }
        if ($null -eq $parentRef) { $parentRef = Get-PropertyValue -Object $v -PropertyName "parentRef" }
        $parentPool = $poolById[[string]$parentRef]
        $isThin = Get-PropertyValue -Object $v -PropertyName "thinProvisioned"
        if ($null -eq $isThin) { $isThin = Get-PropertyValue -Object $v -PropertyName "isThinProvisioned" }
        $thinText = "N/A"
        if ($null -ne $isThin) {
            if ([string]$isThin -eq "True") { $thinText = "thin" } elseif ([string]$isThin -eq "False") { $thinText = "thick" }
        }
        $segmentSizeBytes = Get-PropertyValue -Object $v -PropertyName "segmentSize"
        $segmentSizeKbText = "N/A"
        if ($null -ne $segmentSizeBytes) {
            $segmentSizeBytesNumber = 0
            if ([long]::TryParse([string]$segmentSizeBytes, [ref]$segmentSizeBytesNumber)) {
                $segmentSizeKbText = "{0}KB" -f [math]::Round(($segmentSizeBytesNumber / 1024), 0)
            }
        }
        $volumeRows += ,@(
            (& $nameOf $v),
            (Convert-BytesToTiB -Bytes (Get-PropertyValue -Object $v -PropertyName "capacity")),
            (Get-PropertyValue -Object $v -PropertyName "raidLevel"),
            $thinText,
            $segmentSizeKbText,
            (& $nameOf $parentPool)
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Name", "Capacity (TiB)", "RAID", "Thick/Thin", "Segment Size", "Storage Pool") -Rows $volumeRows

    $markdownLines.Add("")
    $markdownLines.Add("## Host Groups")
    $markdownLines.Add("Host Groups are collections of hosts that can be used to manage access to volumes. Host groups can be used to simplify the management of host access to volumes, allowing volumes to be mapped to a group and access to be controlled at the group level.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the host groups in the E-Series Storage System.")
    $markdownLines.Add("")
    $hostGroupRows = @()
    foreach ($hg in $hostGroups) {
        $groupId = Get-PropertyValue -Object $hg -PropertyName "id"
        $groupHosts = New-Object System.Collections.Generic.List[string]
        foreach ($h in $hosts) {
            $grpRef = Get-PropertyValue -Object $h -PropertyName "groupRef"
            if ($null -eq $grpRef) { $grpRef = Get-PropertyValue -Object $h -PropertyName "clusterRef" }
            if ([string]$grpRef -eq [string]$groupId) {
                $groupHosts.Add((& $nameOf $h))
            }
        }
        if ($groupHosts.Count -gt 0) {
            $sortedGroupHosts = @($groupHosts | Sort-Object)
            for ($hostIndex = 0; $hostIndex -lt $sortedGroupHosts.Count; $hostIndex++) {
                $hostGroupRows += ,@($(if ($hostIndex -eq 0) { (& $nameOf $hg) } else { "" }), [string]$sortedGroupHosts[$hostIndex])
            }
        }
        else {
            $hostGroupRows += ,@((& $nameOf $hg), "N/A")
        }
    }
    & $emitOpenXmlMergedFirstColumnTable -Headers @("Name", "Hosts") -Rows $hostGroupRows -GridWidths @(3120, 4680)

    $markdownLines.Add("")
    $markdownLines.Add("Hosts are the systems that connect to the E-Series Storage System to access the volumes. Hosts can be physical or virtual machines, and they can run a variety of operating systems. Hosts are identified by their host type, which can be Windows, Linux, VMware, or other supported operating systems.")
    $markdownLines.Add("")
    $markdownLines.Add("Hosts that connect to E-Series Storage Systems typically use multipathing software to provide redundancy and load balancing for the storage connections. Multipathing software allows hosts to use multiple paths to access the same volume, providing failover in case of a path failure and improving performance by distributing I/O across multiple paths.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the hosts in the E-Series Storage System.")
    $markdownLines.Add("")
    $markdownLines.Add("## Hosts")
    $hostRows = @()
    foreach ($h in $hosts) {
        $hostTypeIndex = Get-PropertyValue -Object $h -PropertyName "hostTypeIndex"
        if ($null -eq $hostTypeIndex) { $hostTypeIndex = Get-PropertyValue -Object $h -PropertyName "host_type_index" }
        $grpRef = Get-PropertyValue -Object $h -PropertyName "groupRef"
        if ($null -eq $grpRef) { $grpRef = Get-PropertyValue -Object $h -PropertyName "clusterRef" }
        $hostRows += ,@((& $nameOf $h), $hostTypeByIndex[[string]$hostTypeIndex], (& $nameOf $hostGroupById[[string]$grpRef]))
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Name", "Host Type", "Group Name") -Rows $hostRows

    $markdownLines.Add("")
    $markdownLines.Add("## Volume Mappings")
    $markdownLines.Add("Volume mappings link volumes to hosts or host groups, allowing the hosts to access the storage volumes. For most host operating systems, the volume mappings are presented as logical unit numbers (LUNs), however for NVMe this is presented as Namespace IDs (NSIDs). The host operating system uses the LUN or NSID to access the volume and read/write data.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the volume mappings in the E-Series Storage System.")
    $markdownLines.Add("")
    $mappingRows = @()
    foreach ($v in $volumes) {
        foreach ($m in @((Get-PropertyValue -Object $v -PropertyName "listOfMappings"))) {
            $mapRef = Get-PropertyValue -Object $m -PropertyName "mapRef"
            $mapType = ([string](Get-PropertyValue -Object $m -PropertyName "type")).ToLowerInvariant()
            $targetName = "N/A"
            if ($mapType -eq "cluster") {
                $targetName = (& $nameOf $hostGroupById[[string]$mapRef])
            }
            else {
                $targetName = (& $nameOf $hostById[[string]$mapRef])
            }
            $mappingRows += [pscustomobject]@{
                Target = [string]$targetName
                Volume = [string](& $nameOf $v)
                LUN    = (Get-PropertyValue -Object $m -PropertyName "lun")
            }
        }
    }
    $mappingRows = @($mappingRows | Sort-Object Target, LUN, Volume)
    $mappingDisplayRows = @()
    $lastTarget = ""
    foreach ($mappingRow in $mappingRows) {
        $isFirstForTarget = ([string]$mappingRow.Target -ne [string]$lastTarget)
        $mappingDisplayRows += [pscustomobject]@{
            Target = $(if ($isFirstForTarget) { [string]$mappingRow.Target } else { "" })
            Volume = [string]$mappingRow.Volume
            LUN    = $mappingRow.LUN
        }
        $lastTarget = [string]$mappingRow.Target
    }
    $mappingMergedRows = @()
    foreach ($mappingDisplayRow in $mappingDisplayRows) {
        $mappingMergedRows += ,@($mappingDisplayRow.Target, $mappingDisplayRow.Volume, $mappingDisplayRow.LUN)
    }
    & $emitOpenXmlMergedFirstColumnTable -Headers @("Host or Host Group", "Volume Name", "LUN ID") -Rows $mappingMergedRows -GridWidths @(3120, 4680, 1080)

    $markdownLines.Add("")
    $markdownLines.Add("## Snapshot Images")
    $markdownLines.Add("A snapshot image is a logical copy of volume data captured at a point in time. Snapshot images allow rollback to a known-good data set.")
    $markdownLines.Add("")
    $markdownLines.Add("The snapshots feature uses copy-on-write (COW) technology. After the snapshot is taken, the first write to modified blocks copies original data to reserved capacity before the new write is committed.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides snapshot image details and cross-references each snapshot back to its snapshot group and source volume.")
    $markdownLines.Add("")
    $markdownLines.Add("Total snapshot images: $($snapshotImages.Count)")
    $markdownLines.Add("")
    $snapshotImageRows = @()
    foreach ($s in $snapshotImages) {
        $group = $snapshotGroupById[[string](Get-PropertyValue -Object $s -PropertyName "pitGroupRef")]
        $baseVol = $volumeById[[string](Get-PropertyValue -Object $s -PropertyName "baseVol")]
        $snapshotImageRows += ,@(
            (& $nameOf $baseVol),
            (& $nameOf $group),
            (Get-PropertyValue -Object $s -PropertyName "status"),
            (& $boolText (Get-PropertyValue -Object $s -PropertyName "activeCOW")),
            (& $epochToUtc (Get-PropertyValue -Object $s -PropertyName "pitTimestamp"))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Volume Name", "Snapshot Group", "Status", "Active COW", "Creation Time (UTC)") -Rows $snapshotImageRows

    $markdownLines.Add("")
    $markdownLines.Add("## Snapshot Groups (All)")
    $markdownLines.Add("Snapshot groups bind volumes to a snapshot schedule and define retention policies for snapshot images.")
    $markdownLines.Add("")
    $markdownLines.Add("Total snapshot groups: $($snapshotGroups.Count)")
    $markdownLines.Add("")
    $snapshotGroupRows = @()
    foreach ($g in $snapshotGroups) {
        $groupId = [string](Get-PropertyValue -Object $g -PropertyName "id")
        $baseVol = $volumeById[[string](Get-PropertyValue -Object $g -PropertyName "baseVolume")]
        $imageCount = 0
        foreach ($img in $snapshotImages) {
            if ([string](Get-PropertyValue -Object $img -PropertyName "pitGroupRef") -eq $groupId) { $imageCount += 1 }
        }
        $scheduleCount = 0
        foreach ($sch in $snapshotSchedules) {
            if ([string](Get-PropertyValue -Object $sch -PropertyName "targetObject") -eq $groupId) { $scheduleCount += 1 }
        }
        $snapshotGroupRows += ,@((& $nameOf $g), (& $nameOf $baseVol), (Get-PropertyValue -Object $g -PropertyName "snapshotCount"), $imageCount, $scheduleCount, (Get-PropertyValue -Object $g -PropertyName "status"))
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Snapshot Group", "Volume Name", "Snapshot Count", "Image Count", "Schedule Count", "Status") -Rows $snapshotGroupRows

    $markdownLines.Add("")
    $markdownLines.Add("### Snapshot Group Summary (Non-Consistency Groups)")
    $markdownLines.Add("The following table details retention configuration for snapshot groups that are not associated with a consistency group.")
    $markdownLines.Add("")
    $standaloneRows = @()
    foreach ($g in $snapshotGroups) {
        $groupId = [string](Get-PropertyValue -Object $g -PropertyName "id")
        $isConsistency = $false
        foreach ($mv in $snapshotConsistencyGroupMemberVolumes) {
            if ([string](Get-PropertyValue -Object $mv -PropertyName "pitGroupId") -eq $groupId) { $isConsistency = $true; break }
        }
        if (-not $isConsistency) {
            $standaloneRows += ,@((& $nameOf $g), (Get-PropertyValue -Object $g -PropertyName "fullWarnThreshold"), (Get-PropertyValue -Object $g -PropertyName "autoDeleteLimit"), (Get-PropertyValue -Object $g -PropertyName "repFullPolicy"))
        }
    }
    $standaloneRows = @($standaloneRows | Sort-Object { [string]$_[0] })
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Snapshot Group", "Full Warn Threshold", "Auto Delete Limit", "Rep Full Policy") -Rows $standaloneRows

    $markdownLines.Add("")
    $markdownLines.Add("## Snapshot Consistency Groups")
    $markdownLines.Add("Snapshot consistency groups coordinate point-in-time snapshots across multiple member volumes. The following table details each consistency group, member volume, and associated snapshot group.")
    $markdownLines.Add("")
    $markdownLines.Add("Total snapshot consistency groups: $($snapshotConsistencyGroups.Count)")
    $markdownLines.Add("")
    $scgRows = @()
    foreach ($cg in $snapshotConsistencyGroups) {
        $cgId = [string](Get-PropertyValue -Object $cg -PropertyName "id")
        $hasMember = $false
        foreach ($mv in $snapshotConsistencyGroupMemberVolumes) {
            if ([string](Get-PropertyValue -Object $mv -PropertyName "consistencyGroupId") -ne $cgId) { continue }
            $hasMember = $true
            $memberVolumeName = Get-PropertyValue -Object $mv -PropertyName "baseVolumeName"
            if ([string]::IsNullOrWhiteSpace([string]$memberVolumeName)) {
                $mvVol = $volumeById[[string](Get-PropertyValue -Object $mv -PropertyName "volumeId")]
                $memberVolumeName = (& $nameOf $mvVol)
            }
            $linkedSnapshotGroup = $snapshotGroupById[[string](Get-PropertyValue -Object $mv -PropertyName "pitGroupId")]
            $scgRows += ,@((& $nameOf $cg), $memberVolumeName, (& $nameOf $linkedSnapshotGroup))
        }
        if (-not $hasMember) {
            $scgRows += ,@((& $nameOf $cg), "N/A", "N/A")
        }
    }
    $scgRows = @($scgRows | Sort-Object { [string]$_[0] }, { [string]$_[1] })
    $scgDisplayRows = @()
    $lastCgName = ""
    foreach ($scgRow in $scgRows) {
        $currentCgName = [string]$scgRow[0]
        $isFirstForCg = ($currentCgName -ne $lastCgName)
        $scgDisplayRows += ,@($(if ($isFirstForCg) { $currentCgName } else { "" }), $scgRow[1], $scgRow[2])
        $lastCgName = $currentCgName
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Consistency Group", "Member Volume", "Snapshot Group") -Rows $scgDisplayRows

    $markdownLines.Add("")
    $markdownLines.Add("### Snapshot Consistency Group Summary")
    $markdownLines.Add("The following table details retention configuration for snapshot groups associated with a consistency group.")
    $markdownLines.Add("")
    $scgSummaryRows = @()
    foreach ($cg in $snapshotConsistencyGroups) {
        $cgId = [string](Get-PropertyValue -Object $cg -PropertyName "id")
        $memberCount = 0
        foreach ($mv in $snapshotConsistencyGroupMemberVolumes) {
            if ([string](Get-PropertyValue -Object $mv -PropertyName "consistencyGroupId") -eq $cgId) { $memberCount += 1 }
        }
        $scgSummaryRows += ,@((& $nameOf $cg), $memberCount, (Get-PropertyValue -Object $cg -PropertyName "fullWarnThreshold"), (Get-PropertyValue -Object $cg -PropertyName "autoDeleteLimit"), (Get-PropertyValue -Object $cg -PropertyName "repFullPolicy"))
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Consistency Group", "Member Volume Count", "Full Warn Threshold", "Auto Delete Limit", "Rep Full Policy") -Rows $scgSummaryRows

    $markdownLines.Add("")
    $markdownLines.Add("## Snapshot Schedules")
    $markdownLines.Add("Snapshot schedules target either snapshot groups or snapshot consistency groups. The following table maps each schedule target to associated source volume names.")
    $markdownLines.Add("")
    $markdownLines.Add("Total snapshot schedules: $($snapshotSchedules.Count)")
    $markdownLines.Add("")
    $scheduleRows = @()
    foreach ($sch in $snapshotSchedules) {
        $target = [string](Get-PropertyValue -Object $sch -PropertyName "targetObject")
        $targetName = "N/A"
        $volumeNames = New-Object System.Collections.Generic.List[string]

        if ($snapshotGroupById.ContainsKey($target)) {
            $group = $snapshotGroupById[$target]
            $targetName = (& $nameOf $group)
            $baseVol = $volumeById[[string](Get-PropertyValue -Object $group -PropertyName "baseVolume")]
            $volumeNames.Add((& $nameOf $baseVol))
        }
        elseif ($snapshotConsistencyGroupById.ContainsKey($target)) {
            $cg = $snapshotConsistencyGroupById[$target]
            $targetName = (& $nameOf $cg)
            foreach ($mv in $snapshotConsistencyGroupMemberVolumes) {
                if ([string](Get-PropertyValue -Object $mv -PropertyName "consistencyGroupId") -eq $target) {
                    $memberVolName = Get-PropertyValue -Object $mv -PropertyName "baseVolumeName"
                    if ([string]::IsNullOrWhiteSpace([string]$memberVolName)) {
                        $memberVol = $volumeById[[string](Get-PropertyValue -Object $mv -PropertyName "volumeId")]
                        $memberVolName = (& $nameOf $memberVol)
                    }
                    if (-not [string]::IsNullOrWhiteSpace([string]$memberVolName)) {
                        $volumeNames.Add([string]$memberVolName)
                    }
                }
            }
        }

        $scheduleMethod = Get-PropertyValue -Object (Get-PropertyValue -Object (Get-PropertyValue -Object $sch -PropertyName "schedule") -PropertyName "calendar") -PropertyName "scheduleMethod"
        $scheduleRows += ,@(
            $targetName,
            ($(if ($volumeNames.Count -gt 0) { ($volumeNames | Select-Object -Unique) -join ", " } else { "N/A" })),
            (Get-PropertyValue -Object $sch -PropertyName "scheduleStatus"),
            $scheduleMethod,
            (& $epochToUtc (Get-PropertyValue -Object $sch -PropertyName "nextRunTime")),
            (& $epochToUtc (Get-PropertyValue -Object $sch -PropertyName "lastRunTime"))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Target Group Name", "Volume Name(s)", "Schedule Status", "Schedule Method", "Next Run (UTC)", "Last Run (UTC)") -Rows $scheduleRows

    $markdownLines.Add("")
    $markdownLines.Add("# System Settings and Configuration")
    $markdownLines.Add("This section captures DNS/NTP configuration, authentication, alerting, audit logging, and AutoSupport settings.")
    $markdownLines.Add("")
    $markdownLines.Add("## DNS")
    $markdownLines.Add("DNS servers resolve hostnames to IP addresses. The E-Series Storage System can be configured with manual or automatic DNS settings and one or more DNS servers.")
    $markdownLines.Add("")
    $dnsRows = @()
    foreach ($c in $controllers) {
        $ctrlLabel = Get-PropertyValue -Object (Get-PropertyValue -Object $c -PropertyName "physicalLocation") -PropertyName "label"
        $dnsAcqObj = Get-PropertyValue -Object (Get-PropertyValue -Object (Get-PropertyValue -Object $c -PropertyName "networkSettings") -PropertyName "dnsProperties") -PropertyName "acquisitionProperties"
        $dnsAcq = Get-PropertyValue -Object $dnsAcqObj -PropertyName "dnsAcquisitionType"
        $dnsServers = @((Get-PropertyValue -Object $dnsAcqObj -PropertyName "dnsServers"))
        $dns1 = "N/A"
        $dns2 = "N/A"
        if ($dnsServers.Count -gt 0) { $dns1 = Get-PropertyValue -Object $dnsServers[0] -PropertyName "ipv4Address" }
        if ($dnsServers.Count -gt 1) { $dns2 = Get-PropertyValue -Object $dnsServers[1] -PropertyName "ipv4Address" }
        $dnsRows += [pscustomobject]@{ Controller=$ctrlLabel; Method=$dnsAcq; Server1=$dns1; Server2=$dns2 }
    }
    $dnsRows = @($dnsRows | Sort-Object Controller)
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Controller", "DNS Config Method", "DNS Server 1", "DNS Server 2") -Rows @(
        $dnsRows | ForEach-Object { ,@($_.Controller, $_.Method, $_.Server1, $_.Server2) }
    )

    $markdownLines.Add("")
    $markdownLines.Add("## NTP")
    $markdownLines.Add("NTP enables the storage array to automatically synchronize the controller's clocks with an external host using Simple Network Time Protocol (SNTP). The controller periodically queries the configured NTP server, and then uses the results to update its internal time-of-day clock.")
    $markdownLines.Add("")
    $markdownLines.Add("If only one controller has NTP enabled, the alternate controller periodically synchronizes its clock with the controller that has NTP enabled. If neither controller has NTP enabled, the controllers periodically synchronize their clocks with each other.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the NTP settings in the E-Series Storage System.")
    $markdownLines.Add("")
    $ntpRows = @()
    foreach ($c in $controllers) {
        $ctrlLabel = Get-PropertyValue -Object (Get-PropertyValue -Object $c -PropertyName "physicalLocation") -PropertyName "label"
        $ntpAcqObj = Get-PropertyValue -Object (Get-PropertyValue -Object (Get-PropertyValue -Object $c -PropertyName "networkSettings") -PropertyName "ntpProperties") -PropertyName "acquisitionProperties"
        $ntpMethod = Get-PropertyValue -Object $ntpAcqObj -PropertyName "ntpAcquisitionType"
        $ntpServers = @((Get-PropertyValue -Object $ntpAcqObj -PropertyName "ntpServers"))
        $ntp1 = "N/A"
        $ntp2 = "N/A"
        if ($ntpServers.Count -gt 0) { $ntp1 = Get-PropertyValue -Object $ntpServers[0] -PropertyName "domainName" }
        if ($ntpServers.Count -gt 1) { $ntp2 = Get-PropertyValue -Object $ntpServers[1] -PropertyName "domainName" }
        $ntpRows += [pscustomobject]@{ Controller=$ctrlLabel; Method=$ntpMethod; Server1=$ntp1; Server2=$ntp2 }
    }
    $ntpRows = @($ntpRows | Sort-Object Controller)
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Controller", "NTP Config Method", "NTP Server 1", "NTP Server 2") -Rows @(
        $ntpRows | ForEach-Object { ,@($_.Controller, $_.Method, $_.Server1, $_.Server2) }
    )

    $markdownLines.Add("")
    $markdownLines.Add("## Array Settings")
    $markdownLines.Add("E-Series systems include some general array settings that can be configured to control the behaviour of the storage array. These settings include Access Volume (used for legacy in-band storage management), ASUP, Host Connectivity Reporting, and Automatic Load Balancing.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the general array settings in the E-Series Storage System.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("System Setting", "Configuration") -Rows @(
        @("Access Volume Enabled", (& $boolText (Get-PropertyValue -Object (Get-PropertyValue -Object $storageSystem -PropertyName "accessVolume") -PropertyName "enabled"))),
        @("ASUP Enabled", (& $boolText (Get-PropertyValue -Object $storageSystem -PropertyName "asupEnabled"))),
        @("Host Connectivity Reporting", (& $boolText (Get-PropertyValue -Object $storageSystem -PropertyName "hostConnectivityReportingEnabled"))),
        @("Automatic Load Balancing", (& $boolText (Get-PropertyValue -Object $storageSystem -PropertyName "autoLoadBalancingEnabled")))
    )

    $markdownLines.Add("")
    $markdownLines.Add("## Login Banner")
    $markdownLines.Add("The login banner is presented to users before they establish sessions in SANtricity System Manager. The banner can include an advisory notice and a consent message.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the login banner settings in the E-Series Storage System.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Login Banner Setting", "Configuration") -Rows @(
        @("Enabled", (& $boolText (Get-PropertyValue -Object $loginBanner -PropertyName "enabled"))),
        @("Message", (Get-PropertyValue -Object $loginBanner -PropertyName "message"))
    )

    $markdownLines.Add("")
    $markdownLines.Add("## SMTP Alerts")
    $markdownLines.Add("Alerts can be sent to an SMTP server, which can forward notifications to one or more recipients. SMTP security settings can include encryption and authentication.")
    $markdownLines.Add("")
    $markdownLines.Add("The following tables provide details of the SMTP alert settings in the E-Series Storage System.")
    $markdownLines.Add("")
    $markdownLines.Add("### SMTP Server")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("SMTP Server Setting", "Configuration") -Rows @(
        @("Alerting Enabled", (& $boolText (Get-PropertyValue -Object $deviceAlerts -PropertyName "alertingEnabled"))),
        @("Server Address", (Get-PropertyValue -Object $deviceAlerts -PropertyName "emailServerAddress")),
        @("Server Port", (Get-PropertyValue -Object $deviceAlerts -PropertyName "emailServerPort"))
    )

    $markdownLines.Add("")
    $markdownLines.Add("### SMTP Server Security")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("SMTP Server Security Setting", "Configuration") -Rows @(
        @("Encryption", (Get-PropertyValue -Object $deviceAlerts -PropertyName "emailServerEncryption")),
        @("Username", (Get-PropertyValue -Object $deviceAlerts -PropertyName "emailServerUsername")),
        @("Password", "***")
    )

    $markdownLines.Add("")
    $markdownLines.Add("### Email Settings")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Email Setting", "Configuration") -Rows @(
        @("Sender Address", (Get-PropertyValue -Object $deviceAlerts -PropertyName "emailSenderAddress")),
        @("Recipients", (@((Get-PropertyValue -Object $deviceAlerts -PropertyName "recipientEmailAddresses")) -join "<br>")),
        @("Contact Information", (Get-PropertyValue -Object $deviceAlerts -PropertyName "additionalContactInformation"))
    )

    $markdownLines.Add("")
    $markdownLines.Add("## SNMP Alerts")
    $markdownLines.Add("Alerts can be sent to a Simple Network Management Protocol (SNMP) trap destination. The configuration requires a community name or user name, and an IP address for the server. SNMP MIB Variables can also be configured to provide additional information about the storage array.")
    $markdownLines.Add("")
    $markdownLines.Add("The following tables provide details of the SNMP alert settings in the E-Series Storage System.")
    $markdownLines.Add("")
    $markdownLines.Add("### SNMP Destinations")
    $markdownLines.Add("")
    $snmpDestinationRows = @()
    foreach ($community in @((Get-PropertyValue -Object $snmpAlertSettings -PropertyName "communities"))) {
        $communityName = Get-PropertyValue -Object $community -PropertyName "name"
        foreach ($trap in @((Get-PropertyValue -Object $community -PropertyName "trapDestinations"))) {
            $receiver = Get-PropertyValue -Object $trap -PropertyName "receiverAddress"
            if (-not [string]::IsNullOrWhiteSpace([string]$receiver)) {
                $snmpDestinationRows += ,@($receiver, $communityName, (& $boolText (Get-PropertyValue -Object $trap -PropertyName "sendAuthenticationFailureTraps")))
            }
        }
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Receiver Address", "Community", "Send Authentication Failure Traps") -Rows $snmpDestinationRows

    $markdownLines.Add("")
    $markdownLines.Add("### SNMP MIB Variables")
    $markdownLines.Add("")
    $markdownLines.Add("Configuring MIB variables is optional. If configured, the storage array will return these variables in response to GetRequests.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("SNMP MIB Variable", "Configuration") -Rows @(
        @("Name", (Get-PropertyValue -Object $snmpAlertSettings -PropertyName "sysName")),
        @("Location", (Get-PropertyValue -Object $snmpAlertSettings -PropertyName "sysLocation")),
        @("Contact", (Get-PropertyValue -Object $snmpAlertSettings -PropertyName "sysContact"))
    )

    $markdownLines.Add("")
    $markdownLines.Add("## Syslog Alerts")
    $markdownLines.Add("Alerts can be sent to a Syslog server whenever an alertable event occurs. The Syslog server can be configured with a server address and a UDP port number. Typically, the UDP port for syslog is 514.")
    $markdownLines.Add("")
    $markdownLines.Add("The default syslog settings can also be configured, including the RFC format, default facility, and default tag.")
    $markdownLines.Add("")
    $markdownLines.Add("The following tables provide details of the Syslog alert settings in the E-Series Storage System.")
    $markdownLines.Add("")
    $markdownLines.Add("### Syslog Servers")
    $markdownLines.Add("")
    $syslogRows = @()
    $syslogReceivers = @((Get-PropertyValue -Object $deviceAlertSyslogSettings -PropertyName "syslogReceivers"))
    if ($syslogReceivers.Count -eq 0) {
        $syslogReceivers = @((Get-PropertyValue -Object $deviceAlertSyslogSettings -PropertyName "sysLogReceivers"))
    }
    foreach ($receiver in $syslogReceivers) {
        $syslogRows += ,@((Get-PropertyValue -Object $receiver -PropertyName "serverName"), (Get-PropertyValue -Object $receiver -PropertyName "portNumber"))
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Server Address", "UDP Port") -Rows $syslogRows

    $markdownLines.Add("")
    $markdownLines.Add("### Default Syslog Settings")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Default Syslog Setting", "Configuration") -Rows @(
        @("RFC Format", (Get-PropertyValue -Object $deviceAlertSyslogSettings -PropertyName "syslogFormat")),
        @("Default Facility", (Get-PropertyValue -Object $deviceAlertSyslogSettings -PropertyName "defaultFacility")),
        @("Default Tag", (Get-PropertyValue -Object $deviceAlertSyslogSettings -PropertyName "defaultTag"))
    )

    $markdownLines.Add("")
    $markdownLines.Add("## User Authentication")
    $markdownLines.Add("")
    $markdownLines.Add("### LDAP Configuration")
    $markdownLines.Add("")
    $markdownLines.Add("Authentication to SANtricity System Manager can be configured to use LDAP authentication. Authentication can be managed through an LDAP (Lightweight Directory Access Protocol) server and Directory Services, such as Microsoft's Active Directory.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table details the LDAP configuration settings in the E-Series Storage System.")
    $markdownLines.Add("")
    $ldapRows = @()
    foreach ($domain in @((Get-PropertyValue -Object $ldapSettings -PropertyName "ldapDomains"))) {
        $ldapRows += ,@(
            (@((Get-PropertyValue -Object $domain -PropertyName "names")) -join ", "),
            (Get-PropertyValue -Object $domain -PropertyName "ldapUrl"),
            (Get-PropertyValue -Object $domain -PropertyName "searchBase"),
            (Get-PropertyValue -Object $domain -PropertyName "userAttribute"),
            (@((Get-PropertyValue -Object $domain -PropertyName "groupAttributes")) -join ", ")
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Domain", "LDAP URL", "Search Base", "User Attribute", "Group Attributes") -Rows $ldapRows

    $markdownLines.Add("")
    $markdownLines.Add("### LDAP Bind Lookup User")
    $markdownLines.Add("")
    $markdownLines.Add("The LDAP Bind Lookup User is used to authenticate and search for users in the LDAP directory. The bind user credentials are required to establish a connection to the LDAP server and perform searches for user authentication, and must be supplied in the LDAP-type format.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the LDAP Bind Lookup User configuration in the E-Series Storage System.")
    $markdownLines.Add("")
    $ldapBindRows = @()
    foreach ($domain in @((Get-PropertyValue -Object $ldapSettings -PropertyName "ldapDomains"))) {
        $ldapBindRows += ,@((Get-PropertyValue -Object (Get-PropertyValue -Object $domain -PropertyName "bindLookupUser") -PropertyName "user"), "***")
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("LDAP Bind User", "Password") -Rows $ldapBindRows

    $markdownLines.Add("")
    $markdownLines.Add("### LDAP Group Mappings")
    $markdownLines.Add("Administrators can configure LDAP group mappings to assign roles based on directory group membership. This enables centralized access management and consistent role assignment in SANtricity System Manager.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table provides details of the LDAP group mappings in the E-Series Storage System.")
    $markdownLines.Add("")
    $ldapMapRows = @()
    foreach ($domain in @((Get-PropertyValue -Object $ldapSettings -PropertyName "ldapDomains"))) {
        foreach ($mapping in @((Get-PropertyValue -Object $domain -PropertyName "roleMapCollection"))) {
            $ldapMapRows += [pscustomobject]@{
                Group = [string](Get-PropertyValue -Object $mapping -PropertyName "groupRegex")
                Role  = [string](Get-PropertyValue -Object $mapping -PropertyName "name")
            }
        }
    }
    $ldapMapRows = @($ldapMapRows | Sort-Object Group, Role)
    $ldapMapDisplayRows = @()
    $lastLdapGroup = ""
    foreach ($ldapMapRow in $ldapMapRows) {
        $showGroup = ([string]$ldapMapRow.Group -ne [string]$lastLdapGroup)
        $ldapMapDisplayRows += ,@($(if ($showGroup) { $ldapMapRow.Group } else { "" }), $ldapMapRow.Role)
        $lastLdapGroup = [string]$ldapMapRow.Group
    }
    & $emitOpenXmlMergedFirstColumnTable -Headers @("LDAP Group", "Roles") -Rows $ldapMapDisplayRows -GridWidths @(3120, 4680)

    $markdownLines.Add("")
    $markdownLines.Add("## Audit Log Configuration")
    $markdownLines.Add("")
    $markdownLines.Add("### Audit Log Syslog Destinations")
    $markdownLines.Add("SANtricity records user administrative actions and security events. Audit logs can be forwarded to a Syslog server for centralized logging, monitoring, and analysis.")
    $markdownLines.Add("")
    $markdownLines.Add("Audit log Syslog destinations can use secure protocols such as TLS. A CA certificate must be provided to establish a secure connection with the Syslog server.")
    $markdownLines.Add("")
    $auditRows = @()
    foreach ($destination in $auditLogSyslogSettings) {
        $isAudit = $false
        foreach ($component in @((Get-PropertyValue -Object $destination -PropertyName "components"))) {
            if (([string](Get-PropertyValue -Object $component -PropertyName "type")).ToLowerInvariant() -eq "auditlog") {
                $isAudit = $true
                break
            }
        }
        if ($isAudit) {
            $auditRows += ,@(
                (Get-PropertyValue -Object $destination -PropertyName "serverAddress"),
                (Get-PropertyValue -Object $destination -PropertyName "protocol"),
                (Get-PropertyValue -Object $destination -PropertyName "port")
            )
        }
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Server Address", "Protocol", "Port") -Rows $auditRows

    $warningThreshold = Get-PropertyValue -Object $auditLogConfig -PropertyName "auditLogWarningThresholdPct"
    if ($null -eq $warningThreshold) { $warningThreshold = Get-PropertyValue -Object $auditLogConfig -PropertyName "auditLogwarningThreshold" }
    if ($null -eq $warningThreshold) { $warningThreshold = Get-PropertyValue -Object $auditLogConfig -PropertyName "auditLogWarningThreshold" }

    $markdownLines.Add("")
    $markdownLines.Add("### Audit Log System Configuration")
    $markdownLines.Add("Audit log system configuration controls log level, full-log behavior, maximum records, and warning threshold.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Audit Log Setting", "Configuration") -Rows @(
        @("Log Level", (Get-PropertyValue -Object $auditLogConfig -PropertyName "auditLogLevel")),
        @("Log Full Policy", (Get-PropertyValue -Object $auditLogConfig -PropertyName "auditLogFullPolicy")),
        @("Maximum Records", (Get-PropertyValue -Object $auditLogConfig -PropertyName "auditLogMaxRecords")),
        @("Warning Threshold", $(if ($null -ne $warningThreshold) { "$warningThreshold%" } else { "N/A" }))
    )

    $markdownLines.Add("")
    $markdownLines.Add("## AutoSupport")
    $markdownLines.Add("The AutoSupport (ASUP) feature monitors the health of a storage array and sends automatic dispatches to technical support.")
    $markdownLines.Add("")
    $markdownLines.Add("Technical support uses the AutoSupport data reactively to speed the diagnosis and resolution of customer issues and proactively to detect and avoid potential issues.")
    $markdownLines.Add("")
    $markdownLines.Add("AutoSupport data includes information about a storage array's configuration, status, performance, and system events. The AutoSupport data does not contain any user data. Dispatches can be sent immediately or on a schedule (daily and weekly).")
    $markdownLines.Add("")
    $markdownLines.Add("Some key benefits of the AutoSupport feature include:")
    $markdownLines.Add("")
    $markdownLines.Add("- Expedited case resolution")
    $markdownLines.Add("- Sophisticated monitoring for faster incident management")
    $markdownLines.Add("- Automated reporting according to a schedule, as well as automated reporting about critical events")
    $markdownLines.Add("- Automated hardware replacement requests for select components")
    $markdownLines.Add("- Nonintrusive alerting to notify you of a problem and provide information for technical support to initiate corrective action")
    $markdownLines.Add("- AutoSupport analysis tools that monitor dispatches for known configuration issues")
    $markdownLines.Add("")
    $markdownLines.Add("The AutoSupport feature is made up of three individual features that you enable separately.")
    $markdownLines.Add("")
    $markdownLines.Add("- **Basic AutoSupport** - Allows your storage array to automatically collect and send data to technical support.")
    $markdownLines.Add("- **AutoSupport OnDemand** - Allows technical support to request retransmission of a previous AutoSupport dispatch when needed for troubleshooting an issue. All transmissions are initiated from the storage array, not from the AutoSupport server. The storage array checks in periodically with the AutoSupport server to determine if there are any pending retransmission requests and responds accordingly.")
    $markdownLines.Add("- **Remote Diagnostics** - Allows technical support to request a new, up-to-date AutoSupport dispatch when needed for troubleshooting an issue. All transmissions are initiated from the storage array, not from the AutoSupport server. The storage array checks in periodically with the AutoSupport server to determine if there are any pending new requests and responds accordingly.")
    $markdownLines.Add("")
    $markdownLines.Add("### AutoSupport Feature Configuration")
    $markdownLines.Add("")
    $markdownLines.Add("The following table details the AutoSupport Feature configuration settings in the E-Series Storage System.")
    $markdownLines.Add("")
    $delivery = Get-PropertyValue -Object $deviceAsup -PropertyName "delivery"
    $asupMethod = ([string](Get-PropertyValue -Object $delivery -PropertyName "method")).ToLowerInvariant()
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("ASUP Feature", "Configuration") -Rows @(
        @("ASUP Enabled", (& $boolText (Get-PropertyValue -Object $deviceAsup -PropertyName "asupEnabled"))),
        @("ASUP On-Demand", (& $boolText (Get-PropertyValue -Object $deviceAsup -PropertyName "onDemandEnabled"))),
        @("Remote Diagnostics", (& $boolText (Get-PropertyValue -Object $deviceAsup -PropertyName "remoteDiagsEnabled"))),
        @("Delivery Method", (Get-PropertyValue -Object $delivery -PropertyName "method"))
    )

    if ($asupMethod -eq "https") {
        $markdownLines.Add("")
        $markdownLines.Add("### ASUP HTTPS Delivery")
        $markdownLines.Add("")
        $markdownLines.Add("The HTTPS delivery method for ASUP allows you to connect directly, or via a proxy server, to the destination technical support server using HTTPS. If you want to enable either AutoSupport OnDemand or Remote Diagnostics, the AutoSupport delivery method must be set to HTTPS.")
        $markdownLines.Add("")
        $markdownLines.Add("The following table provides details of the ASUP HTTPS delivery settings in the E-Series Storage System.")
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("ASUP HTTPS Delivery Setting", "Configuration") -Rows @(
            @("Routing Type", (Get-PropertyValue -Object $delivery -PropertyName "routingType")),
            @("Proxy Host", (Get-PropertyValue -Object $delivery -PropertyName "proxyHost")),
            @("Proxy Port", (Get-PropertyValue -Object $delivery -PropertyName "proxyPort")),
            @("Proxy Script", (Get-PropertyValue -Object $delivery -PropertyName "proxyScript")),
            @("Proxy Username", (Get-PropertyValue -Object $delivery -PropertyName "proxyUsername")),
            @("Proxy Password", "***"),
            @("HTTPS Max Size Limit", $(if ($null -ne (Get-PropertyValue -Object $delivery -PropertyName "maxSizeLimitHttps")) { "{0} MiB" -f [math]::Round(([double](Get-PropertyValue -Object $delivery -PropertyName "maxSizeLimitHttps") / 1048576), 2) } else { "N/A" }))
        )
    }
    elseif ($asupMethod -eq "smtp") {
        $markdownLines.Add("")
        $markdownLines.Add("### ASUP SMTP Delivery")
        $markdownLines.Add("")
        $markdownLines.Add("The SMTP delivery method for ASUP allows you to send AutoSupport messages via an SMTP server.")
        $markdownLines.Add("")
        $markdownLines.Add("The Email delivery method, which uses SMTP, has some important differences from the HTTPs delivery method.")
        $markdownLines.Add("- The size of the dispatches for the Email method are limited to 5MB, which means that some ASUP data collections will not be dispatched")
        $markdownLines.Add("- The AutoSupport OnDemand feature is available only when using the HTTPS delivery method.")
        $markdownLines.Add("")
        $markdownLines.Add("The following table provides details of the ASUP SMTP delivery settings in the E-Series Storage System.")
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("ASUP SMTP Delivery Setting", "Configuration") -Rows @(
            @("Sender Email Address", (Get-PropertyValue -Object $delivery -PropertyName "mailSenderAddress")),
            @("Additional ASUP Email Recipients", (@((Get-PropertyValue -Object $delivery -PropertyName "destinationAddressList")) -join "<br>")),
            @("SMTP Server", (Get-PropertyValue -Object $delivery -PropertyName "mailRelayServer")),
            @("SMTP Server Port", (Get-PropertyValue -Object $delivery -PropertyName "emailServerPort")),
            @("SMTP Server Encryption", (Get-PropertyValue -Object $delivery -PropertyName "emailServerEncryption")),
            @("SMTP Server Username", (Get-PropertyValue -Object $delivery -PropertyName "emailServerUsername")),
            @("SMTP Server Password", "***"),
            @("SMTP Max Size Limit", $(if ($null -ne (Get-PropertyValue -Object $delivery -PropertyName "maxSizeLimitSmtp")) { "{0} MiB" -f [math]::Round(([double](Get-PropertyValue -Object $delivery -PropertyName "maxSizeLimitSmtp") / 1048576), 2) } else { "N/A" }))
        )
    }

    $markdownLines.Add("")
    $markdownLines.Add('```{=openxml}')
    $markdownLines.Add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    $markdownLines.Add('```')
    $markdownLines.Add("")
    $markdownLines.Add("# References")
    $markdownLines.Add("")
    $markdownLines.Add("## Online Support Resources")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Name", "URL") -Rows @(
        @("Support Portal", "[https://mysupport.netapp.com](https://mysupport.netapp.com)"),
        @("Knowledge Base", "[https://kb.netapp.com](https://kb.netapp.com)"),
        @("E-Series Family Documentation", "[https://docs.netapp.com/us-en/e-series-family/](https://docs.netapp.com/us-en/e-series-family/)"),
        @("SANtricity Software Documentation", "[https://docs.netapp.com/us-en/e-series-santricity/index.html](https://docs.netapp.com/us-en/e-series-santricity/index.html)"),
        @("Product Documentation", "[https://docs.netapp.com](https://docs.netapp.com)"),
        @("Community Forums", "[https://community.netapp.com](https://community.netapp.com)"),
        @("Product Security", "[https://security.netapp.com](https://security.netapp.com)"),
        @("NetApp Console", "[https://console.netapp.com](https://console.netapp.com)"),
        @("NetApp Data Infrastructure Insights", "[https://www.netapp.com/data-insights](https://www.netapp.com/data-insights)"),
        @("Product Support Portals", "[https://mysupport.netapp.com/site/products/all](https://mysupport.netapp.com/site/products/all)"),
        @("Interoperability Matrix Tool", "[https://imt.netapp.com](https://imt.netapp.com)"),
        @("Hardware Universe", "[https://hwu.netapp.com](https://hwu.netapp.com)"),
        @("Support Communications", "[https://mysupport.netapp.com/info/communications/](https://mysupport.netapp.com/info/communications/)"),
        @("NetApp Learning Services", "[https://www.netapp.com/support-and-training/netapp-learning-services/](https://www.netapp.com/support-and-training/netapp-learning-services/)"),
        @("Hands-On Labs", "[https://labondemand.netapp.com/](https://labondemand.netapp.com/)"),
        @("Active IQ", "[https://activeiq.netapp.com](https://activeiq.netapp.com)"),
        @("Support Tools", "[https://mysupport.netapp.com/site/tools](https://mysupport.netapp.com/site/tools)")
    )

    if ($null -ne $RestQueryFailures -and @($RestQueryFailures).Count -gt 0) {
        $markdownLines.Add("")
        $markdownLines.Add('```{=openxml}')
        $markdownLines.Add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        $markdownLines.Add('```')
        $markdownLines.Add("")
        $markdownLines.Add("# API Collection Failures")
        $markdownLines.Add("")
        $markdownLines.Add("The following API requests did not complete successfully during collection. A successful response with an empty payload is not listed here.")
        $markdownLines.Add("")
        $warningRows = @()
        foreach ($failure in @($RestQueryFailures | Sort-Object -Property endpoint)) {
            $warningRows += ,@(
                (Get-PropertyValue -Object $failure -PropertyName "endpoint"),
                (Get-PropertyValue -Object $failure -PropertyName "status"),
                (Get-PropertyValue -Object $failure -PropertyName "msg")
            )
        }
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("API Endpoint", "Response Code", "Detail") -Rows $warningRows
    }

    [System.IO.File]::WriteAllLines($OutputPath, [string[]]$markdownLines)
}

function ConvertTo-Hashtable {
    param([Parameter(Mandatory = $true)]$InputObject)

    if ($null -eq $InputObject) {
        return @{}
    }

    if ($InputObject -is [hashtable]) {
        return $InputObject
    }

    if ($InputObject -is [pscustomobject]) {
        $ht = @{}
        foreach ($p in $InputObject.PSObject.Properties) {
            $ht[$p.Name] = $p.Value
        }
        return $ht
    }

    return @{}
}

function Invoke-EseriesApi {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$Endpoint,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $true)][string]$AuthorizationHeader,
        [Parameter(Mandatory = $false)][string]$AcceptHeader = "application/json"
    )

    $url = "$BaseUrl$Endpoint"
    $invokeParams = @{
        Uri         = $url
        Method      = "GET"
        Credential  = $Credential
        Headers     = @{
            Accept        = $AcceptHeader
            Authorization = $AuthorizationHeader
        }
        ErrorAction = "Stop"
    }
    if ($script:UsePowerShell7SkipCertificateCheck) {
        $invokeParams.SkipCertificateCheck = $true
    }

    try {
        $response = Invoke-RestMethod @invokeParams
        $normalizedPayload = ConvertTo-NormalizedApiPayload -Payload $response
        return [pscustomobject]@{
            Endpoint = $Endpoint
            Failed   = $false
            Status   = 200
            Data     = $normalizedPayload
            Message  = ""
        }
    }
    catch {
        $status = "N/A"
        $exceptionType = $_.Exception.GetType().FullName
        try {
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $status = [int]$_.Exception.Response.StatusCode
            }
        }
        catch {
            $status = "N/A"
        }

        return [pscustomobject]@{
            Endpoint = $Endpoint
            Failed   = $true
            Status   = $status
            Data     = @{}
            Message  = "Type=$exceptionType; " + (Get-ExceptionMessageChain -Exception $_.Exception)
        }
    }
}

$isJsonInputMode = -not [string]::IsNullOrWhiteSpace($JsonInputPath)
$asbuiltExport = $null
$restQueryFailures = @()
$jsonResolvedInput = $null

if ($isJsonInputMode) {
    $jsonResolvedInput = Resolve-ExistingFilePath -Path $JsonInputPath
    if ($null -eq $jsonResolvedInput) {
        throw "JSON input file not found: $JsonInputPath"
    }
    Write-Host "JSON input mode: loading data from $jsonResolvedInput"
    $asbuiltExport = Get-Content -Path $jsonResolvedInput -Raw | ConvertFrom-Json -ErrorAction Stop
    if ($null -eq $asbuiltExport.PSObject.Properties['storage_array_facts']) {
        throw "JSON input does not contain E-Series collection data: $JsonInputPath"
    }
}

if (-not $isJsonInputMode -and [string]::IsNullOrWhiteSpace($Target)) {
    $Target = Read-Host "Enter the E-Series Web Services target (e.g. santricity.example.com, santricity.example.com:8443, or https://santricity.example.com)"
}

if (-not $isJsonInputMode -and $null -eq $Credential) {
    $resolvedApiPassword = $null
    if ($ApiPasswordSecure -is [System.Security.SecureString]) {
        $resolvedApiPassword = $ApiPasswordSecure
    }
    elseif (-not [string]::IsNullOrWhiteSpace($LegacyApiSecret)) {
        $resolvedApiPassword = ConvertTo-SecureString -String $LegacyApiSecret -AsPlainText -Force
    }

    if (-not [string]::IsNullOrWhiteSpace($ApiUsername) -and $null -ne $resolvedApiPassword) {
        Write-Warning "ApiUsername/ApiPassword parameters are a legacy fallback and may expose secrets in shell history. Prefer -Credential or interactive Get-Credential prompt."
        $Credential = New-Object System.Management.Automation.PSCredential(
            $ApiUsername,
            $resolvedApiPassword
        )
    }
    else {
        $Credential = Get-Credential -Message "Enter E-Series API credentials"
    }
}

if ($PromptForTitlePageFields) {
    if ([string]::IsNullOrWhiteSpace($CustomerName)) {
        $CustomerName = Read-Host "Enter title page Customer Name"
    }
    if ([string]::IsNullOrWhiteSpace($CustomerLocation)) {
        $CustomerLocation = Read-Host "Enter title page Customer Location"
    }
    if ([string]::IsNullOrWhiteSpace($ProjectName)) {
        $ProjectName = Read-Host "Enter title page Project Name"
    }
}

if ($null -eq $TitlePageFields) {
    $TitlePageFields = @{}
}

if (-not [string]::IsNullOrWhiteSpace($TitlePageFieldsJson)) {
    $parsed = ConvertFrom-Json -InputObject $TitlePageFieldsJson -ErrorAction Stop
    $parsedFields = ConvertTo-Hashtable -InputObject $parsed
    foreach ($k in $parsedFields.Keys) {
        $TitlePageFields[$k] = $parsedFields[$k]
    }
}

if (-not [string]::IsNullOrWhiteSpace($CustomerName)) {
    $TitlePageFields["<<CUSTOMER_NAME>>"] = $CustomerName
}
if (-not [string]::IsNullOrWhiteSpace($CustomerLocation)) {
    $TitlePageFields["<<CUSTOMER_LOCATION>>"] = $CustomerLocation
}
if (-not [string]::IsNullOrWhiteSpace($ProjectName)) {
    $TitlePageFields["<<PROJECT_NAME>>"] = $ProjectName
}

$outputDirResolved = Resolve-OptionalPath -Path $OutputDir
Initialize-Path -Path $outputDirResolved

$timestamp = (Get-Date).ToUniversalTime().ToString("yyyyMMddTHHmmss")
$iso8601 = (Get-Date).ToUniversalTime().ToString("o")

if ($isJsonInputMode) {
    $apiUrl = "(loaded from $([System.IO.Path]::GetFileName($jsonResolvedInput)))"
    $storageSystemData = Get-PropertyValue -Object $asbuiltExport.storage_array_facts -PropertyName 'netapp_storage_systems_details'
    $responses = @{
        storage_systems = [pscustomobject]@{ Data = $storageSystemData }
    }
    if ([string]::IsNullOrWhiteSpace($Ssid)) {
        $Ssid = 'Loaded from JSON'
    }
}
else {
    # Default bare host or host:port input to HTTPS; retain an explicit HTTP or HTTPS scheme.
    $Target = $Target.Trim().TrimEnd('/')
    if ($Target -notmatch '^[a-zA-Z][a-zA-Z0-9+.-]*://') {
        $Target = "https://$Target"
    }
    try {
        $targetUri = [uri]$Target
    }
    catch {
        throw "Invalid E-Series Web Services target: $Target"
    }
    if ($targetUri.Scheme -notin @('http', 'https')) {
        throw "E-Series Web Services target must use HTTP or HTTPS: $Target"
    }
    if (-not [string]::IsNullOrWhiteSpace($targetUri.Query) -or -not [string]::IsNullOrWhiteSpace($targetUri.Fragment)) {
        throw "E-Series Web Services target must not include a query string or fragment: $Target"
    }
    if ($targetUri.AbsolutePath -notin @('/', '/devmgr/v2', '/devmgr/v2/')) {
        throw "E-Series Web Services target must be the server root or /devmgr/v2: $Target"
    }

    $authority = if ($targetUri.HostNameType -eq [System.UriHostNameType]::IPv6) { "[$($targetUri.Host)]" } else { $targetUri.Host }
    $hasExplicitUrlPort = $Target -match '^[a-zA-Z][a-zA-Z0-9+.-]*://(?:\[[^\]]+\]|[^/:]+):\d+(?:/|$)'
    $resolvedPort = if ($hasExplicitUrlPort) { $targetUri.Port } elseif ($PSBoundParameters.ContainsKey('Port')) { $Port } else { 8443 }
    $apiUrl = "{0}://{1}:{2}/devmgr/v2" -f $targetUri.Scheme, $authority, $resolvedPort
    Set-TlsPolicy -ValidateCertificates $ValidateCerts.IsPresent
    Set-ProxyPolicy -UseProxy $UseSystemProxy.IsPresent
    $authorizationHeader = New-BasicAuthHeader -Credential $Credential

Write-Host "Querying /storage-systems ..."
$storageSystemsResp = Invoke-EseriesApi -BaseUrl $apiUrl -Endpoint "/storage-systems" -Credential $Credential -AuthorizationHeader $authorizationHeader -AcceptHeader "application/json"

if ($storageSystemsResp.Failed) {
    $statusText = if ($null -ne $storageSystemsResp.Status) { "HTTP $($storageSystemsResp.Status)" } else { "request failure" }
    throw "E-Series storage-system inventory failed ($statusText): $($storageSystemsResp.Message)"
}

if ([string]::IsNullOrWhiteSpace($Ssid)) {
    if (-not $storageSystemsResp.Failed -and $null -ne $storageSystemsResp.Data) {
        if ($storageSystemsResp.Data -is [array] -and $storageSystemsResp.Data.Count -gt 0 -and $storageSystemsResp.Data[0].id) {
            $Ssid = [string]$storageSystemsResp.Data[0].id
        }
        elseif ($storageSystemsResp.Data.id) {
            $Ssid = [string]$storageSystemsResp.Data.id
        }
    }
}
if ([string]::IsNullOrWhiteSpace($Ssid)) {
    $Ssid = Read-Host "Could not auto-detect SSID. Enter SSID"
}

$endpointMap = @(
    @{ Key = "storage_pools"; Endpoint = "/storage-systems/$Ssid/storage-pools" },
    @{ Key = "volumes"; Endpoint = "/storage-systems/$Ssid/volumes" },
    @{ Key = "hosts"; Endpoint = "/storage-systems/$Ssid/hosts" },
    @{ Key = "host_types"; Endpoint = "/storage-systems/$Ssid/host-types" },
    @{ Key = "host_groups"; Endpoint = "/storage-systems/$Ssid/host-groups" },
    @{ Key = "snapshot_groups"; Endpoint = "/storage-systems/$Ssid/snapshot-groups" },
    @{ Key = "snapshot_space_utilization"; Endpoint = "/storage-systems/$Ssid/snapshot-groups/repository-utilization" },
    @{ Key = "snapshot_images"; Endpoint = "/storage-systems/$Ssid/snapshot-images" },
    @{ Key = "snapshot_schedules"; Endpoint = "/storage-systems/$Ssid/snapshot-schedules" },
    @{ Key = "snapshot_consistency_groups"; Endpoint = "/storage-systems/$Ssid/consistency-groups" },
    @{ Key = "consistency_group_member_volumes"; Endpoint = "/storage-systems/$Ssid/consistency-groups/member-volumes" },
    @{ Key = "embedded_firmware_versions"; Endpoint = "/firmware/embedded-firmware/$Ssid/versions" },
    @{ Key = "ldap_settings"; Endpoint = "/storage-systems/$Ssid/ldap" },
    @{ Key = "device_alerts"; Endpoint = "/storage-systems/$Ssid/device-alerts" },
    @{ Key = "device_asup"; Endpoint = "/device-asup" },
    @{ Key = "audit_log_config"; Endpoint = "/storage-systems/$Ssid/audit-log/config" },
    @{ Key = "audit_log_syslog_settings"; Endpoint = "/storage-systems/$Ssid/syslog" },
    @{ Key = "snmp_alert_settings"; Endpoint = "/storage-systems/$Ssid/snmp" },
    @{ Key = "hardware_inventory"; Endpoint = "/storage-systems/$Ssid/hardware-inventory" },
    @{ Key = "device_alert_syslog"; Endpoint = "/storage-systems/$Ssid/device-alerts/alert-syslog"; Accept = "application/json" },
    @{ Key = "login_banner"; Endpoint = "/storage-systems/$Ssid/login-banner"; Accept = "*/*" }
)

$responses = @{}
$responses["storage_systems"] = $storageSystemsResp

foreach ($entry in $endpointMap) {
    Write-Host "Querying $($entry.Endpoint) ..."
    $acceptHeader = "application/json"
    if ($entry -is [hashtable] -and $entry.ContainsKey("Accept") -and -not [string]::IsNullOrWhiteSpace([string]$entry["Accept"])) {
        $acceptHeader = [string]$entry["Accept"]
    }
    $responses[$entry.Key] = Invoke-EseriesApi -BaseUrl $apiUrl -Endpoint $entry.Endpoint -Credential $Credential -AuthorizationHeader $authorizationHeader -AcceptHeader $acceptHeader
}

$arrayPayloadKeys = @(
    "storage_pools",
    "volumes",
    "hosts",
    "host_types",
    "host_groups",
    "snapshot_groups",
    "snapshot_space_utilization",
    "snapshot_images",
    "snapshot_schedules",
    "snapshot_consistency_groups",
    "consistency_group_member_volumes",
    "audit_log_syslog_settings"
)

foreach ($arrayKey in $arrayPayloadKeys) {
    if ($responses.ContainsKey($arrayKey)) {
        $responses[$arrayKey].Data = Convert-ToArrayPayload -Payload $responses[$arrayKey].Data
    }
}

$loginBannerRaw = $responses["login_banner"]
$loginBannerMessage = ""
if ($loginBannerRaw.Data -is [string]) {
    $loginBannerMessage = [string]$loginBannerRaw.Data
}
elseif ($loginBannerRaw.Data -is [pscustomobject] -or $loginBannerRaw.Data -is [hashtable]) {
    if ($loginBannerRaw.Data -is [hashtable]) {
        if ($loginBannerRaw.Data.ContainsKey("message") -and -not [string]::IsNullOrWhiteSpace([string]$loginBannerRaw.Data["message"])) {
            $loginBannerMessage = [string]$loginBannerRaw.Data["message"]
        }
        elseif ($loginBannerRaw.Data.ContainsKey("banner") -and -not [string]::IsNullOrWhiteSpace([string]$loginBannerRaw.Data["banner"])) {
            $loginBannerMessage = [string]$loginBannerRaw.Data["banner"]
        }
        elseif ($loginBannerRaw.Data.ContainsKey("text") -and -not [string]::IsNullOrWhiteSpace([string]$loginBannerRaw.Data["text"])) {
            $loginBannerMessage = [string]$loginBannerRaw.Data["text"]
        }
    }
    else {
        $messageProp = $loginBannerRaw.Data.PSObject.Properties["message"]
        $bannerProp = $loginBannerRaw.Data.PSObject.Properties["banner"]
        $textProp = $loginBannerRaw.Data.PSObject.Properties["text"]

        if ($null -ne $messageProp -and -not [string]::IsNullOrWhiteSpace([string]$messageProp.Value)) {
            $loginBannerMessage = [string]$messageProp.Value
        }
        elseif ($null -ne $bannerProp -and -not [string]::IsNullOrWhiteSpace([string]$bannerProp.Value)) {
            $loginBannerMessage = [string]$bannerProp.Value
        }
        elseif ($null -ne $textProp -and -not [string]::IsNullOrWhiteSpace([string]$textProp.Value)) {
            $loginBannerMessage = [string]$textProp.Value
        }
    }
}

$normalizedLoginBanner = [ordered]@{
    status_code  = $loginBannerRaw.Status
    content_type = "application/json"
    enabled      = (-not $loginBannerRaw.Failed) -and -not [string]::IsNullOrWhiteSpace($loginBannerMessage)
    message      = $loginBannerMessage
    raw_json     = $loginBannerRaw.Data
    raw_text     = $loginBannerMessage
}

$asbuiltExport = [ordered]@{
    changed            = $false
    failed             = $false
    msg                = "REST-collected data"
    storage_array_facts = [ordered]@{
        netapp_storage_systems_details                              = Get-ResponseData -Responses $responses -Key "storage_systems"
        netapp_storage_pools_details                                = Get-ResponseData -Responses $responses -Key "storage_pools" -AsArray
        netapp_volumes_details                                      = Get-ResponseData -Responses $responses -Key "volumes" -AsArray
        netapp_hosts_details                                        = Get-ResponseData -Responses $responses -Key "hosts" -AsArray
        netapp_host_types_details                                   = Get-ResponseData -Responses $responses -Key "host_types" -AsArray
        netapp_host_groups_details                                  = Get-ResponseData -Responses $responses -Key "host_groups" -AsArray
        netapp_snapshot_groups_details                              = Get-ResponseData -Responses $responses -Key "snapshot_groups" -AsArray
        netapp_snapshot_space_utilization_details                   = Get-ResponseData -Responses $responses -Key "snapshot_space_utilization" -AsArray
        netapp_snapshot_images_details                              = Get-ResponseData -Responses $responses -Key "snapshot_images" -AsArray
        netapp_snapshot_schedules_details                           = Get-ResponseData -Responses $responses -Key "snapshot_schedules" -AsArray
        netapp_snapshot_consistency_groups_details                  = Get-ResponseData -Responses $responses -Key "snapshot_consistency_groups" -AsArray
        netapp_snapshot_consistency_group_member_volumes_details    = Get-ResponseData -Responses $responses -Key "consistency_group_member_volumes" -AsArray
        netapp_embedded_firmware_versions                           = Get-ResponseData -Responses $responses -Key "embedded_firmware_versions"
        netapp_ldap_settings                                        = Get-ResponseData -Responses $responses -Key "ldap_settings"
        netapp_device_alerts_settings                               = Get-ResponseData -Responses $responses -Key "device_alerts"
        netapp_device_asup_settings                                 = Get-ResponseData -Responses $responses -Key "device_asup"
        netapp_audit_log_config                                     = Get-ResponseData -Responses $responses -Key "audit_log_config"
        netapp_audit_log_syslog_settings                            = Get-ResponseData -Responses $responses -Key "audit_log_syslog_settings" -AsArray
        netapp_snmp_alert_settings                                  = Get-ResponseData -Responses $responses -Key "snmp_alert_settings"
        netapp_hardware_inventory                                   = Get-ResponseData -Responses $responses -Key "hardware_inventory"
        netapp_device_alert_syslog_settings                         = Get-ResponseData -Responses $responses -Key "device_alert_syslog"
        netapp_login_banner_settings                                = $normalizedLoginBanner
    }
}

$arrayFactKeys = @(
    "netapp_storage_pools_details",
    "netapp_volumes_details",
    "netapp_hosts_details",
    "netapp_host_types_details",
    "netapp_host_groups_details",
    "netapp_snapshot_groups_details",
    "netapp_snapshot_space_utilization_details",
    "netapp_snapshot_images_details",
    "netapp_snapshot_schedules_details",
    "netapp_snapshot_consistency_groups_details",
    "netapp_snapshot_consistency_group_member_volumes_details",
    "netapp_audit_log_syslog_settings"
)

foreach ($factKey in $arrayFactKeys) {
    if (-not $asbuiltExport.storage_array_facts.Contains($factKey)) {
        $asbuiltExport.storage_array_facts[$factKey] = [object[]]@()
        continue
    }

    $normalized = Convert-ToArrayPayload -Payload $asbuiltExport.storage_array_facts[$factKey]
    $asbuiltExport.storage_array_facts[$factKey] = [object[]]@($normalized)
}

$restQueryFailures = @()
foreach ($key in $responses.Keys) {
    if ($responses[$key].Failed) {
        $restQueryFailures += [pscustomobject]@{
            endpoint = $responses[$key].Endpoint
            status   = $responses[$key].Status
            msg      = $responses[$key].Message
        }
    }
}

if ($restQueryFailures.Count -gt 0) {
    Write-Warning "One or more REST endpoint queries failed; affected report sections may show N/A."
    $restQueryFailures | Format-Table -AutoSize | Out-String | Write-Host
}
} # end live collection

$storageSystemObject = $responses["storage_systems"].Data
$reportSystemNameSource = "unknown"
if ($storageSystemObject -is [array] -and $storageSystemObject.Count -gt 0) {
    $firstSystem = $storageSystemObject[0]
    if ($firstSystem -is [hashtable]) {
        if ($firstSystem.ContainsKey("name") -and -not [string]::IsNullOrWhiteSpace([string]$firstSystem["name"])) {
            $reportSystemNameSource = [string]$firstSystem["name"]
        }
    }
    elseif ($null -ne $firstSystem.PSObject.Properties["name"] -and -not [string]::IsNullOrWhiteSpace([string]$firstSystem.PSObject.Properties["name"].Value)) {
        $reportSystemNameSource = [string]$firstSystem.PSObject.Properties["name"].Value
    }
}
elseif ($storageSystemObject -is [hashtable]) {
    if ($storageSystemObject.ContainsKey("name") -and -not [string]::IsNullOrWhiteSpace([string]$storageSystemObject["name"])) {
        $reportSystemNameSource = [string]$storageSystemObject["name"]
    }
}
elseif ($null -ne $storageSystemObject -and $null -ne $storageSystemObject.PSObject.Properties["name"] -and -not [string]::IsNullOrWhiteSpace([string]$storageSystemObject.PSObject.Properties["name"].Value)) {
    $reportSystemNameSource = [string]$storageSystemObject.PSObject.Properties["name"].Value
}
$reportSystemName = Convert-ToSafeName -Value $reportSystemNameSource

$reportBaseName = if ([string]::IsNullOrWhiteSpace($ReportPrefix)) {
    "${reportSystemName}_${timestamp}"
}
else {
    "${ReportPrefix}_${reportSystemName}_${timestamp}"
}

$jsonPath = if ($isJsonInputMode) { $null } else { Join-Path -Path $outputDirResolved -ChildPath "${reportBaseName}.json" }
$mdPath = Join-Path -Path $outputDirResolved -ChildPath "${reportBaseName}.md"
$docxPath = Join-Path -Path $outputDirResolved -ChildPath "${reportBaseName}.docx"

if (-not $isJsonInputMode) {
    $asbuiltExport | ConvertTo-Json -Depth 100 | Set-Content -Path $jsonPath -Encoding UTF8
}

if ($CollectionOnly) {
    if ($isJsonInputMode) {
        Write-Host "Collection-only mode has no effect when using -JsonInputPath."
    }
    else {
        Write-Host "Collection-only mode enabled. JSON export generated at: $jsonPath"
    }
    return
}

Write-Host "Rendering markdown report ..."
Write-EseriesMarkdown -Export $asbuiltExport -ApiUrl $apiUrl -Ssid $Ssid -Iso8601 $iso8601 -OutputPath $mdPath -RestQueryFailures $restQueryFailures -TitleFields $TitlePageFields -IncludeTitlePage ([bool]$DocxEnableTitlePage)

$docxReferenceTemplate = Resolve-ExistingFilePath -Path ".\docx\reference-template.docx"

Write-Host "Converting markdown to DOCX ..."
Convert-AsBuiltMarkdownToDocx -MarkdownPath $mdPath -DocxPath $docxPath -ReferenceDocumentPath $docxReferenceTemplate -TableStyleName $DocxTableStyleName -NumberSections ([bool]$EnableDocxNumberSections) | Out-Null

$reportMetadata = Read-ReportMetadata -MetadataPath ".\docx\report-metadata-e-series.yml"

$docPropCustomerName = ""
if ($TitlePageFields.ContainsKey("<<CUSTOMER_NAME>>") -and -not [string]::IsNullOrWhiteSpace([string]$TitlePageFields["<<CUSTOMER_NAME>>"])) {
    $docPropCustomerName = [string]$TitlePageFields["<<CUSTOMER_NAME>>"]
}

$docPropCustomerLocation = ""
if ($TitlePageFields.ContainsKey("<<CUSTOMER_LOCATION>>") -and -not [string]::IsNullOrWhiteSpace([string]$TitlePageFields["<<CUSTOMER_LOCATION>>"])) {
    $docPropCustomerLocation = [string]$TitlePageFields["<<CUSTOMER_LOCATION>>"]
}

$docPropProjectName = ""
if ($TitlePageFields.ContainsKey("<<PROJECT_NAME>>") -and -not [string]::IsNullOrWhiteSpace([string]$TitlePageFields["<<PROJECT_NAME>>"])) {
    $docPropProjectName = [string]$TitlePageFields["<<PROJECT_NAME>>"]
}

$docPropSystemName = ""
if (-not [string]::IsNullOrWhiteSpace($reportSystemNameSource)) {
    $docPropSystemName = [string]$reportSystemNameSource
}

$docPropPlaceholder = '<<PLEASE UPDATE via DOCUMENT PROPERTIES>>'
$docPropCustomerName = if ([string]::IsNullOrWhiteSpace($docPropCustomerName)) { $docPropPlaceholder } else { $docPropCustomerName }
$docPropCustomerLocation = if ([string]::IsNullOrWhiteSpace($docPropCustomerLocation)) { $docPropPlaceholder } else { $docPropCustomerLocation }
$docPropProjectName = if ([string]::IsNullOrWhiteSpace($docPropProjectName)) { $docPropPlaceholder } else { $docPropProjectName }
$docPropSystemName = if ([string]::IsNullOrWhiteSpace($docPropSystemName)) { $null } else { $docPropSystemName }

Set-DocxMetadataProperties `
    -DocxPath $docxPath `
    -Title $(if ([string]::IsNullOrWhiteSpace([string]$reportMetadata.title)) { 'E-Series As-Built' } else { [string]$reportMetadata.title }) `
    -Subject $(if ([string]::IsNullOrWhiteSpace([string]$reportMetadata.subject)) { 'E-Series Configuration Summary' } else { [string]$reportMetadata.subject }) `
    -Author $(if ([string]::IsNullOrWhiteSpace([string]$reportMetadata.author)) { 'NetApp' } else { [string]$reportMetadata.author }) `
    -CustomerName $docPropCustomerName `
    -CustomerLocation $docPropCustomerLocation `
    -ProjectName $docPropProjectName `
    -SystemName $docPropSystemName

Set-DocxTableStyle -DocxPath $docxPath -TableStyleId $DocxTableStyleName -HeaderParagraphStyle $DocxTableHeaderParagraphStyle -BodyParagraphStyle $DocxTableBodyParagraphStyle -AutofitToWindow ([bool]$DocxTableAutofitToWindow)

if ($CleanupIntermediateOutputs -and -not $KeepIntermediateOutputs) {
    $pathsToRemove = @($mdPath)
    if (-not [string]::IsNullOrWhiteSpace($jsonPath)) { $pathsToRemove += $jsonPath }
    Remove-Item -Path $pathsToRemove -Force -ErrorAction SilentlyContinue
}

Write-Host "DOCX report generated at: $docxPath"

}

Export-ModuleMember -Function 'Invoke-ESeriesAsBuilt'
