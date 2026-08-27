Set-StrictMode -Version Latest
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'AsBuilt.Common.psm1') -Force

function Get-StorageGridDiscoveryPropertyValue {
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

function Get-StorageGridSantricityCollectionMode {
    param(
        [Parameter(Mandatory = $false)][bool]$CollectSantricityAppliances,
        [Parameter(Mandatory = $false)][bool]$CollectParameterSpecified,
        [Parameter(Mandatory = $false)][bool]$NonInteractive
    )

    $explicitOptOut = $CollectParameterSpecified -and -not $CollectSantricityAppliances
    return [pscustomobject][ordered]@{
        DetectionEnabled = -not $explicitOptOut
        ForcedCollection = $CollectParameterSpecified -and $CollectSantricityAppliances
        PromptAllowed    = -not $NonInteractive
        ExplicitOptOut   = $explicitOptOut
    }
}

function Read-StorageGridSantricityYesNo {
    param(
        [Parameter(Mandatory = $true)][string]$Heading,
        [Parameter(Mandatory = $false)][string[]]$Message = @(),
        [Parameter(Mandatory = $false)][scriptblock]$ChoicePromptCommand
    )

    while ($true) {
        Write-Host ''
        Write-Host '======================================================================' -ForegroundColor DarkCyan
        Write-Host (" {0}" -f $Heading) -ForegroundColor Yellow
        Write-Host '======================================================================' -ForegroundColor DarkCyan
        foreach ($line in @($Message)) {
            if ([string]::IsNullOrWhiteSpace([string]$line)) {
                Write-Host ''
            }
            else {
                Write-Host $line -ForegroundColor White
            }
        }
        Write-Host ''
        Write-Host '[Y] Yes' -ForegroundColor Cyan
        Write-Host '[N] No' -ForegroundColor Cyan
        Write-Host ''
        $answer = if ($null -ne $ChoicePromptCommand) {
            & $ChoicePromptCommand
        }
        else {
            Read-Host 'Enter Y or N'
        }
        $normalizedAnswer = ([string]$answer).Trim().ToUpperInvariant()
        if ($normalizedAnswer -in @('Y', 'N')) {
            Write-Host ''
            return $normalizedAnswer
        }
        Write-Warning "Please enter Y or N. The response '$answer' was not recognized."
    }
}

function Get-StorageGridSantricityApplianceCandidates {
    param(
        [Parameter(Mandatory = $false)]$NetworkTopology,
        [Parameter(Mandatory = $false)]$NodeAttributes,
        [Parameter(Mandatory = $false)]$NodeHealth,
        [Parameter(Mandatory = $false)]$ApplianceStorageMetrics
    )

    $attributeByNodeId = @{}
    $nodesPayload = Get-StorageGridDiscoveryPropertyValue -Object $NodeAttributes -PropertyName 'nodes'
    foreach ($attributeNode in @($nodesPayload)) {
        $nodeId = [string](Get-StorageGridDiscoveryPropertyValue -Object $attributeNode -PropertyName 'nodeId')
        if (-not [string]::IsNullOrWhiteSpace($nodeId)) {
            $attributeByNodeId[$nodeId] = $attributeNode
        }
    }

    $healthByNodeId = @{}
    foreach ($healthNode in @($NodeHealth)) {
        $nodeId = [string](Get-StorageGridDiscoveryPropertyValue -Object $healthNode -PropertyName 'id')
        if (-not [string]::IsNullOrWhiteSpace($nodeId)) {
            $healthByNodeId[$nodeId] = $healthNode
        }
    }

    $storageMetricsByNodeName = @{}
    $metricsPayload = Get-StorageGridDiscoveryPropertyValue -Object $ApplianceStorageMetrics -PropertyName 'nodes'
    foreach ($metricNode in @($metricsPayload)) {
        $nodeName = [string](Get-StorageGridDiscoveryPropertyValue -Object $metricNode -PropertyName 'nodeName')
        if (-not [string]::IsNullOrWhiteSpace($nodeName)) {
            $storageMetricsByNodeName[$nodeName.ToLowerInvariant()] = $metricNode
        }
    }

    $gridNodesPayload = Get-StorageGridDiscoveryPropertyValue -Object $NetworkTopology -PropertyName 'gridNodes'
    $gridNodesArray = @($gridNodesPayload)
    $candidates = @()

    foreach ($topologyNode in $gridNodesArray) {
        $nodeConfig = Get-StorageGridDiscoveryPropertyValue -Object $topologyNode -PropertyName 'nodeConfig'
        if ($null -eq $nodeConfig) { continue }
        if ([string](Get-StorageGridDiscoveryPropertyValue -Object $nodeConfig -PropertyName 'platform') -ne 'SGA') { continue }

        $nodeId = [string](Get-StorageGridDiscoveryPropertyValue -Object $nodeConfig -PropertyName 'nodeId')
        $nodeName = [string](Get-StorageGridDiscoveryPropertyValue -Object $nodeConfig -PropertyName 'hostname')
        if ([string]::IsNullOrWhiteSpace($nodeId) -or [string]::IsNullOrWhiteSpace($nodeName)) { continue }
        if (-not $attributeByNodeId.ContainsKey($nodeId)) { continue }

        $attributeNode = $attributeByNodeId[$nodeId]
        $candidateIps = @()
        foreach ($propertyName in @('controllerAIp', 'controllerBIp')) {
            $ipText = [string](Get-StorageGridDiscoveryPropertyValue -Object $attributeNode -PropertyName $propertyName)
            $parsedIp = $null
            if ([System.Net.IPAddress]::TryParse($ipText, [ref]$parsedIp) -and -not $candidateIps.Contains($ipText)) {
                $candidateIps += $ipText
            }
        }
        if ($candidateIps.Count -eq 0) { continue }

        $healthNode = if ($healthByNodeId.ContainsKey($nodeId)) { $healthByNodeId[$nodeId] } else { $null }
        $metricNode = if ($storageMetricsByNodeName.ContainsKey($nodeName.ToLowerInvariant())) { $storageMetricsByNodeName[$nodeName.ToLowerInvariant()] } else { $null }
        $siteName = [string](Get-StorageGridDiscoveryPropertyValue -Object $healthNode -PropertyName 'siteName')
        if ([string]::IsNullOrWhiteSpace($siteName)) { $siteName = [string](Get-StorageGridDiscoveryPropertyValue -Object $attributeNode -PropertyName 'siteName') }
        if ([string]::IsNullOrWhiteSpace($siteName)) { $siteName = 'N/A' }

        $santricityVersion = [string](Get-StorageGridDiscoveryPropertyValue -Object $metricNode -PropertyName 'santricityVersion')
        if ([string]::IsNullOrWhiteSpace($santricityVersion)) { $santricityVersion = 'N/A' }
        $storageSerial = [string](Get-StorageGridDiscoveryPropertyValue -Object $metricNode -PropertyName 'storageSerialNumber')
        if ([string]::IsNullOrWhiteSpace($storageSerial)) { $storageSerial = 'N/A' }

        $candidates += [pscustomobject][ordered]@{
            NodeId            = $nodeId
            NodeName          = $nodeName
            SiteName          = $siteName
            NodeType          = [string](Get-StorageGridDiscoveryPropertyValue -Object $nodeConfig -PropertyName 'nodeType')
            Platform          = 'SGA'
            HasEseries        = Get-StorageGridDiscoveryPropertyValue -Object $nodeConfig -PropertyName 'hasEseries'
            ApplianceModel    = [string](Get-StorageGridDiscoveryPropertyValue -Object $attributeNode -PropertyName 'applianceModel')
            ControllerName    = [string](Get-StorageGridDiscoveryPropertyValue -Object $attributeNode -PropertyName 'controllerName')
            CandidateIps      = [object[]]$candidateIps
            StorageSerial     = $storageSerial
            SantricityVersion = $santricityVersion
            Eligible          = $true
        }
    }

    return ,@($candidates | Sort-Object SiteName, NodeName)
}

function Invoke-StorageGridSantricityApplianceCollection {
    param(
        [Parameter(Mandatory = $true)][object[]]$Candidates,
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [Parameter(Mandatory = $true)][string]$WorkspaceRoot,
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$SharedCredential,
        [Parameter(Mandatory = $false)][hashtable]$CredentialMap,
        [Parameter(Mandatory = $false)][switch]$ValidateCerts,
        [Parameter(Mandatory = $false)][switch]$UseSystemProxy,
        [Parameter(Mandatory = $false)][switch]$KeepIntermediateOutputs,
        [Parameter(Mandatory = $false)][switch]$DocxEnableTitlePage,
        [Parameter(Mandatory = $false)][string]$CustomerName,
        [Parameter(Mandatory = $false)][string]$CustomerLocation,
        [Parameter(Mandatory = $false)][string]$ProjectName,
        [Parameter(Mandatory = $false)][hashtable]$TitlePageFields,
        [Parameter(Mandatory = $false)][scriptblock]$CollectorCommand
    )

    if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
    }

    $results = @()
    foreach ($candidate in @($Candidates)) {
        $credential = $SharedCredential
        $credentialKey = $null
        foreach ($key in @($candidate.NodeName) + @($candidate.CandidateIps)) {
            if ($null -ne $CredentialMap -and $CredentialMap.ContainsKey([string]$key)) {
                $credential = $CredentialMap[[string]$key]
                $credentialKey = [string]$key
                break
            }
        }

        $result = [ordered]@{
            NodeName      = [string]$candidate.NodeName
            CandidateIps  = [object[]]@($candidate.CandidateIps)
            SuccessfulIp  = $null
            Status        = 'Failed'
            Error         = $null
            CredentialKey = $credentialKey
        }
        if ($null -eq $credential) {
            $result.Error = 'No SANtricity credential was supplied.'
            $results += [pscustomobject]$result
            continue
        }

        foreach ($ip in @($candidate.CandidateIps)) {
            try {
                Write-Host "Collecting SANtricity data for $($candidate.NodeName) via $ip ..."
                $invokeParameters = @{
                    Target                   = [string]$ip
                    Credential               = $credential
                    OutputDir                = $OutputDirectory
                    ReportPrefix             = [string]$candidate.NodeName
                    WorkspaceRoot            = $WorkspaceRoot
                    ValidateCerts            = $ValidateCerts.IsPresent
                    UseSystemProxy           = $UseSystemProxy.IsPresent
                    KeepIntermediateOutputs = $KeepIntermediateOutputs.IsPresent
                }
                if ($DocxEnableTitlePage.IsPresent) {
                    $invokeParameters.DocxEnableTitlePage = $true
                }
                if (-not [string]::IsNullOrWhiteSpace($CustomerName)) {
                    $invokeParameters.CustomerName = $CustomerName
                }
                if (-not [string]::IsNullOrWhiteSpace($CustomerLocation)) {
                    $invokeParameters.CustomerLocation = $CustomerLocation
                }
                if (-not [string]::IsNullOrWhiteSpace($ProjectName)) {
                    $invokeParameters.ProjectName = $ProjectName
                }
                if ($null -ne $TitlePageFields -and $TitlePageFields.Count -gt 0) {
                    $invokeParameters.TitlePageFields = $TitlePageFields
                }
                if ($null -ne $CollectorCommand) {
                    & $CollectorCommand @invokeParameters
                }
                else {
                    Invoke-ESeriesAsBuilt @invokeParameters
                }
                $result.SuccessfulIp = [string]$ip
                $result.Status = 'Succeeded'
                $result.Error = $null
                break
            }
            catch {
                $result.Error = $_.Exception.Message
                Write-Warning "SANtricity collection failed for $($candidate.NodeName) via $ip`: $($result.Error)"
            }
        }
        $results += [pscustomobject]$result
    }

    return ,@($results)
}

function Resolve-StorageGridSantricityCredentials {
    param(
        [Parameter(Mandatory = $true)][object[]]$Candidates,
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$SharedCredential,
        [Parameter(Mandatory = $false)][hashtable]$CredentialMap,
        [Parameter(Mandatory = $false)][switch]$NonInteractive,
        [Parameter(Mandatory = $false)][scriptblock]$ChoicePromptCommand,
        [Parameter(Mandatory = $false)][scriptblock]$CredentialPromptCommand
    )

    if ($null -ne $CredentialMap -or $null -ne $SharedCredential) {
        return [pscustomobject]@{
            SharedCredential = $SharedCredential
            CredentialMap    = $CredentialMap
        }
    }
    if ($NonInteractive) {
        return [pscustomobject]@{
            SharedCredential = $null
            CredentialMap    = $null
        }
    }

    $choice = Read-StorageGridSantricityYesNo -Heading 'SANTRICITY CREDENTIALS' -Message @(
        ("{0} appliances will be collected." -f @($Candidates).Count)
        ''
        'Use one SANtricity credential for all appliances?'
        ''
        'Y uses one shared credential; N prompts separately for each appliance.'
    ) -ChoicePromptCommand $ChoicePromptCommand
    if ($choice -eq 'Y') {
        $credential = if ($null -ne $CredentialPromptCommand) {
            & $CredentialPromptCommand 'all appliances'
        }
        else {
            Get-Credential -Message 'Enter shared SANtricity credentials for all detected appliances'
        }
        return [pscustomobject]@{
            SharedCredential = $credential
            CredentialMap    = $null
        }
    }

    $perApplianceMap = @{}
    $candidateIndex = 0
    $candidateCount = @($Candidates).Count
    foreach ($candidate in @($Candidates)) {
        $candidateIndex++
        Write-Host ''
        Write-Host '======================================================================' -ForegroundColor DarkCyan
        Write-Host (" SANTRICITY CREDENTIALS [{0} of {1}]" -f $candidateIndex, $candidateCount) -ForegroundColor Yellow
        Write-Host '======================================================================' -ForegroundColor DarkCyan
        Write-Host ''
        Write-Host ("Enter credentials for {0}" -f $candidate.NodeName) -ForegroundColor White
        $credential = if ($null -ne $CredentialPromptCommand) {
            & $CredentialPromptCommand ([string]$candidate.NodeName)
        }
        else {
            Get-Credential -Message ("Enter SANtricity credentials for {0}" -f $candidate.NodeName)
        }
        if ($null -ne $credential) {
            $perApplianceMap[[string]$candidate.NodeName] = $credential
        }
    }

    return [pscustomobject]@{
        SharedCredential = $null
        CredentialMap    = $perApplianceMap
    }
}

function Invoke-StorageGridAsBuilt {
[CmdletBinding()]
param(
    # Full base URL of the StorageGRID Management API, e.g. https://admin.example.com or https://admin.example.com:8443
    [Parameter(Mandatory = $false)]
    [string]$Target,

    [Parameter(Mandatory = $false)]
    [string]$ApiUsername,

    [Parameter(Mandatory = $false)]
    [System.Security.SecureString]$ApiPasswordSecure,

    # Backward-compatible alias for legacy plain-text password input; -Credential is preferred.
    [Parameter(Mandatory = $false)]
    [Alias("ApiPassword")]
    [string]$LegacyApiSecret,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [switch]$ValidateCerts,

    [Parameter(Mandatory = $false)]
    [switch]$UseSystemProxy,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir = ".\asbuilt_output_storagegrid",

    [Parameter(Mandatory = $false)]
    [switch]$CollectionOnly,

    [Parameter(Mandatory = $false)]
    [switch]$ExportCapacityDiagnostics,

    [Parameter(Mandatory = $false)]
    [switch]$CollectSantricityAppliances,

    [Parameter(Mandatory = $false)]
    [switch]$NonInteractive,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$SantricityCredential,

    [Parameter(Mandatory = $false)]
    [string]$SantricityAuthMapPath,

    # Load previously collected JSON instead of running a live collection (skips Target/Credential requirements)
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
    [switch]$PromptForTitlePageFields
,

    [Parameter(Mandatory = $true)]
    [string]$WorkspaceRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$collectSantricitySpecified = $PSBoundParameters.ContainsKey('CollectSantricityAppliances')
$santricityMode = Get-StorageGridSantricityCollectionMode -CollectSantricityAppliances ([bool]$CollectSantricityAppliances) -CollectParameterSpecified ([bool]$collectSantricitySpecified) -NonInteractive ([bool]$NonInteractive)
$santricityDetectionEnabled = $santricityMode.DetectionEnabled
$santricityCollectionRequested = $santricityMode.ForcedCollection
$santricityApplianceCandidates = @()
$santricityCollectionResults = @()
Import-Module (Join-Path -Path $WorkspaceRoot -ChildPath 'Modules\AsBuilt.Common.psm1') -Force
# ChangeLog:
# 2026.08.27.1 - Moved detailed node-attribute and private API diagnostics to verbose output, made the response type diagnostic null-safe, and report completion of the StorageGRID DOCX before prompting for, or collecting, SANtricity appliance reports.
# 2026.08.26.1 - Replaced Pandoc-based DOCX generation with a direct DocumentFormat.OpenXml SDK pipeline (Lib\OpenXml); no external DOCX conversion binary is required. Fixed title-page section header/footer/titlePg preservation and single-column table parsing, added selective hyperlink support for markdown [text](url) content, and removed all Pandoc-related code and dependencies.
$ScriptVersion = "2026.08.27.1"

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

            # Tables authored with an explicit fixed layout (e.g. our merged-cell Field/Value tables) keep their
            # own dxa column widths and must be excluded from the autofit-to-window pct conversion below.
            $existingTblLayout = $tblPr.SelectSingleNode('./w:tblLayout', $namespaceManager)
            $isFixedLayoutTable = $null -ne $existingTblLayout -and $existingTblLayout.GetAttribute('type', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') -eq 'fixed'

            if ($AutofitToWindow -and -not $isFixedLayoutTable) {
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

            # Re-assert each fixed-layout column's dxa width from the tblGrid, since some Word clients
            # can otherwise normalize a cell's tcW to "auto".
            if ($isFixedLayoutTable) {
                $gridCols = @($table.SelectNodes('./w:tblGrid/w:gridCol', $namespaceManager))
                if ($gridCols.Count -gt 0 -and $rows.Count -gt 0) {
                    foreach ($row in $rows) {
                        $rowCells = @($row.SelectNodes('./w:tc', $namespaceManager))
                        for ($cellIndex = 0; $cellIndex -lt $rowCells.Count -and $cellIndex -lt $gridCols.Count; $cellIndex++) {
                            $colWidth = $gridCols[$cellIndex].GetAttribute('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                            $tcPr = Get-OrCreate-XmlChildNode -Parent $rowCells[$cellIndex] -ChildName 'tcPr' -InsertAtStart $true
                            $tcW = Get-OrCreate-XmlChildNode -Parent $tcPr -ChildName 'tcW'
                            [void]$tcW.SetAttribute('type', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'dxa')
                            [void]$tcW.SetAttribute('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', $colWidth)
                        }
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
    param([Parameter(Mandatory = $false)][string]$MetadataPath = ".\docx\report-metadata-storagegrid.yml")

    $reportMetadata = @{
        title = "As-Built Documentation"
        subject = "NetApp StorageGRID"
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

        if (-not [string]::IsNullOrWhiteSpace([string]$CustomerName)) {
            Set-CustomStringProperty -Xml $customXml -Root $customRoot -Name 'CustomerName' -Value ([string]$CustomerName)
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$CustomerLocation)) {
            Set-CustomStringProperty -Xml $customXml -Root $customRoot -Name 'CustomerLocation' -Value ([string]$CustomerLocation)
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$ProjectName)) {
            Set-CustomStringProperty -Xml $customXml -Root $customRoot -Name 'ProjectName' -Value ([string]$ProjectName)
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$SystemName)) {
            Set-CustomStringProperty -Xml $customXml -Root $customRoot -Name 'SystemName' -Value ([string]$SystemName)
        }

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

    # StorageGRID API v4 wraps all responses in { "data": ... , "apiVersion": ... }
    # Use PSObject.Properties directly to avoid PS5.1 unrolling empty arrays through Get-PropertyValue
    $apiVersionProp = $Payload.PSObject.Properties["apiVersion"]
    if ($null -ne $apiVersionProp) {
        $dataProp = $Payload.PSObject.Properties["data"]
        if ($null -ne $dataProp) {
            # Comma operator prevents PS5.1 from unrolling an empty array return to $null
            if ($dataProp.Value -is [array]) { return ,$dataProp.Value }
            return $dataProp.Value
        }
        return $null
    }

    $valueProperty = Get-PropertyValue -Object $Payload -PropertyName "value"
    $countProperty = Get-PropertyValue -Object $Payload -PropertyName "Count"
    if ($null -ne $valueProperty -and $null -ne $countProperty -and -not ($Payload -is [string])) {
        return $valueProperty
    }

    return $Payload
}

function ConvertTo-X509CertificateFromPem {
    param([Parameter(Mandatory = $false)][string]$Pem)

    if ([string]::IsNullOrWhiteSpace($Pem)) { return $null }
    try {
        $base64 = [regex]::Replace($Pem, '-----BEGIN CERTIFICATE-----|-----END CERTIFICATE-----|\s', '')
        $certificateBytes = [Convert]::FromBase64String($base64)
        return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList (,$certificateBytes)
    }
    catch {
        return $null
    }
}

function Get-X509CertificateDisplayDetails {
    param([Parameter(Mandatory = $false)][string]$Pem)

    $certificate = ConvertTo-X509CertificateFromPem -Pem $Pem
    if ($null -eq $certificate) {
        return [pscustomobject]@{ CommonName = "N/A"; SubjectAltNames = @("N/A"); Expiry = "N/A" }
    }

    $commonName = $certificate.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName, $false)
    if ([string]::IsNullOrWhiteSpace($commonName)) { $commonName = "N/A" }

    $sanValues = @()
    foreach ($extension in $certificate.Extensions) {
        if ([string]$extension.Oid.Value -ne "2.5.29.17") { continue }
        $formattedSans = $extension.Format($true)
        foreach ($sanLine in @($formattedSans -split "`r?`n")) {
            $sanText = [string]$sanLine
            if ([string]::IsNullOrWhiteSpace($sanText)) { continue }
            $sanValues += ($sanText -replace '^\s*(DNS Name|IP Address|RFC822 Name|URL)=', '$1:')
        }
    }
    if ($sanValues.Count -eq 0) { $sanValues = @("N/A") }

    return [pscustomobject]@{
        CommonName     = [string]$commonName
        SubjectAltNames = [object[]]@($sanValues)
        Expiry         = $certificate.NotAfter.ToUniversalTime().ToString("dd-MM-yyyy")
    }
}

function Convert-ToArrayPayload {
    param([Parameter(Mandatory = $false)]$Payload)

    # Comma operator on all array returns prevents PS5.1 from producing AutomationNull for empty arrays
    if ($null -eq $Payload) {
        return ,@()
    }

    if ($Payload -is [array]) {
        return ,$Payload
    }

    if (
        $Payload -is [System.Collections.IEnumerable] -and
        -not ($Payload -is [string]) -and
        -not ($Payload -is [hashtable]) -and
        -not ($Payload -is [pscustomobject])
    ) {
        $items = @($Payload)
        return ,$items
    }

    if ($Payload -is [hashtable]) {
        if ((@($Payload.GetEnumerator())).Count -eq 0) {
            return ,@()
        }
        return ,([object[]]@($Payload))
    }

    $propNames = @($Payload.PSObject.Properties.Name)
    if ($propNames.Count -gt 0) {
        return ,([object[]]@($Payload))
    }

    return ,([object[]]@($Payload))
}

function Get-ResponseData {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Responses,
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $false)][switch]$AsArray
    )

    if (-not $Responses.ContainsKey($Key) -or $null -eq $Responses[$Key]) {
        if ($AsArray) {
            return ,@()
        }
        return @{}
    }

    $data = $Responses[$Key].Data
    if ($AsArray) {
        $normalized = Convert-ToArrayPayload -Payload $data
        return ,$normalized
    }

    if ($null -eq $data) {
        return @{}
    }

    return $data
}

function Get-StorageGridToken {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential
    )

    $plainPassword = [System.Net.NetworkCredential]::new("", $Credential.Password).Password
    $body = @{
        username    = $Credential.UserName
        password    = $plainPassword
        cookie      = $false
        csrfToken   = $false
    } | ConvertTo-Json

    $invokeParams = @{
        Uri         = "$BaseUrl/api/v4/authorize"
        Method      = "POST"
        Body        = $body
        ContentType = "application/json"
        Headers     = @{ Accept = "application/json" }
        ErrorAction = "Stop"
    }
    if ($script:UsePowerShell7SkipCertificateCheck) {
        $invokeParams.SkipCertificateCheck = $true
    }

    try {
        $response = Invoke-RestMethod @invokeParams
        $token = Get-PropertyValue -Object $response -PropertyName "data"
        if ([string]::IsNullOrWhiteSpace([string]$token)) {
            throw "Authorization response did not contain a token in the 'data' field."
        }
        return [string]$token
    }
    catch {
        throw "StorageGRID authentication failed: " + (Get-ExceptionMessageChain -Exception $_.Exception)
    }
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
    $text = $text -replace "(?i)<br\s*/?>", "<br />"
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

function Convert-BytesToDecimalTB {
    param([Parameter(Mandatory = $false)]$Bytes)

    if ($null -eq $Bytes) {
        return "N/A"
    }

    $n = 0
    if (-not [double]::TryParse([string]$Bytes, [ref]$n)) {
        return "N/A"
    }

    return ([math]::Round(($n / 1000000000000), 1)).ToString("0.0")
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

function Write-StorageGridMarkdown {
    param(
        [Parameter(Mandatory = $true)]$Export,
        [Parameter(Mandatory = $true)][string]$ApiUrl,
        [Parameter(Mandatory = $true)][string]$Iso8601,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)]$RestQueryFailures,
        [Parameter(Mandatory = $true)][hashtable]$TitleFields,
        [Parameter(Mandatory = $true)][bool]$IncludeTitlePage
    )

    $facts = $Export.storagegrid_facts
    $productVersion     = Get-PropertyValue -Object $facts -PropertyName "sg_product_version"
    $gridConfig         = Get-PropertyValue -Object $facts -PropertyName "sg_grid_config"
    $gridHealth         = Get-PropertyValue -Object $facts -PropertyName "sg_grid_health"
    $nodeHealth         = @((Get-PropertyValue -Object $facts -PropertyName "sg_node_health"))
    $gridRegions        = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $facts -PropertyName "sg_grid_regions")
    $networkTopology    = Get-PropertyValue -Object $facts -PropertyName "sg_network_topology"
    $gridNetworks       = Get-PropertyValue -Object $facts -PropertyName "sg_grid_networks"
    $dnsServers         = Get-PropertyValue -Object $facts -PropertyName "sg_dns_servers"
    $ntpServers         = Get-PropertyValue -Object $facts -PropertyName "sg_ntp_servers"
    $browserTimeout     = Get-PropertyValue -Object $facts -PropertyName "sg_browser_timeout"
    $gridLicense        = Get-PropertyValue -Object $facts -PropertyName "sg_grid_license"
    $capacityMetrics    = Get-PropertyValue -Object $facts -PropertyName "sg_capacity_metrics"
    $applianceStorageMetrics = Get-PropertyValue -Object $facts -PropertyName "sg_appliance_storage_metrics"
    $nodeAttributeDetails = Get-PropertyValue -Object $facts -PropertyName "sg_node_attribute_details"
    $santricityApplianceCandidates = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $facts -PropertyName "sg_santricity_appliance_candidates")
    $autosupport        = Get-PropertyValue -Object $facts -PropertyName "sg_autosupport"
    $identitySource     = Get-PropertyValue -Object $facts -PropertyName "sg_identity_source"
    $ssoConfig          = Get-PropertyValue -Object $facts -PropertyName "sg_sso"
    $gridGroups         = @((Get-PropertyValue -Object $facts -PropertyName "sg_grid_groups"))
    $gridUsers          = @((Get-PropertyValue -Object $facts -PropertyName "sg_grid_users"))
    $auditConfig        = Get-PropertyValue -Object $facts -PropertyName "sg_audit_config"
    $auditDestinations  = Get-PropertyValue -Object $facts -PropertyName "sg_audit_destinations"
    $alertReceivers     = @((Get-PropertyValue -Object $facts -PropertyName "sg_alert_receivers"))
    $alertRules         = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $facts -PropertyName "sg_alert_rules")
    $alertSilences      = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $facts -PropertyName "sg_alert_silences")
    $snmpConfig         = Get-PropertyValue -Object $facts -PropertyName "sg_snmp"
    $complianceGlobal   = Get-PropertyValue -Object $facts -PropertyName "sg_compliance_global"
    $storageWatermarks  = Get-PropertyValue -Object $facts -PropertyName "sg_storage_watermarks"
    $mgmtCert           = Get-PropertyValue -Object $facts -PropertyName "sg_management_certificate"
    $internalCaCert     = Get-PropertyValue -Object $facts -PropertyName "sg_internal_ca_cert"
    $clientCertificates = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $facts -PropertyName "sg_client_certificates")
    $untrustedClientNet = Get-PropertyValue -Object $facts -PropertyName "sg_untrusted_client_network"
    $externalLoadBalancers = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $facts -PropertyName "sg_external_load_balancers")
    $firewallBlockedPorts  = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $facts -PropertyName "sg_firewall_blocked_ports")
    $firewallExternalPorts = Get-PropertyValue -Object $facts -PropertyName "sg_firewall_external_ports"
    $firewallPrivilegedIps = Get-PropertyValue -Object $facts -PropertyName "sg_firewall_privileged_ips"
    $ciphersConfig      = Get-PropertyValue -Object $facts -PropertyName "sg_ciphers"
    $bmcConfig          = Get-PropertyValue -Object $facts -PropertyName "sg_bmc"
    $domainNames        = Get-PropertyValue -Object $facts -PropertyName "sg_domain_names"
    $managementCors    = Get-PropertyValue -Object $facts -PropertyName "sg_management_cors"
    $haGroups           = @((Get-PropertyValue -Object $facts -PropertyName "sg_ha_groups"))
    $gatewayConfigs     = @((Get-PropertyValue -Object $facts -PropertyName "sg_gateway_configs"))
    $lbServerConfigs    = @((Get-PropertyValue -Object $facts -PropertyName "sg_lb_server_configs"))
    $trafficPolicies    = @((Get-PropertyValue -Object $facts -PropertyName "sg_traffic_policies"))
    $trafficPolicyDetails = @((Get-PropertyValue -Object $facts -PropertyName "sg_traffic_policy_details"))
    $vlanInterfaces     = @((Get-PropertyValue -Object $facts -PropertyName "sg_vlan_interfaces"))
    $ilmRules           = @((Get-PropertyValue -Object $facts -PropertyName "sg_ilm_rules"))
    $ilmPolicies        = @((Get-PropertyValue -Object $facts -PropertyName "sg_ilm_policies"))
    $ecProfiles         = @((Get-PropertyValue -Object $facts -PropertyName "sg_ec_profiles"))
    $ilmPools           = @((Get-PropertyValue -Object $facts -PropertyName "sg_ilm_pools"))
    $ilmPolicyTags      = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $facts -PropertyName "sg_ilm_policy_tags")
    $ilmGrades          = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $facts -PropertyName "sg_ilm_grades")
    $ilmGradeSite       = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $facts -PropertyName "sg_ilm_grade_site")
    $cloudStoragePools  = @((Get-PropertyValue -Object $facts -PropertyName "sg_cloud_storage_pools"))
    $kmipClusters       = @((Get-PropertyValue -Object $facts -PropertyName "sg_kmip_clusters"))
    $storageProxy       = Get-PropertyValue -Object $facts -PropertyName "sg_storage_proxy"
    $adminProxy         = Get-PropertyValue -Object $facts -PropertyName "sg_admin_proxy"
    $gridAccounts       = @((Get-PropertyValue -Object $facts -PropertyName "sg_grid_accounts"))
    $gridAccountUsage   = @((Get-PropertyValue -Object $facts -PropertyName "sg_grid_account_usage"))
    $gridFederationConnections = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $facts -PropertyName "sg_grid_federation")
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
            $n = Get-PropertyValue -Object $Obj -PropertyName "displayName"
        }
        if ([string]::IsNullOrWhiteSpace([string]$n)) {
            return "N/A"
        }
        return [string]$n
    }

    $capacityBytesText = {
        param($Bytes)
        if ($null -eq $Bytes) { return "N/A" }
        return (Convert-BytesToTiB -Bytes $Bytes)
    }

    $capacityPercentText = {
        param($Value)
        if ($null -eq $Value) { return "N/A" }
        return (([math]::Round([double]$Value, 2)).ToString("0.00") + "%")
    }

    $capacityObjectMetricsAreValid = {
        param($Record)
        return (Test-StorageGridObjectCapacityRecordValidity -Record $Record)
    }

    $capacityObjectBytesText = {
        param(
            $Record,
            [string]$PropertyName,
            [bool]$RequireValidatedObjectMetrics = $false
        )

        if ($RequireValidatedObjectMetrics -and -not (& $capacityObjectMetricsAreValid $Record)) {
            return "N/A"
        }

        return (& $capacityBytesText (Get-PropertyValue -Object $Record -PropertyName $PropertyName))
    }

    $capacityObjectPercentText = {
        param(
            $Record,
            [string]$PropertyName,
            [bool]$RequireValidatedObjectMetrics = $false
        )

        if ($RequireValidatedObjectMetrics -and -not (& $capacityObjectMetricsAreValid $Record)) {
            return "N/A"
        }

        return (& $capacityPercentText (Get-PropertyValue -Object $Record -PropertyName $PropertyName))
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

    $gridName = [string](Get-PropertyValue -Object $gridLicense -PropertyName "systemName")
    if ([string]::IsNullOrWhiteSpace($gridName)) {
        $gridName = [string](Get-PropertyValue -Object $gridLicense -PropertyName "displayName")
    }
    if ([string]::IsNullOrWhiteSpace($gridName)) { $gridName = "StorageGRID" }

    $reportMetadata = Read-ReportMetadata -MetadataPath ".\docx\report-metadata-storagegrid.yml"

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
            [Parameter(Mandatory = $false)][int[]]$GridWidths = @(),
            [Parameter(Mandatory = $false)][switch]$AutofitToWindow
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
        $tableTotalWidth = 0
        foreach ($gridWidth in $GridWidths) { $tableTotalWidth += $gridWidth }

        $markdownLines.Add('```{=openxml}')
        $tableWidthXml = if ($AutofitToWindow.IsPresent) {
            '<w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="autofit"/>'
        } else {
            '<w:tblW w:w="' + [string]$tableTotalWidth + '" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        }
        $markdownLines.Add('<w:tbl>')
        $markdownLines.Add('<w:tblPr><w:tblStyle w:val="NetAppTable1"/>' + $tableWidthXml + '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>')
        $markdownLines.Add('<w:tblGrid>' + $tableGridXml + '</w:tblGrid>')

        $headerXml = '<w:tr><w:trPr><w:tblHeader/></w:trPr>'
        for ($headerIndex = 0; $headerIndex -lt $Headers.Count; $headerIndex++) {
            $escapedHeader = [System.Security.SecurityElement]::Escape([string]$Headers[$headerIndex])
            $headerXml += '<w:tc><w:tcPr><w:tcW w:w="' + [string]$GridWidths[$headerIndex] + '" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableHeading"/></w:pPr><w:r><w:t xml:space="preserve">' + $escapedHeader + '</w:t></w:r></w:p></w:tc>'
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
                $firstCellRunsXml = ($firstCellRaw -split "`n" | ForEach-Object { '<w:r><w:t xml:space="preserve">' + [System.Security.SecurityElement]::Escape($_) + '</w:t></w:r>' }) -join '<w:r><w:br/></w:r>'
                $rowXml += '<w:tc><w:tcPr><w:tcW w:w="' + [string]$GridWidths[0] + '" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableText"/></w:pPr>' + $firstCellRunsXml + '</w:p></w:tc>'
            }
            else {
                $rowXml += '<w:tc><w:tcPr><w:tcW w:w="' + [string]$GridWidths[0] + '" w:type="dxa"/><w:vMerge/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableText"/></w:pPr></w:p></w:tc>'
            }

            for ($colIndex = 1; $colIndex -lt $columnCount; $colIndex++) {
                $cellText = [string]$cells[$colIndex]
                if ([string]::IsNullOrWhiteSpace($cellText)) {
                    $cellText = "N/A"
                }
                $cellRunsXml = ($cellText -split "`n" | ForEach-Object { '<w:r><w:t xml:space="preserve">' + [System.Security.SecurityElement]::Escape($_) + '</w:t></w:r>' }) -join '<w:r><w:br/></w:r>'
                $rowXml += '<w:tc><w:tcPr><w:tcW w:w="' + [string]$GridWidths[$colIndex] + '" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableText"/></w:pPr>' + $cellRunsXml + '</w:p></w:tc>'
            }

            $rowXml += '</w:tr>'
            $markdownLines.Add($rowXml)
        }

        $markdownLines.Add('</w:tbl>')
        $markdownLines.Add('```')
    }

    $emitOpenXmlMergedColumnsTable = {
        param(
            [Parameter(Mandatory = $true)][string[]]$Headers,
            [Parameter(Mandatory = $false)][object[]]$Rows = @(),
            [Parameter(Mandatory = $false)][int[]]$GridWidths = @(),
            [Parameter(Mandatory = $false)][int[]]$MergeColumns = @(),
            [Parameter(Mandatory = $false)][switch]$AutofitToWindow
        )

        $columnCount = $Headers.Count
        if ($columnCount -lt 2) { throw "Merged-column table emitter requires at least two columns." }
        if ($GridWidths.Count -ne $columnCount) {
            $GridWidths = @()
            for ($widthIndex = 0; $widthIndex -lt $columnCount; $widthIndex++) { $GridWidths += 2400 }
        }

        $tableGridXml = (($GridWidths | ForEach-Object { '<w:gridCol w:w="' + [string]$_ + '"/>' }) -join '')
        $tableTotalWidth = ($GridWidths | Measure-Object -Sum).Sum
        $mergeLookup = @{}
        foreach ($mergeColumn in $MergeColumns) { $mergeLookup[$mergeColumn] = $true }

        $markdownLines.Add('```{=openxml}')
        $tableWidthXml = if ($AutofitToWindow.IsPresent) {
            '<w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="autofit"/>'
        } else {
            '<w:tblW w:w="' + [string]$tableTotalWidth + '" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        }
        $markdownLines.Add('<w:tbl>')
        $markdownLines.Add('<w:tblPr><w:tblStyle w:val="NetAppTable1"/>' + $tableWidthXml + '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>')
        $markdownLines.Add('<w:tblGrid>' + $tableGridXml + '</w:tblGrid>')

        $headerXml = '<w:tr><w:trPr><w:tblHeader/></w:trPr>'
        for ($headerIndex = 0; $headerIndex -lt $columnCount; $headerIndex++) {
            $headerXml += '<w:tc><w:tcPr><w:tcW w:w="' + [string]$GridWidths[$headerIndex] + '" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableHeading"/></w:pPr><w:r><w:t xml:space="preserve">' + [System.Security.SecurityElement]::Escape([string]$Headers[$headerIndex]) + '</w:t></w:r></w:p></w:tc>'
        }
        $markdownLines.Add($headerXml + '</w:tr>')

        $rowsToRender = @($Rows)
        if ($rowsToRender.Count -eq 0) { $rowsToRender = ,(@(1..$columnCount | ForEach-Object { "N/A" })) }
        foreach ($rowObject in $rowsToRender) {
            $cells = @($rowObject)
            while ($cells.Count -lt $columnCount) { $cells += $null }
            $rowXml = '<w:tr>'
            for ($columnIndex = 0; $columnIndex -lt $columnCount; $columnIndex++) {
                $cellText = [string]$cells[$columnIndex]
                $isMerged = $mergeLookup.ContainsKey($columnIndex)
                $isRestart = -not [string]::IsNullOrWhiteSpace($cellText)
                if ([string]::IsNullOrWhiteSpace($cellText) -and -not $isMerged) { $cellText = "N/A" }
                $tcMergeXml = if ($isMerged) { if ($isRestart) { '<w:vMerge w:val="restart"/>' } else { '<w:vMerge/>' } } else { '' }
                $runsXml = if ($isRestart -or -not $isMerged) { (($cellText -split "`n" | ForEach-Object { '<w:r><w:t xml:space="preserve">' + [System.Security.SecurityElement]::Escape($_) + '</w:t></w:r>' }) -join '<w:r><w:br/></w:r>') } else { '' }
                $rowXml += '<w:tc><w:tcPr><w:tcW w:w="' + [string]$GridWidths[$columnIndex] + '" w:type="dxa"/>' + $tcMergeXml + '</w:tcPr><w:p><w:pPr><w:pStyle w:val="TableText"/></w:pPr>' + $runsXml + '</w:p></w:tc>'
            }
            $markdownLines.Add($rowXml + '</w:tr>')
        }
        $markdownLines.Add('</w:tbl>')
        $markdownLines.Add('```')
    }

    $markdownLines = New-Object System.Collections.Generic.List[string]

    if ($IncludeTitlePage) {
        $reportDateDisplay = (Get-Date $Iso8601).ToString("dd/MM/yyyy")
        $systemNameForTitle = $gridName
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
        & $emitOpenXmlParagraph -StyleId "BodyText" -Text "This document provides a comprehensive configuration summary of the NetApp StorageGRID system for which this report was generated."
        $markdownLines.Add('<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>')
        $markdownLines.Add('```')
        $markdownLines.Add("")

        if ($EnableDocxToc) {
            $markdownLines.Add('```{=openxml}')
            & $emitOpenXmlParagraph -StyleId "TOCHeading" -Text $reportMetadata["toc-title"]
            $markdownLines.Add('<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/><w:instrText xml:space="preserve">TOC \o "1-' + $DocxTocDepth + '" \h \z \u</w:instrText><w:fldChar w:fldCharType="separate"/><w:fldChar w:fldCharType="end"/></w:r></w:p>')
            $markdownLines.Add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
            $markdownLines.Add('```')
            $markdownLines.Add("")
        }
    }

    # --- Document Version Information ---
    $markdownLines.Add("# Document Version Information")
    $markdownLines.Add("")
    $markdownLines.Add("This table records the date and time this as-built report was generated from the live grid, along with the version of the collection script used, so the report's contents can always be correlated back to the exact state of the grid at that point in time.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Grid Name", "Data Reference Date", "Script Version") -Rows @(
        ,@($gridName, $Iso8601, $ScriptVersion)
    )

    $markdownLines.Add("")
    $markdownLines.Add('```{=openxml}')
    $markdownLines.Add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    $markdownLines.Add('```')

    # --- Pre-compute derived display values ---
    $uniqueSiteIds = @(
        $nodeHealth |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string](Get-PropertyValue -Object $_ -PropertyName "siteId")) } |
            ForEach-Object { [string](Get-PropertyValue -Object $_ -PropertyName "siteId") } |
            Sort-Object -Unique
    )

    $softwareVersion      = [string](Get-PropertyValue -Object $productVersion -PropertyName "productVersion")
    $gridDisplayName      = [string](Get-PropertyValue -Object $gridLicense -PropertyName "displayName")
    if ([string]::IsNullOrWhiteSpace($gridDisplayName)) { $gridDisplayName = $gridName }
    
    # Normalize counts - ensure we can safely call .Count on arrays and scalars
    $siteCountValue       = $uniqueSiteIds.Count
    
    # Display values - normalize to array to handle both scalar and array returns from API
    $dnsItems             = [array](@($dnsServers) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $dnsDisplay           = if ($dnsItems.Count -gt 0) { ($dnsItems | ForEach-Object { [string]$_ }) -join "<br>" } else { "N/A" }
    
    $ntpItems             = [array](@($ntpServers) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $ntpDisplay           = if ($ntpItems.Count -gt 0) { ($ntpItems | ForEach-Object { [string]$_ }) -join "<br>" } else { "N/A" }
    
    $gridNetItems         = [array](@($gridNetworks) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $gridNetDisplay       = if ($gridNetItems.Count -gt 0) { ($gridNetItems | ForEach-Object { [string]$_ }) -join "<br>" } else { "N/A" }
    
    $guiTimeoutRaw        = Get-PropertyValue -Object $browserTimeout -PropertyName "browserInactivityTimeout"

    # --- Overview ---
    $markdownLines.Add("")
    $markdownLines.Add("# Overview")
    $markdownLines.Add("")
    $markdownLines.Add("This report will provide the following information:")
    $markdownLines.Add("")
    $markdownLines.Add("- Hardware inventory")
    $markdownLines.Add("- Firmware and software revisions")
    $markdownLines.Add("- StorageGRID software configuration details")
    $markdownLines.Add("- Storage capacity and provisioning")
    $markdownLines.Add("- Secure multi-tenant configuration")
    $markdownLines.Add("- Network configuration")
    $markdownLines.Add("- ILM configuration")
    $markdownLines.Add("- Reference material")

    # --- Grid Overview ---
        $markdownLines.Add("")
    $markdownLines.Add("# Grid Overview")
    $markdownLines.Add("")
    $markdownLines.Add("StorageGRID is NetApp's software-defined object storage platform, deployed as a grid of interconnected virtual or appliance-based nodes that can span one or more physical data centers, called sites. This table summarizes the identity and scale of the deployment covered by this report: the grid name, the installed StorageGRID software version, and the number of sites, nodes, and tenant accounts currently configured.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Grid Name", "StorageGRID Version", "Site Count", "Node Count", "Tenant Count") -Rows @(
        ,@($gridName, $softwareVersion, $siteCountValue, @($nodeHealth).Count, @($gridAccounts).Count)
    )

    # --- Grid Passwords ---
    $markdownLines.Add("")
    $markdownLines.Add("## Grid Passwords")
    $markdownLines.Add("")
    $markdownLines.Add("StorageGRID requires the Provisioning Passphrase, which is defined during installation, to facilitate any topology changes to the grid, such as adding/removing nodes and sites, routing changes and applying software updates. Any such change requires the generation of an updated Grid Recovery Package, which is a ZIP file that contains critical recovery data for the grid and is recommended to be stored in a secure location in a minimum of two data centers for safety.")
    $markdownLines.Add("")
    $markdownLines.Add("StorageGRID has a default user, root, which is the default user with access to the grid management interface. The password for this user is entered during the grid installation.")
    $markdownLines.Add("")
    $markdownLines.Add("Each grid node also has an admin user used for SSH access to the node for maintenance and troubleshooting. These passwords can be randomly generated by the system and are stored in the Passwords.txt file in the Grid Recovery Package.")

    $markdownLines.Add("")
    # --- Health Summary ---
    $markdownLines.Add("## Health Summary")
    $markdownLines.Add("")
    $markdownLines.Add("This health summary shows the number of active alerts and node connection states across the grid at the time this report was generated. Alerts are StorageGRID's primary tool for monitoring operational issues and are triggered at Critical, Major, or Minor severity when an alert rule's condition evaluates as true. Node connection states show whether each grid node is currently connected, administratively down (intentionally taken offline for maintenance), or in an unknown state.")
    $markdownLines.Add("")
    $gridAlerts = Get-PropertyValue -Object $gridHealth -PropertyName "alerts"
    $gridHealthNodes = Get-PropertyValue -Object $gridHealth -PropertyName "nodes"
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Category", "Metric", "Count") -Rows @(
        @("Alerts", "Critical", (& $boolText (Get-PropertyValue -Object $gridAlerts -PropertyName "critical"))),
        @("Alerts", "Major", (& $boolText (Get-PropertyValue -Object $gridAlerts -PropertyName "major"))),
        @("Alerts", "Minor", (& $boolText (Get-PropertyValue -Object $gridAlerts -PropertyName "minor"))),
        @("Nodes", "Connected", (& $boolText (Get-PropertyValue -Object $gridHealthNodes -PropertyName "connected"))),
        @("Nodes", "Administratively Down", (& $boolText (Get-PropertyValue -Object $gridHealthNodes -PropertyName "administratively-down"))),
        @("Nodes", "Unknown", (& $boolText (Get-PropertyValue -Object $gridHealthNodes -PropertyName "unknown")))
    )

    $markdownLines.Add("")
    # --- Configuration Details ---
    $markdownLines.Add("## Configuration Details")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows core grid-wide network configuration. StorageGRID requires accurate, synchronized time across every node for TLS certificate validation, object metadata consistency, and audit log correlation, so the configured Network Time Protocol (NTP) servers are a foundational dependency for grid health. DNS servers are used by nodes to resolve external hostnames, such as identity providers, key management servers, and NTP servers. The Grid Networks list shows the CIDR subnets that make up the trusted internal Grid Network (eth0), which nodes use for node-to-node and intra-grid communication.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Grid Name", "Display Name", "DNS Servers", "NTP Servers", "Grid Networks") -Rows @(
        ,@($gridName, $gridDisplayName, $dnsDisplay, $ntpDisplay, $gridNetDisplay)
    )

    # --- Grid License ---
    $markdownLines.Add("")
    $markdownLines.Add("## Grid License")
    $markdownLines.Add("")
    
    $licensePackage = Get-PropertyValue -Object $gridLicense -PropertyName "licensePackage"
    $licenseType = Get-PropertyValue -Object $gridLicense -PropertyName "licenseType"
    $licenseCapacity = Get-PropertyValue -Object $gridLicense -PropertyName "licenseCapacity"
    $supportCapacity = Get-PropertyValue -Object $gridLicense -PropertyName "supportCapacity"
    $licenseEnd = Get-PropertyValue -Object $gridLicense -PropertyName "licenseEnd"
    $supportEnd = Get-PropertyValue -Object $gridLicense -PropertyName "supportEnd"
    
    $licenseCapacityDisplay = if ($licenseCapacity -gt 0) { "$licenseCapacity TB" } else { "Unlimited / Not Specified" }
    $supportCapacityDisplay = if ($supportCapacity -gt 0) { "$supportCapacity TB" } else { "N/A" }
    
    $markdownLines.Add("StorageGRID requires a valid license to activate object storage capacity on the grid. The license defines the license package (feature tier), the license type, the total storage capacity purchased, and the license and support expiration dates. Exceeding the licensed capacity does not prevent the grid from ingesting new objects, but NetApp recommends monitoring actual usage against the licensed capacity so that any required license or support renewal can be planned ahead of the expiration date.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Property", "Value") -Rows @(
        @("License Package", $licensePackage),
        @("License Type", $licenseType),
        @("Licensed Capacity", $licenseCapacityDisplay),
        @("Support Capacity", $supportCapacityDisplay),
        @("License Expiration", $licenseEnd),
        @("Support Expiration", $supportEnd)
    )

    $gridCapacityMetrics = Get-PropertyValue -Object $capacityMetrics -PropertyName "grid"
    $siteCapacityMetrics = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $capacityMetrics -PropertyName "sites")
    $nodeCapacityMetrics = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $capacityMetrics -PropertyName "nodes")
    $nodeCapacityMetrics = @(
        $nodeCapacityMetrics | ForEach-Object { Repair-StorageGridNodeObjectCapacityRecord -Record $_ }
    )
    $hasGridCapacityMetrics = $null -ne $gridCapacityMetrics -and (
        $null -ne (Get-PropertyValue -Object $gridCapacityMetrics -PropertyName "overallTotalBytes") -or
        $null -ne (Get-PropertyValue -Object $gridCapacityMetrics -PropertyName "objectTotalBytes") -or
        $null -ne (Get-PropertyValue -Object $gridCapacityMetrics -PropertyName "metadataAllowedBytes")
    )

    $fallbackNodeObjectCapacityRows = @(
        $nodeCapacityMetrics | Where-Object { -not [string]::IsNullOrWhiteSpace([string](Get-PropertyValue -Object $_ -PropertyName "objectFallbackSource")) }
    )

    # --- Grid Sites ---
    $sortedNodeHealth = @(
        $nodeHealth | Sort-Object -Property `
            @{ Expression = { [string](Get-PropertyValue -Object $_ -PropertyName "siteName") } }, `
            @{ Expression = { [string](Get-PropertyValue -Object $_ -PropertyName "name") } }
    )
    $networkNodes = @((Get-PropertyValue -Object $networkTopology -PropertyName "gridNodes"))
    $sites = @($nodeHealth | ForEach-Object { Get-PropertyValue -Object $_ -PropertyName "siteName" } | Sort-Object -Unique)

    $markdownLines.Add("")
    $markdownLines.Add("## Grid Sites")
    $markdownLines.Add("")
    $markdownLines.Add("A StorageGRID site represents a physical or logical data center location containing one or more grid nodes. Sites are a key building block for data protection and ILM placement: erasure coding and multi-copy replication rules can be configured to store object copies across multiple sites for geographic redundancy. This table shows each site's display name, its link cost group (used to influence routing preference between sites when multiple paths exist), and the nodes located at that site.")
    $markdownLines.Add("")

    $siteRowsDetailed = @()
    foreach ($site in ($sites | Sort-Object)) {
        $siteNodes = @($sortedNodeHealth | Where-Object { [string](Get-PropertyValue -Object $_ -PropertyName "siteName") -eq $site })
        $siteDisplayName = ""
        if ($siteNodes.Count -gt 0) {
            $siteDisplayCandidate = [string](Get-PropertyValue -Object $siteNodes[0] -PropertyName "siteDisplayName")
            if (-not [string]::IsNullOrWhiteSpace($siteDisplayCandidate)) { $siteDisplayName = $siteDisplayCandidate }
        }

        $siteMembers = @(
            $siteNodes |
                ForEach-Object { [string](Get-PropertyValue -Object $_ -PropertyName "name") } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object
        )
        $siteMembersDisplay = if ($siteMembers.Count -gt 0) { $siteMembers -join "`n" } else { "N/A" }

        $siteNodeIds = @(
            $siteNodes |
                ForEach-Object { [string](Get-PropertyValue -Object $_ -PropertyName "id") } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        $siteLinkCosts = @()
        foreach ($nn in $networkNodes) {
            $cfg = Get-PropertyValue -Object $nn -PropertyName "nodeConfig"
            if ($null -eq $cfg) { continue }
            $cfgId = [string](Get-PropertyValue -Object $cfg -PropertyName "nodeId")
            if ($siteNodeIds -contains $cfgId) {
                $lc = [string](Get-PropertyValue -Object $cfg -PropertyName "linkCostGroup")
                if (-not [string]::IsNullOrWhiteSpace($lc)) { $siteLinkCosts += $lc }
            }
        }
        $siteLinkCosts = @($siteLinkCosts | Sort-Object -Unique)
        $siteLinkCostDisplay = if ($siteLinkCosts.Count -gt 0) { $siteLinkCosts -join ", " } else { "N/A" }

        $siteRowsDetailed += ,@([string]$site, [string]$siteDisplayName, [string]$siteLinkCostDisplay, [string]$siteMembersDisplay)
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Site Name", "Display Name", "Link Cost Group", "Site Members") -Rows $siteRowsDetailed

    $markdownLines.Add("")
    # --- Grid Capacity ---
    $markdownLines.Add("## Grid Capacity")
    $markdownLines.Add("")
    if ($hasGridCapacityMetrics) {
        $markdownLines.Add("StorageGRID stores object data and object metadata separately: object data (replicated copies and erasure-coded fragments) consumes the bulk of usable storage space, while a fixed amount of space on each Storage Node's first storage volume is reserved for object metadata in a distributed Cassandra database. This table shows validated capacity values derived from StorageGRID's Prometheus storage-utilization metrics: Overall Grid reflects total usable space across all storage volumes, Objects reflects space consumed specifically by object data, and Metadata reflects the allowed and used space for the object metadata database. Monitoring these values over time helps determine when to expand the grid with additional Storage Nodes or storage volumes before usable capacity is exhausted.")
        $markdownLines.Add("")
        $markdownLines.Add("This table shows the total capacity for the Grid.")
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Capacity Type", "Total (TiB)", "Used (TiB)", "Available (TiB)", "Used (%)") -Rows @(
            @(
                "Overall Grid",
                (& $capacityBytesText (Get-PropertyValue -Object $gridCapacityMetrics -PropertyName "overallTotalBytes")),
                (& $capacityBytesText (Get-PropertyValue -Object $gridCapacityMetrics -PropertyName "overallUsedBytes")),
                (& $capacityBytesText (Get-PropertyValue -Object $gridCapacityMetrics -PropertyName "overallAvailableBytes")),
                (& $capacityPercentText (Get-PropertyValue -Object $gridCapacityMetrics -PropertyName "overallUsedPct"))
            ),
            @(
                "Objects",
                (& $capacityObjectBytesText $gridCapacityMetrics "objectTotalBytes"),
                (& $capacityObjectBytesText $gridCapacityMetrics "objectUsedBytes"),
                (& $capacityObjectBytesText $gridCapacityMetrics "objectAvailableBytes"),
                (& $capacityObjectPercentText $gridCapacityMetrics "objectUsedPct")
            ),
            @(
                "Metadata",
                (& $capacityBytesText (Get-PropertyValue -Object $gridCapacityMetrics -PropertyName "metadataAllowedBytes")),
                (& $capacityBytesText (Get-PropertyValue -Object $gridCapacityMetrics -PropertyName "metadataUsedBytes")),
                (& $capacityBytesText (Get-PropertyValue -Object $gridCapacityMetrics -PropertyName "metadataAvailableBytes")),
                (& $capacityPercentText (Get-PropertyValue -Object $gridCapacityMetrics -PropertyName "metadataUsedPct"))
            )
        )
    }
    else {
        $markdownLines.Add("Validated storage-utilization metrics were not available at collection time.")
    }

    if (@($siteCapacityMetrics).Count -gt 0) {
        $markdownLines.Add("")
        # --- Site Storage Capacity ---
        $markdownLines.Add("## Site Storage Capacity")
        $markdownLines.Add("")
        $markdownLines.Add("The following tables break down the same Total, Object, and Metadata capacity values by site, so that capacity trends can be compared across data centers. Because StorageGRID stores a full copy of all object metadata at every site, the site with the least available metadata capacity effectively limits the object-count ceiling for the entire grid, even if other sites have more free space.")
        $markdownLines.Add("")

        $siteTotalCapacityRows = @(
            $siteCapacityMetrics | Sort-Object -Property name | ForEach-Object {
                ,@(
                    [string](Get-PropertyValue -Object $_ -PropertyName "name"),
                    (& $capacityBytesText (Get-PropertyValue -Object $_ -PropertyName "overallTotalBytes")),
                    (& $capacityBytesText (Get-PropertyValue -Object $_ -PropertyName "overallUsedBytes")),
                    (& $capacityBytesText (Get-PropertyValue -Object $_ -PropertyName "overallAvailableBytes")),
                    (& $capacityPercentText (Get-PropertyValue -Object $_ -PropertyName "overallUsedPct"))
                )
            }
        )
        # --- Site Total Capacity ---
        $markdownLines.Add("### Total Capacity")
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Site", "Total (TiB)", "Used (TiB)", "Available (TiB)", "Used (%)") -Rows $siteTotalCapacityRows

        $siteObjectCapacityRows = @(
            $siteCapacityMetrics | Sort-Object -Property name | ForEach-Object {
                ,@(
                    [string](Get-PropertyValue -Object $_ -PropertyName "name"),
                    (& $capacityObjectBytesText $_ "objectTotalBytes"),
                    (& $capacityObjectBytesText $_ "objectUsedBytes"),
                    (& $capacityObjectBytesText $_ "objectAvailableBytes"),
                    (& $capacityObjectPercentText $_ "objectUsedPct")
                )
            }
        )
        $markdownLines.Add("")
        # --- Site Object Capacity ---
        $markdownLines.Add("### Object Capacity")
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Site", "Total (TiB)", "Used (TiB)", "Available (TiB)", "Used (%)") -Rows $siteObjectCapacityRows

        $siteMetadataCapacityRows = @(
            $siteCapacityMetrics | Sort-Object -Property name | ForEach-Object {
                ,@(
                    [string](Get-PropertyValue -Object $_ -PropertyName "name"),
                    (& $capacityBytesText (Get-PropertyValue -Object $_ -PropertyName "metadataAllowedBytes")),
                    (& $capacityBytesText (Get-PropertyValue -Object $_ -PropertyName "metadataUsedBytes")),
                    (& $capacityBytesText (Get-PropertyValue -Object $_ -PropertyName "metadataAvailableBytes")),
                    (& $capacityPercentText (Get-PropertyValue -Object $_ -PropertyName "metadataUsedPct"))
                )
            }
        )
        $markdownLines.Add("")
        # --- Site Metadata Capacity ---
        $markdownLines.Add("### Metadata Capacity")
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Site", "Allowed (TiB)", "Used (TiB)", "Available (TiB)", "Used (%)") -Rows $siteMetadataCapacityRows
    }

    if (@($nodeCapacityMetrics).Count -gt 0) {
        $markdownLines.Add("")
        # --- Node Storage Capacity ---
        $markdownLines.Add("## Node Storage Capacity")
        $markdownLines.Add("")
        $markdownLines.Add("The following tables show the same capacity metrics broken down per Storage Node, using the Prometheus node label exposed by the grid. Reviewing capacity at the node level helps identify individual nodes that are approaching their usable space or metadata limits, since object metadata is distributed evenly across all Storage Nodes at a site and the smallest or fullest node can constrain the capacity of the entire site.")
        $markdownLines.Add("")

        $nodeTotalCapacityRows = @(
            $nodeCapacityMetrics | Sort-Object -Property siteName, name | ForEach-Object {
                ,@(
                    [string](Get-PropertyValue -Object $_ -PropertyName "siteName"),
                    [string](Get-PropertyValue -Object $_ -PropertyName "name"),
                    [string](Get-PropertyValue -Object $_ -PropertyName "labelValue"),
                    (& $capacityBytesText (Get-PropertyValue -Object $_ -PropertyName "overallTotalBytes")),
                    (& $capacityBytesText (Get-PropertyValue -Object $_ -PropertyName "overallUsedBytes")),
                    (& $capacityBytesText (Get-PropertyValue -Object $_ -PropertyName "overallAvailableBytes")),
                    (& $capacityPercentText (Get-PropertyValue -Object $_ -PropertyName "overallUsedPct"))
                )
            }
        )
        # --- Node Total Capacity ---
        $markdownLines.Add("### Total Capacity")
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Site", "Node", "Metric Label", "Total (TiB)", "Used (TiB)", "Available (TiB)", "Used (%)") -Rows $nodeTotalCapacityRows

        $nodeObjectCapacityRows = @(
            $nodeCapacityMetrics | Sort-Object -Property siteName, name | ForEach-Object {
                ,@(
                    [string](Get-PropertyValue -Object $_ -PropertyName "siteName"),
                    [string](Get-PropertyValue -Object $_ -PropertyName "name"),
                    [string](Get-PropertyValue -Object $_ -PropertyName "labelValue"),
                    (& $capacityObjectBytesText $_ "objectTotalBytes" $true),
                    (& $capacityObjectBytesText $_ "objectUsedBytes"),
                    (& $capacityObjectBytesText $_ "objectAvailableBytes" $true),
                    (& $capacityObjectPercentText $_ "objectUsedPct" $true)
                )
            }
        )
        $markdownLines.Add("")
        # --- Node Object Capacity ---
        $markdownLines.Add("### Object Capacity")
        $markdownLines.Add("")
        if (@($fallbackNodeObjectCapacityRows).Count -gt 0) {
            $markdownLines.Add("Node rows whose object usable-space metrics were inconsistent now fall back to node total-space metrics for Object Total, Available, and Used (%). This preserves 11.9 node reporting without showing impossible percentages.")
            $markdownLines.Add("")
        }
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Site", "Node", "Metric Label", "Total (TiB)", "Used (TiB)", "Available (TiB)", "Used (%)") -Rows $nodeObjectCapacityRows

        $nodeMetadataCapacityRows = @(
            $nodeCapacityMetrics | Sort-Object -Property siteName, name | ForEach-Object {
                ,@(
                    [string](Get-PropertyValue -Object $_ -PropertyName "siteName"),
                    [string](Get-PropertyValue -Object $_ -PropertyName "name"),
                    [string](Get-PropertyValue -Object $_ -PropertyName "labelValue"),
                    (& $capacityBytesText (Get-PropertyValue -Object $_ -PropertyName "metadataAllowedBytes")),
                    (& $capacityBytesText (Get-PropertyValue -Object $_ -PropertyName "metadataUsedBytes")),
                    (& $capacityBytesText (Get-PropertyValue -Object $_ -PropertyName "metadataAvailableBytes")),
                    (& $capacityPercentText (Get-PropertyValue -Object $_ -PropertyName "metadataUsedPct"))
                )
            }
        )
        $markdownLines.Add("")
        # --- Node Metadata Capacity ---
        $markdownLines.Add("### Metadata Capacity")
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Site", "Node", "Metric Label", "Allowed (TiB)", "Used (TiB)", "Available (TiB)", "Used (%)") -Rows $nodeMetadataCapacityRows
    }

    # --- Grid Nodes ---
    $markdownLines.Add("")
    $markdownLines.Add("# Grid Nodes")
    $markdownLines.Add("")
    
    # Get unique sites from node health data
    $sites = @($nodeHealth | ForEach-Object { Get-PropertyValue -Object $_ -PropertyName "siteName" } | Sort-Object -Unique)
    
    # Build site overview table
    $markdownLines.Add("The following table shows the grid sites and node type distribution. A StorageGRID node is a virtual machine, container, or dedicated appliance running one or more StorageGRID services.")
    $markdownLines.Add("")
    $markdownLines.Add("Admin Nodes host the Grid Manager and Tenant Manager interfaces and the Grid Management API; Storage Nodes store object data and metadata; Gateway Nodes run the Load Balancer service for S3 client traffic; and Archive Nodes (deprecated) target external archival storage.")
    $markdownLines.Add("")
    $markdownLines.Add("A well-distributed mix of node types per site supports redundancy and load distribution.")
    $markdownLines.Add("")
    
    $siteRows = @()
    foreach ($site in $sites) {
        $siteNodes = @($nodeHealth | Where-Object { (Get-PropertyValue -Object $_ -PropertyName "siteName") -eq $site })
        $adminNodeCount = @($siteNodes | Where-Object { (Get-PropertyValue -Object $_ -PropertyName "type") -eq "adminNode" }).Count
        $storageNodeCount = @($siteNodes | Where-Object { (Get-PropertyValue -Object $_ -PropertyName "type") -eq "storageNode" }).Count
        $gatewayNodeCount = @($siteNodes | Where-Object { (Get-PropertyValue -Object $_ -PropertyName "type") -eq "apiGatewayNode" }).Count
        $archiveNodeCount = @($siteNodes | Where-Object { (Get-PropertyValue -Object $_ -PropertyName "type") -eq "archiveNode" }).Count
        $totalCount = $siteNodes.Count
        
        $siteRows += ,@($site, $totalCount, $adminNodeCount, $storageNodeCount, $gatewayNodeCount, $archiveNodeCount)
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Site Name", "Total Nodes", "Admin", "Storage", "Gateway", "Archive") -Rows $siteRows

    # --- Grid Node Details ---
    $markdownLines.Add("")
    $markdownLines.Add("## Node Details")
    $markdownLines.Add("")
    $markdownLines.Add("This table provides detailed information for each node in the grid, including its site, role, and the platform type (StorageGRID Appliance or vSphere or Baremetal)")
    $markdownLines.Add("")

    $sortedNodeHealth = @(
        $nodeHealth | Sort-Object -Property `
            @{ Expression = { [string](Get-PropertyValue -Object $_ -PropertyName "siteName") } }, `
            @{ Expression = { [string](Get-PropertyValue -Object $_ -PropertyName "name") } }
    )

    $networkNodes = @((Get-PropertyValue -Object $networkTopology -PropertyName "gridNodes"))
    $networkById = @{}
    $networkByHostname = @{}
    foreach ($netNode in $networkNodes) {
        $nodeConfig = Get-PropertyValue -Object $netNode -PropertyName "nodeConfig"
        if ($null -eq $nodeConfig) { continue }

        $cfgNodeId = [string](Get-PropertyValue -Object $nodeConfig -PropertyName "nodeId")
        $cfgHostname = [string](Get-PropertyValue -Object $nodeConfig -PropertyName "hostname")

        if (-not [string]::IsNullOrWhiteSpace($cfgNodeId)) { $networkById[$cfgNodeId] = $nodeConfig }
        if (-not [string]::IsNullOrWhiteSpace($cfgHostname)) { $networkByHostname[$cfgHostname] = $nodeConfig }
    }

    $nodeRows = @()
    $nodeIpRows = @()
    $nodeNetworkRouteRows = @()
    foreach ($node in $sortedNodeHealth) {
        $nodeId = [string](Get-PropertyValue -Object $node -PropertyName "id")
        $nodeName = [string](Get-PropertyValue -Object $node -PropertyName "name")
        $siteName = [string](Get-PropertyValue -Object $node -PropertyName "siteName")
        $role = [string](Get-PropertyValue -Object $node -PropertyName "type")
        $isPrimaryAdmin = [bool](Get-PropertyValue -Object $node -PropertyName "isPrimaryAdmin")
        $typeDisplay = switch ($role) {
            "adminNode" { "Admin" }
            "storageNode" { "Storage" }
            "apiGatewayNode" { "Gateway" }
            "archiveNode" { "Archive" }
            default { $role }
        }
        $nodeTypeDisplay = if ($isPrimaryAdmin) { "$typeDisplay (Primary)" } else { $typeDisplay }

        $nodeConfig = $null
        if (-not [string]::IsNullOrWhiteSpace($nodeId) -and $networkById.ContainsKey($nodeId)) {
            $nodeConfig = $networkById[$nodeId]
        } elseif (-not [string]::IsNullOrWhiteSpace($nodeName) -and $networkByHostname.ContainsKey($nodeName)) {
            $nodeConfig = $networkByHostname[$nodeName]
        }

        $platform = "N/A"
        $nodeTypeDetail = "N/A"
        $gridNetworksDetail = "N/A"
        $gridIp = "N/A"
        $adminIp = "N/A"
        $clientIp = "N/A"
        $adminNetworksDetail = "N/A"
        $clientNetworksDetail = "N/A"

        if ($null -ne $nodeConfig) {
            $platformRaw = Get-PropertyValue -Object $nodeConfig -PropertyName "platform"
            if (-not [string]::IsNullOrWhiteSpace([string]$platformRaw)) { $platform = [string]$platformRaw }

            $nodeTypeRaw = Get-PropertyValue -Object $nodeConfig -PropertyName "nodeType"
            if (-not [string]::IsNullOrWhiteSpace([string]$nodeTypeRaw)) { $nodeTypeDetail = [string]$nodeTypeRaw }

            $networking = Get-PropertyValue -Object $nodeConfig -PropertyName "networking"
            if ($null -ne $networking) {
                $gridNet = Get-PropertyValue -Object $networking -PropertyName "grid"
                $adminNet = Get-PropertyValue -Object $networking -PropertyName "admin"
                $clientNet = Get-PropertyValue -Object $networking -PropertyName "client"

                $routeSubnets = @()
                if ($null -ne $gridNet) {
                    $routes = @((Get-PropertyValue -Object $gridNet -PropertyName "routes"))
                    foreach ($route in $routes) {
                        $subnets = @((Get-PropertyValue -Object $route -PropertyName "subnets"))
                        foreach ($subnet in $subnets) {
                            $subnetText = [string]$subnet
                            if (-not [string]::IsNullOrWhiteSpace($subnetText)) {
                                $routeSubnets += $subnetText
                            }
                        }
                    }
                }
                $routeSubnets = @($routeSubnets | Sort-Object -Unique)
                if ($routeSubnets.Count -gt 0) {
                    $gridNetworksDetail = $routeSubnets -join "`n"
                } else {
                    $gridCidrFallback = [string](Get-PropertyValue -Object $gridNet -PropertyName "cidr")
                    if (-not [string]::IsNullOrWhiteSpace($gridCidrFallback)) { $gridNetworksDetail = $gridCidrFallback }
                }

                $gridCidr = [string](Get-PropertyValue -Object $gridNet -PropertyName "cidr")
                $adminCidr = [string](Get-PropertyValue -Object $adminNet -PropertyName "cidr")
                $clientCidr = [string](Get-PropertyValue -Object $clientNet -PropertyName "cidr")

                if (-not [string]::IsNullOrWhiteSpace($gridCidr)) { $gridIp = $gridCidr }
                if (-not [string]::IsNullOrWhiteSpace($adminCidr)) { $adminIp = $adminCidr }
                if (-not [string]::IsNullOrWhiteSpace($clientCidr)) { $clientIp = $clientCidr }

                $adminRouteSubnets = @()
                if ($null -ne $adminNet) {
                    foreach ($adminRoute in @((Get-PropertyValue -Object $adminNet -PropertyName "routes"))) {
                        foreach ($adminSubnet in @((Get-PropertyValue -Object $adminRoute -PropertyName "subnets"))) {
                            $adminSubnetText = [string]$adminSubnet
                            if (-not [string]::IsNullOrWhiteSpace($adminSubnetText)) {
                                $adminRouteSubnets += $adminSubnetText
                            }
                        }
                    }
                }
                $adminRouteSubnets = @($adminRouteSubnets | Sort-Object -Unique)
                if ($adminRouteSubnets.Count -gt 0) {
                    $adminNetworksDetail = $adminRouteSubnets -join "`n"
                } elseif (-not [string]::IsNullOrWhiteSpace($adminCidr)) {
                    $adminNetworksDetail = $adminCidr
                }

                $clientRouteSubnets = @()
                if ($null -ne $clientNet) {
                    foreach ($clientRoute in @((Get-PropertyValue -Object $clientNet -PropertyName "routes"))) {
                        foreach ($clientSubnet in @((Get-PropertyValue -Object $clientRoute -PropertyName "subnets"))) {
                            $clientSubnetText = [string]$clientSubnet
                            if (-not [string]::IsNullOrWhiteSpace($clientSubnetText)) {
                                $clientRouteSubnets += $clientSubnetText
                            }
                        }
                    }
                }
                $clientRouteSubnets = @($clientRouteSubnets | Sort-Object -Unique)
                if ($clientRouteSubnets.Count -gt 0) {
                    $clientNetworksDetail = $clientRouteSubnets -join "`n"
                }
            }
        }

        $nodeRows += ,@(
            [string]$nodeName,
            [string]$siteName,
            [string]$nodeTypeDisplay,
            [string]$nodeTypeDetail,
            [string]$platform
        )

        $nodeIpRows += ,@(
            [string]$nodeName,
            [string]$siteName,
            [string]$gridIp,
            [string]$adminIp,
            [string]$clientIp
        )

        $nodeNetworkRouteRows += ,@(
            [string]$nodeName,
            [string]$siteName,
            [string]$gridNetworksDetail,
            [string]$adminNetworksDetail,
            [string]$clientNetworksDetail
        )
    }

    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Node Name", "Site", "Node Type", "Node Role", "Platform") -Rows $nodeRows

    $markdownLines.Add("")
    # --- Node Networks ---
    $markdownLines.Add("## Node Networks")
    $markdownLines.Add("")
    # --- Node IP Addresses ---
    $markdownLines.Add("### Node IP Addresses")
    $markdownLines.Add("")
    $markdownLines.Add("The Node IP Addresses table shows the address assigned to each node on the Grid, Admin, and Client networks. A value is shown only in the column for the network to which it belongs; unavailable or unconfigured networks are reported as N/A.")
    $markdownLines.Add("")
    $markdownLines.Add("- Grid Network (eth0): Used for all internal, trusted communication between nodes and required for every node. The Grid Network is each node's default route (0.0.0.0/0) when the Client Network is not configured. ")
    $markdownLines.Add("- Admin Network (eth1): This is an optional network typically used for administrative and maintenance access;")
    $markdownLines.Add("- Client Network (eth2): This is an optional network used for client (S3) application traffic so it can be isolated from internal grid traffic. The client network becomes the default route for the individual node when configured")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Node", "Site", "Grid", "Admin", "Client") -Rows $nodeIpRows

    $markdownLines.Add("")
    # --- Node Network Routes ---
    $markdownLines.Add("### Node Network Routes")
    $markdownLines.Add("")
    $markdownLines.Add("The Node Network Routes table lists the routed network subnets configured for each node. Grid Networks are the routed Grid Network subnets, Admin Networks are the Admin Network route subnets, and Client Networks are the Client Network route subnets. A Client Network default route is shown as 0.0.0.0/0; multiple values are separated by soft returns.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Node", "Site", "Grid Networks", "Admin Networks", "Client Networks") -Rows $nodeNetworkRouteRows

    # --- Grid Appliances ---
    $markdownLines.Add("")
    $markdownLines.Add("## Grid Appliances")
    $markdownLines.Add("")
    $markdownLines.Add("StorageGRID appliances are purpose-built hardware platforms (the SG1xx/SG1xxx series for Admin and Gateway services (referred to as Services Appliances), and the SG57xx / SG58xx / SG6xxx / SGF6xxx series for Storage Nodes that combine a compute controller running StorageGRID software with, for storage appliances, either an embedded E-Series storage controller managed by NetApp SANtricity software or internal SSD storage.")
    $markdownLines.Add("")
    $markdownLines.Add("This table identifies each appliance node's hardware model and the serial numbers of its compute controller and (for E-Series based storage appliances) its attached E-Series storage controller, along with the appliance's baseboard management controller (BMC) IP address used for out-of-band hardware management.")
    $markdownLines.Add("")

    $nodeHealthById = @{}
    foreach ($nh in $sortedNodeHealth) {
        $nhId = [string](Get-PropertyValue -Object $nh -PropertyName "id")
        if (-not [string]::IsNullOrWhiteSpace($nhId)) {
            $nodeHealthById[$nhId] = $nh
        }
    }

    $applianceRows = @()
    $applianceStorageRows = @()
    $applianceControllerRows = @()
    $applianceShelfRows = @()
    $applianceStorageByNode = @{}
    $applianceStorageNodes = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $applianceStorageMetrics -PropertyName "nodes")
    foreach ($metricNode in @($applianceStorageNodes)) {
        $metricNodeName = [string](Get-PropertyValue -Object $metricNode -PropertyName "nodeName")
        if ([string]::IsNullOrWhiteSpace($metricNodeName)) { continue }
        $applianceStorageByNode[$metricNodeName.ToLowerInvariant()] = $metricNode
    }

    $nodeAttributeByNodeId = @{}
    $nodeAttributeNodes = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $nodeAttributeDetails -PropertyName "nodes")
    foreach ($attributeNode in @($nodeAttributeNodes)) {
        $attributeNodeId = [string](Get-PropertyValue -Object $attributeNode -PropertyName "nodeId")
        if ([string]::IsNullOrWhiteSpace($attributeNodeId)) { continue }
        $nodeAttributeByNodeId[$attributeNodeId] = $attributeNode
    }

    $applianceNodes = @(
        $networkNodes |
            Where-Object {
                $cfg = Get-PropertyValue -Object $_ -PropertyName "nodeConfig"
                $null -ne $cfg -and ([string](Get-PropertyValue -Object $cfg -PropertyName "platform")) -eq "SGA"
            }
    )

    $sortedApplianceNodes = @(
        $applianceNodes | Sort-Object -Property @(
            @{ Expression = {
                    $cfg = Get-PropertyValue -Object $_ -PropertyName "nodeConfig"
                    $nid = [string](Get-PropertyValue -Object $cfg -PropertyName "nodeId")
                    if ($nodeHealthById.ContainsKey($nid)) {
                        [string](Get-PropertyValue -Object $nodeHealthById[$nid] -PropertyName "siteName")
                    } else {
                        ""
                    }
                }
            },
            @{ Expression = {
                    $cfg = Get-PropertyValue -Object $_ -PropertyName "nodeConfig"
                    [string](Get-PropertyValue -Object $cfg -PropertyName "hostname")
                }
            }
        )
    )

    foreach ($apNode in $sortedApplianceNodes) {
        $cfg = Get-PropertyValue -Object $apNode -PropertyName "nodeConfig"
        if ($null -eq $cfg) { continue }

        $nodeId = [string](Get-PropertyValue -Object $cfg -PropertyName "nodeId")
        $hostName = [string](Get-PropertyValue -Object $cfg -PropertyName "hostname")

        $siteName = "N/A"
        if ($nodeHealthById.ContainsKey($nodeId)) {
            $siteNameRaw = [string](Get-PropertyValue -Object $nodeHealthById[$nodeId] -PropertyName "siteName")
            if (-not [string]::IsNullOrWhiteSpace($siteNameRaw)) { $siteName = $siteNameRaw }
        }

        $sgaModel = "N/A"
        $computeSerial = "N/A"
        $bmcIp = "N/A"
        $driveSize = "N/A"
        $driveType = "N/A"
        $raidMode = "N/A"
        $controllerName = "N/A"
        $controllerAIp = "N/A"
        $controllerBIp = "N/A"
        $shelfSerials = @()
        $shelfIds = @()

        $nodeAttributeNode = $null
        if ($nodeAttributeByNodeId.ContainsKey($nodeId)) {
            $nodeAttributeNode = $nodeAttributeByNodeId[$nodeId]
        }
        if ($null -ne $nodeAttributeNode) {
            $computeSerialValue = [string](Get-PropertyValue -Object $nodeAttributeNode -PropertyName "computeSerial")
            if (-not [string]::IsNullOrWhiteSpace($computeSerialValue)) { $computeSerial = $computeSerialValue }

            $bmcIpValue = [string](Get-PropertyValue -Object $nodeAttributeNode -PropertyName "bmcIp")
            if (-not [string]::IsNullOrWhiteSpace($bmcIpValue)) { $bmcIp = $bmcIpValue }

            $applianceModelValue = [string](Get-PropertyValue -Object $nodeAttributeNode -PropertyName "applianceModel")
            if (-not [string]::IsNullOrWhiteSpace($applianceModelValue)) { $sgaModel = $applianceModelValue }

            $driveSizeValue = [string](Get-PropertyValue -Object $nodeAttributeNode -PropertyName "driveSize")
            if (-not [string]::IsNullOrWhiteSpace($driveSizeValue)) { $driveSize = Convert-BytesToDecimalTB -Bytes $driveSizeValue }

            $driveTypeValue = [string](Get-PropertyValue -Object $nodeAttributeNode -PropertyName "driveType")
            if (-not [string]::IsNullOrWhiteSpace($driveTypeValue)) { $driveType = $driveTypeValue }

            $raidModeValue = [string](Get-PropertyValue -Object $nodeAttributeNode -PropertyName "raidMode")
            if (-not [string]::IsNullOrWhiteSpace($raidModeValue)) { $raidMode = $raidModeValue }

            $controllerNameValue = [string](Get-PropertyValue -Object $nodeAttributeNode -PropertyName "controllerName")
            if (-not [string]::IsNullOrWhiteSpace($controllerNameValue)) { $controllerName = $controllerNameValue }

            $controllerAIpValue = [string](Get-PropertyValue -Object $nodeAttributeNode -PropertyName "controllerAIp")
            if (-not [string]::IsNullOrWhiteSpace($controllerAIpValue)) { $controllerAIp = $controllerAIpValue }

            $controllerBIpValue = [string](Get-PropertyValue -Object $nodeAttributeNode -PropertyName "controllerBIp")
            if (-not [string]::IsNullOrWhiteSpace($controllerBIpValue)) { $controllerBIp = $controllerBIpValue }

            $shelfSerials = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $nodeAttributeNode -PropertyName "shelfSerials")
            $shelfIds = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $nodeAttributeNode -PropertyName "shelfIds")
        }

        $metricNode = $null
        $metricKey = $hostName.ToLowerInvariant()
        if ($applianceStorageByNode.ContainsKey($metricKey)) {
            $metricNode = $applianceStorageByNode[$metricKey]
        }

        $sgaSerial = "N/A"
        $santricityVersion = "N/A"
        $nvsramVersion = "N/A"
        if ($null -ne $metricNode) {
            $sgaModelMetricValue = [string](Get-PropertyValue -Object $metricNode -PropertyName "sgaModel")
            if (-not [string]::IsNullOrWhiteSpace($sgaModelMetricValue)) { $sgaModel = $sgaModelMetricValue }

            $sgaSerialValue = [string](Get-PropertyValue -Object $metricNode -PropertyName "storageSerialNumber")
            if (-not [string]::IsNullOrWhiteSpace($sgaSerialValue)) { $sgaSerial = $sgaSerialValue }

            $santricityVersionValue = [string](Get-PropertyValue -Object $metricNode -PropertyName "santricityVersion")
            if (-not [string]::IsNullOrWhiteSpace($santricityVersionValue)) { $santricityVersion = $santricityVersionValue }

            $nvsramVersionValue = [string](Get-PropertyValue -Object $metricNode -PropertyName "nvsramVersion")
            if (-not [string]::IsNullOrWhiteSpace($nvsramVersionValue)) { $nvsramVersion = $nvsramVersionValue }
        }

        $applianceRows += ,@(
            [string]$hostName,
            [string]$siteName,
            [string]$sgaModel,
            [string]$computeSerial,
            [string]$sgaSerial,
            [string]$bmcIp
        )

        $applianceStorageRows += ,@(
            [string]$hostName,
            [string]$siteName,
            [string]$santricityVersion,
            [string]$nvsramVersion,
            [string]$driveSize,
            [string]$driveType,
            [string]$raidMode
        )

        $applianceControllerRows += ,@(
            [string]$hostName,
            [string]$siteName,
            [string]$controllerName,
            [string]$controllerAIp,
            [string]$controllerBIp
        )

        $shelfCount = [math]::Max($shelfIds.Count, $shelfSerials.Count)
        if ($shelfCount -eq 0) {
            $applianceShelfRows += ,@([string]$hostName, [string]$siteName, "N/A", "N/A")
        } else {
            for ($shelfIndex = 0; $shelfIndex -lt $shelfCount; $shelfIndex++) {
                $shelfId = if ($shelfIndex -lt $shelfIds.Count) { [string]$shelfIds[$shelfIndex] } else { "N/A" }
                $shelfSerial = if ($shelfIndex -lt $shelfSerials.Count) { [string]$shelfSerials[$shelfIndex] } else { "N/A" }
                $shelfNodeLabel = if ($shelfIndex -eq 0) { [string]$hostName } else { "" }
                $applianceShelfRows += ,@($shelfNodeLabel, [string]$siteName, $shelfId, $shelfSerial)
            }
        }
    }

    if ($applianceRows.Count -eq 0) {
        $applianceRows += ,@("none", "N/A", "N/A", "N/A", "N/A", "N/A")
        $applianceStorageRows += ,@("none", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A")
        $applianceControllerRows += ,@("none", "N/A", "N/A", "N/A", "N/A")
        $applianceShelfRows += ,@("none", "N/A", "N/A", "N/A")
    }

    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Node", "Site Name", "SGA Model", "Compute Serial", "Storage Serial", "BMC IP") -Rows $applianceRows

    $markdownLines.Add("")
    # --- Grid Appliance Storage ---
    $markdownLines.Add("## Grid Appliance Storage")
    $markdownLines.Add("")
    $markdownLines.Add("Certain models of storage StorageGRID storage appliances use an embedded E-Series storage controller(s) running NetApp SANtricity OS to manage the appliance's disk shelves and present storage to the StorageGRID compute and software layer.")
    $markdownLines.Add("")
    $markdownLines.Add("This table lists the SANtricity software version and NVSRAM (non-volatile configuration memory) version running on each appliance's storage controller, along with the representative drive size reported for that node's storage.")
    $markdownLines.Add("")
    $markdownLines.Add("Keeping SANtricity and NVSRAM versions current and consistent across appliances is important for supportability and for applying NetApp-recommended firmware and drive qualifications.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Node", "Site Name", "SANtricity Version", "NVSRAM", "Drive Size (TB)", "Drive Type", "RAID Mode") -Rows $applianceStorageRows

    $markdownLines.Add("")
    # --- Grid Appliance Storage Controllers ---
    $markdownLines.Add("## Grid Appliance Storage Controllers")
    $markdownLines.Add("")
    $markdownLines.Add("Each StorageGRID storage appliance's embedded E-Series storage controller is managed independently via NetApp SANtricity System Manager and is identified by its own name and dedicated Controller A and Controller B management IP addresses (for dual-controller appliances).")
    $markdownLines.Add("")
    $markdownLines.Add("This table records the storage controller identity and management IPs for each appliance node, which are useful for locating and directly accessing the SANtricity management interface during hardware maintenance or troubleshooting.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Node", "Site Name", "Storage Controller Name", "Controller A Mgmt IP", "Controller B Mgmt IP") -Rows $applianceControllerRows

    # --- SANtricity Appliance Discovery ---
    $markdownLines.Add("")
    $markdownLines.Add("## SANtricity Appliance Discovery")
    $markdownLines.Add("")
    $markdownLines.Add("This table identifies StorageGRID appliances with one or more valid embedded SANtricity controller management IP addresses. These appliances are eligible for optional follow-on E-Series collection; discovery does not initiate SANtricity API access.")
    $markdownLines.Add("")
    $santricityDiscoveryRows = @()
    foreach ($candidate in @($santricityApplianceCandidates)) {
        $candidateIps = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $candidate -PropertyName "CandidateIps")
        $santricityDiscoveryRows += ,@(
            (Get-PropertyValue -Object $candidate -PropertyName "NodeName"),
            (Get-PropertyValue -Object $candidate -PropertyName "SiteName"),
            (Get-PropertyValue -Object $candidate -PropertyName "NodeType"),
            (Get-PropertyValue -Object $candidate -PropertyName "ApplianceModel"),
            (Get-PropertyValue -Object $candidate -PropertyName "ControllerName"),
            ($candidateIps -join "`n"),
            (& $boolText (Get-PropertyValue -Object $candidate -PropertyName "HasEseries")),
            (Get-PropertyValue -Object $candidate -PropertyName "SantricityVersion"),
            (& $boolText (Get-PropertyValue -Object $candidate -PropertyName "Eligible"))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Node", "Site", "Node Type", "Model", "Controller Name", "SANtricity Management IPs", "hasEseries", "SANtricity Version", "Eligible") -Rows $santricityDiscoveryRows

    $markdownLines.Add("")
    # --- Grid Appliance Shelves ---
    $markdownLines.Add("## Grid Appliance Shelves")
    $markdownLines.Add("")
    $markdownLines.Add("Larger StorageGRID storage appliance models support one or more attached expansion shelves to increase the number of drives available to the E-Series storage controller.")
    $markdownLines.Add("")
    $markdownLines.Add("This table lists the storage shelves reported for each appliance storage node, identified by shelf ID and chassis serial number; a node with multiple rows in this table has one or more expansion shelves attached in addition to its base controller shelf.")
    $markdownLines.Add("")
    & $emitOpenXmlMergedFirstColumnTable -Headers @("Node", "Site Name", "Shelf ID", "Shelf Serial Number") -Rows $applianceShelfRows -GridWidths @(2160, 1800, 1440, 2520) -AutofitToWindow

    # --- Grid Manager: Support ---
    $markdownLines.Add("")
    $markdownLines.Add("# Support")
    $markdownLines.Add("")
    $markdownLines.Add("## Grid AutoSupport")
    $markdownLines.Add("")
    $markdownLines.Add("The AutoSupport feature enables your StorageGRID system to send periodic and event-driven health and status messages to technical support to allow proactive monitoring and troubleshooting. StorageGRID AutoSupport also enables the use of Active IQ for predictive recommendations. Weekly AutoSupport messages provide a regular health snapshot, while event-triggered messages are sent automatically when specific system events occur, helping technical support identify and resolve issues faster.")
    $markdownLines.Add("")

    $asupProtocol = [string](Get-PropertyValue -Object $autosupport -PropertyName "transport")
    $asupCertValidation = [string](Get-PropertyValue -Object $autosupport -PropertyName "certEnable")
    $asupWeekly = [string](Get-PropertyValue -Object $autosupport -PropertyName "weeklyEnable")
    $asupEvent = [string](Get-PropertyValue -Object $autosupport -PropertyName "eventEnable")
    $asupDestinations = @((Get-PropertyValue -Object $autosupport -PropertyName "destinations"))

    $asupDestinationItems = @()
    foreach ($dest in $asupDestinations) {
        if ($dest -is [string]) {
            if (-not [string]::IsNullOrWhiteSpace([string]$dest)) { $asupDestinationItems += [string]$dest }
            continue
        }

        $destName = [string](Get-PropertyValue -Object $dest -PropertyName "name")
        if ([string]::IsNullOrWhiteSpace($destName)) {
            $destName = [string](Get-PropertyValue -Object $dest -PropertyName "destination")
        }
        if ([string]::IsNullOrWhiteSpace($destName)) {
            $destName = [string](Get-PropertyValue -Object $dest -PropertyName "type")
        }
        if (-not [string]::IsNullOrWhiteSpace($destName)) { $asupDestinationItems += $destName }
    }
    $asupDestinationDisplay = if ($asupDestinationItems.Count -gt 0) { ($asupDestinationItems | Sort-Object -Unique) -join ", " } else { "N/A" }

    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Protocol", "Certificate Validation", "Weekly ASUP", "Event Triggered ASUP", "ASUP Destinations") -Rows @(
        ,@((& $boolText $asupProtocol), (& $boolText $asupCertValidation), (& $boolText $asupWeekly), (& $boolText $asupEvent), [string]$asupDestinationDisplay)
    )

    # --- Configuration: Access Control ---
    $markdownLines.Add("")
    $markdownLines.Add("# Configuration - Access Control")
    $markdownLines.Add("")
    # --- Identity Source ---
    $markdownLines.Add("## Identity Federation")
    $markdownLines.Add("")
    $markdownLines.Add("You can configure identity federation to allow admin groups and users to be managed in a centralised identity system, such as Active Directory, Azure AD, OpenLDAP, or Oracle Directory Server, rather than as local StorageGRID accounts.")
    $markdownLines.Add("")
    $markdownLines.Add("Once configured, StorageGRID periodically synchronizes federated groups and users from the identity source, allowing users to sign in with their existing organizational credentials and letting administrators manage access centrally in the identity provider.")
    $markdownLines.Add("")

    $idDisable = Get-PropertyValue -Object $identitySource -PropertyName "disable"
    $idLdapServiceType = Get-PropertyValue -Object $identitySource -PropertyName "ldapServiceType"
    $idHost = Get-PropertyValue -Object $identitySource -PropertyName "hostname"
    $idPort = Get-PropertyValue -Object $identitySource -PropertyName "port"
    $idUsername = Get-PropertyValue -Object $identitySource -PropertyName "username"
    $idGroupBaseDn = Get-PropertyValue -Object $identitySource -PropertyName "baseGroupDn"
    $idUserBaseDn = Get-PropertyValue -Object $identitySource -PropertyName "baseUserDn"
    $idDisableTls = Get-PropertyValue -Object $identitySource -PropertyName "disableTLS"
    $idEnableLdaps = Get-PropertyValue -Object $identitySource -PropertyName "enableLDAPS"

    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Identity Federation Disabled", "LDAP Type", "Hostname", "Port", "Disable TLS", "Enable LDAPS") -Rows @(
        ,@(
            (& $boolText $idDisable),
            (& $boolText $idLdapServiceType),
            (& $boolText $idHost),
            (& $boolText $idPort),
            (& $boolText $idDisableTls),
            (& $boolText $idEnableLdaps)
        )
    )

    $markdownLines.Add("")
    # --- Identity Federation - Authentication and Search Bases ---
    $markdownLines.Add("## Identity Federation - Authentication and Search Bases")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows the service account and LDAP search base configuration used to query the identity source. The bind username is the account StorageGRID uses to authenticate to the LDAP server and must have permission to list groups, users, and their key attributes.")
    $markdownLines.Add("")
    $markdownLines.Add("The Group Base DN and User Base DN define the LDAP subtrees that StorageGRID searches when importing federated groups and users; unique name values only need to be unique within their respective base DN.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Username", "Group Base DN", "User Base DN") -Rows @(
        ,@(
            (& $boolText $idUsername),
            (& $boolText $idGroupBaseDn),
            (& $boolText $idUserBaseDn)
        )
    )

    # --- Single Sign-On ---
    $markdownLines.Add("")
    $markdownLines.Add("## Single Sign-On")
    $markdownLines.Add("")
    $markdownLines.Add("Single sign-on (SSO) lets StorageGRID authenticate Grid Manager and Tenant Manager users through an external identity provider using the SAML 2.0 standard or OpenID Connect, instead of local StorageGRID passwords.")
    $markdownLines.Add("")
    $markdownLines.Add("When SSO is enabled, local (non-federated) users can no longer sign in; users must be authenticated by the configured identity provider and belong to a federated group with StorageGRID access permissions. Sandbox mode allows administrators to test and validate an SSO configuration for federated groups before it is enforced for all users.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
        @("SSO Disabled", (& $boolText (Get-PropertyValue -Object $ssoConfig -PropertyName "disable"))),
        @("Sandbox Mode", (& $boolText (Get-PropertyValue -Object $ssoConfig -PropertyName "sandbox"))),
        @("Identity Provider", (& $boolText (Get-PropertyValue -Object $ssoConfig -PropertyName "identityProvider"))),
        @("Disable TLS", (& $boolText (Get-PropertyValue -Object $ssoConfig -PropertyName "disableTLS")))
    )

    # --- Admin Groups ---
    $markdownLines.Add("")
    $markdownLines.Add("## Admin Groups")
    $markdownLines.Add("")
    $markdownLines.Add("Admin groups (local or federated) determine which features and operations in the Grid Manager and the Grid Management API that group members can access.")
    $markdownLines.Add("")
    $markdownLines.Add("Every StorageGRID user must belong to at least one group to sign in, since permissions are always granted at the group level rather than to individual users.")
    $markdownLines.Add("")
    $markdownLines.Add("Each group's Access mode controls whether its members have read-write or read-only access to the management permissions assigned to that group.")
    $markdownLines.Add("")

    $groupRows = @()
    $sortedGroups = @(
        $gridGroups | Sort-Object -Property @(
            @{ Expression = { [string](Get-PropertyValue -Object $_ -PropertyName "displayName") } },
            @{ Expression = { [string](Get-PropertyValue -Object $_ -PropertyName "uniqueName") } }
        )
    )

    foreach ($group in $sortedGroups) {
        $displayName = [string](Get-PropertyValue -Object $group -PropertyName "displayName")
        $uniqueName = [string](Get-PropertyValue -Object $group -PropertyName "uniqueName")
        $federated = Get-PropertyValue -Object $group -PropertyName "federated"

        $managementPermissions = "N/A"
        $policies = Get-PropertyValue -Object $group -PropertyName "policies"
        $managementPolicy = Get-PropertyValue -Object $policies -PropertyName "management"
        if ($null -ne $managementPolicy) {
            $enabledPerms = @()
            foreach ($perm in @($managementPolicy.PSObject.Properties.Name)) {
                $permEnabled = Get-PropertyValue -Object $managementPolicy -PropertyName $perm
                if ($permEnabled -eq $true) { $enabledPerms += [string]$perm }
            }
            if ($enabledPerms.Count -gt 0) {
                $managementPermissions = ($enabledPerms | Sort-Object) -join "`n"
            }
        }

        $groupRows += ,@(
            (& $boolText $displayName),
            (& $boolText $uniqueName),
            (& $boolText $federated),
            [string]$managementPermissions
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Display Name", "Unique Name", "Federated", "Management Permissions") -Rows $groupRows

    # --- Admin Users ---
    $markdownLines.Add("")
    $markdownLines.Add("## Admin Users")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows grid users (local and federated) and the admin groups each user belongs to.")
    $markdownLines.Add("")
    $markdownLines.Add("A user's actual permissions in the Grid Manager are the union of the management permissions assigned to all groups they belong to; a user with no group memberships has no access to the Grid Manager.")
    $markdownLines.Add("")
    $markdownLines.Add("Federated users are managed in the configured identity source, while local users are created and managed directly in StorageGRID.")
    $markdownLines.Add("")

    $groupDisplayNameById = @{}
    foreach ($group in $sortedGroups) {
        $groupId = [string](Get-PropertyValue -Object $group -PropertyName "id")
        if ([string]::IsNullOrWhiteSpace($groupId)) { continue }
        $groupDisplayName = [string](Get-PropertyValue -Object $group -PropertyName "displayName")
        if (-not [string]::IsNullOrWhiteSpace($groupDisplayName)) {
            $groupDisplayNameById[$groupId] = $groupDisplayName
        }
    }

    $userRows = @()
    $sortedUsers = @(
        $gridUsers | Sort-Object -Property @(
            @{ Expression = { [string](Get-PropertyValue -Object $_ -PropertyName "fullName") } },
            @{ Expression = { [string](Get-PropertyValue -Object $_ -PropertyName "uniqueName") } }
        )
    )
    foreach ($user in $sortedUsers) {
        $memberOfIds = @((Get-PropertyValue -Object $user -PropertyName "memberOf"))
        $memberGroupNames = @()
        foreach ($memberId in $memberOfIds) {
            $idText = [string]$memberId
            if ([string]::IsNullOrWhiteSpace($idText)) { continue }
            if ($groupDisplayNameById.ContainsKey($idText)) {
                $memberGroupNames += [string]$groupDisplayNameById[$idText]
            }
        }
        $memberGroupNames = @($memberGroupNames | Sort-Object -Unique)
        $memberGroupDisplay = if ($memberGroupNames.Count -gt 0) { $memberGroupNames -join "`n" } else { "N/A" }

        $userRows += ,@(
            (& $boolText (Get-PropertyValue -Object $user -PropertyName "fullName")),
            (& $boolText (Get-PropertyValue -Object $user -PropertyName "uniqueName")),
            (& $boolText (Get-PropertyValue -Object $user -PropertyName "federated")),
            (& $boolText (Get-PropertyValue -Object $user -PropertyName "disable")),
            [string]$memberGroupDisplay
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Full Name", "Unique Name", "Federated", "Disabled", "Grid Groups") -Rows $userRows

    # --- Configuration - Monitoring and Logging ---
    $markdownLines.Add("")
    $markdownLines.Add("# Configuration - Monitoring and Logging")
    $markdownLines.Add("")
    # --- Audit Logging ---
    $markdownLines.Add("## Audit Logging")
    $markdownLines.Add("")
    $markdownLines.Add("StorageGRID generates audit messages for system, storage, management, ILM, client read, client write, and cross-grid replication activity to support monitoring, security review, and troubleshooting.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows the configured audit level for each message category:")
    $markdownLines.Add("- Off: Disables logging for the category")
    $markdownLines.Add("- Error: Logs only failed operations")
    $markdownLines.Add("- Normal: Logs the full set of standard transactional messages (which also includes all Error-level messages)")
    $markdownLines.Add("")
    $markdownLines.Add("Reducing categories such as Client Reads to Error can significantly decrease audit log volume for high-traffic S3 workloads.")
    $markdownLines.Add("")

    $auditLevelsObject = Get-PropertyValue -Object $auditConfig -PropertyName "levels"
    $auditLevelRows = @()
    if ($null -ne $auditLevelsObject) {
        foreach ($p in @($auditLevelsObject.PSObject.Properties)) {
            $auditLevelRows += ,@([string]$p.Name, (& $boolText $p.Value))
        }
    }
    $auditLevelRows = @($auditLevelRows | Sort-Object { [string]$_[0] })
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Audit Setting", "Value") -Rows $auditLevelRows

    # --- Logged Headers ---
    $markdownLines.Add("")
    $markdownLines.Add("## Logged Headers")
    $markdownLines.Add("")
    $loggedHeadersList = @((Get-PropertyValue -Object $auditConfig -PropertyName "loggedHeaders"))
    $loggedHeadersDisplay = if ($loggedHeadersList.Count -gt 0) { $loggedHeadersList -join ", " } else { "N/A" }
    $markdownLines.Add("You can optionally configure specific HTTP request headers to be captured in client read and write audit messages for S3 requests.")
    $markdownLines.Add("")
    $markdownLines.Add("When present in a client request, these headers are recorded in the audit message under the HTRH field, which can help correlate object storage operations back to specific client applications, sessions, or request identifiers for troubleshooting or compliance purposes.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Logged Headers") -Rows @(
        ,@([string]$loggedHeadersDisplay)
    )

    $auditDefaults = Get-PropertyValue -Object $auditDestinations -PropertyName "defaults"
    $syslogDestination = Get-PropertyValue -Object $auditDefaults -PropertyName "remoteSyslogServerA"
    if ($null -eq $syslogDestination) {
        $syslogDestination = Get-PropertyValue -Object $auditDefaults -PropertyName "remoteSyslogServerATest"
    }

    # --- Syslog - Destination ---
    $markdownLines.Add("")
    $markdownLines.Add("## Syslog - Destination")
    $markdownLines.Add("")
    $markdownLines.Add("An external syslog server lets StorageGRID save audit logs, application logs, and security event logs to a location outside the grid, which is especially useful for large grids or where centralized, long-term log retention is required.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows the syslog server connection details, including the protocol used to send audit information; using TLS or RELP/TLS is recommended over plain UDP or TCP because they provide encryption and, for RELP, more reliable delivery if the syslog server needs to restart.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Enabled", "Protocol", "Hostname", "Port", "Insecure TLS") -Rows @(
        ,@(
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "enabled")),
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "protocol")),
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "hostname")),
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "port")),
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "insecureTLS"))
        )
    )

    # --- Syslog - Audit Logs ---
    $markdownLines.Add("")
    $markdownLines.Add("## Syslog - Audit Logs")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows whether StorageGRID audit log messages are forwarded to the external syslog server, and the severity and facility values applied to those messages.")
    $markdownLines.Add("")
    $markdownLines.Add("Setting a fixed severity or facility (rather than Passthrough) can help aggregate and filter these logs more easily in downstream log management tooling, at the cost of losing StorageGRID's own per-message severity distinctions.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Send", "Facility", "Severity") -Rows @(
        ,@(
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "auditLogsSend")),
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "auditLogsFacility")),
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "auditLogsSeverity"))
        )
    )

    # --- Syslog - Security Events ---
    $markdownLines.Add("")
    $markdownLines.Add("## Syslog - Security Events")
    $markdownLines.Add("")
    $markdownLines.Add("Security events include activity such as unauthorized sign-in attempts or a user signing in as root, and are useful for security monitoring and incident investigation. This table shows whether these events are forwarded to the external syslog server and the severity/facility values applied to them.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Send", "Facility", "Severity") -Rows @(
        ,@(
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "authEventsSend")),
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "authEventsFacility")),
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "authEventsSeverity"))
        )
    )

    # --- Syslog - Application Logs ---
    $markdownLines.Add("")
    $markdownLines.Add("## Syslog - Application Logs")
    $markdownLines.Add("")
    $markdownLines.Add("Application logs include StorageGRID software log files, such as bycast.log, nms.log, and prometheus.log, that are useful for troubleshooting. This table shows whether these logs are forwarded to the external syslog server and the severity/facility values applied to them.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Send", "Facility", "Severity") -Rows @(
        ,@(
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "applicationLogsSend")),
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "applicationLogsFacility")),
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "applicationLogsSeverity"))
        )
    )

    # --- Syslog - Access Logs ---
    $markdownLines.Add("")
    $markdownLines.Add("## Syslog - Access Logs")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows whether access-related log information is forwarded to the external syslog server, along with the severity and facility values applied to those messages.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Send", "Facility", "Severity") -Rows @(
        ,@(
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "accessLogsSend")),
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "accessLogsFacility")),
            (& $boolText (Get-PropertyValue -Object $syslogDestination -PropertyName "accessLogsSeverity"))
        )
    )

    # --- Grid SNMP ---
    $markdownLines.Add("")
    $markdownLines.Add("## Grid SNMP")
    $markdownLines.Add("")
    $markdownLines.Add("StorageGRID can act as a Simple Network Management Protocol (SNMP) agent, allowing standard SNMP monitoring tools to poll basic grid health status and receive trap or inform notifications when alerts are triggered.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows whether the SNMP agent is enabled, its authentication trap and system identification settings, and the number of configured community strings, trap destinations, and USM (User-based Security Model) users for SNMPv3.")
    $markdownLines.Add("")
    $snmpCommunityStrings = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $snmpConfig -PropertyName "community_strings")
    $snmpTrapDestinations = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $snmpConfig -PropertyName "trap_destinations")
    $snmpUsmUsers = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $snmpConfig -PropertyName "usm_users")
    $snmpCommunityStringsDisplay = if (@($snmpCommunityStrings).Count -gt 0) { [string]@($snmpCommunityStrings).Count } else { "N/A" }
    $snmpTrapDestinationsDisplay = if (@($snmpTrapDestinations).Count -gt 0) { [string]@($snmpTrapDestinations).Count } else { "N/A" }
    $snmpUsmUsersDisplay = if (@($snmpUsmUsers).Count -gt 0) { [string]@($snmpUsmUsers).Count } else { "N/A" }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("SNMP Setting", "Value") -Rows @(
        @("Enabled", (& $boolText (Get-PropertyValue -Object $snmpConfig -PropertyName "enable_snmp"))),
        @("Auth Trap Enabled", (& $boolText (Get-PropertyValue -Object $snmpConfig -PropertyName "authtrapenable"))),
        @("Disable Notifications", (& $boolText (Get-PropertyValue -Object $snmpConfig -PropertyName "disable_notifications"))),
        @("System Location", (& $boolText (Get-PropertyValue -Object $snmpConfig -PropertyName "sysLocation"))),
        @("System Contact", (& $boolText (Get-PropertyValue -Object $snmpConfig -PropertyName "sysContact"))),
        @("Trap Community", (& $boolText (Get-PropertyValue -Object $snmpConfig -PropertyName "trapcommunity"))),
        @("Community Strings", $snmpCommunityStringsDisplay),
        @("Trap Destinations", $snmpTrapDestinationsDisplay),
        @("USM Users", $snmpUsmUsersDisplay)
    )

    # --- Configuration - System and Storage ---
    $markdownLines.Add("")
    $markdownLines.Add("# Configuration - System and Storage")
    $markdownLines.Add("")
    # --- General Options ---
    $markdownLines.Add("## General Options")
    $markdownLines.Add("")
    $markdownLines.Add("These grid-wide options control default object behavior and network protocol settings applied across the grid.")
    $markdownLines.Add("")
    $markdownLines.Add("Object Compression and Object Encryption apply StorageGRID-side compression/encryption to new objects ingested into the Grid. Prevent Client Modification blocks S3 clients from altering certain bucket or object settings; and Bucket Default Consistency Level sets the default data metadata consistency guarantee applied to new S3 buckets unless overridden at the bucket level.")
    $markdownLines.Add("")

    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Option", "Value") -Rows @(
        @("Object Compression", (& $boolText (Get-PropertyValue -Object $gridConfig -PropertyName "objectCompression"))),
        @("Object Encryption", (& $boolText (Get-PropertyValue -Object $gridConfig -PropertyName "objectEncryption"))),
        @("Object Hashing", (& $boolText (Get-PropertyValue -Object $gridConfig -PropertyName "objectHashing"))),
        @("Prevent Client Modification", (& $boolText (Get-PropertyValue -Object $gridConfig -PropertyName "preventClientModification"))),
        @("HTTP Connection", (& $boolText (Get-PropertyValue -Object $gridConfig -PropertyName "httpConnection"))),
        @("Network Transfer Encryption", (& $boolText (Get-PropertyValue -Object $gridConfig -PropertyName "networkTransferEncryption"))),
        @("Bucket Default Consistency Level", (& $boolText (Get-PropertyValue -Object $gridConfig -PropertyName "consistencyLevel"))),
        @("GUI Inactivity Timeout (s)", (& $boolText $guiTimeoutRaw))
    )

    # --- S3 Object Lock ---
    $markdownLines.Add("")
    $markdownLines.Add("## S3 Object Lock")
    $markdownLines.Add("")
    $markdownLines.Add("S3 Object Lock is a compliance feature that lets S3 buckets prevent object versions from being deleted or overwritten for a specified retention period, similar to write-once-read-many (WORM) storage. This grid-wide setting must be enabled before any tenant can create an S3 Object Lock-enabled bucket; it cannot be disabled once any bucket has been created with Object Lock enabled.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Option", "Value") -Rows @(
        ,@("Compliance Enabled", (& $boolText (Get-PropertyValue -Object $complianceGlobal -PropertyName "complianceEnabled")))
    )

    # --- Storage Options ---
    $markdownLines.Add("")
    $markdownLines.Add("## Storage Options")
    $markdownLines.Add("")
    $markdownLines.Add("Storage volume watermarks control when a Storage Node's storage volumes transition between read-write, soft read-only, and hard read-only states as they approach capacity, helping to prevent object data loss from a completely full volume.")
    $markdownLines.Add("")
    $markdownLines.Add("The Metadata Reserved Space setting is a grid-wide value that determines how much space on storage volume 0 (rangedb0) of every Storage Node is reserved for the distributed object-metadata database; the actual reserved space on a given node is capped by the size of that node's volume 0.")
    $markdownLines.Add("")
    $metadataReservedRaw = Get-PropertyValue -Object $storageWatermarks -PropertyName "metadataReservedSpace"
    $metadataReservedDisplay = "N/A"
    if ($null -ne $metadataReservedRaw) {
        $metadataReservedBytes = 0.0
        if ([double]::TryParse([string]$metadataReservedRaw, [ref]$metadataReservedBytes)) {
            $metadataReservedDisplay = ([math]::Round($metadataReservedBytes / 1000000000000)).ToString() + " TB"
        }
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Option", "Value") -Rows @(
        @("Storage Volume Read-Write Watermark (%)", (& $boolText (Get-PropertyValue -Object $storageWatermarks -PropertyName "storageVolumeReadWriteWatermark"))),
        @("Storage Volume Soft Read-Only Watermark (%)", (& $boolText (Get-PropertyValue -Object $storageWatermarks -PropertyName "storageVolumeSoftReadOnlyWatermark"))),
        @("Storage Volume Hard Read-Only Watermark (%)", (& $boolText (Get-PropertyValue -Object $storageWatermarks -PropertyName "storageVolumeHardReadOnlyWatermark"))),
        @("Metadata Reserved Space", $metadataReservedDisplay)
    )

    # --- Configuration - Security and Certificates ---
    $markdownLines.Add("")
    $markdownLines.Add("# Configuration - Security and Certificates")
    $markdownLines.Add("")
    # --- Management Certificate ---
    $markdownLines.Add("## Management Certificate")
    $markdownLines.Add("")
    $markdownLines.Add("The Grid Manager and Tenant Manager management interface certificate is presented to browsers and API clients connecting to the Grid Manager, Tenant Manager, Grid Management API, and Tenant Management API.")
    $markdownLines.Add("")
    $mgmtServerCert = [string](Get-PropertyValue -Object $mgmtCert -PropertyName "serverCertificateEncoded")
    $mgmtCaBundle = [string](Get-PropertyValue -Object $mgmtCert -PropertyName "caBundleEncoded")
    $mgmtCertificateRows = @()
    foreach ($certificateRole in @(
        [pscustomobject]@{ Name = "Management Server Certificate"; Pem = $mgmtServerCert },
        [pscustomobject]@{ Name = "Management CA Bundle"; Pem = $mgmtCaBundle }
    )) {
        $details = Get-X509CertificateDisplayDetails -Pem $certificateRole.Pem
        for ($sanIndex = 0; $sanIndex -lt $details.SubjectAltNames.Count; $sanIndex++) {
            $mgmtCertificateRows += ,@(
                $(if ($sanIndex -eq 0) { $certificateRole.Name } else { "" }),
                $(if ($sanIndex -eq 0) { $details.CommonName } else { "" }),
                [string]$details.SubjectAltNames[$sanIndex],
                $(if ($sanIndex -eq 0) { $details.Expiry } else { "" })
            )
        }
    }
    & $emitOpenXmlMergedColumnsTable -Headers @("Certificate", "Common Name", "Subject Alternative Name", "Expiry") -Rows $mgmtCertificateRows -GridWidths @(3000, 2160, 3000, 1440) -MergeColumns @(0, 1, 3) -AutofitToWindow

    # ---Storage API Certificate ---
    $markdownLines.Add("")
    $markdownLines.Add("## Storage API Certificate")
    $markdownLines.Add("")
    $markdownLines.Add("The Storage API CA certificate secures internal grid communication for the S3 storage API endpoints between grid nodes. This certificate is internal to the Grid and not recommended to be exposed outside of the private Grid Network.")
    $markdownLines.Add("")
    $storageApiCa = [string](Get-PropertyValue -Object $internalCaCert -PropertyName "caCertificateEncoded")
    $storageApiCertificateDetails = Get-X509CertificateDisplayDetails -Pem $storageApiCa
    $storageApiCertificateRows = @()
    for ($sanIndex = 0; $sanIndex -lt $storageApiCertificateDetails.SubjectAltNames.Count; $sanIndex++) {
        $storageApiCertificateRows += ,@(
            $(if ($sanIndex -eq 0) { "Storage API CA Certificate" } else { "" }),
            $(if ($sanIndex -eq 0) { $storageApiCertificateDetails.CommonName } else { "" }),
            [string]$storageApiCertificateDetails.SubjectAltNames[$sanIndex],
            $(if ($sanIndex -eq 0) { $storageApiCertificateDetails.Expiry } else { "" })
        )
    }
    & $emitOpenXmlMergedColumnsTable -Headers @("Certificate", "Common Name", "Subject Alternative Name", "Expiry") -Rows $storageApiCertificateRows -GridWidths @(3000, 2160, 3000, 1440) -MergeColumns @(0, 1, 3) -AutofitToWindow

    # --- Load Balancer Endpoint Certificates ---
    $markdownLines.Add("")
    $markdownLines.Add("## Load Balancer Endpoint Certificates")
    $markdownLines.Add("")
    $markdownLines.Add("Each Load Balancer Endpoint can use a server certificate for TLS termination of client S3 and management connections.")
    $markdownLines.Add("")
    $endpointNameById = @{}
    foreach ($configuredEndpoint in $gatewayConfigs) {
        $configuredEndpointId = [string](Get-PropertyValue -Object $configuredEndpoint -PropertyName "id")
        if (-not [string]::IsNullOrWhiteSpace($configuredEndpointId)) {
            $endpointNameById[$configuredEndpointId] = [string](Get-PropertyValue -Object $configuredEndpoint -PropertyName "displayName")
        }
    }
    $lbCertRows = @()
    foreach ($lbConfig in $lbServerConfigs) {
        $endpointId = [string](Get-PropertyValue -Object $lbConfig -PropertyName "__endpointId")
        $endpointName = if ($endpointNameById.ContainsKey($endpointId)) { $endpointNameById[$endpointId] } else { $endpointId }
        $serviceType = [string](Get-PropertyValue -Object $lbConfig -PropertyName "defaultServiceType")
        $plaintextData = Get-PropertyValue -Object $lbConfig -PropertyName "plaintextCertData"
        $metadata = Get-PropertyValue -Object $plaintextData -PropertyName "metadata"
        $serverCertDetails = Get-PropertyValue -Object $metadata -PropertyName "serverCertificateDetails"
        $certificateDetails = $null
        if ($null -ne $serverCertDetails) {
            $subject = [string](Get-PropertyValue -Object $serverCertDetails -PropertyName "subject")
            $commonName = "N/A"
            if ($subject -match '(^|/)CN=([^/]+)') { $commonName = [string]$matches[2] }
            $sanValues = @([string](Get-PropertyValue -Object $serverCertDetails -PropertyName "subjectAltNames") -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($sanValues.Count -eq 0) { $sanValues = @("N/A") }
            $expiry = "N/A"
            try { $expiry = ([datetimeoffset]::Parse([string](Get-PropertyValue -Object $serverCertDetails -PropertyName "notAfter"))).UtcDateTime.ToString("dd-MM-yyyy") } catch { $expiry = "N/A" }
            $certificateDetails = [pscustomobject]@{ CommonName = $commonName; SubjectAltNames = $sanValues; Expiry = $expiry }
        } elseif ($serviceType -match '^(management|tenant)$') {
            $certificateDetails = Get-X509CertificateDisplayDetails -Pem $mgmtServerCert
        }
        if ($null -eq $certificateDetails) {
            $certificateDetails = [pscustomobject]@{ CommonName = "N/A"; SubjectAltNames = @("N/A"); Expiry = "N/A" }
        }
        for ($sanIndex = 0; $sanIndex -lt $certificateDetails.SubjectAltNames.Count; $sanIndex++) {
            $lbCertRows += ,@(
                $(if ($sanIndex -eq 0) { (& $boolText $endpointName) } else { "" }),
                $(if ($sanIndex -eq 0) { (& $boolText $serviceType) } else { "" }),
                $(if ($sanIndex -eq 0) { $certificateDetails.CommonName } else { "" }),
                [string]$certificateDetails.SubjectAltNames[$sanIndex],
                $(if ($sanIndex -eq 0) { $certificateDetails.Expiry } else { "" })
            )
        }
    }
    & $emitOpenXmlMergedColumnsTable -Headers @("Endpoint Name", "Service Type", "Common Name", "Subject Alternative Name", "Expiry") -Rows $lbCertRows -GridWidths @(2160, 1800, 2160, 2880, 1440) -MergeColumns @(0, 1, 2, 4) -AutofitToWindow

    # --- Client Certificates ---
    $markdownLines.Add("")
    $markdownLines.Add("## Client Certificates")
    $markdownLines.Add("")
    $markdownLines.Add("Client certificates allow external tools, such as monitoring or automation systems, to authenticate to the Grid Management API or Tenant Management API using a certificate instead of a username and password. Each client certificate is bound to a specific StorageGRID user account and can optionally be permitted to bypass single sign-on (SSO) when SSO is enabled for the grid.")
    $markdownLines.Add("")
    $clientCertRows = @()
    foreach ($clientCert in $clientCertificates) {
        $clientCertRows += ,@(
            (& $boolText (& $nameOf $clientCert)),
            (& $boolText (Get-PropertyValue -Object $clientCert -PropertyName "userUUID")),
            (& $boolText (Get-PropertyValue -Object $clientCert -PropertyName "allowSso")),
            (& $boolText (Get-PropertyValue -Object $clientCert -PropertyName "expirationDate"))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Certificate Name", "Bound User", "Allow SSO Bypass", "Expiration Date") -Rows $clientCertRows

    # --- KMIP Cluster ---
    $markdownLines.Add("")
    $markdownLines.Add("## KMIP Cluster")
    $markdownLines.Add("")
    $markdownLines.Add("A Key Management Interoperability Protocol (KMIP) server cluster lets StorageGRID retrieve encryption keys from an external, customer-managed key management system rather than storing them locally, providing an additional layer of control over appliance node encryption. Each KMIP cluster can be associated with one or more sites or appliance nodes; if the external KMIP server becomes unavailable and a node loses power, that node cannot decrypt its own data until connectivity to the key server is restored.")
    $markdownLines.Add("")
    $kmipRows = @()
    foreach ($kmip in $kmipClusters) {
        $kmipServers = @((Get-PropertyValue -Object $kmip -PropertyName "servers"))
        $kmipServerNames = @()
        foreach ($server in $kmipServers) {
            $serverHost = [string](Get-PropertyValue -Object $server -PropertyName "hostname")
            if (-not [string]::IsNullOrWhiteSpace($serverHost)) { $kmipServerNames += $serverHost }
        }
        $kmipRows += ,@(
            (& $boolText (& $nameOf $kmip)),
            $(if ($kmipServerNames.Count -gt 0) { $kmipServerNames -join ", " } else { "N/A" }),
            (& $boolText (Get-PropertyValue -Object $kmip -PropertyName "port")),
            (& $boolText (Get-PropertyValue -Object $kmip -PropertyName "enabled"))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Cluster Name", "Servers", "Port", "Enabled") -Rows $kmipRows

    # --- Storage Proxy ---
    $markdownLines.Add("## Storage Proxy")
    $markdownLines.Add("")
    $markdownLines.Add("A storage proxy allows Storage Nodes to reach an external HTTP or HTTPS proxy server to communicate with an external S3-compatible endpoint, such as a Cloud Storage Pool target, when direct outbound access is not available.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows whether a storage proxy is configured and, if so, its connection details.")
    $markdownLines.Add("")
    $storageProxyConfig = Get-PropertyValue -Object $storageProxy -PropertyName "proxy"
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Enabled", "Hostname", "Port", "Username") -Rows @(
        ,@(
            (& $boolText (Get-PropertyValue -Object $storageProxyConfig -PropertyName "enable")),
            (& $boolText (Get-PropertyValue -Object $storageProxyConfig -PropertyName "hostname")),
            (& $boolText (Get-PropertyValue -Object $storageProxyConfig -PropertyName "hostPort")),
            (& $boolText (Get-PropertyValue -Object $storageProxyConfig -PropertyName "username"))
        )
    )

    # --- Admin Proxy ---
    $markdownLines.Add("")
    $markdownLines.Add("## Admin Proxy")
    $markdownLines.Add("")
    $markdownLines.Add("An admin proxy allows Admin Nodes to reach an external HTTP or HTTPS proxy server for outbound connections such as AutoSupport messages to technical support or NTP/DNS lookups where direct internet access is restricted.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows whether an admin proxy is configured and, if so, its connection details and whether a custom CA bundle has been supplied.")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Enabled", "Hostname", "Port", "Username", "CA Bundle") -Rows @(
        ,@(
            (& $boolText (Get-PropertyValue -Object $adminProxy -PropertyName "enable")),
            (& $boolText (Get-PropertyValue -Object $adminProxy -PropertyName "hostname")),
            (& $boolText (Get-PropertyValue -Object $adminProxy -PropertyName "hostPort")),
            (& $boolText (Get-PropertyValue -Object $adminProxy -PropertyName "username")),
            $(if ([string]::IsNullOrWhiteSpace([string](Get-PropertyValue -Object $adminProxy -PropertyName "caBundle"))) { "Absent" } else { "Present" })
        )
    )

    # --- Untrusted Client Network ---
    $markdownLines.Add("")
    $markdownLines.Add("## Untrusted Client Network")
    $markdownLines.Add("")
    $markdownLines.Add("By default, the optional Client Network on each node is untrusted, meaning inbound connections are limited to load-balancer endpoint ports only, to reduce the network attack surface exposed to client applications.")
    $markdownLines.Add("")
    $markdownLines.Add("This table lists the nodes whose Client Network is configured as untrusted. A node not in this list is considered trusted, meaning it can accept client connections on any of its configured load-balancer endpoint ports (including management ports).")
    $markdownLines.Add("")
    $untrustedNodeIds = @((Get-PropertyValue -Object $untrustedClientNet -PropertyName "untrustedNodes"))
    $nodeHealthByIdForUntrusted = @{}
    foreach ($healthNode in $nodeHealth) {
        $healthNodeId = [string](Get-PropertyValue -Object $healthNode -PropertyName "id")
        if (-not [string]::IsNullOrWhiteSpace($healthNodeId)) {
            $nodeHealthByIdForUntrusted[$healthNodeId] = $healthNode
        }
    }

    $untrustedNodeRows = @()
    $mappedUntrustedNodes = [object[]]@(
        @(
            foreach ($untrustedNodeId in $untrustedNodeIds) {
                $untrustedNodeIdText = [string]$untrustedNodeId
                if ($nodeHealthByIdForUntrusted.ContainsKey($untrustedNodeIdText)) {
                    $healthNode = $nodeHealthByIdForUntrusted[$untrustedNodeIdText]
                    [pscustomobject]@{
                        SiteName = [string](Get-PropertyValue -Object $healthNode -PropertyName "siteName")
                        NodeName = [string](Get-PropertyValue -Object $healthNode -PropertyName "name")
                    }
                }
            }
        ) | Sort-Object SiteName, NodeName
    )

    if ($mappedUntrustedNodes.Count -gt 0) {
        foreach ($mappedNodeIndex in 0..($mappedUntrustedNodes.Count - 1)) {
            $mappedNode = $mappedUntrustedNodes[$mappedNodeIndex]
            $defaultModeCell = if ($mappedNodeIndex -eq 0) {
                (& $boolText (Get-PropertyValue -Object $untrustedClientNet -PropertyName "default"))
            } else {
                ""
            }
            $untrustedNodeRows += ,@($defaultModeCell, [string]$mappedNode.SiteName, [string]$mappedNode.NodeName)
        }
    }
    if ($untrustedNodeRows.Count -eq 0) {
        $untrustedNodeRows = ,@((& $boolText (Get-PropertyValue -Object $untrustedClientNet -PropertyName "default")), "N/A", "N/A")
    }
    & $emitOpenXmlMergedFirstColumnTable -Headers @("Default Mode", "Grid Site", "Node Name") -Rows $untrustedNodeRows -GridWidths @(1800, 2400, 3600) -AutofitToWindow

    # --- External Load Balancers ---
    $markdownLines.Add("")
    $markdownLines.Add("## External Load Balancers")
    $markdownLines.Add("")
    $markdownLines.Add("The list of trusted external load balancer IP addresses identifies load balancers that StorageGRID trusts as intermediaries for client connections. This allows StorageGRID to trust the client IP address values in X-Forwarded-For and X-Real-IP headers set by these external load balancers, so that traffic classification, audit logs, and untrusted client network rules can accurately reflect the original client IP address rather than the load balancer's own address.")
    $markdownLines.Add("")
    $externalLbRows = @()
    foreach ($externalLb in $externalLoadBalancers) {
        $externalLbRows += ,@([string]$externalLb)
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Trusted External Load Balancer IP") -Rows $externalLbRows

    # --- Configuration - Firewall and Appliance Security ---
    $markdownLines.Add("")    
    $markdownLines.Add("# Configuration - Firewall and Appliance Security")

    # --- Firewall Blocked Ports ---
    $markdownLines.Add("")
    $markdownLines.Add("## Firewall Blocked Ports")
    $markdownLines.Add("")
    $markdownLines.Add("The firewall configuration allows administrators to explicitly block inbound ports on all nodes in the grid, in addition to the ports that are blocked by default based on node type. Blocking unused ports reduces the network attack surface of the grid.")
    $markdownLines.Add("")
    $firewallBlockedRows = @()
    foreach ($blockedEntry in $firewallBlockedPorts) {
        $blockedTcp = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $blockedEntry -PropertyName "tcpPorts")
        $blockedUdp = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $blockedEntry -PropertyName "udpPorts")
        $firewallBlockedRows += ,@(
            (& $boolText (Get-PropertyValue -Object $blockedEntry -PropertyName "id")),
            $(if ($blockedTcp.Count -gt 0) { ($blockedTcp -join "`n") } else { "N/A" }),
            $(if ($blockedUdp.Count -gt 0) { ($blockedUdp -join "`n") } else { "N/A" })
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Node Id", "Blocked TCP Ports", "Blocked UDP Ports") -Rows $firewallBlockedRows

    $markdownLines.Add("")
    $markdownLines.Add("## Firewall External Ports")
    $markdownLines.Add("")
    $markdownLines.Add("The external firewall configuration controls which TCP and UDP ports remain open to external networks (the Client and Admin networks). Ports not explicitly opened here are blocked by the internal StorageGRID firewall, reducing the exposed network attack surface while still allowing required client and management access.")
    $markdownLines.Add("")
    $externalTcpPorts = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $firewallExternalPorts -PropertyName "externalTcpPorts")
    $externalUdpPorts = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $firewallExternalPorts -PropertyName "externalUdpPorts")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
        @("External TCP Ports", $(if ($externalTcpPorts.Count -gt 0) { ($externalTcpPorts -join "`n") } else { "N/A" })),
        @("External UDP Ports", $(if ($externalUdpPorts.Count -gt 0) { ($externalUdpPorts -join "`n") } else { "N/A" }))
    )

    $markdownLines.Add("")
    $markdownLines.Add("## Firewall Privileged IPs")
    $markdownLines.Add("")
    $markdownLines.Add("Privileged IP addresses and subnets are exempted from certain firewall restrictions and are always permitted to reach internal-only ports used for grid administration and internal communication. The Grid Internal Access setting also controls whether nodes on the grid network are automatically treated as privileged.")
    $markdownLines.Add("")
    $privilegedIps = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $firewallPrivilegedIps -PropertyName "privilegedIps")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
        @("Privileged IPs", $(if ($privilegedIps.Count -gt 0) { ($privilegedIps -join "`n") } else { "N/A" })),
        @("Grid Internal Access", (& $boolText (Get-PropertyValue -Object $firewallPrivilegedIps -PropertyName "gridInternalAccess")))
    )

    # --- TLS and SSH ---
    $markdownLines.Add("")
    $markdownLines.Add("## TLS and SSH")
    $markdownLines.Add("")
    $markdownLines.Add("This configuration defines the FIPS 140 compliance mode and the allowed TLS and SSH cipher suites and protocol versions used for the grid's inbound, internal, and outbound network connections. Restricting the allowed ciphers and protocols is part of hardening the grid to meet organizational or regulatory security requirements.")
    $markdownLines.Add("")
    $tlsInbound = Get-PropertyValue -Object $ciphersConfig -PropertyName "tlsInbound"
    $tlsInternal = Get-PropertyValue -Object $ciphersConfig -PropertyName "tlsInternal"
    $tlsOutbound = Get-PropertyValue -Object $ciphersConfig -PropertyName "tlsOutbound"
    $sshCiphers = Get-PropertyValue -Object $ciphersConfig -PropertyName "ssh"
    $cipherProtocolsDisplay = {
        param($Section)
        $protocols = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $Section -PropertyName "protocols")
        if ($protocols.Count -gt 0) { $protocols -join ", " } else { "N/A" }
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
        @("FIPS Mode", (& $boolText (Get-PropertyValue -Object $ciphersConfig -PropertyName "fipsMode"))),
        @("TLS Inbound Protocols", (& $cipherProtocolsDisplay $tlsInbound)),
        @("TLS Internal Protocols", (& $cipherProtocolsDisplay $tlsInternal)),
        @("TLS Outbound Protocols", (& $cipherProtocolsDisplay $tlsOutbound)),
        @("SSH Ciphers Configured", (& $boolText ((Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $sshCiphers -PropertyName "ciphers")).Count -gt 0)))
    )

    # --- Remote IPMI Access ---
    $markdownLines.Add("")
    $markdownLines.Add("## Remote IPMI Access")
    $markdownLines.Add("")
    $markdownLines.Add("This setting controls whether remote IPMI (Intelligent Platform Management Interface) access to appliance baseboard management controllers (BMCs) is enabled grid-wide.")
    $markdownLines.Add("")
    $markdownLines.Add("The BMC provides out-of-band hardware management, such as power control and hardware health monitoring, independent of the StorageGRID software stack; restricting IPMI access reduces the attack surface for appliance hardware management interfaces.")
    $markdownLines.Add("")
    $bmcNetworking = Get-PropertyValue -Object $bmcConfig -PropertyName "networking"
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
        ,@("Remote IPMI Access Enabled", (& $boolText (Get-PropertyValue -Object $bmcNetworking -PropertyName "remoteIpmiAccessEnabled")))
    )

    # --- Configuration - Network ---
    $markdownLines.Add("")
    $markdownLines.Add("## Configuration - Network")
    $markdownLines.Add("")
    # --- S3 Endpoint Domain Names ---
    $markdownLines.Add("## S3 Endpoint Domain Names")
    $markdownLines.Add("")
    $markdownLines.Add("Endpoint domain names are DNS wildcard names configured for the grid so that S3 clients can use virtual-hosted-style requests (for example, bucket.s3.example.com) instead of path-style requests.")
    $markdownLines.Add("")
    $markdownLines.Add("This table lists the domain names registered for this purpose; each domain typically requires a corresponding wildcard DNS entry pointing to the grid's Load Balancer endpoints.")
    $markdownLines.Add("")
    $domainRows = @()
    foreach ($domain in @($domainNames)) {
        if ([string]::IsNullOrWhiteSpace([string]$domain)) { continue }
        $domainRows += ,@([string]$domain)
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Domain Name") -Rows $domainRows

    # --- HA Groups ---
    $markdownLines.Add("")
    $markdownLines.Add("## HA Groups")
    $markdownLines.Add("")
    $markdownLines.Add("A high-availability (HA) group binds one or more virtual IP addresses to a set of network interfaces on one or more Admin or Gateway Nodes, using the VRRP protocol to automatically fail over the virtual IP to a backup interface if the active node or interface becomes unavailable.")
    $markdownLines.Add("")
    $markdownLines.Add("HA groups are commonly used to provide a stable, highly available IP address for load-balanced S3 client traffic or Grid Manager access.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows each HA group's virtual IP addresses, member node interfaces, and configured gateway.")
    $markdownLines.Add("")
    $haRows = @()
    foreach ($ha in $haGroups) {
        $virtualIps = @((Get-PropertyValue -Object $ha -PropertyName "virtualIps"))
        $interfaces = @((Get-PropertyValue -Object $ha -PropertyName "interfaces"))
        $interfaceItems = @()
        foreach ($iface in $interfaces) {
            $ifaceNodeId = [string](Get-PropertyValue -Object $iface -PropertyName "nodeId")
            $ifaceName = [string](Get-PropertyValue -Object $iface -PropertyName "interface")
            $ifaceNodeName = $ifaceNodeId
            if ($nodeHealthById.ContainsKey($ifaceNodeId)) {
                $ifaceNodeName = [string](Get-PropertyValue -Object $nodeHealthById[$ifaceNodeId] -PropertyName "name")
            }
            $interfaceItems += "${ifaceNodeName}:${ifaceName}"
        }
        $haRows += ,@(
            (& $boolText (Get-PropertyValue -Object $ha -PropertyName "name")),
            (& $boolText (Get-PropertyValue -Object $ha -PropertyName "description")),
            $(if ($virtualIps.Count -gt 0) { $virtualIps -join ", " } else { "N/A" }),
            $(if ($interfaceItems.Count -gt 0) { $interfaceItems -join ", " } else { "N/A" }),
            (& $boolText (Get-PropertyValue -Object $ha -PropertyName "gatewayCidr"))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("HA Group", "Description", "Virtual IPs", "Interfaces", "Gateway") -Rows $haRows

    # --- Load Balancer Endpoints ---
    $markdownLines.Add("")
    $markdownLines.Add("## Load Balancer Endpoints")
    $markdownLines.Add("")
    $markdownLines.Add("A Load Balancer endpoint defines a listening port on the StorageGRID Load Balancer service (running on Admin and/or Gateway Nodes) that accepts S3 client connections and forwards them to the appropriate Storage Nodes.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows each endpoint's port, whether it uses TLS (secure), whether it is restricted to a specific tenant account, whether it is pinned to specific HA groups, and whether it is automatically closed on the untrusted Client Network.")
    $markdownLines.Add("")
    $accountById = @{}
    foreach ($acct in $gridAccounts) {
        $acctId = [string](Get-PropertyValue -Object $acct -PropertyName "id")
        if (-not [string]::IsNullOrWhiteSpace($acctId)) {
            $accountById[$acctId] = $acct
        }
    }

    $haNameById = @{}
    foreach ($ha in $haGroups) {
        $haId = [string](Get-PropertyValue -Object $ha -PropertyName "id")
        if (-not [string]::IsNullOrWhiteSpace($haId)) {
            $haNameById[$haId] = [string](Get-PropertyValue -Object $ha -PropertyName "name")
        }
    }

    $lbEndpointRows = @()
    $endpointHaRows = @()
    foreach ($endpoint in $gatewayConfigs) {
        $pinTargets = Get-PropertyValue -Object $endpoint -PropertyName "pinTargets"
        $pinHaGroups = @((Get-PropertyValue -Object $pinTargets -PropertyName "haGroups"))
        $endpointName = [string](Get-PropertyValue -Object $endpoint -PropertyName "displayName")
        $endpointAccountId = [string](Get-PropertyValue -Object $endpoint -PropertyName "accountId")
        $endpointAccountName = $endpointAccountId
        if ($accountById.ContainsKey($endpointAccountId)) {
            $endpointAccountName = [string](Get-PropertyValue -Object $accountById[$endpointAccountId] -PropertyName "name")
        }

        $lbEndpointRows += ,@(
            (& $boolText $endpointName),
            (& $boolText (Get-PropertyValue -Object $endpoint -PropertyName "port")),
            (& $boolText (Get-PropertyValue -Object $endpoint -PropertyName "secure")),
            (& $boolText $endpointAccountName),
            (& $boolText (Get-PropertyValue -Object $endpoint -PropertyName "closedOnUntrustedClientNetwork"))
        )

        if ($pinHaGroups.Count -eq 0) {
            $endpointHaRows += ,@((& $boolText $endpointName), "N/A")
        } else {
            for ($haIndex = 0; $haIndex -lt $pinHaGroups.Count; $haIndex++) {
                $haId = [string]$pinHaGroups[$haIndex]
                $haName = if ($haNameById.ContainsKey($haId)) { [string]$haNameById[$haId] } else { "N/A - Unresolved HA Group ID" }
                $endpointHaRows += ,@($(if ($haIndex -eq 0) { (& $boolText $endpointName) } else { "" }), $haName)
            }
        }
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Endpoint Name", "Port", "Secure", "Account", "Closed On Untrusted") -Rows $lbEndpointRows

    $markdownLines.Add("")
    # --- Endpoint HA Group Mapping ---
    $markdownLines.Add("### Endpoint HA Group Mapping")
    $markdownLines.Add("")
    $markdownLines.Add("This table maps each Load Balancer endpoint to the HA groups through which the endpoint is accessible. Endpoints without pinned HA groups are shown as N/A; an endpoint with no pinned groups can be available through all eligible Load Balancer nodes, subject to its network and endpoint configuration.")
    $markdownLines.Add("")
    & $emitOpenXmlMergedFirstColumnTable -Headers @("Endpoint Name", "HA Group Name") -Rows $endpointHaRows -GridWidths @(3600, 4200) -AutofitToWindow

    # --- Traffic Classification Policies ---
    $markdownLines.Add("")
    $markdownLines.Add("## Traffic Classification Policies")
    $markdownLines.Add("")
    $markdownLines.Add("Traffic classification policies let administrators identify specific types of Load Balancer traffic, based on matchers such as bucket name, tenant account, CIDR, or endpoint, and apply limits to that traffic, such as bandwidth or concurrent connection limits, or relative traffic-shaping priority.")
    $markdownLines.Add("")
    $markdownLines.Add("This is commonly used to prevent a single noisy tenant or bucket from monopolizing Load Balancer resources at the expense of other workloads.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows the number of matchers and limits configured for each policy; see the following sections for the full matcher and limit details.")
    $markdownLines.Add("")
    $trafficDetailById = @{}
    foreach ($tpd in $trafficPolicyDetails) {
        $tpdId = [string](Get-PropertyValue -Object $tpd -PropertyName "id")
        if (-not [string]::IsNullOrWhiteSpace($tpdId)) { $trafficDetailById[$tpdId] = $tpd }
    }
    $trafficRows = @()
    foreach ($policy in $trafficPolicies) {
        $policyId = [string](Get-PropertyValue -Object $policy -PropertyName "id")
        $detail = $null
        if ($trafficDetailById.ContainsKey($policyId)) { $detail = $trafficDetailById[$policyId] }
        $trafficRows += ,@(
            (& $boolText (Get-PropertyValue -Object $policy -PropertyName "name")),
            (& $boolText (Get-PropertyValue -Object $policy -PropertyName "description")),
            [string]@((Get-PropertyValue -Object $detail -PropertyName "matchers")).Count,
            [string]@((Get-PropertyValue -Object $detail -PropertyName "limits")).Count
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Policy Name", "Description", "Matcher Count", "Limit Count") -Rows $trafficRows

    $markdownLines.Add("")
    $markdownLines.Add("## Traffic Classification Policy Details")
    $markdownLines.Add("")
    $markdownLines.Add("The following sections show the matching rules and traffic limits configured for each traffic classification policy. Matcher members are resolved to tenant names where the matcher type is tenant; other matcher members are shown as returned by the StorageGRID API.")

    if ($trafficPolicies.Count -eq 0) {
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
            ,@("Policy Name", "N/A")
        )
    }

    foreach ($policy in $trafficPolicies) {
        $policyId = [string](Get-PropertyValue -Object $policy -PropertyName "id")
        $policyName = [string](& $boolText (Get-PropertyValue -Object $policy -PropertyName "name"))
        $policyDetail = $null
        if ($trafficDetailById.ContainsKey($policyId)) {
            $policyDetail = $trafficDetailById[$policyId]
        }

        $markdownLines.Add("")
        $markdownLines.Add("### Traffic Classification Policy Details - $policyName")
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
            @("Policy Name", $policyName),
            @("Description", (& $boolText (Get-PropertyValue -Object $policy -PropertyName "description")))
        )

        $markdownLines.Add("")
        $markdownLines.Add("#### Matchers")
        $markdownLines.Add("")
        $matcherTypes = @()
        $tenantNames = @()
        $tenantIds = @()
        $matcherValues = @()
        foreach ($matcher in (Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $policyDetail -PropertyName "matchers"))) {
            $matcherType = [string](& $boolText (Get-PropertyValue -Object $matcher -PropertyName "type"))
            if (-not [string]::IsNullOrWhiteSpace($matcherType)) {
                $matcherTypes += $matcherType
            }
            $members = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $matcher -PropertyName "members")
            foreach ($member in $members) {
                $memberId = [string]$member
                if ([string]::IsNullOrWhiteSpace($memberId)) { continue }

                if ($matcherType -eq "tenant") {
                    $tenantName = "N/A"
                    if ($accountById.ContainsKey($memberId)) {
                        $tenantName = [string](Get-PropertyValue -Object $accountById[$memberId] -PropertyName "name")
                        if ([string]::IsNullOrWhiteSpace($tenantName)) { $tenantName = "N/A" }
                    }
                    $tenantNames += $tenantName
                    $tenantIds += $memberId
                } else {
                    $matcherValues += $memberId
                }
            }
        }
        $matcherRows = @(
            ,@(
                $(if ($matcherTypes.Count -gt 0) { $matcherTypes -join "`n" } else { "N/A" }),
                $(if ($tenantNames.Count -gt 0) { $tenantNames -join "`n" } else { "N/A" }),
                $(if ($tenantIds.Count -gt 0) { $tenantIds -join "`n" } else { "N/A" }),
                $(if ($matcherValues.Count -gt 0) { $matcherValues -join "`n" } else { "N/A" })
            )
        )
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Matcher Type", "Tenant Account Name", "Tenant Account ID", "Match Value") -Rows $matcherRows

        $markdownLines.Add("")
        $markdownLines.Add("#### Limits")
        $markdownLines.Add("")
        $limitValuesByName = @{}
        foreach ($limit in (Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $policyDetail -PropertyName "limits"))) {
            foreach ($limitProperty in @($limit.PSObject.Properties | Where-Object { $_.Name -ne "id" })) {
                $limitPropertyName = [string]$limitProperty.Name
                if (-not $limitValuesByName.ContainsKey($limitPropertyName)) {
                    $limitValuesByName[$limitPropertyName] = @()
                }

                $limitValueItems = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $limit -PropertyName $limitPropertyName)
                foreach ($limitValueItem in $limitValueItems) {
                    $limitValueText = [string]$limitValueItem
                    if (-not [string]::IsNullOrWhiteSpace($limitValueText)) {
                        $limitValuesByName[$limitPropertyName] += $limitValueText
                    }
                }
            }
        }
        $limitRows = @()
        foreach ($limitPropertyName in @($limitValuesByName.Keys | Sort-Object)) {
            $limitValues = @($limitValuesByName[$limitPropertyName])
            $limitRows += ,@(
                $limitPropertyName,
                $(if ($limitValues.Count -gt 0) { $limitValues -join "`n" } else { "N/A" })
            )
        }
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Limit Setting", "Value") -Rows $limitRows
    }

    # --- VLAN Interfaces ---
    $markdownLines.Add("")
    $markdownLines.Add("## VLAN Interfaces")
    $markdownLines.Add("")
    $markdownLines.Add("You can create virtual LAN (VLAN) interfaces on Admin Nodes and Gateway Nodes and use them in HA groups and load balancer endpoints to isolate and partition traffic for security, flexibility, and performance. The selected nodes in the HA group can use the VLAN interfaces to share up to 10 virtual IP addresses, so that if a node goes down, another node takes over traffic to and from the virtual IP addresses.")
    $markdownLines.Add("")
    $markdownLines.Add("Each VLAN is identified by a numeric ID or tag.")
    $markdownLines.Add("")
    $markdownLines.Add("You can use the Grid Manager to create VLAN interfaces that allow clients to access StorageGRID on a specific VLAN. When you create VLAN interfaces, you specify the VLAN ID and select parent (trunk) interfaces on one or more nodes.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table lists each configured VLAN interface.")
    $markdownLines.Add("")
    $vlanNodeNameById = @{}
    foreach ($healthNode in $nodeHealth) {
        $healthNodeId = [string](Get-PropertyValue -Object $healthNode -PropertyName "id")
        $healthNodeName = [string](Get-PropertyValue -Object $healthNode -PropertyName "name")
        if (-not [string]::IsNullOrWhiteSpace($healthNodeId) -and -not [string]::IsNullOrWhiteSpace($healthNodeName)) {
            $vlanNodeNameById[$healthNodeId] = $healthNodeName
        }
    }
    $vlanRows = @()
    foreach ($vlan in $vlanInterfaces) {
        $vlanNodeId = [string](Get-PropertyValue -Object $vlan -PropertyName "nodeId")
        $vlanNodeName = [string](Get-PropertyValue -Object $vlan -PropertyName "nodeName")
        if ([string]::IsNullOrWhiteSpace($vlanNodeName) -and $vlanNodeNameById.ContainsKey($vlanNodeId)) {
            $vlanNodeName = [string]$vlanNodeNameById[$vlanNodeId]
        }
        $vlanRows += ,@(
            (& $boolText (Get-PropertyValue -Object $vlan -PropertyName "name")),
            (& $boolText (Get-PropertyValue -Object $vlan -PropertyName "id")),
            (& $boolText (Get-PropertyValue -Object $vlan -PropertyName "vlanId")),
            (& $boolText (Get-PropertyValue -Object $vlan -PropertyName "parentInterface")),
            (& $boolText $vlanNodeName)
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("VLAN Name", "VLAN Id", "Tag", "Parent Interface", "Node Name") -Rows $vlanRows

    # --- Management Interface CORS ---
    $markdownLines.Add("")
    $markdownLines.Add("## Management Interface CORS")
    $markdownLines.Add("")
    $markdownLines.Add("Cross-origin resource sharing (CORS) controls which browser-based applications hosted on other origins can make requests to the StorageGRID Grid Manager and Tenant Manager interfaces. Management Interface CORS was introduced in StorageGRID 12.0.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows whether cross-origin requests are allowed for each management interface and lists the explicitly permitted origins.")
    $markdownLines.Add("")
    $allowedCorsOrigins = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $managementCors -PropertyName "allowedOrigins")
    $allowedCorsOriginsDisplay = if ($allowedCorsOrigins.Count -gt 0) { $allowedCorsOrigins -join "`n" } else { "N/A" }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Setting", "Value") -Rows @(
        @("Allow Grid Manager", (& $boolText (Get-PropertyValue -Object $managementCors -PropertyName "allowGridManager"))),
        @("Allow Tenant Manager", (& $boolText (Get-PropertyValue -Object $managementCors -PropertyName "allowTenantManager"))),
        @("Allowed Origins", $allowedCorsOriginsDisplay)
    )

    # --- Alerts ---
    $markdownLines.Add("")
    $markdownLines.Add("# Alerts")
    $markdownLines.Add("")
    # --- Alert Setup ---
    $markdownLines.Add("## Alert Setup")
    $markdownLines.Add("")
    $markdownLines.Add("Alert notifications are sent by email to configured recipients when an alert reaches a specified severity, keeping administrators informed without requiring them to actively monitor the Grid Manager dashboard.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows the configured email settings, recipient list, and the minimum alert severity that triggers a notification to that receiver.")
    $markdownLines.Add("")
    foreach ($receiver in $alertReceivers) {
        $toEmails = @((Get-PropertyValue -Object $receiver -PropertyName "toEmails"))
        $username = Get-PropertyValue -Object $receiver -PropertyName "username"
        $usernameDisplay = if ([string]::IsNullOrWhiteSpace([string]$username)) { "N/A" } else { [string]$username }
        $fromEmail = Get-PropertyValue -Object $receiver -PropertyName "fromEmail"
        $fromEmailDisplay = if ([string]::IsNullOrWhiteSpace([string]$fromEmail)) { "N/A" } else { [string]$fromEmail }

        $caCert = Get-PropertyValue -Object $receiver -PropertyName "caCert"
        $caCertDisplay = if ([string]::IsNullOrWhiteSpace([string]$caCert)) { "N/A" } else { "True" }
        $clientCert = Get-PropertyValue -Object $receiver -PropertyName "clientCert"
        $clientCertDisplay = if ([string]::IsNullOrWhiteSpace([string]$clientCert)) { "N/A" } else { "True" }
        $clientKey = Get-PropertyValue -Object $receiver -PropertyName "clientKey"
        $clientKeyDisplay = if ([string]::IsNullOrWhiteSpace([string]$clientKey)) { "N/A" } else { "True" }

        $receiverRows = @(
            @("Type", (& $boolText (Get-PropertyValue -Object $receiver -PropertyName "type"))),
            @("Enabled", (& $boolText (Get-PropertyValue -Object $receiver -PropertyName "enable"))),
            @("SMTP Host", (& $boolText (Get-PropertyValue -Object $receiver -PropertyName "smtpHost"))),
            @("SMTP Port", (& $boolText (Get-PropertyValue -Object $receiver -PropertyName "smtpPort"))),
            @("Username", $usernameDisplay),
            @("From Address", $fromEmailDisplay),
            @("Minimum Severity", (& $boolText (Get-PropertyValue -Object $receiver -PropertyName "minimumSeverity")))
        )

        if ($toEmails.Count -eq 0) {
            $receiverRows += ,@("Recipients", "N/A")
        } else {
            for ($emailIndex = 0; $emailIndex -lt $toEmails.Count; $emailIndex++) {
                $recipientLabel = if ($emailIndex -eq 0) { "Recipients" } else { "" }
                $receiverRows += ,@($recipientLabel, [string]$toEmails[$emailIndex])
            }
        }

        $receiverRows += ,@("CA Bundle Present", $caCertDisplay)
        $receiverRows += ,@("Client Cert Present", $clientCertDisplay)
        $receiverRows += ,@("Client Key Present", $clientKeyDisplay)

        $markdownLines.Add("")
        & $emitOpenXmlMergedFirstColumnTable -Headers @("Field", "Value") -Rows $receiverRows -GridWidths @(3120, 4680) -AutofitToWindow
    }
    if ($alertReceivers.Count -eq 0) {
        $markdownLines.Add("")
        & $emitOpenXmlMergedFirstColumnTable -Headers @("Field", "Value") -Rows @() -GridWidths @(3120, 4680) -AutofitToWindow
    }

    # --- Custom Alert Rules ---
    $markdownLines.Add("")
    $markdownLines.Add("## Custom Alert Rules")
    $markdownLines.Add("")
    $markdownLines.Add("Alert rules define the conditions that trigger StorageGRID alerts. The system includes a comprehensive set of default alert rules covering appliance, network, storage, and service health, which can be used as-is, modified, or disabled.")
    $markdownLines.Add("")
    $markdownLines.Add("Administrators can also create custom alert rules using Prometheus expressions to detect conditions specific to their environment.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table lists only the custom alert rules defined on this grid; default alert rules are omitted for brevity.")
    $markdownLines.Add("")
    $customAlertRuleRows = @()
    foreach ($rule in $alertRules) {
        $isCustomRule = Get-PropertyValue -Object $rule -PropertyName "custom"
        if ($isCustomRule -ne $true) { continue }

        $expressionsObject = Get-PropertyValue -Object $rule -PropertyName "expressions"
        $expressionParts = @()
        if ($null -ne $expressionsObject) {
            foreach ($expressionProp in @($expressionsObject.PSObject.Properties)) {
                $expressionParts += "$($expressionProp.Name): $($expressionProp.Value)"
            }
        }
        $expressionDisplay = if ($expressionParts.Count -gt 0) { $expressionParts -join "`n" } else { "N/A" }

        $customAlertRuleRows += ,@(
            (& $boolText (Get-PropertyValue -Object $rule -PropertyName "name")),
            (& $boolText (Get-PropertyValue -Object $rule -PropertyName "enable")),
            (& $boolText (Get-PropertyValue -Object $rule -PropertyName "for")),
            $expressionDisplay,
            (& $boolText (Get-PropertyValue -Object (Get-PropertyValue -Object $rule -PropertyName "annotations") -PropertyName "description"))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Rule Name", "Enabled", "Duration", "Conditions", "Description") -Rows $customAlertRuleRows

    # --- Alert Silences ---
    $markdownLines.Add("")
    $markdownLines.Add("## Alert Silences")
    $markdownLines.Add("")
    $markdownLines.Add("Alert silences allow administrators to temporarily suppress notifications for a specific alert rule (or all alert rules) across the entire grid, a single site, or a single node, for one or more severities. Silences are typically used to prevent notification noise during planned maintenance or for alerts triggered by an intentional configuration.")
    $markdownLines.Add("")
    $markdownLines.Add("Silencing an alert does not disable the underlying condition monitoring; the alert is still evaluated and shown in the Grid Manager, but email and SNMP notifications are suppressed for its duration.")
    $markdownLines.Add("")
    $alertSilenceRows = @()
    foreach ($silence in $alertSilences) {
        $alertSilenceRows += ,@(
            (& $boolText (Get-PropertyValue -Object $silence -PropertyName "description")),
            (& $boolText (Get-PropertyValue -Object $silence -PropertyName "ruleUniqueName")),
            (& $boolText (Get-PropertyValue -Object $silence -PropertyName "severities")),
            (& $boolText (Get-PropertyValue -Object $silence -PropertyName "nodes")),
            (& $boolText (Get-PropertyValue -Object $silence -PropertyName "ends"))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Description", "Alert Rule", "Severities", "Scope", "Ends") -Rows $alertSilenceRows


    # --- ILM - Storage Configuration ---
    $markdownLines.Add("")
    $markdownLines.Add("# ILM - Storage Configuration")
    $markdownLines.Add("")
    # --- Storage Pools ---
    $markdownLines.Add("## Storage Pools")
    $markdownLines.Add("")
    $markdownLines.Add("A storage pool is a logical group of Storage Nodes, defined by site and/or storage grade, that ILM rules use as a placement target for object data. Storage pools let administrators direct object copies to specific hardware tiers (for example, faster or slower disk grades) or specific sites without needing to reference individual nodes directly in ILM rules.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows each storage pool's member storage node count and the site/grade criteria used to select its member nodes.")
    $markdownLines.Add("")

    $ilmSiteNameById = @{}
    foreach ($nh in $sortedNodeHealth) {
        $nhSiteId = [string](Get-PropertyValue -Object $nh -PropertyName "siteId")
        $nhSiteName = [string](Get-PropertyValue -Object $nh -PropertyName "siteName")
        if (-not [string]::IsNullOrWhiteSpace($nhSiteId) -and -not [string]::IsNullOrWhiteSpace($nhSiteName)) {
            $ilmSiteNameById[$nhSiteId] = $nhSiteName
        }
    }

    $ilmGradeNameById = @{}
    foreach ($grade in $ilmGrades) {
        $gradeId = [string](Get-PropertyValue -Object $grade -PropertyName "id")
        $gradeName = [string](Get-PropertyValue -Object $grade -PropertyName "name")
        if (-not [string]::IsNullOrWhiteSpace($gradeId) -and -not [string]::IsNullOrWhiteSpace($gradeName)) {
            $ilmGradeNameById[$gradeId] = $gradeName
        }
    }

    $ilmGradeSiteNameById = @{}
    foreach ($gradeSite in $ilmGradeSite) {
        $gradeSiteId = [string](Get-PropertyValue -Object $gradeSite -PropertyName "id")
        $gradeSiteName = [string](Get-PropertyValue -Object $gradeSite -PropertyName "name")
        if (-not [string]::IsNullOrWhiteSpace($gradeSiteId) -and -not [string]::IsNullOrWhiteSpace($gradeSiteName)) {
            $ilmGradeSiteNameById[$gradeSiteId] = $gradeSiteName
        }
    }

    $ilmPoolById = @{}
    $ilmPoolRows = @()
    foreach ($pool in $ilmPools) {
        $poolId = [string](Get-PropertyValue -Object $pool -PropertyName "id")
        if (-not [string]::IsNullOrWhiteSpace($poolId)) { $ilmPoolById[$poolId] = $pool }

        $diskEntries = @((Get-PropertyValue -Object $pool -PropertyName "disks"))
        $diskLayout = @()
        foreach ($disk in $diskEntries) {
            $diskSiteId = [string](Get-PropertyValue -Object $disk -PropertyName "siteId")
            $diskSiteName = if ($ilmSiteNameById.ContainsKey($diskSiteId)) { $ilmSiteNameById[$diskSiteId] } else { "N/A - Unresolved Site ID" }

            $diskGradeId = [string](Get-PropertyValue -Object $disk -PropertyName "grade")
            $diskGradeName = if ($ilmGradeNameById.ContainsKey($diskGradeId)) { $ilmGradeNameById[$diskGradeId] } else { "N/A - Unresolved Grade ID" }

            $diskGroupId = [string](Get-PropertyValue -Object $disk -PropertyName "group")
            $diskGroupName = if ($ilmGradeSiteNameById.ContainsKey($diskGroupId)) { $ilmGradeSiteNameById[$diskGroupId] } else { "N/A - Unresolved Site Group ID" }

            $diskLayout += "Site: $diskSiteName"
            $diskLayout += "Grade: $diskGradeName"
            $diskLayout += "Site Group: $diskGroupName"
        }

        $ilmPoolRows += ,@(
            (& $boolText (Get-PropertyValue -Object $pool -PropertyName "displayName")),
            (& $boolText (Get-PropertyValue -Object $pool -PropertyName "storageNodeCount")),
            [string]@((Get-PropertyValue -Object $pool -PropertyName "archiveNodes")).Count,
            $(if ($diskLayout.Count -gt 0) { $diskLayout -join "`n" } else { "N/A" })
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Pool Name", "Storage Node Count", "Archive Node Count", "Criteria") -Rows $ilmPoolRows

    # --- Storage Grades ---
    $markdownLines.Add("")
    $markdownLines.Add("## Storage Grades")
    $markdownLines.Add("")
    $markdownLines.Add("A storage grade is an optional label that can be assigned to a Storage Node's storage volumes to distinguish nodes with different underlying disk performance or type (for example, to separate flash from spinning disk).")
    $markdownLines.Add("")
    $markdownLines.Add("Storage grades allow storage pools and ILM rules to target specific classes of hardware rather than treating all Storage Nodes as identical.")
    $markdownLines.Add("")
    $markdownLines.Add("This table lists the storage grade identifiers and names currently defined on the grid.")
    $markdownLines.Add("")
    $ilmGradeRows = @()
    foreach ($grade in $ilmGrades) {
        $ilmGradeRows += ,@(
            (& $boolText (Get-PropertyValue -Object $grade -PropertyName "id")),
            (& $boolText (Get-PropertyValue -Object $grade -PropertyName "name"))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Grade Id", "Grade Name") -Rows $ilmGradeRows

    # --- Storage Grade Node Counts ---
    $markdownLines.Add("")
    $markdownLines.Add("## Storage Grade Node Counts")
    $markdownLines.Add("")
    $markdownLines.Add("This table cross-references each site with the number of Storage Nodes (and Archive Nodes) reporting each configured storage grade, which is useful for verifying that a storage pool's site/grade criteria will resolve to the expected number of member nodes.")
    $markdownLines.Add("")
    $ilmGradeSiteRows = @()
    foreach ($gradeSite in $ilmGradeSite) {
        $gradeSiteName = (& $boolText (Get-PropertyValue -Object $gradeSite -PropertyName "name"))
        $archiveCount = (& $boolText (Get-PropertyValue -Object $gradeSite -PropertyName "archiveCount"))
        $gradesObject = Get-PropertyValue -Object $gradeSite -PropertyName "grades"
        if ($null -eq $gradesObject) {
            $ilmGradeSiteRows += ,@($gradeSiteName, "N/A", "N/A", $archiveCount)
            continue
        }
        foreach ($gradeProp in @($gradesObject.PSObject.Properties)) {
            $gradeLabel = if ($gradeProp.Name -eq "allDisks") {
                "All Disks"
            } elseif ($ilmGradeNameById.ContainsKey([string]$gradeProp.Name)) {
                $ilmGradeNameById[[string]$gradeProp.Name]
            } else {
                "N/A - Unresolved Grade ID ($($gradeProp.Name))"
            }
            $ilmGradeSiteRows += ,@($gradeSiteName, $gradeLabel, [string]$gradeProp.Value, $archiveCount)
        }
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Site Name", "Grade", "Storage Node Count", "Archive Node Count") -Rows $ilmGradeSiteRows

    # --- Cloud Storage Pools ---
    $markdownLines.Add("")
    $markdownLines.Add("## Cloud Storage Pools")
    $markdownLines.Add("")
    $markdownLines.Add("A Cloud Storage Pool is an ILM placement target that stores object copies on an external S3-compatible bucket or Microsoft Azure Blob container instead of on grid Storage Nodes, which can be used for cost-effective long-term retention or as part of a hybrid-cloud data protection strategy.")
    $markdownLines.Add("")
    $markdownLines.Add("This table lists the cloud storage pools configured on the grid; see the following section for each pool's full connection details.")
    $markdownLines.Add("")
    $cloudPoolRows = @()
    foreach ($cloudPool in $cloudStoragePools) {
        $cloudPoolRows += ,@(
            (& $boolText (& $nameOf $cloudPool))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Cloud Pool Name") -Rows $cloudPoolRows

    # --- Cloud Storage Pool Details ---
    $markdownLines.Add("")
    $markdownLines.Add("## Cloud Storage Pool Details")
    $markdownLines.Add("")
    $markdownLines.Add("The following sections provide the full connection configuration for each cloud storage pool, including the target bucket, the S3 or Azure Blob protocol used, the endpoint URI, the authentication type, and whether the target bucket has S3 Object Lock enabled (which is required if the cloud storage pool is used as a placement target for Object Lock-protected objects).")

    if ($cloudStoragePools.Count -eq 0) {
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
            ,@("Cloud Pool Name", "N/A")
        )
    }

    foreach ($cloudPool in $cloudStoragePools) {
        $cloudPoolDisplayName = [string](& $boolText (& $nameOf $cloudPool))

        $markdownLines.Add("")
        $markdownLines.Add("### Cloud Storage Pool Details - $cloudPoolDisplayName")
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
            @("Cloud Pool Name", $cloudPoolDisplayName),
            @("Bucket", (& $boolText (Get-PropertyValue -Object $cloudPool -PropertyName "bucket"))),
            @("Protocol", (& $boolText (Get-PropertyValue -Object $cloudPool -PropertyName "protocol"))),
            @("Endpoint URI", (& $boolText (Get-PropertyValue -Object $cloudPool -PropertyName "endpointURI"))),
            @("Auth Type", (& $boolText (Get-PropertyValue -Object $cloudPool -PropertyName "authType"))),
            @("Object Lock Enabled Target", (& $boolText (Get-PropertyValue -Object $cloudPool -PropertyName "objectLockEnabledTarget")))
        )
    }

    # --- Erasure Coding Profile ---
    $markdownLines.Add("")
    $markdownLines.Add("## Erasure Coding Profile")
    $markdownLines.Add("")
    $markdownLines.Add("Erasure coding (EC) chunks object data into a configurable number of data and parity fragments, distributed across storage pool member nodes, providing storage efficiency and resiliency comparable to or better than full-copy replication while using significantly less raw capacity.")
    $markdownLines.Add("")
    $markdownLines.Add("An erasure coding profile defines a specific EC scheme (data-fragment-count and parity-fragment-count, such as 6+3) applied to a chosen storage pool.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows each defined erasure coding profile and the storage pool it is associated with.")
    $markdownLines.Add("")
    $ecRows = @()
    foreach ($ec in $ecProfiles) {
        $poolId = [string](Get-PropertyValue -Object $ec -PropertyName "poolId")
        $poolName = "N/A"
        if ($ilmPoolById.ContainsKey($poolId)) {
            $poolName = [string](Get-PropertyValue -Object $ilmPoolById[$poolId] -PropertyName "displayName")
        }
        $ecRows += ,@(
            (& $boolText (Get-PropertyValue -Object $ec -PropertyName "name")),
            (& $boolText (Get-PropertyValue -Object $ec -PropertyName "id")),
            (& $boolText $poolName),
            (& $boolText (Get-PropertyValue -Object $ec -PropertyName "schemeId")),
            (& $boolText (Get-PropertyValue -Object $ec -PropertyName "gridNodeRedundancy")),
            (& $boolText (Get-PropertyValue -Object $ec -PropertyName "active"))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Profile Name", "Profile Id", "Storage Pool", "Scheme Id", "Node Redundancy", "Active") -Rows $ecRows

    # --- Regions ---
    $markdownLines.Add("")
    $markdownLines.Add("## Regions")
    $markdownLines.Add("")
    $markdownLines.Add("StorageGRID regions are logical identifiers used by S3 clients when creating and addressing buckets. A region name does not identify a physical StorageGRID site; object placement and data protection are controlled separately through ILM policies and rules. ILM rules can be configured to use the region to make object placement instructiions.")
    $markdownLines.Add("")
    $markdownLines.Add("The following table lists the regions configured for S3 client use.")
    $markdownLines.Add("")
    $gridRegionRows = @()
    foreach ($region in $gridRegions) {
        $regionText = [string]$region
        if (-not [string]::IsNullOrWhiteSpace($regionText)) {
            $gridRegionRows += ,@($regionText)
        }
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Region") -Rows $gridRegionRows

    # --- ILM - Rules and Policies ---
    $markdownLines.Add("")
    $markdownLines.Add("# ILM - Rules and Policies")
    # --- ILM Rule Usage ---
    $markdownLines.Add("")
    $markdownLines.Add("## ILM Rule Usage")
    $markdownLines.Add("")
    $markdownLines.Add("Information Lifecycle Management (ILM) rules define which objects to protect (through filters), how to protect them (through placement instructions such as replication or erasure coding), and where to place their copies (across storage pools and sites). A single ILM rule can be reused by more than one ILM policy. This table cross-references each ILM rule with the ILM policies that reference it, which helps determine the impact of modifying or removing a rule.")
    $markdownLines.Add("")
    $ruleById = @{}
    foreach ($rule in $ilmRules) {
        $ruleId = [string](Get-PropertyValue -Object $rule -PropertyName "id")
        if (-not [string]::IsNullOrWhiteSpace($ruleId)) { $ruleById[$ruleId] = $rule }
    }
    $sortedIlmRules = @(
        $ilmRules | Sort-Object { [string](Get-PropertyValue -Object $_ -PropertyName "displayName") }
    )
    $ruleUsageRows = @()
    foreach ($rule in $sortedIlmRules) {
        $ruleId = [string](Get-PropertyValue -Object $rule -PropertyName "id")
        $usingPolicies = @()
        foreach ($policy in $ilmPolicies) {
            $policyRuleIds = @((Get-PropertyValue -Object $policy -PropertyName "rules"))
            if ($policyRuleIds -contains $ruleId) {
                $usingPolicies += [string](Get-PropertyValue -Object $policy -PropertyName "name")
            }
        }
        $usingPolicies = @($usingPolicies | Sort-Object -Unique)
        $ruleUsageRows += ,@(
            (& $boolText (Get-PropertyValue -Object $rule -PropertyName "displayName")),
            $(if ($usingPolicies.Count -gt 0) { $usingPolicies -join "`n" } else { "N/A" }),
            [string]$usingPolicies.Count
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Rule Name", "Used By Policies", "Policy Count") -Rows $ruleUsageRows

    # --- ILM Rules ---
    $markdownLines.Add("")
    $markdownLines.Add("## ILM Rules")
    $markdownLines.Add("")
    $markdownLines.Add("This table summarizes each ILM rule's compliance compatibility (whether the rule is eligible for use with S3 Object Lock-protected buckets) and the number of filters and placement instructions it contains. See the ILM Rule Details section below for the full configuration of each rule.")
    $markdownLines.Add("")
    $ruleRows = @()
    foreach ($rule in $sortedIlmRules) {
        $ruleRows += ,@(
            (& $boolText (Get-PropertyValue -Object $rule -PropertyName "displayName")),
            (& $boolText (Get-PropertyValue -Object $rule -PropertyName "complianceCompatible")),
            [string]@((Get-PropertyValue -Object $rule -PropertyName "filters")).Count,
            [string]@((Get-PropertyValue -Object $rule -PropertyName "placements")).Count
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Rule Name", "Compliant", "Filter Count", "Placement Count") -Rows $ruleRows

    $ecProfileById = @{}
    foreach ($ec in $ecProfiles) {
        $ecId = [string](Get-PropertyValue -Object $ec -PropertyName "id")
        if (-not [string]::IsNullOrWhiteSpace($ecId)) { $ecProfileById[$ecId] = $ec }
    }

    $markdownLines.Add("")
    # --- ILM Rule Details ---
    $markdownLines.Add("## ILM Rule Details")
    $markdownLines.Add("")
    $markdownLines.Add("The following sections provide the filter and placement configuration for each ILM rule.")
    $markdownLines.Add("")
    $markdownLines.Add("Filters determine which objects a rule applies to, based on criteria such as bucket name, object size, or metadata/tag values; a rule with no filters applies to all objects.")
    $markdownLines.Add("")
    $markdownLines.Add("Placement instructions determine how and where matching objects are protected: replicated placements store a specified number of full copies in a storage pool, while erasure-coded placements split object data into data and parity fragments distributed across a storage pool using a specific erasure coding profile.")
    $markdownLines.Add("")
    $markdownLines.Add("The Retention After value shows how long after the reference time (such as ingest time) each placement instruction remains active before the rule evaluates the next placement instruction, if any.")
    $markdownLines.Add("")

    foreach ($rule in $sortedIlmRules) {
        $ruleDisplayName = [string](& $boolText (Get-PropertyValue -Object $rule -PropertyName "displayName"))

        $markdownLines.Add("")
        $markdownLines.Add("### ILM Rule Details - $ruleDisplayName")
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
            @("Rule Name", $ruleDisplayName),
            @("Description", (& $boolText (Get-PropertyValue -Object $rule -PropertyName "description"))),
            @("Version", (& $boolText (Get-PropertyValue -Object $rule -PropertyName "version"))),
            @("Reference Time", (& $boolText (Get-PropertyValue -Object $rule -PropertyName "referenceTime"))),
            @("Ingest Behavior", (& $boolText (Get-PropertyValue -Object $rule -PropertyName "ingestBehavior"))),
            @("API", (& $boolText (Get-PropertyValue -Object $rule -PropertyName "api"))),
            @("Compliance Compatible", (& $boolText (Get-PropertyValue -Object $rule -PropertyName "complianceCompatible")))
        )

        $markdownLines.Add("")
        # --- ILM Rule Filters ---
        $markdownLines.Add("#### Filters")
        $markdownLines.Add("")
        $filterMergedRows = @()
        $filters = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $rule -PropertyName "filters")
        for ($filterIndex = 0; $filterIndex -lt $filters.Count; $filterIndex++) {
            $filterLabel = "Filter " + ($filterIndex + 1)
            $criteria = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $filters[$filterIndex] -PropertyName "criteria")
            if ($criteria.Count -eq 0) {
                $filterMergedRows += ,@($filterLabel, "N/A", "N/A", "N/A", "N/A")
                continue
            }
            for ($criteriaIndex = 0; $criteriaIndex -lt $criteria.Count; $criteriaIndex++) {
                $criterion = $criteria[$criteriaIndex]
                $firstCell = if ($criteriaIndex -eq 0) { $filterLabel } else { "" }
                $filterMergedRows += ,@(
                    $firstCell,
                    (& $boolText (Get-PropertyValue -Object $criterion -PropertyName "metadataName")),
                    (& $boolText (Get-PropertyValue -Object $criterion -PropertyName "operator")),
                    (& $boolText (Get-PropertyValue -Object $criterion -PropertyName "metadataType")),
                    (& $boolText (Get-PropertyValue -Object $criterion -PropertyName "value"))
                )
            }
        }
        if ($filterMergedRows.Count -eq 0) {
            $markdownLines.Add("This rule has no filters configured; it applies to all matching objects.")
            $markdownLines.Add("")
        } else {
            & $emitOpenXmlMergedFirstColumnTable -Headers @("Filter", "Metadata Name", "Operator", "Metadata Type", "Value") -Rows $filterMergedRows -GridWidths @(1440, 2160, 1800, 1800, 2160)
        }

        $markdownLines.Add("")
        # --- ILM Rule Placements ---
        $markdownLines.Add("#### Placements")
        $markdownLines.Add("")
        $placementMergedRows = @()
        $placements = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $rule -PropertyName "placements")
        for ($placementIndex = 0; $placementIndex -lt $placements.Count; $placementIndex++) {
            $placement = $placements[$placementIndex]
            $placementLabel = "Placement " + ($placementIndex + 1)
            $retentionAfter = [string](Get-PropertyValue -Object (Get-PropertyValue -Object $placement -PropertyName "retention") -PropertyName "after")

            $poolEntries = @()
            foreach ($ecEntry in (Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $placement -PropertyName "erasureCoded"))) {
                $ecPoolId = [string](Get-PropertyValue -Object $ecEntry -PropertyName "poolId")
                $ecPoolName = "N/A"
                if ($ilmPoolById.ContainsKey($ecPoolId)) {
                    $ecPoolName = [string](Get-PropertyValue -Object $ilmPoolById[$ecPoolId] -PropertyName "displayName")
                }
                $ecProfileId = [string](Get-PropertyValue -Object $ecEntry -PropertyName "profileId")
                $ecProfileName = $ecProfileId
                if ($ecProfileById.ContainsKey($ecProfileId)) {
                    $ecProfileName = [string](Get-PropertyValue -Object $ecProfileById[$ecProfileId] -PropertyName "name")
                }
                $poolEntries += ,@("Erasure Coded", $ecPoolName, "EC Profile: $ecProfileName")
            }
            foreach ($repEntry in (Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $placement -PropertyName "replicated"))) {
                $repPoolId = [string](Get-PropertyValue -Object $repEntry -PropertyName "poolId")
                $repPoolName = "N/A"
                if ($ilmPoolById.ContainsKey($repPoolId)) {
                    $repPoolName = [string](Get-PropertyValue -Object $ilmPoolById[$repPoolId] -PropertyName "displayName")
                }
                $repCopies = [string](Get-PropertyValue -Object $repEntry -PropertyName "copies")
                $poolEntries += ,@("Replicated", $repPoolName, "$repCopies Copies")
            }

            if ($poolEntries.Count -eq 0) {
                $placementMergedRows += ,@($placementLabel, (& $boolText $retentionAfter), "N/A", "N/A", "N/A")
                continue
            }
            for ($poolIndex = 0; $poolIndex -lt $poolEntries.Count; $poolIndex++) {
                $firstCell = if ($poolIndex -eq 0) { $placementLabel } else { "" }
                $retentionCell = (& $boolText $retentionAfter)
                $placementMergedRows += ,@($firstCell, $retentionCell, $poolEntries[$poolIndex][0], $poolEntries[$poolIndex][1], $poolEntries[$poolIndex][2])
            }
        }
        if ($placementMergedRows.Count -eq 0) {
            $markdownLines.Add("This rule has no placements configured.")
            $markdownLines.Add("")
        } else {
            & $emitOpenXmlMergedFirstColumnTable -Headers @("Placement", "Retention After", "Storage Type", "Pool Name", "Profile / Copies") -Rows $placementMergedRows -GridWidths @(1620, 1620, 1620, 1620, 2520)
        }
    }

    # --- ILM Policy ---
    $markdownLines.Add("")
    $markdownLines.Add("## ILM Policy")
    $markdownLines.Add("")
    $markdownLines.Add("An ILM policy is an ordered set of ILM rules that determines how StorageGRID manages and protects all object data in the grid.")
    $markdownLines.Add("")
    $markdownLines.Add("Every object is evaluated against the rules in the active policy from top to bottom, and any object that does not match any rule's filter is protected by the policy's default rule, which must not have a filter.")
    $markdownLines.Add("")
    $markdownLines.Add("This table shows each defined policy, whether it is the currently active policy, its rule count, and its default rule.")
    $markdownLines.Add("")
    $sortedIlmPolicies = @(
        $ilmPolicies | Sort-Object -Property @(
            @{ Expression = { -not [bool](Get-PropertyValue -Object $_ -PropertyName "active") } },
            @{ Expression = { [string](Get-PropertyValue -Object $_ -PropertyName "name") } }
        )
    )
    $policyRows = @()
    foreach ($policy in $sortedIlmPolicies) {
        $defaultRuleId = [string](Get-PropertyValue -Object $policy -PropertyName "defaultRule")
        $defaultRuleName = $defaultRuleId
        if ($ruleById.ContainsKey($defaultRuleId)) {
            $defaultRuleName = [string](Get-PropertyValue -Object $ruleById[$defaultRuleId] -PropertyName "displayName")
        }
        $policyRows += ,@(
            (& $boolText (Get-PropertyValue -Object $policy -PropertyName "name")),
            (& $boolText (Get-PropertyValue -Object $policy -PropertyName "active")),
            [string]@((Get-PropertyValue -Object $policy -PropertyName "rules")).Count,
            (& $boolText $defaultRuleName)
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Policy Name", "Active", "Rule Count", "Default Rule") -Rows $policyRows

    $userDisplayNameById = @{}
    foreach ($gridUser in $gridUsers) {
        $gridUserId = [string](Get-PropertyValue -Object $gridUser -PropertyName "id")
        if ([string]::IsNullOrWhiteSpace($gridUserId)) { continue }
        $gridUserDisplayName = [string](Get-PropertyValue -Object $gridUser -PropertyName "fullName")
        if ([string]::IsNullOrWhiteSpace($gridUserDisplayName)) {
            $gridUserDisplayName = [string](Get-PropertyValue -Object $gridUser -PropertyName "uniqueName")
        }
        if (-not [string]::IsNullOrWhiteSpace($gridUserDisplayName)) {
            $userDisplayNameById[$gridUserId] = $gridUserDisplayName
        }
    }

    $markdownLines.Add("")
    # --- ILM Policy Details ---
    $markdownLines.Add("## ILM Policy Details")
    $markdownLines.Add("")
    $markdownLines.Add("The following sections provide the full rule membership and details for each ILM policy, including which rules are included (in evaluation order) and which is the Default Rule.")
    $markdownLines.Add("")
    $markdownLines.Add("Reviewing the rule order is important because StorageGRID always evaluates an active policy's rules from top to bottom and applies the first rule whose filter matches an object.")
    $markdownLines.Add("")

    foreach ($policy in $sortedIlmPolicies) {
        $policyDisplayName = [string](& $boolText (Get-PropertyValue -Object $policy -PropertyName "name"))
        $defaultRuleId = [string](Get-PropertyValue -Object $policy -PropertyName "defaultRule")
        $defaultRuleName = $defaultRuleId
        if ($ruleById.ContainsKey($defaultRuleId)) {
            $defaultRuleName = [string](Get-PropertyValue -Object $ruleById[$defaultRuleId] -PropertyName "displayName")
        }
        $activatedByList = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $policy -PropertyName "activatedBy")
        $activatedByNames = @()
        foreach ($activatedById in $activatedByList) {
            $activatedByIdText = [string]$activatedById
            if ([string]::IsNullOrWhiteSpace($activatedByIdText)) { continue }
            if ($userDisplayNameById.ContainsKey($activatedByIdText)) {
                $activatedByNames += $userDisplayNameById[$activatedByIdText]
            } else {
                $activatedByNames += "N/A - Unresolved User ID"
            }
        }
        $activatedByDisplay = if ($activatedByNames.Count -gt 0) { $activatedByNames -join "`n" } else { "N/A" }

        $markdownLines.Add("")
        $markdownLines.Add("### ILM Policy Details - $policyDisplayName")
        $markdownLines.Add("")
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Field", "Value") -Rows @(
            @("Policy Name", $policyDisplayName),
            @("Description", (& $boolText (Get-PropertyValue -Object $policy -PropertyName "reason"))),
            @("Active", (& $boolText (Get-PropertyValue -Object $policy -PropertyName "active"))),
            @("Default Rule", (& $boolText $defaultRuleName)),
            @("Activated By", $activatedByDisplay)
        )

        $markdownLines.Add("")
        # --- ILM Policy Rules ---
        $markdownLines.Add("#### Rules")
        $markdownLines.Add("")
        $policyRuleRows = @()
        foreach ($policyRuleId in (Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $policy -PropertyName "rules"))) {
            $policyRuleIdText = [string]$policyRuleId
            $policyRuleName = $policyRuleIdText
            if ($ruleById.ContainsKey($policyRuleIdText)) {
                $policyRuleName = [string](Get-PropertyValue -Object $ruleById[$policyRuleIdText] -PropertyName "displayName")
            }
            $isDefault = $policyRuleIdText -eq $defaultRuleId
            $policyRuleRows += ,@($policyRuleName, [string]$isDefault)
        }
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Rule Name", "Is Default Rule") -Rows $policyRuleRows
    }

    # --- ILM Policy Tags ---
    $markdownLines.Add("")
    $markdownLines.Add("## ILM Policy Tags")
    $markdownLines.Add("")
    $markdownLines.Add("If you want to allow tenants to easily switch between multiple data protection policies on a per-bucket basis, you can use multiple ILM policies with ILM policy tags. You assign each ILM policy to a tag, then tenants tag a bucket to apply the policy to that bucket. You can set ILM policy tags on S3 buckets only.")
    $markdownLines.Add("")
    $markdownLines.Add("For example, you might have three tags named Gold, Silver, and Bronze. You can assign an ILM policy to each tag, based on how long and where that policy stores objects. Tenants can choose which policy to use by tagging their buckets. A bucket tagged Gold is managed by the Gold policy and receives the Gold level of data protection and performance.")
    $markdownLines.Add("")
    $markdownLines.Add("A default ILM policy tag is automatically created when you install StorageGRID. Every grid must have one active policy that is assigned to the Default tag. The default policy applies to any untagged S3 buckets.")
    $markdownLines.Add("")
    $ilmPolicyById = @{}
    foreach ($policy in $ilmPolicies) {
        $policyIdKey = [string](Get-PropertyValue -Object $policy -PropertyName "id")
        if (-not [string]::IsNullOrWhiteSpace($policyIdKey)) { $ilmPolicyById[$policyIdKey] = $policy }
    }
    $ilmPolicyTagRows = @()
    foreach ($policyTag in $ilmPolicyTags) {
        $policyTagPolicyId = [string](Get-PropertyValue -Object $policyTag -PropertyName "policyId")
        $policyTagPolicyName = "N/A - Unresolved Policy ID"
        if ($ilmPolicyById.ContainsKey($policyTagPolicyId)) {
            $policyTagPolicyName = [string](Get-PropertyValue -Object $ilmPolicyById[$policyTagPolicyId] -PropertyName "name")
        }
        $ilmPolicyTagRows += ,@(
            (& $boolText (Get-PropertyValue -Object $policyTag -PropertyName "name")),
            (& $boolText (Get-PropertyValue -Object $policyTag -PropertyName "description")),
            $policyTagPolicyName
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Tag Name", "Description", "Policy Name") -Rows $ilmPolicyTagRows

    # --- Tenants ---
    $markdownLines.Add("")
    $markdownLines.Add("# Tenants")
    $markdownLines.Add("")
    # --- Tenant Accounts ---
    $markdownLines.Add("## Tenant Accounts")
    $markdownLines.Add("")
    $markdownLines.Add("A tenant account is an isolated S3 client account within the grid, with its own set of users, buckets/containers, and permissions, managed through the Tenant Manager.")
    $markdownLines.Add("")
    $markdownLines.Add("Tenants allow a single StorageGRID grid to be securely shared by multiple, independent, client organizations or applications, each unaware of the others' data.")
    $markdownLines.Add("")
    $markdownLines.Add("This table lists the tenant accounts defined on the grid; see the following section for each tenant's configuration.")
    $markdownLines.Add("")
    $accountRows = @()
    foreach ($account in $gridAccounts) {
        $accountRows += ,@(
            (& $boolText (Get-PropertyValue -Object $account -PropertyName "name")),
            (& $boolText (Get-PropertyValue -Object $account -PropertyName "description"))
        )
    }
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Account Name", "Description") -Rows $accountRows

    $gridFederationNameById = @{}
    foreach ($connection in $gridFederationConnections) {
        $connectionId = [string](Get-PropertyValue -Object $connection -PropertyName "id")
        if ([string]::IsNullOrWhiteSpace($connectionId)) { continue }
        $connectionName = [string](Get-PropertyValue -Object $connection -PropertyName "name")
        if ([string]::IsNullOrWhiteSpace($connectionName)) {
            $connectionName = [string](Get-PropertyValue -Object $connection -PropertyName "displayName")
        }
        if (-not [string]::IsNullOrWhiteSpace($connectionName)) {
            $gridFederationNameById[$connectionId] = $connectionName
        }
    }

    $markdownLines.Add("")
    # --- Tenant Details ---
    $markdownLines.Add("## Tenant Details")
    $markdownLines.Add("")
    $markdownLines.Add("The following sections provide the grid-level configuration for each tenant account, including its capabilities (such as S3 access and self-service management), whether it uses its own identity source for tenant users, platform services (such as CloudMirror replication or event notifications) are permitted, logical storage quota, and any S3 Object Lock compliance-mode retention limits configured for the account.")

    foreach ($account in $gridAccounts) {
        $accountDisplayName = [string](& $boolText (Get-PropertyValue -Object $account -PropertyName "name"))
        $capabilities = @((Get-PropertyValue -Object $account -PropertyName "capabilities"))
        $capabilitiesDisplay = if ($capabilities.Count -gt 0) { $capabilities -join "`n" } else { "N/A" }
        $policy = Get-PropertyValue -Object $account -PropertyName "policy"

        $federationIds = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $policy -PropertyName "allowedGridFederationConnections")
        $federationNames = @()
        foreach ($federationId in $federationIds) {
            $federationIdText = [string]$federationId
            if ([string]::IsNullOrWhiteSpace($federationIdText)) { continue }
            if ($gridFederationNameById.ContainsKey($federationIdText)) {
                $federationNames += $gridFederationNameById[$federationIdText]
            } else {
                $federationNames += "N/A - Unresolved Connection ID"
            }
        }
        $federationDisplay = if ($federationNames.Count -gt 0) { $federationNames -join "`n" } else { "N/A" }

        $quotaBytes = Get-PropertyValue -Object $policy -PropertyName "quotaObjectBytes"
        $quotaDisplay = if ($null -eq $quotaBytes) { "N/A" } else { Convert-BytesToTiB -Bytes $quotaBytes }

        $markdownLines.Add("")
        $markdownLines.Add("### Tenant Details - $accountDisplayName")
        $markdownLines.Add("")
        & $emitOpenXmlMergedFirstColumnTable -Headers @("Field", "Value") -Rows @(
            @("Account Name", $accountDisplayName),
            @("Description", (& $boolText (Get-PropertyValue -Object $account -PropertyName "description"))),
            @("Capabilities", $capabilitiesDisplay),
            @("Use own identity source", (& $boolText (Get-PropertyValue -Object $policy -PropertyName "useAccountIdentitySource"))),
            @("Allow Platform Services", (& $boolText (Get-PropertyValue -Object $policy -PropertyName "allowPlatformServices"))),
            @("Allow S3 Select", (& $boolText (Get-PropertyValue -Object $policy -PropertyName "allowSelectObjectContent"))),
            @("Grid Federation Connections", $federationDisplay),
            @("Storage Quota (TiB)", $quotaDisplay),
            @("Allow S3 Object Lock Compliance Mode", (& $boolText (Get-PropertyValue -Object $policy -PropertyName "allowComplianceMode"))),
            @("Max Retention Days", (& $boolText (Get-PropertyValue -Object $policy -PropertyName "maxRetentionDays"))),
            @("Max Retention Years", (& $boolText (Get-PropertyValue -Object $policy -PropertyName "maxRetentionYears")))
        ) -GridWidths @(3600, 4200)
    }

    # --- Tenant Usage ---
    $markdownLines.Add("")
    $markdownLines.Add("## Tenant Usage")
    $markdownLines.Add("")
    $markdownLines.Add("StorageGRID periodically calculates each tenant account's storage usage in the background; this table shows the object count, data capacity consumed, and bucket count reported for each tenant as of the most recent calculation. These usage metrics are also what StorageGRID uses to enforce a tenant's storage quota, if one is configured.")
    $markdownLines.Add("")
    $accountUsageRows = @()
    foreach ($usage in $gridAccountUsage) {
        $usageAccountId = [string](Get-PropertyValue -Object $usage -PropertyName "__accountId")
        $usageAccountName = $usageAccountId
        if ($accountById.ContainsKey($usageAccountId)) {
            $usageAccountName = [string](Get-PropertyValue -Object $accountById[$usageAccountId] -PropertyName "name")
        }

        $dataBytes = Get-PropertyValue -Object $usage -PropertyName "dataBytes"
        $dataBytesSort = 0.0
        $dataTiB = "N/A"
        if ($null -ne $dataBytes -and -not [string]::IsNullOrWhiteSpace([string]$dataBytes)) {
            $dataTiB = Convert-BytesToTiB -Bytes $dataBytes
            [void][double]::TryParse([string]$dataBytes, [ref]$dataBytesSort)
        }

        $objectCount = Get-PropertyValue -Object $usage -PropertyName "objectCount"
        $objectCountDisplay = "N/A"
        $objectCountNumeric = 0.0
        if ($null -ne $objectCount -and [double]::TryParse([string]$objectCount, [ref]$objectCountNumeric)) {
            $objectCountDisplay = $objectCountNumeric.ToString("N0", [System.Globalization.CultureInfo]::InvariantCulture)
        }

        $bucketItems = @((Get-PropertyValue -Object $usage -PropertyName "buckets"))
        $accountUsageRows += ,@(
            (& $boolText $usageAccountName),
            $objectCountDisplay,
            (& $boolText $dataTiB),
            [string]$bucketItems.Count,
            $dataBytesSort
        )
    }
    $accountUsageRows = @($accountUsageRows | Sort-Object -Property @(
            @{ Expression = { $_[4] }; Descending = $true },
            @{ Expression = { [string]$_[0] }; Descending = $false }
        ) | ForEach-Object { ,@($_[0], $_[1], $_[2], $_[3]) })
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Account Name", "Object Count", "Data Capacity (TiB)", "Bucket Count") -Rows $accountUsageRows

    # --- Tenant Usage - Bucket Details ---
    $markdownLines.Add("")
    $markdownLines.Add("## Tenant Usage - Bucket Details")
    $markdownLines.Add("")
    $markdownLines.Add("The following sections provide per-bucket usage and configuration details for each tenant account's S3 buckets, including object count, data capacity consumed, and configuration attributes such as versioning and metadata consistency controls.")

    foreach ($usage in $gridAccountUsage) {
        $usageAccountId = [string](Get-PropertyValue -Object $usage -PropertyName "__accountId")
        $usageAccountName = $usageAccountId
        if ($accountById.ContainsKey($usageAccountId)) {
            $usageAccountName = [string](Get-PropertyValue -Object $accountById[$usageAccountId] -PropertyName "name")
        }

        $markdownLines.Add("")
        $markdownLines.Add("### Bucket Details - $usageAccountName")
        $markdownLines.Add("")

        $bucketItems = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $usage -PropertyName "buckets")
        $bucketRows = @()
        foreach ($bucket in $bucketItems) {
            $bucketObjectCount = Get-PropertyValue -Object $bucket -PropertyName "objectCount"
            $bucketObjectCountDisplay = "N/A"
            $bucketObjectCountNumeric = 0.0
            if ($null -ne $bucketObjectCount -and [double]::TryParse([string]$bucketObjectCount, [ref]$bucketObjectCountNumeric)) {
                $bucketObjectCountDisplay = $bucketObjectCountNumeric.ToString("N0", [System.Globalization.CultureInfo]::InvariantCulture)
            }

            $bucketDataBytes = Get-PropertyValue -Object $bucket -PropertyName "dataBytes"
            $bucketDataTiB = "N/A"
            if ($null -ne $bucketDataBytes -and -not [string]::IsNullOrWhiteSpace([string]$bucketDataBytes)) {
                $bucketDataTiB = Convert-BytesToTiB -Bytes $bucketDataBytes
            }

            $bucketRows += ,@(
                (& $boolText (Get-PropertyValue -Object $bucket -PropertyName "name")),
                $bucketObjectCountDisplay,
                $bucketDataTiB,
                (& $boolText (Get-PropertyValue -Object $bucket -PropertyName "region")),
                (& $boolText (Get-PropertyValue -Object $bucket -PropertyName "consistency")),
                (& $boolText (Get-PropertyValue -Object $bucket -PropertyName "versioningEnabled"))
            )
        }
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Bucket Name", "Object Count", "Data Capacity (TiB)", "Region", "Consistency", "Versioning Enabled") -Rows $bucketRows
    }

    # --- References ---
    $markdownLines.Add("")
    $markdownLines.Add('```{=openxml}')
    $markdownLines.Add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    $markdownLines.Add('```')
    $markdownLines.Add("")
    $markdownLines.Add("# References")
    $markdownLines.Add("")
    $markdownLines.Add("## Online Support Resources")
    $markdownLines.Add("")
    Add-MarkdownTable -LineBuffer $markdownLines -Headers @("Name", "URL") -Rows @(
        @("Support Portal", "[https://mysupport.netapp.com](https://mysupport.netapp.com)"),
        @("Knowledge Base", "[https://kb.netapp.com](https://kb.netapp.com)"),
        @("StorageGRID Software Documentation", "[https://docs.netapp.com/us-en/storagegrid/index.html](https://docs.netapp.com/us-en/storagegrid/index.html)"),
        @("StorageGRID Appliance Documentation", "[https://docs.netapp.com/us-en/storagegrid-appliance/index.html](https://docs.netapp.com/us-en/storagegrid-appliance/index.html)"),
        @("StorageGRID Solutions and Resources", "[https://docs.netapp.com/us-en/storagegrid-enable/index.html](https://docs.netapp.com/us-en/storagegrid-enable/index.html)"),
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

    # --- API Collection Failures ---
    if (@($RestQueryFailures).Count -gt 0) {
        $markdownLines.Add("")
        $markdownLines.Add('```{=openxml}')
        $markdownLines.Add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        $markdownLines.Add('```')
        $markdownLines.Add("")
        $markdownLines.Add("# API Collection Failures")
        $markdownLines.Add("")
        $markdownLines.Add("The following API requests did not complete successfully during collection. A successful response with an empty payload is not listed here.")
        $markdownLines.Add("")
        $apiFailureRows = @()
        foreach ($failure in @($RestQueryFailures | Sort-Object -Property endpoint)) {
            $apiFailureRows += ,@(
                (Get-PropertyValue -Object $failure -PropertyName "endpoint"),
                (Get-PropertyValue -Object $failure -PropertyName "status"),
                (Get-PropertyValue -Object $failure -PropertyName "msg")
            )
        }
        Add-MarkdownTable -LineBuffer $markdownLines -Headers @("API Endpoint", "Response Code", "Detail") -Rows $apiFailureRows
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

function Invoke-StorageGridApi {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$Endpoint,
        [Parameter(Mandatory = $true)][string]$BearerToken,
        [Parameter(Mandatory = $false)][string]$AcceptHeader = "application/json"
    )

    $url = "$BaseUrl/api/v4$Endpoint"
    $invokeParams = @{
        Uri         = $url
        Method      = "GET"
        Headers     = @{
            Accept        = $AcceptHeader
            Authorization = "Bearer $BearerToken"
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

function Invoke-StorageGridMetricQuery {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$BearerToken,
        [Parameter(Mandatory = $true)][string]$Query
    )

    $encodedQuery = [System.Uri]::EscapeDataString($Query)
    return Invoke-StorageGridApi -BaseUrl $BaseUrl -Endpoint "/grid/metric-query?query=$encodedQuery" -BearerToken $BearerToken
}

function Convert-StorageGridMetricSampleValue {
    param([Parameter(Mandatory = $false)]$Sample)

    $sampleParts = @($Sample)
    if ($sampleParts.Count -lt 2) {
        return $null
    }

    $parsedValue = 0.0
    $rawValue = [string]$sampleParts[1]
    if ([double]::TryParse($rawValue, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedValue)) {
        return [double]$parsedValue
    }

    if ([double]::TryParse($rawValue, [ref]$parsedValue)) {
        return [double]$parsedValue
    }

    return $null
}

function Convert-ToNullableDouble {
    param([Parameter(Mandatory = $false)]$Value)

    $candidate = Get-FirstItem -Value $Value
    if ($null -eq $candidate) {
        return $null
    }

    $parsedValue = 0.0
    $rawValue = [string]$candidate
    if ([double]::TryParse($rawValue, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedValue)) {
        return [double]$parsedValue
    }

    if ([double]::TryParse($rawValue, [ref]$parsedValue)) {
        return [double]$parsedValue
    }

    return $null
}

function Convert-ToFlatObjectArray {
    param([Parameter(Mandatory = $false)]$Value)

    if ($null -eq $Value) {
        return ,@()
    }

    $items = @($Value)
    if ($items.Count -eq 1 -and $items[0] -is [array]) {
        return ,@($items[0])
    }

    return ,$items
}

function Get-StorageGridMetricQueryRows {
    param(
        [Parameter(Mandatory = $true)]$QueryResponse,
        [Parameter(Mandatory = $false)][string]$LabelName
    )

    $rows = @()
    if ($null -eq $QueryResponse -or $QueryResponse.Failed) {
        return ,$rows
    }

    $seriesList = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $QueryResponse.Data -PropertyName "result")
    foreach ($series in @($seriesList)) {
        $value = Convert-StorageGridMetricSampleValue -Sample (Get-PropertyValue -Object $series -PropertyName "value")
        if ($null -eq $value) {
            continue
        }

        $metricData = Get-PropertyValue -Object $series -PropertyName "metric"
        $labelValue = if ([string]::IsNullOrWhiteSpace($LabelName)) {
            "grid"
        }
        else {
            [string](Get-PropertyValue -Object $metricData -PropertyName $LabelName)
        }

        if ([string]::IsNullOrWhiteSpace($labelValue)) {
            continue
        }

        $rows += [pscustomobject]@{
            Label = [string]$labelValue
            Value = [double]$value
        }
    }

    return ,$rows
}

function Get-StorageGridMetricAggregateValue {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$BearerToken,
        [Parameter(Mandatory = $true)][string]$MetricName
    )

    $response = Invoke-StorageGridMetricQuery -BaseUrl $BaseUrl -BearerToken $BearerToken -Query ("sum({0})" -f $MetricName)
    $rows = Get-StorageGridMetricQueryRows -QueryResponse $response
    if (@($rows).Count -eq 0) {
        return $null
    }

    return [double]$rows[0].Value
}

function Find-StorageGridMetricGroupingLabel {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$BearerToken,
        [Parameter(Mandatory = $true)][string[]]$MetricNames,
        [Parameter(Mandatory = $true)][string[]]$LabelCandidates
    )

    foreach ($labelCandidate in $LabelCandidates) {
        foreach ($metricName in $MetricNames) {
            $response = Invoke-StorageGridMetricQuery -BaseUrl $BaseUrl -BearerToken $BearerToken -Query ("sum by ({0})({1})" -f $labelCandidate, $metricName)
            $rows = Get-StorageGridMetricQueryRows -QueryResponse $response -LabelName $labelCandidate
            if (@($rows).Count -gt 0) {
                return [string]$labelCandidate
            }
        }
    }

    return $null
}

function Get-StorageGridMetricGroupedValues {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$BearerToken,
        [Parameter(Mandatory = $true)][string]$MetricName,
        [Parameter(Mandatory = $true)][string]$GroupLabel
    )

    $response = Invoke-StorageGridMetricQuery -BaseUrl $BaseUrl -BearerToken $BearerToken -Query ("sum by ({0})({1})" -f $GroupLabel, $MetricName)
    return Get-StorageGridMetricQueryRows -QueryResponse $response -LabelName $GroupLabel
}

function Get-StorageGridMetricSeriesRows {
    param(
        [Parameter(Mandatory = $true)]$QueryResponse
    )

    $rows = @()
    if ($null -eq $QueryResponse -or $QueryResponse.Failed) {
        return ,$rows
    }

    $seriesList = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $QueryResponse.Data -PropertyName "result")
    foreach ($series in @($seriesList)) {
        $value = Convert-StorageGridMetricSampleValue -Sample (Get-PropertyValue -Object $series -PropertyName "value")
        if ($null -eq $value) {
            continue
        }

        $metricData = Get-PropertyValue -Object $series -PropertyName "metric"
        $labels = @{}
        if ($metricData -is [System.Collections.IDictionary]) {
            foreach ($metricKey in $metricData.Keys) {
                $labels[[string]$metricKey] = [string]$metricData[$metricKey]
            }
        }
        else {
            foreach ($prop in @($metricData.PSObject.Properties)) {
                $labels[[string]$prop.Name] = [string]$prop.Value
            }
        }

        $rows += [pscustomobject]@{
            Labels = $labels
            Value  = [double]$value
        }
    }

    return ,$rows
}

function Get-StorageGridMetricLabelText {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Labels,
        [Parameter(Mandatory = $true)][string[]]$Candidates
    )

    foreach ($candidate in $Candidates) {
        if ($Labels.ContainsKey($candidate)) {
            $text = [string]$Labels[$candidate]
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                return $text
            }
        }
    }

    return ""
}

function Get-StorageGridApplianceStorageMetrics {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$BearerToken,
        [Parameter(Mandatory = $false)][object[]]$NodeHealth = @()
    )

    $metricNames = [ordered]@{
        firmware = "storagegrid_appliance_storage_controller_firmware_version"
        nvsram   = "storagegrid_appliance_storage_controller_nvsram_version"
        sensor   = "storagegrid_appliance_storage_controller_hardware_sensor"
    }

    $rowsByNode = @{}
    foreach ($metricKey in $metricNames.Keys) {
        $metricName = [string]$metricNames[$metricKey]
        $response = Invoke-StorageGridMetricQuery -BaseUrl $BaseUrl -BearerToken $BearerToken -Query $metricName
        $seriesRows = Get-StorageGridMetricSeriesRows -QueryResponse $response
        foreach ($seriesRow in @($seriesRows)) {
            $labels = [hashtable](Get-PropertyValue -Object $seriesRow -PropertyName "Labels")
            if ($null -eq $labels) { continue }

            $instanceLabel = Get-StorageGridMetricLabelText -Labels $labels -Candidates @("instance", "node", "node_id")
            if ([string]::IsNullOrWhiteSpace($instanceLabel)) { continue }

            $nodeInfo = Resolve-StorageGridNodeMetricLabel -LabelValue $instanceLabel -NodeHealth $NodeHealth
            $rowKey = [string]$nodeInfo.NodeName
            if ([string]::IsNullOrWhiteSpace($rowKey)) {
                $rowKey = [string]$instanceLabel
            }

            if (-not $rowsByNode.ContainsKey($rowKey)) {
                $rowsByNode[$rowKey] = [ordered]@{
                    nodeName            = [string]$nodeInfo.NodeName
                    siteName            = [string]$nodeInfo.SiteName
                    metricLabel         = [string]$nodeInfo.LabelValue
                    santricityVersion   = ""
                    santricityFirmware  = ""
                    nvsramVersion       = ""
                    storageSerialNumber = ""
                    sgaModel            = ""
                }
            }

            $target = $rowsByNode[$rowKey]
            if ([string]::IsNullOrWhiteSpace([string]$target.nodeName)) {
                $target.nodeName = [string]$nodeInfo.NodeName
            }
            if ([string]::IsNullOrWhiteSpace([string]$target.siteName)) {
                $target.siteName = [string]$nodeInfo.SiteName
            }

            if ($metricKey -eq "firmware") {
                $firmwareText = Get-StorageGridMetricLabelText -Labels $labels -Candidates @("CFW", "firmware_version", "controller_firmware_version", "firmware", "version")
                if (-not [string]::IsNullOrWhiteSpace($firmwareText)) {
                    $target.santricityFirmware = $firmwareText
                }

                $santricityText = Get-StorageGridMetricLabelText -Labels $labels -Candidates @("BundleDisplayVersion", "santricity_version", "santricity", "os_version", "version")
                if (-not [string]::IsNullOrWhiteSpace($santricityText)) {
                    $target.santricityVersion = $santricityText
                }
            }
            elseif ($metricKey -eq "nvsram") {
                $nvsramText = Get-StorageGridMetricLabelText -Labels $labels -Candidates @("NVSRAM", "nvsram_version", "version")
                if (-not [string]::IsNullOrWhiteSpace($nvsramText)) {
                    $target.nvsramVersion = $nvsramText
                }
            }
            elseif ($metricKey -eq "sensor") {
                $serialText = Get-StorageGridMetricLabelText -Labels $labels -Candidates @("chassis_serial_number", "serial_number", "serial", "controller_serial_number", "model_serial")
                if (-not [string]::IsNullOrWhiteSpace($serialText)) {
                    $target.storageSerialNumber = $serialText
                }

                $modelText = Get-StorageGridMetricLabelText -Labels $labels -Candidates @("appliance_model", "model", "product", "controller_model")
                if (-not [string]::IsNullOrWhiteSpace($modelText)) {
                    $target.sgaModel = $modelText
                }
            }
        }
    }

    $rows = @()
    foreach ($nodeKey in @($rowsByNode.Keys | Sort-Object)) {
        $entry = $rowsByNode[$nodeKey]
        $rows += [pscustomobject][ordered]@{
            nodeName            = if ([string]::IsNullOrWhiteSpace([string]$entry.nodeName)) { [string]$nodeKey } else { [string]$entry.nodeName }
            siteName            = if ([string]::IsNullOrWhiteSpace([string]$entry.siteName)) { "N/A" } else { [string]$entry.siteName }
            metricLabel         = [string]$entry.metricLabel
            santricityVersion   = [string]$entry.santricityVersion
            santricityFirmware  = [string]$entry.santricityFirmware
            nvsramVersion       = [string]$entry.nvsramVersion
            storageSerialNumber = [string]$entry.storageSerialNumber
            sgaModel            = [string]$entry.sgaModel
        }
    }

    return [pscustomobject][ordered]@{
        available   = [bool]($rows.Count -gt 0)
        metricNames = [pscustomobject]$metricNames
        nodes       = [object[]]@($rows)
    }
}

function Get-StorageGridNodeAttributeDetails {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$BearerToken,
        [Parameter(Mandatory = $false)]$ServiceIdsData,
        [Parameter(Mandatory = $false)]$NetworkTopologyData,
        [Parameter(Mandatory = $false)][object[]]$NodeHealth = @()
    )

    $attributeCodes = @("CNSN", "CBMC", "SAMD", "DDSZ", "DDTP", "RDMD", "SFSN", "SFID", "SANM", "SAIP", "SAIQ")
    $attributeQuery = [System.Uri]::EscapeDataString(($attributeCodes -join ","))

    # Restrict attribute lookups to appliance nodes (Platform == "SGA") when topology data is available.
    $applianceNodeIds = New-Object System.Collections.Generic.HashSet[string]
    if ($null -ne $NetworkTopologyData) {
        $gridNodesRaw = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $NetworkTopologyData -PropertyName "gridNodes")
        foreach ($gridNode in @($gridNodesRaw)) {
            $cfg = Get-PropertyValue -Object $gridNode -PropertyName "nodeConfig"
            if ($null -eq $cfg) { continue }
            if (([string](Get-PropertyValue -Object $cfg -PropertyName "platform")) -eq "SGA") {
                $cfgNodeId = [string](Get-PropertyValue -Object $cfg -PropertyName "nodeId")
                if (-not [string]::IsNullOrWhiteSpace($cfgNodeId)) {
                    [void]$applianceNodeIds.Add($cfgNodeId)
                }
            }
        }
        Write-Verbose "[NodeAttr] Found $($applianceNodeIds.Count) SGA appliance nodes in topology"
    }

    $ssmIdByNodeId = @{}
    if ($null -ne $ServiceIdsData) {
        Write-Verbose "[NodeAttr] Processing service IDs ..."
        $serviceEntries = @()
        if ($ServiceIdsData -is [System.Collections.IDictionary]) {
            foreach ($key in $ServiceIdsData.Keys) {
                $serviceEntries += [pscustomobject]@{ NodeId = [string]$key; NodeEntry = $ServiceIdsData[$key] }
            }
        }
        else {
            foreach ($prop in @($ServiceIdsData.PSObject.Properties)) {
                $serviceEntries += [pscustomobject]@{ NodeId = [string]$prop.Name; NodeEntry = $prop.Value }
            }
        }
        Write-Verbose "[NodeAttr] Found $($serviceEntries.Count) service entries total"

        foreach ($serviceEntry in $serviceEntries) {
            if ($applianceNodeIds.Count -gt 0 -and -not $applianceNodeIds.Contains($serviceEntry.NodeId)) {
                continue
            }

            $services = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $serviceEntry.NodeEntry -PropertyName "services")
            foreach ($service in @($services)) {
                if (([string](Get-PropertyValue -Object $service -PropertyName "name")) -ne "SSM") { continue }
                $ssmId = [string](Get-PropertyValue -Object $service -PropertyName "id")
                if (-not [string]::IsNullOrWhiteSpace($ssmId)) {
                    $ssmIdByNodeId[$serviceEntry.NodeId] = $ssmId
                    Write-Verbose "[NodeAttr] Found SSM service for node $($serviceEntry.NodeId): $ssmId"
                }
                break
            }
        }
        Write-Verbose "[NodeAttr] Total SSM nodes to query: $($ssmIdByNodeId.Count)"
    }

    $nodeRows = @()
    foreach ($nodeId in @($ssmIdByNodeId.Keys | Sort-Object)) {
        $ssmId = $ssmIdByNodeId[$nodeId]
        $nodeInfo = Resolve-StorageGridNodeMetricLabel -LabelValue $nodeId -NodeHealth $NodeHealth

        Write-Verbose "Querying /private/attributes/$attributeQuery/node/$ssmId ..."
        $response = Invoke-StorageGridApi -BaseUrl $BaseUrl -Endpoint "/private/attributes/$attributeQuery/node/$ssmId" -BearerToken $BearerToken

        $dataType = if ($null -eq $response.Data) { "<null>" } else { $response.Data.GetType().FullName }
        Write-Verbose "[NodeAttr] Response Failed: $($response.Failed), Data type: $dataType, Data is null: $($null -eq $response.Data)"

        $computeSerial = "N/A"
        $bmcIp = "N/A"
        $applianceModel = "N/A"
        $driveSize = "N/A"
        $driveType = "N/A"
        $raidMode = "N/A"
        $controllerName = "N/A"
        $controllerAIp = "N/A"
        $controllerBIp = "N/A"
        $shelfSerials = @()
        $shelfIds = @()

        if (-not $response.Failed) {
            $attributeItems = Convert-ToArrayPayload -Payload $response.Data
            Write-Verbose "[NodeAttr]   Attribute items count: $(@($attributeItems).Count)"
            foreach ($attributeItem in @($attributeItems)) {
                $code = [string](Get-PropertyValue -Object $attributeItem -PropertyName "code")
                $value = [string](Get-PropertyValue -Object $attributeItem -PropertyName "value")
                if ([string]::IsNullOrWhiteSpace($value)) { continue }
                switch ($code) {
                    "CNSN" { $computeSerial = $value }
                    "CBMC" { $bmcIp = $value }
                    "SAMD" { $applianceModel = $value }
                    "DDSZ" { $driveSize = $value }
                    "DDTP" { $driveType = $value }
                    "RDMD" { $raidMode = $value }
                    "SFSN" { $shelfSerials += $value }
                    "SFID" { $shelfIds += $value }
                    "SANM" { $controllerName = $value }
                    "SAIP" { $controllerAIp = $value }
                    "SAIQ" { $controllerBIp = $value }
                }
            }
            Write-Verbose "[NodeAttr]   Collected: model=$applianceModel, controllerAIp=$controllerAIp, controllerBIp=$controllerBIp"
        }
        else {
            Write-Verbose "[NodeAttr]   Response failed: $($response.Message)"
        }

        $nodeRows += [pscustomobject][ordered]@{
            nodeId         = [string]$nodeId
            nodeName       = [string]$nodeInfo.NodeName
            siteName       = [string]$nodeInfo.SiteName
            computeSerial  = $computeSerial
            bmcIp          = $bmcIp
            applianceModel = $applianceModel
            driveSize      = $driveSize
            driveType      = $driveType
            raidMode       = $raidMode
            controllerName = $controllerName
            controllerAIp  = $controllerAIp
            controllerBIp  = $controllerBIp
            shelfSerials   = [object[]]@($shelfSerials)
            shelfIds       = [object[]]@($shelfIds)
        }
    }
    Write-Verbose "[NodeAttr] Total rows collected: $($nodeRows.Count)"

    return [pscustomobject][ordered]@{
        available      = [bool]($nodeRows.Count -gt 0)
        attributeCodes = $attributeCodes
        nodes          = [object[]]@($nodeRows)
    }
}

function New-StorageGridCapacityRecord {
    param([Parameter(Mandatory = $true)][hashtable]$MetricValues)

    $overallTotalBytes = $MetricValues["overallTotalBytes"]
    $overallAvailableBytes = $MetricValues["overallAvailableBytes"]
    $objectUsedBytes = $MetricValues["objectUsedBytes"]
    $objectUsableBytes = $MetricValues["objectTotalBytes"]
    $metadataUsedBytes = $MetricValues["metadataUsedBytes"]
    $metadataAllowedBytes = $MetricValues["metadataAllowedBytes"]

    $overallUsedBytes = if ($null -ne $overallTotalBytes -and $null -ne $overallAvailableBytes) {
        [math]::Max(0.0, ([double]$overallTotalBytes - [double]$overallAvailableBytes))
    }
    else {
        $null
    }

    $overallUsedPct = if ($null -ne $overallUsedBytes -and $null -ne $overallTotalBytes -and [double]$overallTotalBytes -gt 0) {
        [math]::Round((100.0 * [double]$overallUsedBytes / [double]$overallTotalBytes), 2)
    }
    else {
        $null
    }

    $objectTotalBytes = if ($null -ne $objectUsableBytes -and $null -ne $objectUsedBytes) {
        [math]::Max(0.0, ([double]$objectUsableBytes + [double]$objectUsedBytes))
    }
    elseif ($null -ne $objectUsableBytes) {
        [double]$objectUsableBytes
    }
    else {
        $null
    }

    $objectAvailableBytes = if ($null -ne $objectUsableBytes) {
        [math]::Max(0.0, [double]$objectUsableBytes)
    }
    elseif ($null -ne $objectTotalBytes -and $null -ne $objectUsedBytes) {
        [math]::Max(0.0, ([double]$objectTotalBytes - [double]$objectUsedBytes))
    }
    else {
        $null
    }

    $objectUsedPct = if ($null -ne $objectUsedBytes -and $null -ne $objectTotalBytes -and [double]$objectTotalBytes -gt 0) {
        [math]::Round((100.0 * [double]$objectUsedBytes / [double]$objectTotalBytes), 2)
    }
    else {
        $null
    }

    $metadataAvailableBytes = if ($null -ne $metadataAllowedBytes -and $null -ne $metadataUsedBytes) {
        [math]::Max(0.0, ([double]$metadataAllowedBytes - [double]$metadataUsedBytes))
    }
    else {
        $null
    }

    $metadataUsedPct = if ($null -ne $metadataUsedBytes -and $null -ne $metadataAllowedBytes -and [double]$metadataAllowedBytes -gt 0) {
        [math]::Round((100.0 * [double]$metadataUsedBytes / [double]$metadataAllowedBytes), 2)
    }
    else {
        $null
    }

    return [pscustomobject][ordered]@{
        overallTotalBytes      = $overallTotalBytes
        overallUsedBytes       = $overallUsedBytes
        overallAvailableBytes  = $overallAvailableBytes
        overallUsedPct         = $overallUsedPct
        objectTotalBytes       = $objectTotalBytes
        objectUsedBytes        = $objectUsedBytes
        objectAvailableBytes   = $objectAvailableBytes
        objectUsedPct          = $objectUsedPct
        metadataAllowedBytes   = $metadataAllowedBytes
        metadataUsedBytes      = $metadataUsedBytes
        metadataAvailableBytes = $metadataAvailableBytes
        metadataUsedPct        = $metadataUsedPct
    }
}

function Test-StorageGridObjectCapacityRecordValidity {
    param([Parameter(Mandatory = $false)]$Record)

    if ($null -eq $Record) {
        return $false
    }

    $objectTotalBytes = Get-PropertyValue -Object $Record -PropertyName "objectTotalBytes"
    $objectUsedBytes = Get-PropertyValue -Object $Record -PropertyName "objectUsedBytes"
    $objectUsedPct = Get-PropertyValue -Object $Record -PropertyName "objectUsedPct"

    if ($null -eq $objectTotalBytes -or $null -eq $objectUsedBytes) {
        return $false
    }

    if ([double]$objectUsedBytes -gt [double]$objectTotalBytes) {
        return $false
    }

    if ($null -ne $objectUsedPct -and [double]$objectUsedPct -gt 100.0) {
        return $false
    }

    return $true
}

function Repair-StorageGridNodeObjectCapacityRecord {
    param([Parameter(Mandatory = $true)]$Record)

    if (Test-StorageGridObjectCapacityRecordValidity -Record $Record) {
        return $Record
    }

    $fallbackTotalBytes = Get-PropertyValue -Object $Record -PropertyName "overallTotalBytes"
    $fallbackUsedBytes = Get-PropertyValue -Object $Record -PropertyName "objectUsedBytes"
    if ($null -eq $fallbackTotalBytes -or $null -eq $fallbackUsedBytes -or [double]$fallbackTotalBytes -lt [double]$fallbackUsedBytes -or [double]$fallbackTotalBytes -le 0) {
        return $Record
    }

    $fallbackAvailableBytes = [math]::Max(0.0, ([double]$fallbackTotalBytes - [double]$fallbackUsedBytes))
    $fallbackUsedPct = [math]::Round((100.0 * [double]$fallbackUsedBytes / [double]$fallbackTotalBytes), 2)

    return [pscustomobject][ordered]@{
        name                   = Get-PropertyValue -Object $Record -PropertyName "name"
        siteName               = Get-PropertyValue -Object $Record -PropertyName "siteName"
        labelValue             = Get-PropertyValue -Object $Record -PropertyName "labelValue"
        overallTotalBytes      = Get-PropertyValue -Object $Record -PropertyName "overallTotalBytes"
        overallUsedBytes       = Get-PropertyValue -Object $Record -PropertyName "overallUsedBytes"
        overallAvailableBytes  = Get-PropertyValue -Object $Record -PropertyName "overallAvailableBytes"
        overallUsedPct         = Get-PropertyValue -Object $Record -PropertyName "overallUsedPct"
        objectTotalBytes       = $fallbackTotalBytes
        objectUsedBytes        = $fallbackUsedBytes
        objectAvailableBytes   = $fallbackAvailableBytes
        objectUsedPct          = $fallbackUsedPct
        metadataAllowedBytes   = Get-PropertyValue -Object $Record -PropertyName "metadataAllowedBytes"
        metadataUsedBytes      = Get-PropertyValue -Object $Record -PropertyName "metadataUsedBytes"
        metadataAvailableBytes = Get-PropertyValue -Object $Record -PropertyName "metadataAvailableBytes"
        metadataUsedPct        = Get-PropertyValue -Object $Record -PropertyName "metadataUsedPct"
        objectFallbackSource   = "overallTotalBytes"
    }
}

function Resolve-StorageGridNodeMetricLabel {
    param(
        [Parameter(Mandatory = $true)][string]$LabelValue,
        [Parameter(Mandatory = $false)][object[]]$NodeHealth = @()
    )

    $hostCandidate = [string]$LabelValue
    if ($hostCandidate.Contains(':')) {
        $hostCandidate = $hostCandidate.Split(':')[0]
    }

    foreach ($node in @($NodeHealth)) {
        $nodeName = [string](Get-PropertyValue -Object $node -PropertyName "name")
        $nodeId = [string](Get-PropertyValue -Object $node -PropertyName "id")
        if ($nodeName -eq $LabelValue -or $nodeName -eq $hostCandidate -or $nodeId -eq $LabelValue) {
            return [pscustomobject]@{
                NodeName = $nodeName
                SiteName = [string](Get-PropertyValue -Object $node -PropertyName "siteName")
                LabelValue = [string]$LabelValue
            }
        }
    }

    return [pscustomobject]@{
        NodeName = $hostCandidate
        SiteName = "N/A"
        LabelValue = [string]$LabelValue
    }
}

function Merge-StorageGridCapacityMetricRows {
    param(
        [Parameter(Mandatory = $true)][hashtable]$MetricDefinitions,
        [Parameter(Mandatory = $true)][string]$Scope,
        [Parameter(Mandatory = $true)][string]$GroupLabel,
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$BearerToken,
        [Parameter(Mandatory = $false)][object[]]$NodeHealth = @()
    )

    $rowMap = @{}
    foreach ($metricKey in $MetricDefinitions.Keys) {
        $metricName = [string]$MetricDefinitions[$metricKey]
        $groupRows = Get-StorageGridMetricGroupedValues -BaseUrl $BaseUrl -BearerToken $BearerToken -MetricName $metricName -GroupLabel $GroupLabel
        foreach ($groupRow in @($groupRows)) {
            $labelValue = [string]$groupRow.Label
            if ([string]::IsNullOrWhiteSpace($labelValue)) {
                continue
            }

            if (-not $rowMap.ContainsKey($labelValue)) {
                $rowMap[$labelValue] = @{}
            }

            $rowMap[$labelValue][$metricKey] = [double]$groupRow.Value
        }
    }

    $mergedRows = @()
    foreach ($labelValue in @($rowMap.Keys | Sort-Object)) {
        $derived = New-StorageGridCapacityRecord -MetricValues $rowMap[$labelValue]
        if ($Scope -eq "node") {
            $nodeInfo = Resolve-StorageGridNodeMetricLabel -LabelValue $labelValue -NodeHealth $NodeHealth
            $nodeRecord = [pscustomobject][ordered]@{
                name                   = [string]$nodeInfo.NodeName
                siteName               = [string]$nodeInfo.SiteName
                labelValue             = [string]$nodeInfo.LabelValue
                overallTotalBytes      = Get-PropertyValue -Object $derived -PropertyName "overallTotalBytes"
                overallUsedBytes       = Get-PropertyValue -Object $derived -PropertyName "overallUsedBytes"
                overallAvailableBytes  = Get-PropertyValue -Object $derived -PropertyName "overallAvailableBytes"
                overallUsedPct         = Get-PropertyValue -Object $derived -PropertyName "overallUsedPct"
                objectTotalBytes       = Get-PropertyValue -Object $derived -PropertyName "objectTotalBytes"
                objectUsedBytes        = Get-PropertyValue -Object $derived -PropertyName "objectUsedBytes"
                objectAvailableBytes   = Get-PropertyValue -Object $derived -PropertyName "objectAvailableBytes"
                objectUsedPct          = Get-PropertyValue -Object $derived -PropertyName "objectUsedPct"
                metadataAllowedBytes   = Get-PropertyValue -Object $derived -PropertyName "metadataAllowedBytes"
                metadataUsedBytes      = Get-PropertyValue -Object $derived -PropertyName "metadataUsedBytes"
                metadataAvailableBytes = Get-PropertyValue -Object $derived -PropertyName "metadataAvailableBytes"
                metadataUsedPct        = Get-PropertyValue -Object $derived -PropertyName "metadataUsedPct"
            }
            $mergedRows += Repair-StorageGridNodeObjectCapacityRecord -Record $nodeRecord
            continue
        }

        $mergedRows += [pscustomobject][ordered]@{
            name                   = [string]$labelValue
            labelValue             = [string]$labelValue
            overallTotalBytes      = Get-PropertyValue -Object $derived -PropertyName "overallTotalBytes"
            overallUsedBytes       = Get-PropertyValue -Object $derived -PropertyName "overallUsedBytes"
            overallAvailableBytes  = Get-PropertyValue -Object $derived -PropertyName "overallAvailableBytes"
            overallUsedPct         = Get-PropertyValue -Object $derived -PropertyName "overallUsedPct"
            objectTotalBytes       = Get-PropertyValue -Object $derived -PropertyName "objectTotalBytes"
            objectUsedBytes        = Get-PropertyValue -Object $derived -PropertyName "objectUsedBytes"
            objectAvailableBytes   = Get-PropertyValue -Object $derived -PropertyName "objectAvailableBytes"
            objectUsedPct          = Get-PropertyValue -Object $derived -PropertyName "objectUsedPct"
            metadataAllowedBytes   = Get-PropertyValue -Object $derived -PropertyName "metadataAllowedBytes"
            metadataUsedBytes      = Get-PropertyValue -Object $derived -PropertyName "metadataUsedBytes"
            metadataAvailableBytes = Get-PropertyValue -Object $derived -PropertyName "metadataAvailableBytes"
            metadataUsedPct        = Get-PropertyValue -Object $derived -PropertyName "metadataUsedPct"
        }
    }

    return ,$mergedRows
}

function Get-StorageGridCapacitySummary {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$BearerToken,
        [Parameter(Mandatory = $false)][object[]]$NodeHealth = @()
    )

    $metricDefinitions = [ordered]@{
        overallTotalBytes     = "storagegrid_storage_utilization_total_space_bytes"
        overallAvailableBytes = "storagegrid_storage_volume_available_space"
        objectUsedBytes       = "storagegrid_storage_utilization_data_bytes"
        objectTotalBytes      = "storagegrid_storage_utilization_usable_space_bytes"
        metadataUsedBytes     = "storagegrid_storage_utilization_metadata_bytes"
        metadataAllowedBytes  = "storagegrid_storage_utilization_metadata_allowed_bytes"
    }

    $gridMetricValues = @{}
    foreach ($metricKey in $metricDefinitions.Keys) {
        $gridMetricValues[$metricKey] = Get-StorageGridMetricAggregateValue -BaseUrl $BaseUrl -BearerToken $BearerToken -MetricName ([string]$metricDefinitions[$metricKey])
    }

    $siteLabel = Find-StorageGridMetricGroupingLabel -BaseUrl $BaseUrl -BearerToken $BearerToken -MetricNames @($metricDefinitions.Values) -LabelCandidates @("site_name", "site_id")
    $nodeLabel = Find-StorageGridMetricGroupingLabel -BaseUrl $BaseUrl -BearerToken $BearerToken -MetricNames @($metricDefinitions.Values) -LabelCandidates @("instance", "node", "node_id")

    $siteRows = @()
    if (-not [string]::IsNullOrWhiteSpace($siteLabel)) {
        $siteRows = Merge-StorageGridCapacityMetricRows -MetricDefinitions $metricDefinitions -Scope "site" -GroupLabel $siteLabel -BaseUrl $BaseUrl -BearerToken $BearerToken
    }

    $nodeRows = @()
    if (-not [string]::IsNullOrWhiteSpace($nodeLabel)) {
        $nodeRows = Merge-StorageGridCapacityMetricRows -MetricDefinitions $metricDefinitions -Scope "node" -GroupLabel $nodeLabel -BaseUrl $BaseUrl -BearerToken $BearerToken -NodeHealth $NodeHealth
    }

    $gridRecord = New-StorageGridCapacityRecord -MetricValues $gridMetricValues
    $hasCapacityData = $false
    foreach ($propertyName in @("overallTotalBytes", "objectTotalBytes", "metadataAllowedBytes")) {
        if ($null -ne (Get-PropertyValue -Object $gridRecord -PropertyName $propertyName)) {
            $hasCapacityData = $true
            break
        }
    }

    return [pscustomobject][ordered]@{
        available = [bool]$hasCapacityData
        siteLabel = [string]$siteLabel
        nodeLabel = [string]$nodeLabel
        metricNames = [pscustomobject]$metricDefinitions
        grid = $gridRecord
        sites = [object[]]@($siteRows)
        nodes = [object[]]@($nodeRows)
    }
}

function Get-StorageGridCapacityDiagnostics {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$BearerToken,
        [Parameter(Mandatory = $false)][object[]]$NodeHealth = @()
    )

    $candidateMetrics = @(
        "storagegrid_storage_utilization_total_space_bytes",
        "storagegrid_storage_volume_total_space",
        "storagegrid_storage_utilization_usable_space_bytes",
        "storagegrid_storage_volume_available_space",
        "storagegrid_storage_utilization_data_bytes",
        "storagegrid_storage_pool_utilization_data_bytes",
        "storagegrid_storage_pool_utilization_usable_space_bytes",
        "storagegrid_storage_volume_object_data_percent",
        "storagegrid_storage_utilization_ec_data_bytes",
        "storagegrid_storage_utilization_metadata_allowed_bytes",
        "storagegrid_storage_utilization_metadata_bytes",
        "storagegrid_storage_utilization_metadata_reserved_bytes",
        "storagegrid_volume_device_free_space"
    )

    $siteLabel = Find-StorageGridMetricGroupingLabel -BaseUrl $BaseUrl -BearerToken $BearerToken -MetricNames $candidateMetrics -LabelCandidates @("site_name", "site_id")
    $nodeLabel = Find-StorageGridMetricGroupingLabel -BaseUrl $BaseUrl -BearerToken $BearerToken -MetricNames $candidateMetrics -LabelCandidates @("instance", "node", "node_id")

    $metricRows = @()
    foreach ($metricName in $candidateMetrics) {
        $gridValue = Convert-ToNullableDouble -Value (Get-StorageGridMetricAggregateValue -BaseUrl $BaseUrl -BearerToken $BearerToken -MetricName $metricName)
        $siteRows = @()
        if (-not [string]::IsNullOrWhiteSpace($siteLabel)) {
            $siteRows = Convert-ToFlatObjectArray -Value (Get-StorageGridMetricGroupedValues -BaseUrl $BaseUrl -BearerToken $BearerToken -MetricName $metricName -GroupLabel $siteLabel)
        }

        $nodeRowsRaw = @()
        if (-not [string]::IsNullOrWhiteSpace($nodeLabel)) {
            $nodeRowsRaw = Convert-ToFlatObjectArray -Value (Get-StorageGridMetricGroupedValues -BaseUrl $BaseUrl -BearerToken $BearerToken -MetricName $metricName -GroupLabel $nodeLabel)
        }

        $nodeRows = @()
        foreach ($nodeRow in $nodeRowsRaw) {
            $nodeValue = Convert-ToNullableDouble -Value (Get-PropertyValue -Object $nodeRow -PropertyName "Value")
            if ($null -eq $nodeValue) {
                continue
            }

            $nodeInfo = Resolve-StorageGridNodeMetricLabel -LabelValue ([string]$nodeRow.Label) -NodeHealth $NodeHealth
            $nodeRows += [pscustomobject][ordered]@{
                node = [string]$nodeInfo.NodeName
                siteName = [string]$nodeInfo.SiteName
                metricLabel = [string]$nodeInfo.LabelValue
                value = $nodeValue
            }
        }

        $metricRows += [pscustomobject][ordered]@{
            metric = [string]$metricName
            gridValue = $gridValue
            siteLabel = [string]$siteLabel
            nodeLabel = [string]$nodeLabel
            sites = [object[]]@($siteRows | Sort-Object -Property Label)
            nodes = [object[]]@($nodeRows | Sort-Object -Property siteName, node)
        }
    }

    return [pscustomobject][ordered]@{
        generatedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
        siteLabel = [string]$siteLabel
        nodeLabel = [string]$nodeLabel
        metrics = [object[]]@($metricRows)
    }
}

function Convert-StorageGridCapacityDiagnosticsToFlatRows {
    param([Parameter(Mandatory = $true)]$Diagnostics)

    $rows = @()
    $metrics = Convert-ToArrayPayload -Payload (Get-PropertyValue -Object $Diagnostics -PropertyName "metrics")
    foreach ($metricEntry in @($metrics)) {
        $metricName = [string](Get-PropertyValue -Object $metricEntry -PropertyName "metric")
        $gridValue = Get-PropertyValue -Object $metricEntry -PropertyName "gridValue"
        $rows += [pscustomobject][ordered]@{
            scope = "grid"
            metric = $metricName
            group = "grid"
            metricLabel = "grid"
            node = ""
            siteName = ""
            value = $gridValue
        }

        $siteRows = Convert-ToFlatObjectArray -Value (Get-PropertyValue -Object $metricEntry -PropertyName "sites")
        foreach ($siteRow in @($siteRows)) {
            $rows += [pscustomobject][ordered]@{
                scope = "site"
                metric = $metricName
                group = [string](Get-PropertyValue -Object $siteRow -PropertyName "Label")
                metricLabel = [string](Get-PropertyValue -Object $siteRow -PropertyName "Label")
                node = ""
                siteName = [string](Get-PropertyValue -Object $siteRow -PropertyName "Label")
                value = Get-PropertyValue -Object $siteRow -PropertyName "Value"
            }
        }

        $nodeRows = Convert-ToFlatObjectArray -Value (Get-PropertyValue -Object $metricEntry -PropertyName "nodes")
        foreach ($nodeRow in @($nodeRows)) {
            $rows += [pscustomobject][ordered]@{
                scope = "node"
                metric = $metricName
                group = [string](Get-PropertyValue -Object $nodeRow -PropertyName "node")
                metricLabel = [string](Get-PropertyValue -Object $nodeRow -PropertyName "metricLabel")
                node = [string](Get-PropertyValue -Object $nodeRow -PropertyName "node")
                siteName = [string](Get-PropertyValue -Object $nodeRow -PropertyName "siteName")
                value = Get-PropertyValue -Object $nodeRow -PropertyName "value"
            }
        }
    }

    return ,$rows
}

if ([string]::IsNullOrWhiteSpace($JsonInputPath)) {
    if ([string]::IsNullOrWhiteSpace($Target)) {
        $Target = Read-Host "Enter the StorageGRID Management API target (e.g. admin.example.com, admin.example.com:8443, or https://admin.example.com)"
    }

    # Default bare host or host:port input to HTTPS; retain an explicit HTTP or HTTPS scheme.
    $Target = $Target.Trim()
    if ($Target -notmatch '^[a-zA-Z][a-zA-Z0-9+.-]*://') {
        $Target = "https://$Target"
    }
    $Target = $Target.TrimEnd("/")
    if ($Target -notmatch '^https?://') {
        throw "StorageGRID Management API target must use HTTP or HTTPS: $Target"
    }

    # Strip any /api path segment the user may have included.
    if ($Target -match "^(https?://[^/]+)/api.*$") {
        $Target = $Matches[1]
    }

    if ($null -eq $Credential) {
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
            $Credential = Get-Credential -Message "Enter StorageGRID Management API credentials"
        }
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

$apiUrl = $Target

$outputDirResolved = Resolve-OptionalPath -Path $OutputDir
Initialize-Path -Path $outputDirResolved

$timestamp = (Get-Date).ToUniversalTime().ToString("yyyyMMddTHHmmss")
$iso8601 = (Get-Date).ToUniversalTime().ToString("o")

# -----------------------------------------------------------------------
# JSON INPUT MODE: load a previously collected file, skip live collection
# -----------------------------------------------------------------------
$isJsonInputMode = -not [string]::IsNullOrWhiteSpace($JsonInputPath)

if ($isJsonInputMode) {
    $jsonResolvedInput = Resolve-ExistingFilePath -Path $JsonInputPath
    if ($null -eq $jsonResolvedInput) {
        throw "JSON input file not found: $JsonInputPath"
    }
    Write-Host "JSON input mode: loading data from $jsonResolvedInput"
    $asbuiltExport = Get-Content -Path $jsonResolvedInput -Raw | ConvertFrom-Json -ErrorAction Stop
    $restQueryFailures = @()
    $apiUrl = "(loaded from $([System.IO.Path]::GetFileName($jsonResolvedInput)))"

    $gridLicenseObject = $asbuiltExport.storagegrid_facts.sg_grid_license
    $reportSystemNameSource = "storagegrid"
    $gridNameRaw = Get-PropertyValue -Object $gridLicenseObject -PropertyName "systemName"
    if ([string]::IsNullOrWhiteSpace([string]$gridNameRaw)) {
        $gridNameRaw = Get-PropertyValue -Object $gridLicenseObject -PropertyName "displayName"
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$gridNameRaw)) {
        $reportSystemNameSource = [string]$gridNameRaw
    }
    $reportSystemName = Convert-ToSafeName -Value $reportSystemNameSource
    $reportBaseName = if ([string]::IsNullOrWhiteSpace($ReportPrefix)) {
        "${reportSystemName}_${timestamp}"
    } else {
        "${ReportPrefix}_${reportSystemName}_${timestamp}"
    }
    $jsonPath  = $null   # no JSON output in JSON input mode
    $mdPath    = Join-Path -Path $outputDirResolved -ChildPath "${reportBaseName}.md"
    $docxPath  = Join-Path -Path $outputDirResolved -ChildPath "${reportBaseName}.docx"

} else {
    # -----------------------------------------------------------------------
    # -----------------------------------------------------------------------
    # LIVE COLLECTION MODE
    # -----------------------------------------------------------------------
    Set-TlsPolicy -ValidateCertificates $ValidateCerts.IsPresent
    Set-ProxyPolicy -UseProxy $UseSystemProxy.IsPresent
    $apiUrl = $Target

    Write-Host "Authenticating to StorageGRID Management API ..."
    $bearerToken = Get-StorageGridToken -BaseUrl $apiUrl -Credential $Credential

    # --- Endpoint map ---
$endpointMap = @(
    # --- Public API endpoints ---
    @{ Key = "product_version";           Endpoint = "/grid/config/product-version" },
    @{ Key = "grid_health";               Endpoint = "/grid/health" },
    @{ Key = "node_health";               Endpoint = "/grid/node-health" },
    @{ Key = "service_ids";               Endpoint = "/grid/service-ids" },
    @{ Key = "grid_regions";              Endpoint = "/grid/regions" },
    @{ Key = "grid_networks";             Endpoint = "/grid/grid-networks" },
    @{ Key = "dns_servers";               Endpoint = "/grid/dns-servers" },
    @{ Key = "ntp_servers";               Endpoint = "/grid/ntp-servers" },
    @{ Key = "browser_timeout";           Endpoint = "/grid/browser-inactivity-timeout" },
    @{ Key = "grid_license";              Endpoint = "/grid/license" },
    @{ Key = "identity_source";           Endpoint = "/grid/identity-source" },
    @{ Key = "grid_groups";               Endpoint = "/grid/groups" },
    @{ Key = "grid_users";                Endpoint = "/grid/users" },
    @{ Key = "audit_config";              Endpoint = "/grid/audit" },
    @{ Key = "alert_receivers";           Endpoint = "/grid/alert-receivers" },
    @{ Key = "alert_rules";               Endpoint = "/grid/alert-rules" },
    @{ Key = "alert_silences";            Endpoint = "/grid/alert-silences" },
    @{ Key = "snmp";                      Endpoint = "/grid/snmp" },
    @{ Key = "compliance_global";         Endpoint = "/grid/compliance-global" },
    @{ Key = "storage_watermarks";        Endpoint = "/grid/storage-watermarks" },
    @{ Key = "management_certificate";    Endpoint = "/grid/management-certificate" },
    @{ Key = "internal_ca_cert";          Endpoint = "/grid/internal-ca-certificate" },
    @{ Key = "client_certificates";       Endpoint = "/grid/client-certificates" },
    @{ Key = "untrusted_client_network";  Endpoint = "/grid/untrusted-client-network" },
    @{ Key = "domain_names";              Endpoint = "/grid/domain-names" },
    @{ Key = "management_cors";           Endpoint = "/grid/management-cross-origin-request" },
    @{ Key = "traffic_policies";          Endpoint = "/grid/traffic-classes/policies" },
    @{ Key = "ilm_rules";                 Endpoint = "/grid/ilm-rules?include=compliance" },
    @{ Key = "ilm_policies";              Endpoint = "/grid/ilm-policies" },
    @{ Key = "ilm_policy_tags";           Endpoint = "/grid/ilm-policy-tags" },
    @{ Key = "ec_profiles";               Endpoint = "/grid/ec-profiles" },
    @{ Key = "grid_accounts";             Endpoint = "/grid/accounts" },

    # --- Private API endpoints ---
    @{ Key = "grid_config";               Endpoint = "/private/grid-config" },
    @{ Key = "autosupport";               Endpoint = "/private/autosupport" },
    @{ Key = "audit_destinations";        Endpoint = "/private/audit-destinations" },
    @{ Key = "ha_groups";                 Endpoint = "/private/ha-groups" },
    @{ Key = "gateway_configs";           Endpoint = "/private/gateway-configs" },
    @{ Key = "vlan_interfaces";           Endpoint = "/private/vlan-interfaces" },
    @{ Key = "ilm_pools";                 Endpoint = "/private/ilm-pools?expand=storage-nodes" },
    @{ Key = "ilm_grades";                Endpoint = "/private/ilm-grades" },
    @{ Key = "ilm_grade_site";            Endpoint = "/private/ilm-grade-site" },
    @{ Key = "cloud_storage_pools";       Endpoint = "/private/cloud-storage-pools" },
    @{ Key = "kmip_clusters";             Endpoint = "/private/kmip-clusters" },
    @{ Key = "storage_proxy";             Endpoint = "/private/storage-proxy" },
    @{ Key = "admin_proxy";               Endpoint = "/private/admin-proxy" },
    @{ Key = "sso";                       Endpoint = "/private/single-sign-on" },
    @{ Key = "ciphers";                   Endpoint = "/private/ciphers" },
    @{ Key = "bmc";                       Endpoint = "/private/bmc" },
    @{ Key = "external_load_balancers";   Endpoint = "/private/external-load-balancers" },
    @{ Key = "firewall_blocked_ports";    Endpoint = "/private/firewall-blocked-ports" },
    @{ Key = "firewall_external_ports";   Endpoint = "/private/firewall-external-ports" },
    @{ Key = "firewall_privileged_ips";   Endpoint = "/private/firewall-privileged-ips" },
    @{ Key = "grid_federation";           Endpoint = "/private/grid-federation-connections" },
    @{ Key = "network_topology";          Endpoint = "/private/network-topology" }
)

$responses = @{}

foreach ($entry in $endpointMap) {
    Write-Host "Querying $($entry.Endpoint) ..."
    $responses[$entry.Key] = Invoke-StorageGridApi -BaseUrl $apiUrl -Endpoint $entry.Endpoint -BearerToken $bearerToken
}

# --- Per-account usage ---
$accountUsageList = @()
$accountsData = $responses["grid_accounts"]
if (-not $accountsData.Failed -and $accountsData.Data -is [array]) {
    foreach ($account in $accountsData.Data) {
        $accountId = Get-PropertyValue -Object $account -PropertyName "id"
        if ([string]::IsNullOrWhiteSpace([string]$accountId)) { continue }
        Write-Host "Querying /grid/accounts/$accountId/usage ..."
        $usageResp = Invoke-StorageGridApi -BaseUrl $apiUrl -Endpoint "/grid/accounts/$accountId/usage?includeBucketDetail=true" -BearerToken $bearerToken
        if (-not $usageResp.Failed) {
            $usageData = $usageResp.Data
            if ($null -ne $usageData) {
                $usageData | Add-Member -NotePropertyName "__accountId" -NotePropertyValue ([string]$accountId) -Force -ErrorAction SilentlyContinue
                $accountUsageList += $usageData
            }
        }
    }
}
$responses["grid_account_usage"] = [pscustomobject]@{
    Endpoint = "/grid/accounts/*/usage"
    Failed   = $false
    Status   = 200
    Data     = $accountUsageList
    Message  = ""
}

# --- Per-traffic-policy details ---
$trafficPolicyDetails = @()
$trafficPoliciesData = $responses["traffic_policies"]
if (-not $trafficPoliciesData.Failed -and $trafficPoliciesData.Data -is [array]) {
    foreach ($policy in $trafficPoliciesData.Data) {
        $policyId = Get-PropertyValue -Object $policy -PropertyName "id"
        if ([string]::IsNullOrWhiteSpace([string]$policyId)) { continue }
        Write-Host "Querying /grid/traffic-classes/policies/$policyId ..."
        $policyResp = Invoke-StorageGridApi -BaseUrl $apiUrl -Endpoint "/grid/traffic-classes/policies/$policyId" -BearerToken $bearerToken
        if (-not $policyResp.Failed -and $null -ne $policyResp.Data) {
            $trafficPolicyDetails += $policyResp.Data
        }
    }
}
$responses["traffic_policy_details"] = [pscustomobject]@{
    Endpoint = "/grid/traffic-classes/policies/*"
    Failed   = $false
    Status   = 200
    Data     = $trafficPolicyDetails
    Message  = ""
}

# --- Per-LB-endpoint server configs ---
$lbServerConfigs = @()
$gatewayConfigsData = $responses["gateway_configs"]
if (-not $gatewayConfigsData.Failed -and $gatewayConfigsData.Data -is [array]) {
    foreach ($ep in $gatewayConfigsData.Data) {
        $epId = Get-PropertyValue -Object $ep -PropertyName "id"
        if ([string]::IsNullOrWhiteSpace([string]$epId)) { continue }
        Write-Host "Querying /private/gateway-configs/$epId/server-config ..."
        $scResp = Invoke-StorageGridApi -BaseUrl $apiUrl -Endpoint "/private/gateway-configs/$epId/server-config" -BearerToken $bearerToken
        if (-not $scResp.Failed -and $null -ne $scResp.Data) {
            $scData = $scResp.Data
            $scData | Add-Member -NotePropertyName "__endpointId" -NotePropertyValue ([string]$epId) -Force -ErrorAction SilentlyContinue
            $lbServerConfigs += $scData
        }
    }
}
$responses["lb_server_configs"] = [pscustomobject]@{
    Endpoint = "/private/gateway-configs/*/server-config"
    Failed   = $false
    Status   = 200
    Data     = $lbServerConfigs
    Message  = ""
}

$nodeHealthData = $responses["node_health"]
$capacityMetrics = Get-StorageGridCapacitySummary -BaseUrl $apiUrl -BearerToken $bearerToken -NodeHealth (Convert-ToArrayPayload -Payload $nodeHealthData.Data)
$applianceStorageMetrics = [pscustomobject]@{ available = $false; nodes = [object[]]@() }
$nodeAttributeDetails = [pscustomobject]@{ available = $false; nodes = [object[]]@() }
$santricityApplianceCandidates = @()
if ($santricityDetectionEnabled) {
    $applianceStorageMetrics = Get-StorageGridApplianceStorageMetrics -BaseUrl $apiUrl -BearerToken $bearerToken -NodeHealth (Convert-ToArrayPayload -Payload $nodeHealthData.Data)
    # Must run before the array-conversion loop below mutates responses["service_ids"].Data
    $nodeAttributeDetails = Get-StorageGridNodeAttributeDetails -BaseUrl $apiUrl -BearerToken $bearerToken -ServiceIdsData $responses["service_ids"].Data -NetworkTopologyData $responses["network_topology"].Data -NodeHealth (Convert-ToArrayPayload -Payload $nodeHealthData.Data)
    $santricityApplianceCandidates = Get-StorageGridSantricityApplianceCandidates -NetworkTopology $responses["network_topology"].Data -NodeAttributes $nodeAttributeDetails -NodeHealth (Convert-ToArrayPayload -Payload $nodeHealthData.Data) -ApplianceStorageMetrics $applianceStorageMetrics
}
$santricityCollectionResults = @()
if ($santricityApplianceCandidates.Count -gt 0) {
    Write-Host "Discovered SANtricity-capable appliances:"
    foreach ($candidate in $santricityApplianceCandidates) {
        Write-Host ("  {0} ({1}): {2}" -f $candidate.NodeName, $candidate.ApplianceModel, ($candidate.CandidateIps -join ', '))
    }
}
$capacityDiagnostics = $null
if ($ExportCapacityDiagnostics) {
    Write-Host "Collecting focused capacity diagnostics ..."
    $capacityDiagnostics = Get-StorageGridCapacityDiagnostics -BaseUrl $apiUrl -BearerToken $bearerToken -NodeHealth (Convert-ToArrayPayload -Payload $nodeHealthData.Data)
}

# Keys whose Data payloads are arrays that may need Convert-ToArrayPayload
$arrayPayloadKeys = @(
    "grid_regions",
    "grid_networks",
    "dns_servers",
    "ntp_servers",
    "node_health",
    "service_ids",
    "grid_groups",
    "grid_users",
    "alert_receivers",
    "alert_rules",
    "alert_silences",
    "client_certificates",
    "traffic_policies",
    "ilm_rules",
    "ilm_policies",
    "ilm_policy_tags",
    "ec_profiles",
    "grid_accounts",
    "ha_groups",
    "gateway_configs",
    "vlan_interfaces",
    "ilm_pools",
    "ilm_grades",
    "ilm_grade_site",
    "cloud_storage_pools",
    "kmip_clusters",
    "external_load_balancers",
    "firewall_blocked_ports",
    "domain_names",
    "grid_federation"
)

foreach ($arrayKey in $arrayPayloadKeys) {
    if ($responses.ContainsKey($arrayKey)) {
        $responses[$arrayKey].Data = Convert-ToArrayPayload -Payload $responses[$arrayKey].Data
    }
}

$asbuiltExport = [ordered]@{
    changed           = $false
    failed            = $false
    msg               = "REST-collected data"
    storagegrid_facts = [ordered]@{
        sg_product_version          = Get-ResponseData -Responses $responses -Key "product_version"
        sg_grid_config              = Get-ResponseData -Responses $responses -Key "grid_config"
        sg_grid_health              = Get-ResponseData -Responses $responses -Key "grid_health"
        sg_node_health              = Get-ResponseData -Responses $responses -Key "node_health"          -AsArray
        sg_service_ids              = Get-ResponseData -Responses $responses -Key "service_ids"          -AsArray
        sg_network_topology         = Get-ResponseData -Responses $responses -Key "network_topology"
        sg_grid_regions             = Get-ResponseData -Responses $responses -Key "grid_regions"         -AsArray
        sg_grid_networks            = Get-ResponseData -Responses $responses -Key "grid_networks"          -AsArray
        sg_dns_servers              = Get-ResponseData -Responses $responses -Key "dns_servers"            -AsArray
        sg_ntp_servers              = Get-ResponseData -Responses $responses -Key "ntp_servers"            -AsArray
        sg_browser_timeout          = Get-ResponseData -Responses $responses -Key "browser_timeout"
        sg_grid_license             = Get-ResponseData -Responses $responses -Key "grid_license"
        sg_capacity_metrics         = $capacityMetrics
        sg_appliance_storage_metrics = $applianceStorageMetrics
        sg_node_attribute_details   = $nodeAttributeDetails
        sg_santricity_appliance_candidates = [object[]]@($santricityApplianceCandidates)
        sg_santricity_appliance_collection = [object[]]@($santricityCollectionResults)
        sg_autosupport              = Get-ResponseData -Responses $responses -Key "autosupport"
        sg_identity_source          = Get-ResponseData -Responses $responses -Key "identity_source"
        sg_grid_groups              = Get-ResponseData -Responses $responses -Key "grid_groups"          -AsArray
        sg_grid_users               = Get-ResponseData -Responses $responses -Key "grid_users"           -AsArray
        sg_audit_config             = Get-ResponseData -Responses $responses -Key "audit_config"
        sg_audit_destinations       = Get-ResponseData -Responses $responses -Key "audit_destinations"
        sg_alert_receivers          = Get-ResponseData -Responses $responses -Key "alert_receivers"      -AsArray
        sg_alert_rules              = Get-ResponseData -Responses $responses -Key "alert_rules"          -AsArray
        sg_alert_silences           = Get-ResponseData -Responses $responses -Key "alert_silences"       -AsArray
        sg_snmp                     = Get-ResponseData -Responses $responses -Key "snmp"
        sg_compliance_global        = Get-ResponseData -Responses $responses -Key "compliance_global"
        sg_storage_watermarks       = Get-ResponseData -Responses $responses -Key "storage_watermarks"
        sg_management_certificate   = Get-ResponseData -Responses $responses -Key "management_certificate"
        sg_internal_ca_cert         = Get-ResponseData -Responses $responses -Key "internal_ca_cert"
        sg_client_certificates      = Get-ResponseData -Responses $responses -Key "client_certificates"  -AsArray
        sg_untrusted_client_network = Get-ResponseData -Responses $responses -Key "untrusted_client_network"
        sg_management_cors         = Get-ResponseData -Responses $responses -Key "management_cors"
        sg_domain_names             = Get-ResponseData -Responses $responses -Key "domain_names"           -AsArray
        sg_ha_groups                = Get-ResponseData -Responses $responses -Key "ha_groups"            -AsArray
        sg_gateway_configs          = Get-ResponseData -Responses $responses -Key "gateway_configs"      -AsArray
        sg_lb_server_configs        = [object[]]@($lbServerConfigs)
        sg_traffic_policies         = Get-ResponseData -Responses $responses -Key "traffic_policies"     -AsArray
        sg_traffic_policy_details   = [object[]]@($trafficPolicyDetails)
        sg_vlan_interfaces          = Get-ResponseData -Responses $responses -Key "vlan_interfaces"      -AsArray
        sg_external_load_balancers  = Get-ResponseData -Responses $responses -Key "external_load_balancers" -AsArray
        sg_firewall_blocked_ports   = Get-ResponseData -Responses $responses -Key "firewall_blocked_ports"  -AsArray
        sg_firewall_external_ports  = Get-ResponseData -Responses $responses -Key "firewall_external_ports"
        sg_firewall_privileged_ips  = Get-ResponseData -Responses $responses -Key "firewall_privileged_ips"
        sg_ilm_rules                = Get-ResponseData -Responses $responses -Key "ilm_rules"            -AsArray
        sg_ilm_policies             = Get-ResponseData -Responses $responses -Key "ilm_policies"         -AsArray
        sg_ilm_policy_tags          = Get-ResponseData -Responses $responses -Key "ilm_policy_tags"      -AsArray
        sg_ec_profiles              = Get-ResponseData -Responses $responses -Key "ec_profiles"          -AsArray
        sg_ilm_pools                = Get-ResponseData -Responses $responses -Key "ilm_pools"            -AsArray
        sg_ilm_grades               = Get-ResponseData -Responses $responses -Key "ilm_grades"           -AsArray
        sg_ilm_grade_site           = Get-ResponseData -Responses $responses -Key "ilm_grade_site"       -AsArray
        sg_cloud_storage_pools      = Get-ResponseData -Responses $responses -Key "cloud_storage_pools"  -AsArray
        sg_kmip_clusters            = Get-ResponseData -Responses $responses -Key "kmip_clusters"        -AsArray
        sg_storage_proxy            = Get-ResponseData -Responses $responses -Key "storage_proxy"
        sg_admin_proxy              = Get-ResponseData -Responses $responses -Key "admin_proxy"
        sg_sso                      = Get-ResponseData -Responses $responses -Key "sso"
        sg_ciphers                  = Get-ResponseData -Responses $responses -Key "ciphers"
        sg_bmc                      = Get-ResponseData -Responses $responses -Key "bmc"
        sg_grid_federation          = Get-ResponseData -Responses $responses -Key "grid_federation"      -AsArray
        sg_grid_accounts            = Get-ResponseData -Responses $responses -Key "grid_accounts"        -AsArray
        sg_grid_account_usage       = [object[]]@($accountUsageList)
    }
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

    # Derive report name from grid system name in license data
    $gridLicenseObject = $responses["grid_license"].Data
    $reportSystemNameSource = "storagegrid"
    $gridNameRaw = Get-PropertyValue -Object $gridLicenseObject -PropertyName "systemName"
    if ([string]::IsNullOrWhiteSpace([string]$gridNameRaw)) {
        $gridNameRaw = Get-PropertyValue -Object $gridLicenseObject -PropertyName "displayName"
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$gridNameRaw)) {
        $reportSystemNameSource = [string]$gridNameRaw
    }
    $reportSystemName = Convert-ToSafeName -Value $reportSystemNameSource
    $reportBaseName = if ([string]::IsNullOrWhiteSpace($ReportPrefix)) {
        "${reportSystemName}_${timestamp}"
    } else {
        "${ReportPrefix}_${reportSystemName}_${timestamp}"
    }

    $jsonPath  = Join-Path -Path $outputDirResolved -ChildPath "${reportBaseName}.json"
    $mdPath    = Join-Path -Path $outputDirResolved -ChildPath "${reportBaseName}.md"
    $docxPath  = Join-Path -Path $outputDirResolved -ChildPath "${reportBaseName}.docx"
    $capacityDiagnosticsJsonPath = Join-Path -Path $outputDirResolved -ChildPath "${reportBaseName}_capacity-diagnostics.json"
    $capacityDiagnosticsCsvPath  = Join-Path -Path $outputDirResolved -ChildPath "${reportBaseName}_capacity-diagnostics.csv"

    $asbuiltExport | ConvertTo-Json -Depth 100 | Set-Content -Path $jsonPath -Encoding UTF8

    if ($ExportCapacityDiagnostics -and $null -ne $capacityDiagnostics) {
        $capacityDiagnostics | ConvertTo-Json -Depth 20 | Set-Content -Path $capacityDiagnosticsJsonPath -Encoding UTF8
        $capacityDiagnosticsFlatRows = Convert-StorageGridCapacityDiagnosticsToFlatRows -Diagnostics $capacityDiagnostics
        $capacityDiagnosticsFlatRows | Export-Csv -NoTypeInformation -Path $capacityDiagnosticsCsvPath -Encoding UTF8
        Write-Host "Capacity diagnostics JSON generated at: $capacityDiagnosticsJsonPath"
        Write-Host "Capacity diagnostics CSV generated at: $capacityDiagnosticsCsvPath"
    }

    if ($CollectionOnly) {
        Write-Host "Collection-only mode enabled. JSON export generated at: $jsonPath"
        return
    }
} # end live-collection else block

Write-Host "Rendering markdown report ..."
Write-StorageGridMarkdown -Export $asbuiltExport -ApiUrl $apiUrl -Iso8601 $iso8601 -OutputPath $mdPath -RestQueryFailures $restQueryFailures -TitleFields $TitlePageFields -IncludeTitlePage ([bool]$DocxEnableTitlePage)

$docxReferenceTemplate = Resolve-ExistingFilePath -Path ".\docx\reference-template.docx"

Write-Host "Converting markdown to DOCX ..."
Convert-AsBuiltMarkdownToDocx -MarkdownPath $mdPath -DocxPath $docxPath -ReferenceDocumentPath $docxReferenceTemplate -TableStyleName $DocxTableStyleName -NumberSections ([bool]$EnableDocxNumberSections) | Out-Null

$reportMetadata = Read-ReportMetadata -MetadataPath ".\docx\report-metadata-storagegrid.yml"

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

# Validate and set metadata with fallbacks
$docxMetaTitle = if ([string]::IsNullOrWhiteSpace([string]$reportMetadata.title)) { "StorageGRID As-Built" } else { [string]$reportMetadata.title }
$docxMetaSubject = if ([string]::IsNullOrWhiteSpace([string]$reportMetadata.subject)) { "StorageGRID Configuration Summary" } else { [string]$reportMetadata.subject }
$docxMetaAuthor = if ([string]::IsNullOrWhiteSpace([string]$reportMetadata.author)) { "NetApp" } else { [string]$reportMetadata.author }

# Convert empty strings to a placeholder so the custom document property is still created
$docPropPlaceholder = "<<PLEASE UPDATE via DOCUMENT PROPERTIES>>"
$docPropCustomerName = if ([string]::IsNullOrWhiteSpace($docPropCustomerName)) { $docPropPlaceholder } else { $docPropCustomerName }
$docPropCustomerLocation = if ([string]::IsNullOrWhiteSpace($docPropCustomerLocation)) { $docPropPlaceholder } else { $docPropCustomerLocation }
$docPropProjectName = if ([string]::IsNullOrWhiteSpace($docPropProjectName)) { $docPropPlaceholder } else { $docPropProjectName }
$docPropSystemName = if ([string]::IsNullOrWhiteSpace($docPropSystemName)) { $null } else { $docPropSystemName }

Set-DocxMetadataProperties -DocxPath $docxPath -Title $docxMetaTitle -Subject $docxMetaSubject -Author $docxMetaAuthor -CustomerName $docPropCustomerName -CustomerLocation $docPropCustomerLocation -ProjectName $docPropProjectName -SystemName $docPropSystemName

Set-DocxTableStyle -DocxPath $docxPath -TableStyleId $DocxTableStyleName -HeaderParagraphStyle $DocxTableHeaderParagraphStyle -BodyParagraphStyle $DocxTableBodyParagraphStyle -AutofitToWindow ([bool]$DocxTableAutofitToWindow)

Write-Host "StorageGRID DOCX report generated at: $docxPath"

if ($santricityApplianceCandidates.Count -gt 0 -and $santricityDetectionEnabled) {
    $shouldCollectSantricity = $santricityCollectionRequested
    if (-not $collectSantricitySpecified) {
        if ($NonInteractive) {
            Write-Warning "SANtricity-capable StorageGRID appliances were detected, but -NonInteractive was specified. SANtricity collection was skipped."
        }
        else {
            $answer = Read-StorageGridSantricityYesNo -Heading 'SANTRICITY APPLIANCES DETECTED' -Message @(
                ("{0} StorageGRID appliances contain embedded E-Series/SANtricity storage." -f $santricityApplianceCandidates.Count)
                ''
                'The StorageGRID Markdown and DOCX reports have been generated.'
                ''
                'Generate SANtricity as-built reports for all detected appliances?'
            )
            $shouldCollectSantricity = $answer -eq 'Y'
        }
    }

    if ($shouldCollectSantricity) {
        Import-Module (Join-Path -Path $WorkspaceRoot -ChildPath 'Modules\AsBuilt.ESeries.psm1') -Force
        $credentialMap = $null
        if (-not [string]::IsNullOrWhiteSpace($SantricityAuthMapPath)) {
            $credentialMapPathResolved = Resolve-ExistingFilePath -Path $SantricityAuthMapPath
            if ($null -eq $credentialMapPathResolved) {
                throw "SANtricity auth map file not found: $SantricityAuthMapPath"
            }
            $credentialMap = Import-Clixml -LiteralPath $credentialMapPathResolved
            if (-not ($credentialMap -is [hashtable])) {
                throw "SANtricity auth map must be a CLIXML-exported hashtable: $SantricityAuthMapPath"
            }
        }
        $resolvedSantricityCredentials = Resolve-StorageGridSantricityCredentials -Candidates $santricityApplianceCandidates -SharedCredential $SantricityCredential -CredentialMap $credentialMap -NonInteractive:$NonInteractive
        $SantricityCredential = $resolvedSantricityCredentials.SharedCredential
        $credentialMap = $resolvedSantricityCredentials.CredentialMap
        if ($null -eq $SantricityCredential -and $null -eq $credentialMap) {
            Write-Warning "SANtricity collection was requested, but no credentials were supplied and interactive prompting is unavailable. SANtricity collection was skipped."
            $shouldCollectSantricity = $false
        }

        if ($shouldCollectSantricity) {
            $gridName = [string](Get-PropertyValue -Object $responses["grid_license"].Data -PropertyName 'systemName')
            if ([string]::IsNullOrWhiteSpace($gridName)) { $gridName = 'storagegrid' }
            $santricityOutputDirectory = Join-Path -Path $outputDirResolved -ChildPath ("{0}_santricity_appliances" -f (Convert-ToSafeName -Value $gridName))
            $santricityCollectionResults = Invoke-StorageGridSantricityApplianceCollection -Candidates $santricityApplianceCandidates -OutputDirectory $santricityOutputDirectory -WorkspaceRoot $WorkspaceRoot -SharedCredential $SantricityCredential -CredentialMap $credentialMap -ValidateCerts:$ValidateCerts -UseSystemProxy:$UseSystemProxy -KeepIntermediateOutputs:$KeepIntermediateOutputs -DocxEnableTitlePage:$DocxEnableTitlePage -CustomerName $CustomerName -CustomerLocation $CustomerLocation -ProjectName $ProjectName -TitlePageFields $TitlePageFields
            $santricitySummaryPath = Join-Path -Path $santricityOutputDirectory -ChildPath 'santricity-collection-summary.json'
            [ordered]@{
                GridName       = $gridName
                CollectedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
                Results        = [object[]]@($santricityCollectionResults)
            } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $santricitySummaryPath -Encoding UTF8
            Write-Host "SANtricity appliance collection summary: $santricitySummaryPath"
            Write-Host ("SANtricity appliance collection complete: {0} succeeded, {1} failed." -f @($santricityCollectionResults | Where-Object Status -eq 'Succeeded').Count, @($santricityCollectionResults | Where-Object Status -eq 'Failed').Count)
            if (-not [string]::IsNullOrWhiteSpace($jsonPath) -and (Test-Path -LiteralPath $jsonPath -PathType Leaf)) {
                $asbuiltExport.storagegrid_facts.sg_santricity_appliance_collection = [object[]]@($santricityCollectionResults)
                $asbuiltExport | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
            }
        }
    }
}

if ($CleanupIntermediateOutputs -and -not $KeepIntermediateOutputs) {
    $pathsToRemove = @()
    if (-not [string]::IsNullOrWhiteSpace($jsonPath)) { $pathsToRemove += $jsonPath }
    if (-not [string]::IsNullOrWhiteSpace($mdPath)) { $pathsToRemove += $mdPath }
    if ($pathsToRemove.Count -gt 0) {
        Remove-Item -Path $pathsToRemove -Force -ErrorAction SilentlyContinue
    }
}

}

Export-ModuleMember -Function 'Invoke-StorageGridAsBuilt'
