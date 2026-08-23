Set-StrictMode -Version Latest

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

function Resolve-AsBuiltPandocExecutable {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$DocxDirectory)

    $localPandoc = Join-Path -Path $DocxDirectory -ChildPath 'pandoc.exe'
    if (Test-Path -LiteralPath $localPandoc -PathType Leaf) {
        return (Resolve-Path -LiteralPath $localPandoc).ProviderPath
    }

    $pandocCommand = Get-Command pandoc -ErrorAction SilentlyContinue
    if ($null -ne $pandocCommand) {
        return $pandocCommand.Source
    }

    throw "Pandoc was not found. Place pandoc.exe in '$DocxDirectory' or install Pandoc in PATH."
}

function Convert-AsBuiltMarkdownToDocx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$MarkdownPath,
        [Parameter(Mandatory = $true)][string]$DocxPath,
        [Parameter(Mandatory = $true)][string]$PandocExecutable,
        [Parameter(Mandatory = $true)][string]$MarkdownFormat,
        [Parameter(Mandatory = $false)][string]$ReferenceDocumentPath,
        [Parameter(Mandatory = $false)][string]$LuaFilterPath,
        [Parameter(Mandatory = $false)][string]$TableStyleName,
        [Parameter(Mandatory = $false)][bool]$NumberSections = $false,
        [Parameter(Mandatory = $false)][string]$SoftBreakToken = 'ASBUILT_SOFT_BREAK_TOKEN',
        [Parameter(Mandatory = $false)][string]$TemporaryFilePrefix = 'asbuilt_docx'
    )

    if (-not (Test-Path -LiteralPath $MarkdownPath -PathType Leaf)) {
        throw "Markdown input not found: $MarkdownPath"
    }

    $pandocInputPath = $MarkdownPath
    $temporaryMarkdownPath = $null
    $markdownText = [System.IO.File]::ReadAllText($MarkdownPath)
    if ($markdownText -match '(?i)<br\s*/?>') {
        $tokenizedMarkdown = [System.Text.RegularExpressions.Regex]::Replace($markdownText, '(?i)<br\s*/?>', $SoftBreakToken)
        $temporaryMarkdownPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ('{0}_{1}.md' -f $TemporaryFilePrefix, [System.Guid]::NewGuid().ToString('N')))
        [System.IO.File]::WriteAllText($temporaryMarkdownPath, $tokenizedMarkdown, (New-Object System.Text.UTF8Encoding($false)))
        $pandocInputPath = $temporaryMarkdownPath
    }

    $pandocArguments = @($pandocInputPath, '--from', $MarkdownFormat, '--to', 'docx')
    if (-not [string]::IsNullOrWhiteSpace($LuaFilterPath)) {
        $pandocArguments += @('-M', "docx-table-style=$TableStyleName", '-M', "docx_table_style=$TableStyleName", '--lua-filter', $LuaFilterPath)
    }
    if ($NumberSections) {
        $pandocArguments += '--number-sections'
    }
    if (-not [string]::IsNullOrWhiteSpace($ReferenceDocumentPath)) {
        $pandocArguments += @('--reference-doc', $ReferenceDocumentPath)
    }
    $pandocArguments += @('-o', $DocxPath)

    $pandocLogPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ('pandoc_{0}.log' -f (Get-Random)))
    try {
        & $PandocExecutable @pandocArguments > $pandocLogPath 2>&1
        $pandocExitCode = $LASTEXITCODE
        if ($pandocExitCode -ne 0) {
            throw "Pandoc conversion failed with exit code $pandocExitCode."
        }
    }
    finally {
        if (-not [string]::IsNullOrWhiteSpace($temporaryMarkdownPath) -and (Test-Path -LiteralPath $temporaryMarkdownPath -PathType Leaf)) {
            Remove-Item -LiteralPath $temporaryMarkdownPath -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $pandocLogPath -PathType Leaf) {
            Remove-Item -LiteralPath $pandocLogPath -Force -ErrorAction SilentlyContinue
        }
    }

    if (-not (Test-Path -LiteralPath $DocxPath -PathType Leaf)) {
        throw "Pandoc did not generate output file: $DocxPath"
    }

    Convert-AsBuiltDocxSoftBreakTokens -DocxPath $DocxPath -Token $SoftBreakToken
    return (Resolve-Path -LiteralPath $DocxPath).ProviderPath
}

function Convert-AsBuiltDocxSoftBreakTokens {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$DocxPath,
        [Parameter(Mandatory = $true)][string]$Token
    )

    if ([string]::IsNullOrWhiteSpace($Token) -or -not (Test-Path -LiteralPath $DocxPath -PathType Leaf)) {
        return
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($DocxPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        if ($null -eq $entry) {
            return
        }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        $documentXmlText = $reader.ReadToEnd()
        $reader.Dispose()
        if ($documentXmlText.IndexOf($Token, [System.StringComparison]::Ordinal) -lt 0) {
            return
        }

        $entry.Delete()
        $newEntry = $zip.CreateEntry('word/document.xml', [System.IO.Compression.CompressionLevel]::Optimal)
        $writer = New-Object System.IO.StreamWriter($newEntry.Open())
        $writer.Write($documentXmlText.Replace($Token, '</w:t></w:r><w:r><w:br/></w:r><w:r><w:t xml:space="preserve">'))
        $writer.Dispose()
    }
    finally {
        $zip.Dispose()
    }
}

function Repair-AsBuiltDocxSectionHeaders {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$DocxPath)

    if (-not (Test-Path -LiteralPath $DocxPath -PathType Leaf)) {
        return
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($DocxPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        if ($null -eq $entry) {
            return
        }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        $documentXmlText = $reader.ReadToEnd()
        $reader.Dispose()

        [xml]$documentXml = $documentXmlText
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

        $sectionProperties = @($documentXml.SelectNodes('//w:sectPr', $namespaceManager))
        if ($sectionProperties.Count -lt 2) {
            return
        }

        $coverSection = $sectionProperties[0]
        $bodySection = $sectionProperties[$sectionProperties.Count - 1]
        $sectionNodes = @('footerReference', 'headerReference', 'pgSz', 'pgMar', 'cols', 'titlePg', 'docGrid')

        foreach ($nodeName in $sectionNodes) {
            foreach ($existing in @($coverSection.SelectNodes('./w:' + $nodeName, $namespaceManager))) {
                $null = $coverSection.RemoveChild($existing)
            }
        }
        foreach ($nodeName in $sectionNodes) {
            foreach ($sourceNode in @($bodySection.SelectNodes('./w:' + $nodeName, $namespaceManager))) {
                $null = $coverSection.AppendChild($documentXml.ImportNode($sourceNode, $true))
            }
        }
        foreach ($existing in @($bodySection.SelectNodes('./w:headerReference[@w:type="first"]', $namespaceManager))) {
            $null = $bodySection.RemoveChild($existing)
        }
        foreach ($existing in @($bodySection.SelectNodes('./w:titlePg', $namespaceManager))) {
            $null = $bodySection.RemoveChild($existing)
        }

        $entry.Delete()
        $replacementEntry = $zip.CreateEntry('word/document.xml')
        $writer = New-Object System.IO.StreamWriter($replacementEntry.Open())
        $writer.Write($documentXml.OuterXml)
        $writer.Dispose()
    }
    finally {
        $zip.Dispose()
    }
}

Export-ModuleMember -Function @(
    'Initialize-AsBuiltPath',
    'Resolve-AsBuiltPath',
    'Resolve-AsBuiltExistingFile',
    'ConvertTo-AsBuiltSafeName',
    'Get-AsBuiltPropertyValue',
    'ConvertTo-AsBuiltArray',
    'New-AsBuiltSettings',
    'Resolve-AsBuiltPandocExecutable',
    'Convert-AsBuiltMarkdownToDocx',
    'Convert-AsBuiltDocxSoftBreakTokens',
    'Repair-AsBuiltDocxSectionHeaders'
)