# ChangeLog:
# 2026.10.15.6 - Prompt with supported platform choices when Platform is omitted.
# 2026.10.15.5 - Removed replay, collection-only, capacity diagnostics, system-proxy, and SSID parameters.
# 2026.10.15.4 - Removed structured title-page field inputs; retain named fields and prompting.
# 2026.10.15.3 - Removed legacy plaintext credential inputs, cleanup and DOCX layout overrides, and the E-Series Port parameter.
# 2026.10.15.2 - Added the entry-point script version identifier.
# 2026.10.15.1 - Enabled TLS certificate validation by default; allow explicit opt-out with -ValidateCerts:$false.
[CmdletBinding()]
param(
    [ValidateSet('StorageGRID', 'ESeries')]
    [string]$Platform,

    # E-Series target examples: santricity.example.com:8443 or https://santricity.example.com:8443
    [string]$Target,
    [System.Management.Automation.PSCredential]$Credential,
    [switch]$ValidateCerts = $true,
    [string]$OutputDir,
    [switch]$CollectSantricityAppliances,
    [switch]$NonInteractive,
    [System.Management.Automation.PSCredential]$SantricityCredential,
    [string]$SantricityAuthMapPath,
    [string]$ReportPrefix,
    [switch]$KeepIntermediateOutputs,
    [Alias('TitleCustomerName')]
    [string]$CustomerName,
    [Alias('TitleCustomerLocation')]
    [string]$CustomerLocation,
    [Alias('TitleProjectName')]
    [string]$ProjectName,
    [switch]$PromptForTitlePageFields
)

$ScriptVersion = '2026.10.15.6'
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Write-Verbose "As-built entry point version: $ScriptVersion"

if ([string]::IsNullOrWhiteSpace($Platform)) {
    if ($NonInteractive) {
        throw 'Platform is required in non-interactive mode. Specify -Platform StorageGRID or -Platform ESeries.'
    }

    do {
        $platformSelection = Read-Host 'Select platform (StorageGRID or ESeries)'
        if ($platformSelection -ieq 'StorageGRID') {
            $Platform = 'StorageGRID'
        }
        elseif ($platformSelection -ieq 'ESeries') {
            $Platform = 'ESeries'
        }
        else {
            Write-Warning 'Invalid platform. Enter StorageGRID or ESeries.'
        }
    } while ([string]::IsNullOrWhiteSpace($Platform))
}

$moduleName = if ($Platform -eq 'StorageGRID') { 'AsBuilt.StorageGrid.psm1' } else { 'AsBuilt.ESeries.psm1' }
$commandName = if ($Platform -eq 'StorageGRID') { 'Invoke-StorageGridAsBuilt' } else { 'Invoke-ESeriesAsBuilt' }
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\$moduleName"
if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
    throw "The required $Platform module was not found at '$modulePath'. Re-extract the complete distribution package."
}

try {
    Import-Module $modulePath -Force -ErrorAction Stop
}
catch {
    throw "Failed to load the $Platform module from '$modulePath': $($_.Exception.Message) Restore the module and its dependencies from the distribution package."
}

$commonParameterNames = @(
    'Target',
    'Credential',
    'ValidateCerts',
    'OutputDir',
    'ReportPrefix',
    'KeepIntermediateOutputs',
    'NonInteractive',
    'CustomerName',
    'CustomerLocation',
    'ProjectName',
    'PromptForTitlePageFields'
)

$platformParameterNames = if ($Platform -eq 'StorageGRID') {
    @(
        'CollectSantricityAppliances',
        'SantricityCredential',
        'SantricityAuthMapPath'
    )
}
else {
    @()
}

$invokeParameters = @{ WorkspaceRoot = $PSScriptRoot }
foreach ($parameterName in @($commonParameterNames + $platformParameterNames)) {
    if ($PSBoundParameters.ContainsKey($parameterName)) {
        $invokeParameters[$parameterName] = $PSBoundParameters[$parameterName]
    }
}

try {
    & $commandName @invokeParameters
}
catch {
    Write-Error "As-built collection for $Platform did not complete: $($_.Exception.Message)"
    Write-Error 'Correct the reported problem and run the collection again. Use -Verbose for additional diagnostics.'
    exit 1
}
