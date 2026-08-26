[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('StorageGRID', 'ESeries')]
    [string]$Platform,

    # Preferred E-Series form: https://santricity.example.com or https://santricity.example.com:8443
    [string]$Target,
    [int]$Port = 8443,
    [string]$ApiUsername,
    [System.Security.SecureString]$ApiPasswordSecure,
    [Alias('ApiPassword')]
    [string]$LegacyApiSecret,
    [System.Management.Automation.PSCredential]$Credential,
    [string]$Ssid,
    [switch]$ValidateCerts,
    [switch]$UseSystemProxy,
    [string]$OutputDir,
    [switch]$CollectionOnly,
    [switch]$ExportCapacityDiagnostics,
    [switch]$CollectSantricityAppliances,
    [switch]$NonInteractive,
    [System.Management.Automation.PSCredential]$SantricityCredential,
    [string]$SantricityAuthMapPath,
    [string]$JsonInputPath,
    [string]$ReportPrefix,
    [switch]$CleanupIntermediateOutputs,
    [switch]$KeepIntermediateOutputs,
    [switch]$EnableDocxToc,
    [int]$DocxTocDepth,
    [switch]$EnableDocxNumberSections,
    [string]$DocxTableStyleName,
    [string]$DocxTableHeaderParagraphStyle,
    [string]$DocxTableBodyParagraphStyle,
    [switch]$DocxTableAutofitToWindow,
    [switch]$DocxEnableTitlePage,
    [Alias('TitleCustomerName')]
    [string]$CustomerName,
    [Alias('TitleCustomerLocation')]
    [string]$CustomerLocation,
    [Alias('TitleProjectName')]
    [string]$ProjectName,
    [hashtable]$TitlePageFields,
    [string]$TitlePageFieldsJson,
    [switch]$PromptForTitlePageFields
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$moduleName = if ($Platform -eq 'StorageGRID') { 'AsBuilt.StorageGrid.psm1' } else { 'AsBuilt.ESeries.psm1' }
$commandName = if ($Platform -eq 'StorageGRID') { 'Invoke-StorageGridAsBuilt' } else { 'Invoke-ESeriesAsBuilt' }
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath "Modules\$moduleName") -Force

$commonParameterNames = @(
    'Target',
    'ApiUsername',
    'ApiPasswordSecure',
    'LegacyApiSecret',
    'Credential',
    'ValidateCerts',
    'UseSystemProxy',
    'OutputDir',
    'CollectionOnly',
    'JsonInputPath',
    'ReportPrefix',
    'CleanupIntermediateOutputs',
    'KeepIntermediateOutputs',
    'NonInteractive',
    'EnableDocxToc',
    'DocxTocDepth',
    'EnableDocxNumberSections',
    'DocxTableStyleName',
    'DocxTableHeaderParagraphStyle',
    'DocxTableBodyParagraphStyle',
    'DocxTableAutofitToWindow',
    'DocxEnableTitlePage',
    'CustomerName',
    'CustomerLocation',
    'ProjectName',
    'TitlePageFields',
    'TitlePageFieldsJson',
    'PromptForTitlePageFields'
)

$platformParameterNames = if ($Platform -eq 'StorageGRID') {
    @(
        'ExportCapacityDiagnostics',
        'CollectSantricityAppliances',
        'SantricityCredential',
        'SantricityAuthMapPath'
    )
}
else {
    @('Port', 'Ssid')
}

$invokeParameters = @{ WorkspaceRoot = $PSScriptRoot }
foreach ($parameterName in @($commonParameterNames + $platformParameterNames)) {
    if ($PSBoundParameters.ContainsKey($parameterName)) {
        $invokeParameters[$parameterName] = $PSBoundParameters[$parameterName]
    }
}

& $commandName @invokeParameters
