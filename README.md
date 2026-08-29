# E-Series and StorageGRID As-Built Collector

`collect_asbuilt.ps1` is the supported entry point for both StorageGRID and E-Series/SANtricity as-built reports. It runs on Windows PowerShell 5.1+ and PowerShell 7.

## Requirements

- Network access to the selected StorageGRID or SANtricity Web Services endpoint.
- A read-capable account for the selected platform.
- The `Lib\OpenXml` folder distributed alongside this script (DOCX generation uses the bundled DocumentFormat.OpenXml SDK directly; no Pandoc, no Microsoft Word, and no admin rights are required).
- Run from the repository directory so relative paths resolve correctly.

## Unblock Scripts

PowerShell may block script files from running and require approval multiple times as every script and input file is accessed. You can unblock the files with the following command:

```powershell
Get-ChildItem -Path ".\" -Recurse | Unblock-File
```

## Basic Usage

```powershell
.\collect_asbuilt.ps1 -Platform StorageGRID -Target "sg-admin-node.example.com"
.\collect_asbuilt.ps1 -Platform ESeries -Target "santricity.example.com"
```

For StorageGRID, a bare hostname or an HTTPS URL without a port uses the default Grid Manager HTTPS/TCP port 443:

```powershell
.\collect_asbuilt.ps1 -Platform StorageGRID -Target "sg-admin-node.example.com"
.\collect_asbuilt.ps1 -Platform StorageGRID -Target "https://sg-admin-node.example.com"
```

If the Grid Manager uses a custom port, include that port in `-Target`:

```powershell
.\collect_asbuilt.ps1 -Platform StorageGRID -Target "https://sg-admin-node.example.com:8443"
```

StorageGRID does not use the `-Port` parameter; custom StorageGRID ports belong in the target URL. E-Series defaults to port 8443 and supports the E-Series-specific `-Port` parameter.

Use `-Credential` for secure authentication; otherwise live collection prompts for credentials.

### Verbose Diagnostics

Use PowerShell's common `-Verbose` parameter when troubleshooting a collection:

```powershell
.\collect_asbuilt.ps1 -Platform StorageGRID -Target "sg-admin-node.example.com" -Verbose
```

Normal output shows authentication, high-level collection progress, report generation, output paths, and actionable warnings or failures. Detailed API and appliance-attribute diagnostics, including requests for data that is unavailable on a particular node, are suppressed by default and are shown with `-Verbose`.

## Output Reports

The output as-built documents are created in the locations below.

- StorageGRID: `.\asbuilt_output_storagegrid`
- E-Series: `.\asbuilt_output_eseries`

Use `-KeepIntermediateOutputs` when troubleshooting or retaining source JSON/Markdown. Otherwise successful reports clean intermediate files according to the effective cleanup setting.

Once generated, open the report(s) and refresh the table of contents (TOC) to complete it's generation. A placeholder TOC is generated that must be manually updated and does not automatically refresh. This decision was made as refreshing the TOC triggers a safety warning in Microsoft Word that may cause confusion when opening the report for the first time (this is a default Word action in these circumstances and not due to safety concerns with the reports).

## StorageGRID SANtricity Appliances

StorageGRID collection can identify eligible appliance nodes with embedded E-Series/SANtricity controllers and also genenerate individual as-built documents for eligibe nodes.

`-CollectSantricityAppliances` is tri-state:

| Invocation | Behavior |
| --- | --- |
| Omitted | Detect appliances, generate the StorageGRID report, then ask whether to collect SANtricity reports. If accepted, ask whether to use one shared credential or individual credentials per appliance. |
| `-CollectSantricityAppliances` | Detect and collect without the collection confirmation. Interactive runs still ask shared versus individual credentials unless credentials are supplied. |
| `-CollectSantricityAppliances:$false` | Skip SANtricity detection and collection completely. |
| `-NonInteractive` | Suppress prompts. Omitted collection mode detects and skips with a warning; forced collection requires credentials. |

The interactive workflow is:

1. Detect eligible appliances.
2. Generate and finalize the StorageGRID Markdown and DOCX reports.
3. Announce the completed StorageGRID DOCX path.
4. Ask whether SANtricity reports should be generated.
5. Ask whether one credential should be reused or credentials should be entered per SANtricity login.
6. Collect each appliance independently and continue after failures.

Individual prompts identify each node and show progress:

```text
SANTRICITY CREDENTIALS [1 of 20]
Enter credentials for sg-storage-node-01
```

Credentials entered interactively remain in memory for the run and are not written to reports or JSON.

### Shared Credential

```powershell
$credential = Get-Credential -Message "SANtricity credentials"
.\collect_asbuilt.ps1 -Platform StorageGRID -Target "sg-admin-node.example.com" -CollectSantricityAppliances -SantricityCredential $credential
```

### Per-Appliance CLIXML Map

Create a Windows-protected map keyed by node name or controller IP:

```powershell
$map = @{}
$map['sg-storage-node-01'] = Get-Credential
$map['10.84.5.17'] = Get-Credential
$map | Export-Clixml .\santricity-auth-map.clixml
```

Use it with unattended forced collection:

```powershell
.\collect_asbuilt.ps1 `
    -Platform StorageGRID `
    -Target "sg-admin-node.example.com" `
    -CollectSantricityAppliances `
    -NonInteractive `
    -SantricityAuthMapPath .\santricity-auth-map.clixml
```

The map is checked by node name first, then controller IP. CLIXML credentials are protected for the Windows user/profile that created the file.

### SANtricity Outputs

Per-appliance reports are written below:

```text
.\asbuilt_output_storagegrid\<grid-name>_santricity_appliances\
```

Each successful appliance produces JSON, Markdown, and DOCX files (only the DOCX file is retained unless the -KeepIntermediateOutputs attributes is provided).

The directory also contains `santricity-collection-summary.json`, including node name, candidate IPs, selected IP, status, credential-map key, and error details. A failed appliance does not stop the remaining collections.

## Parameters

### Common

- `-Target`: host, host:port, or URL.
- `-Credential`: secure `PSCredential` for live collection.
- `-ValidateCerts`: enable normal TLS certificate validation.
- `-UseSystemProxy`: use the system proxy.
- `-OutputDir`: output directory.
- `-CollectionOnly`: write collection JSON and stop before report rendering.
- `-JsonInputPath`: replay an existing platform JSON export without API calls.
- `-ReportPrefix`: prefix generated filenames.
- `-KeepIntermediateOutputs`: retain JSON and Markdown sources after DOCX generation. These files are removed after execution by default.
- `-CleanupIntermediateOutputs`: cleanup is enabled by effective default unless explicitly bound.

### StorageGRID

- `-ExportCapacityDiagnostics`: write focused capacity diagnostics JSON and CSV.
- `-CollectSantricityAppliances`: tri-state appliance collection control.
- `-NonInteractive`: suppress prompts.
- `-SantricityCredential`: shared SANtricity credential.
- `-SantricityAuthMapPath`: CLIXML per-appliance credential map.

### E-Series

- `-Port`: default port for host-only targets, normally `8443`.
- `-Ssid`: storage-system ID; auto-detected when possible, otherwise required interactively.

Both platforms support DOCX formatting and title-page parameters including `-EnableDocxToc`, `-DocxEnableTitlePage`, `-CustomerName`, `-CustomerLocation`, and `-ProjectName`.

## Support Files

These files are required for successful creation of the as-built reports. Do not modify these files unless instructed.

- `docx\reference-template.docx` - template input Word Document to supply the title page and styles.
- `docx\report-metadata-storagegrid.yml` - defaults for the report metadata for StorageGRID As-Built reports
- `docx\report-metadata-e-series.yml` - defaults for the report metadata for E-Series As-Built reports
- `Lib\OpenXml\Desktop\` and `Lib\OpenXml\Core\` - the DocumentFormat.OpenXml SDK assemblies used to generate DOCX reports (see `Lib\OpenXml\OPENXML-SDK-NOTICE.txt`). No installation, PSGallery access, or admin rights are required; the script loads them in-process with `Add-Type`.
