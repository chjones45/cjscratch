# E-Series and StorageGRID As-Built Collector

`collect_asbuilt.ps1` is the supported entry point for both StorageGRID and E-Series/SANtricity as-built reports. It runs on Windows PowerShell 5.1+ and PowerShell 7.

## Requirements

- Network access to the selected StorageGRID or SANtricity Web Services endpoint.
- A read-capable account for the selected platform.
- Pandoc and the supporting files in `docx\` for DOCX generation.
- Run from the repository directory so relative paths resolve correctly.
- Tested with pandoc.exe 3.1.3, available at https://github.com/jgm/pandoc/releases/tag/3.1.3. Download manually and place into the .\docx folder

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
2. Generate the StorageGRID Markdown and DOCX reports.
3. Ask whether SANtricity reports should be generated.
4. Ask whether one credential should be reused or credentials should be entered per SANtricty login.
5. Collect each appliance independently and continue after failures.

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

Both platforms support DOCX formatting and title-page parameters including `-DocxDir`, `-EnableDocxToc`, `-DocxEnableTitlePage`, `-CustomerName`, `-CustomerLocation`, and `-ProjectName`.

## Output Defaults

- StorageGRID: `.\asbuilt_output_storagegrid`
- E-Series: `.\asbuilt_output_eseries`
- DOCX support files: `.\docx`

Use `-KeepIntermediateOutputs` when troubleshooting or retaining source JSON/Markdown. Otherwise successful reports clean intermediate files according to the effective cleanup setting.

## Support Files

- `docx\reference-template.docx`
- `docx\pandoc-apply-table-style.lua`
- `docx\report-metadata-storagegrid.yml`
- `docx\report-metadata-e-series.yml`
- `docx\PANDOC-REDISTRIBUTION-NOTICE.txt`
