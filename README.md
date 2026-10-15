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
.\collect_asbuilt.ps1 -Platform StorageGRID -Target "https://sg-admin-node.example.com:443"
.\collect_asbuilt.ps1 -Platform ESeries -Target "https://santricity.example.com:8443"
```

Include the management HTTPS port in every live `-Target`. StorageGRID commonly uses port 443; E-Series/SANtricity commonly uses port 8443:

```powershell
.\collect_asbuilt.ps1 -Platform StorageGRID -Target "https://sg-admin-node.example.com:443"
.\collect_asbuilt.ps1 -Platform ESeries -Target "https://santricity.example.com:8443"
```

For custom ports, place the port directly in the target URL:

```powershell
.\collect_asbuilt.ps1 -Platform StorageGRID -Target "https://sg-admin-node.example.com:8443"
.\collect_asbuilt.ps1 -Platform ESeries -Target "https://santricity.example.com:9443"
```

Use `-Credential` with a `PSCredential` or enter credentials at the interactive prompt. Password string parameters are not supported.

### Verbose Diagnostics

Use PowerShell's common `-Verbose` parameter when troubleshooting a collection:

```powershell
.\collect_asbuilt.ps1 -Platform StorageGRID -Target "https://sg-admin-node.example.com:443" -Verbose
```

Normal output shows authentication, high-level collection progress, report generation, output paths, and actionable warnings or failures. Detailed API and appliance-attribute diagnostics, including requests for data that is unavailable on a particular node, are suppressed by default and are shown with `-Verbose`.

## Output Reports

The output as-built documents are created in the locations below.

- StorageGRID: `.\asbuilt_output_storagegrid`
- E-Series: `.\asbuilt_output_eseries`

By default, successful report generation removes intermediate JSON, Markdown, and collection-summary files, leaving the DOCX reports. Use `-KeepIntermediateOutputs` to retain these generated files alongside the DOCX reports.

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
.\collect_asbuilt.ps1 -Platform StorageGRID -Target "https://sg-admin-node.example.com:443" -CollectSantricityAppliances -SantricityCredential $credential
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
    -Target "https://sg-admin-node.example.com:443" `
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

Each successful appliance produces a DOCX report. JSON and Markdown intermediates are removed by default and retained with `-KeepIntermediateOutputs`.

With `-KeepIntermediateOutputs`, the directory also retains `santricity-collection-summary.json`, including node name, candidate IPs, selected IP, status, credential-map key, and error details. A failed appliance does not stop the remaining collections.

## Parameters

### Common

- `-Target`: host:port or URL with an explicit port, such as `https://sg-admin-node.example.com:443` or `https://santricity.example.com:8443`.
- `-Credential`: secure `PSCredential` for live collection.
- TLS certificate validation is enabled by default. To bypass it for a collection using a known untrusted or self-signed certificate, explicitly pass `-ValidateCerts:$false`. This disables certificate verification and is not recommended outside a trusted network.
- `-OutputDir`: output directory.
- `-ReportPrefix`: prefix generated filenames.
- `-KeepIntermediateOutputs`: retain generated JSON, Markdown, and collection summaries after DOCX generation. They are removed after successful report generation by default.

### StorageGRID

- `-CollectSantricityAppliances`: tri-state appliance collection control.
- `-NonInteractive`: suppress prompts.
- `-SantricityCredential`: shared SANtricity credential.
- `-SantricityAuthMapPath`: CLIXML per-appliance credential map.

### E-Series

The DOCX layout and styles are fixed. Optional title-page values can be supplied with `-CustomerName`, `-CustomerLocation`, and `-ProjectName`. Use `-PromptForTitlePageFields` to prompt for any of those values not supplied as parameters.

## Support Files

These files are required for successful creation of the as-built reports. Do not modify these files unless instructed.

- `docx\reference-template.docx` - template input Word Document to supply the title page and styles.
- `docx\report-metadata-storagegrid.yml` - defaults for the report metadata for StorageGRID As-Built reports
- `docx\report-metadata-e-series.yml` - defaults for the report metadata for E-Series As-Built reports
- `Lib\OpenXml\Desktop\` and `Lib\OpenXml\Core\` - the DocumentFormat.OpenXml SDK assemblies used to generate DOCX reports (see `Lib\OpenXml\OPENXML-SDK-NOTICE.txt`). No installation, PSGallery access, or admin rights are required; the script loads them in-process with `Add-Type`.
