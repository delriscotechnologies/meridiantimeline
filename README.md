<h1 align="center">Meridian Timeline</h1>

<p align="center">
  A local PowerShell utility for building investigation timelines from logs.
</p>

---

Meridian Timeline combines delimited TXT, CSV, and XLSX logs around one investigated identity and its defined aliases, normalizes timestamps, and creates a reviewable Excel timeline.

> Use Meridian Timeline only with logs and systems you are authorized to investigate.

## Install

Use Windows PowerShell 5.1 with Microsoft Excel desktop and permission to read the source files and write the output workbook.

```powershell
git clone https://github.com/delriscotechnologies/meridiantimeline.git
cd meridiantimeline
.\meridiantimeline.ps1
```

## What it does

1. Imports delimited TXT, CSV, and XLSX logs.
2. Matches one investigated identity and explicitly defined aliases.
3. Normalizes timestamps from supported source time zones.
4. Builds an Excel workbook for timeline and evidence review.

## Output

The workbook contains:

| Worksheet | Contents |
| --- | --- |
| Timeline | Ordered analyst-facing activity |
| Evidence | Supporting normalized event data |
| Source Files | Source metadata and SHA-256 hashes |
| Import Issues | Records that need review |

## Demo

1. Add one or more source files.
2. Select the timezone for each file.
3. Enter the investigated user ID or email.
4. Add any exact aliases.
5. Select Build Timeline and save the workbook.

## Scope and limits

- Runs locally and does not upload logs or request credentials.
- Does not modify source files.
- Supports files up to 100 MB and 200,000 rows per file; XLSX inputs are limited to 256 columns and 2,000,000 used cells.
- More than 60 distinct timeline moments triggers a review warning; no matching evidence is discarded.
- Windows event explanations distinguish subject and target accounts; missing roles stay unspecified. Explanations are not maliciousness verdicts.
- Source logs and generated workbooks may contain sensitive information.

See [SECURITY.md](SECURITY.md) for security guidance.

## License

Meridian Timeline is available under the [MIT License](LICENSE.md).
