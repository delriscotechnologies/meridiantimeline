<h1 align="center" meridiantimeline
  
<p align="center">
A small Windows PowerShell 5.1 + WPF utility for building user-centered investigation timelines from TXT, CSV, and Excel logs.

The tool normalizes timestamps from multiple source time zones, correlates activity for one investigated identity and its exact aliases, and creates an Excel workbook that preserves supporting evidence alongside a concise timeline.

## Requirements

- Windows PowerShell 5.1
- Microsoft Excel desktop application
- Windows with WPF support
- Permission to read the selected source files and write the output workbook

The script runs locally. It does not upload logs, request credentials, or connect to a remote service.

## Quick Start

```powershell
git clone https://github.com/delriscotechnologies/meridiantimeline.git
cd meridiantimeline
powershell.exe -NoProfile -STA -File .\meridiantimeline.ps1
```

If your organization restricts PowerShell execution, follow its approved execution-policy and code-signing requirements.

## Workflow

1. Add one or more `.txt`, `.csv`, or `.xlsx` files.
2. Select the source timezone for each file.
3. Enter the investigated user ID or email address.
4. Add any exact aliases, one per line.
5. Optionally define a Central time range.
6. Select **Build Timeline** and choose where to save the workbook.

TXT and CSV inputs must contain a delimited table. Supported delimiters are tabs, commas, semicolons, and pipes.

## Source Profiles

| Profile | Purpose |
| --- | --- |
| Microsoft Entra | Recognizes common sign-in fields such as user, application, resource, status, IP address, and failure information |
| Windows Security | Recognizes common Windows event fields and explains supported event IDs 4624, 4625, 4634, 4648, 4672, and 4688 |
| Generic | Maps common timestamp, user, host, IP, event, action, outcome, and message column names |

Profile detection and event explanations are deterministic. `What Happened` describes supported source data; it is not a maliciousness verdict or a substitute for analyst review.

## Time Handling

Each file can be identified as UTC, Central, Eastern, Mountain, or Pacific. Timestamps with an explicit UTC marker or offset retain that explicit meaning. Other timestamps are interpreted using the selected source timezone.

Output time is converted to Central and labeled automatically as CST or CDT. Ambiguous or nonexistent daylight-saving timestamps are not guessed; they are recorded in **Import Issues**.

## Output Workbook

| Worksheet | Contents |
| --- | --- |
| Timeline | `Time`, `Source`, `Log`, `Time Elapsed Since Last Log`, and `What Happened` |
| Evidence | Normalized UTC and Central timestamps, identity, source details, event fields, explanation rule, source row, and raw event snapshot |
| Source Files | SHA-256, file size, source timezone, detected profile, row count, match count, and import status |
| Import Issues | Unsupported layouts, malformed data, invalid timestamps, limits, and other review items |

Exact duplicate events are consolidated only in **Timeline** and remain available in **Evidence**. If more than 60 distinct timeline moments remain, the workbook preserves all evidence and asks the analyst to narrow the time range.

## Operational Limits

- Maximum file size: 100 MB
- Maximum rows per file: 200,000
- Maximum detailed import issues: 1,000, followed by a suppression summary
- Maximum raw event snapshot stored in one Excel cell: 30,000 characters; the source file and SHA-256 remain available for verification

## Self-Test

Run the integrated parser, timezone, explanation, duplicate-grouping, and Excel formula-safety checks with:

```powershell
powershell.exe -NoProfile -STA -File .\meridiantimeline.ps1 -SelfTest
```

A successful run returns:

```text
SELF_TEST_OK
```

## Scope

meridiantimeline does not modify source files, determine whether activity is malicious, prove source authenticity, or replace a SIEM, forensic platform, or chain-of-custody process.

Investigation logs and generated workbooks may contain sensitive identity, device, network, and security information. Use the tool only with data and systems you are authorized to access.

See [SECURITY.md](SECURITY.md) for security and vulnerability-reporting guidance.

## License

Licensed under the [MIT License](LICENSE.md).
