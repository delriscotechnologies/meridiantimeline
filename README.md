<h1 align="center">Meridian Timeline</h1>

<p align="center">
  A local PowerShell utility for building user-centered investigation timelines from TXT, CSV, and Excel logs.
</p>

---

meridiantimeline is one Windows PowerShell 5.1 script with a WPF interface for organizing activity from multiple log sources around one investigated identity and its explicitly defined aliases.

Each successful run normalizes timestamps across source time zones and creates one Excel workbook containing a concise event timeline, the supporting evidence, source-file metadata, and import issues for analyst review.

> Use meridiantimeline only with logs and systems you are authorized to investigate. Source files and generated workbooks may contain sensitive identity, device, network, and security information.

## Quick Start

You need a Windows host with Windows PowerShell 5.1, WPF support, Microsoft Excel desktop, permission to read the selected source files, and permission to write the output workbook.

```powershell
git clone https://github.com/delriscotechnologies/meridiantimeline.git
cd meridiantimeline
powershell.exe -NoProfile -STA -File .\meridiantimeline.ps1
```

If your organization restricts PowerShell execution, follow its approved execution-policy and code-signing requirements.

## How It Works

1. Add one or more `.txt`, `.csv`, or `.xlsx` files.
2. Select the source timezone for each file.
3. Enter the investigated user ID or email address.
4. Add any exact aliases, one per line.
5. Optionally define a Central time range.
6. Select **Build Timeline** and choose where to save the workbook.

TXT and CSV inputs must contain a delimited table. Supported delimiters are tabs, commas, semicolons, and pipes.

## What You Get

The generated workbook separates the analyst-facing sequence from the records needed to trace and review it.

| Worksheet | Contents |
| --- | --- |
| Timeline | `Time`, `Source`, `Log`, `Time Elapsed Since Last Log`, and `What Happened` |
| Evidence | Normalized UTC and Central timestamps, identity, source details, event fields, explanation rule, source row, and raw event snapshot |
| Source Files | SHA-256, file size, source timezone, detected profile, row count, match count, and import status |
| Import Issues | Unsupported layouts, malformed data, invalid timestamps, limits, and other review items |

Exact duplicate events are consolidated only in **Timeline** and remain available in **Evidence**. If more than 60 distinct timeline moments remain, the workbook preserves all evidence and asks the analyst to narrow the time range.

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

## Operational Limits

| Limit | Value |
| --- | --- |
| File size | 100 MB per file |
| Rows | 200,000 per file |
| Detailed import issues | 1,000, followed by a suppression summary |
| Raw event snapshot | 30,000 characters per Excel cell; the source file and SHA-256 remain available for verification |

## Safety and Privacy

- Runs locally and does not upload logs, request credentials, or connect to a remote service.
- Does not modify source files.
- Does not determine whether activity is malicious or prove source authenticity.
- Does not replace a SIEM, forensic platform, or chain-of-custody process.
- Keeps supporting evidence and source-file hashes available for analyst review.

Store source logs and generated workbooks according to your organization's access-control, retention, and evidence-handling requirements.

See [SECURITY.md](SECURITY.md) for security and vulnerability-reporting guidance.

## License

meridiantimeline is available under the [MIT License](LICENSE.md).
