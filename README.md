# MsgToRecurringICS

Convert multi-day, **all-day Microsoft Outlook `.msg` appointments** into a single **recurring `.ics` event** covering working days (Monday–Friday).

[![Licence: MIT](https://img.shields.io/badge/Licence-MIT-blue.svg)](LICENSE)
![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue)

**MsgToRecurringICS** is a standalone command-line utility written in Python. It reads the appointment's original dates and Outlook category directly from the binary MSG file, then creates an iCalendar appointment with the same time slot on each weekday. It does not need Microsoft Outlook, COM automation, an internet connection or any third-party Python packages.

## Features

- **One recurring event, not multiple appointments:** produces a `VCALENDAR` containing one `VEVENT` with an `RRULE`.
- **Weekdays only:** skips Saturdays and Sundays, even when they fall inside a multi-day event.
- **Configurable daily hours:** defaults to **08:30–16:30**; supports `--start` and `--end` overrides.
- **Subject from the original appointment:** extracts the leading identifier (for example, `TRAIN-761`).
- **Suffix from Outlook categories:** category `Ewald Müller` becomes `(Ewald)`; without a category, no suffix is added.
- **Keeps the text description:** preserves the original plain-text description, including any links.
- **Standalone, read-only input processing:** never modifies the MSG file or your calendar.
- **Batch processing:** handles individual MSG files, wildcard patterns or folders.
- **Safe output:** will not overwrite existing ICS files without `--overwrite`; supports `--dry-run`.

## Requirements

- Python **3.10 or newer**
- A binary Outlook `.msg` file saved from an **all-day calendar appointment**
- An event using the Outlook time zone **`W. Europe Standard Time`** (Central European time, including daylight saving time)

No `pip install` step is necessary. The converter uses the Python standard library only and runs without Outlook.

## Getting started

1. Download or clone this repository.
2. Save one or more Outlook all-day appointment invitations as `.msg` files.
3. Open a terminal in the repository directory and run the script.
4. Import the resulting `.ics` file into your calendar application.

**Windows (PowerShell):**

```powershell
py MsgToRecurringICS.py 'C:\Invites\*.msg' --output-dir 'C:\Appointments'
```

**One MSG file:**

```powershell
py MsgToRecurringICS.py 'C:\Invites\course-invitation.msg'
```

**A whole folder (non-recursive):**

```powershell
py MsgToRecurringICS.py 'C:\Invites' --output-dir 'C:\Appointments'
```

**Set different daily working hours:**

```powershell
py MsgToRecurringICS.py 'C:\Invites\*.msg' --start 09:00 --end 17:00
```

**Check the result without writing any files:**

```powershell
py MsgToRecurringICS.py 'C:\Invites\*.msg' --dry-run
```

**Allow overwriting previously generated ICS files:**

```powershell
py MsgToRecurringICS.py 'C:\Invites\*.msg' --overwrite
```

On macOS or Linux, use `python3` instead of `py`, and adapt the file paths. The input files must still meet the Outlook MSG and time-zone restrictions described below.

To see all command-line options:

```powershell
py MsgToRecurringICS.py --help
```

## How conversion works

Suppose an all-day invitation covers **Friday 16 April through Tuesday 20 April 2027**, inclusive. With the default hours, the output has three occurrences:

| Date | Time |
| --- | --- |
| Friday 16 April | 08:30–16:30 |
| Monday 19 April | 08:30–16:30 |
| Tuesday 20 April | 08:30–16:30 |

The generated ICS recurrence is:

```ics
RRULE:FREQ=DAILY;BYDAY=MO,TU,WE,TH,FR;COUNT=3
```

The recurrence count equals the number of weekdays within the original inclusive date range. This means the series does not continue beyond the source appointment's last weekday. **Public holidays are not excluded.**

### Subject and file name

The subject is the leading appointment identifier, when present, followed by the **first word of the first Outlook category** in parentheses.

| Original subject | First Outlook category | Output subject / filename |
| --- | --- | --- |
| `DEMO-101: Introductory course` | `Alex Taylor` | `DEMO-101 (Alex).ics` |
| `DEMO-101: Introductory course` | *None* | `DEMO-101.ics` |
| `DEMO-101: Introductory course` | `Training / Consulting` | `DEMO-101 (Training).ics` |

> **Note:** The script does not infer a person's name from the appointment organiser or attendees. It always uses the first word of the first Outlook category. Consequently, activity-based categories such as `Training / Consulting` produce an activity-based suffix. If the suffix is incorrect, inspect the category assigned to the original Outlook appointment.

If no leading identifier is found, the script falls back to the portion of the original subject before its first colon.

### Other output properties

| Property | Output behaviour |
| --- | --- |
| Event start/end | From the MSG appointment's all-day date range (exclusive end), converted from UTC to local dates |
| Daily time | `08:30–16:30` by default, or values passed via CLI |
| Recurrence | One daily series restricted to Monday–Friday, ending within the original date range |
| Description | Plain-text MSG description, correctly escaped for iCalendar |
| Visibility and availability | Public, confirmed, busy |
| Time zone | `W. Europe Standard Time` with `VTIMEZONE` definitions |
| Attendees / organiser | Not included; output is a personal calendar entry, not a meeting invitation |
| Location and reminders | Omitted |
| File name | Derived from the output subject, with invalid filesystem characters replaced |
| Event UID | Deterministic, based on the original Outlook GlobalObjectId when available |
| Destination | Next to the source MSG by default, or `--output-dir` |

## Compatibility and limitations

This tool intentionally handles a **specific Outlook use case** rather than being a general-purpose MSG converter.

- **Only all-day appointments** with the MSG message class `IPM.Appointment` are accepted. Timed appointments, e-mails and other MSG types are rejected.
- **Only `W. Europe Standard Time` is supported.** The built-in Central European daylight-saving calculations apply from 1996 onwards. The converter does not guess dates when time-zone metadata is missing or unsupported.
- **Weekends are skipped; public holidays are not.** Exceptions, changed individual occurrences and alternative daily schedules are not supported.
- The **first Outlook category** supplies the suffix; the script does not detect which category denotes a person.
- The original meeting's **organiser, attendees, location, reminders and HTML description** are not reproduced.
- **ICS import is a calendar-client operation**, not part of this script. Some calendar applications may create duplicates on repeated imports, even when the generated UID is stable.
- The converter supports the relevant Microsoft Compound File / MAPI structures used by the tested Outlook samples; it is not a universal parser for all MSG variants.

Incorrect or unsupported MSG content produces an error rather than a guessed conversion. Each file is processed independently; the command returns a non-zero exit status if any conversion fails.

## Examples

The [`examples/`](examples/) directory contains **synthetic ICS appointments** generated with the converter's output builder to illustrate two-day, three-day and four-day weekday recurrences. They contain **no real meeting details, internal URLs or original invitation data**.

The original three reference pairs were used for development and comparison, but the source `.msg` files and actual meeting exports are **not distributed** in this public repository.

## Testing

Run the automated tests:

```powershell
py -m unittest -v test_MsgToRecurringICS.py
```

The test suite checks calendar-day filtering, time-zone conversion, iCalendar escaping and line folding, custom hours, UID stability, safe overwriting and the original reference pairs. Tests that require private sample MSG/ICS files will be **skipped** unless those fixtures are supplied locally.

To run the optional reference-pair integration tests, set `MSG_TEST_FIXTURES` to a local directory containing the original private reference files:

```powershell
$env:MSG_TEST_FIXTURES = 'C:\PrivateFixtures'
py -m unittest -v test_MsgToRecurringICS.py
```

Never commit your private fixture files to a public repository.

## Privacy and security

MSG invitations can contain personal details, private calendar information and internal links. Review generated ICS files before sharing them. Processing is local: the script does not upload data or make network requests.

A [`.gitignore`](.gitignore) is included to discourage accidental commits of raw `.msg` files, generated ICS files outside `examples/`, temporary files and Python caches. Always review the files you stage for a public commit.

## Licence

Licensed under the **MIT Licence**. See [LICENSE](LICENSE) for the full terms.

## Authors

- **Ewald** — project author, GitHub repository owner and maintainer.
- **ChatGPT (OpenAI)** — AI-assisted co-author for implementation and documentation.

The MIT copyright notice names Ewald as the copyright holder; ChatGPT is credited for its contribution to developing the project, not as a separate legal copyright holder.
