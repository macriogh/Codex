# Outlook Availability Summarizer (Windows App)

A dedicated **native Windows desktop app** (WinForms + .NET) that connects to your locally installed Microsoft Outlook and produces a plain-English summary of your availability.

## Features

- Choose a date range (`from` and `to`).
- Choose a daily timeframe window (for example, 9:00 AM to 5:00 PM).
- Read Outlook calendar appointments (including recurring meetings).
- Summarize each day in plain English with:
  - total free and busy time
  - free blocks (time ranges)
- Provide an overall total for the selected date range.

## Requirements

- Windows 10/11
- .NET 8 SDK/runtime (or publish self-contained)
- Microsoft Outlook desktop app installed and configured with your account

## Run

```powershell
dotnet restore
dotnet run
```

## Build a distributable executable

```powershell
dotnet publish -c Release -r win-x64 --self-contained false
```

Published files will be in:

`bin\Release\net8.0-windows\win-x64\publish\`

## Notes

- This is **not** an Outlook plugin/add-in; it is a standalone app that automates Outlook through the COM interop API.
- Depending on Outlook security policy, first-time access may prompt for trust/permissions.
