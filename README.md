# Outlook Availability Summarizer (Windows App)

A dedicated **native Windows desktop app** (WinForms + .NET) that summarizes your calendar availability in plain English.

## Outlook compatibility

The app supports two sources:

1. **Classic Outlook (COM)**
   - Works with classic desktop Outlook profiles.
2. **New Outlook / Microsoft 365 (Graph API)**
   - Recommended for versions like **1.2026.120.300** where COM automation may not be available.
   - Uses Microsoft Graph `calendarView` and delegated sign-in.

## Features

- Choose a date range (`from` and `to`).
- Choose a daily timeframe window (for example, 9:00 AM to 5:00 PM).
- Choose calendar source (Classic Outlook COM or Graph API).
- Summarize each day in plain English with:
  - total free and busy time
  - free blocks (time ranges)
- Provide an overall total for the selected date range.

## Requirements

- Windows 10/11
- .NET 8 SDK/runtime (or publish self-contained)
- For **Classic Outlook mode**: Microsoft Outlook classic desktop installed/configured.
- For **Graph mode (new Outlook)**:
  - Microsoft 365 account
  - Azure/Entra app registration with delegated permission `Calendars.Read`
  - Client ID from that app registration

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

## First-time setup for Graph mode

1. In Microsoft Entra admin center, create an app registration.
2. Add delegated API permission: `Microsoft Graph -> Calendars.Read`.
3. For public client flows, enable mobile/desktop flow.
4. Copy the app's **Client ID**.
5. In this app, choose **New Outlook / Microsoft 365 (Graph API)** and paste Client ID.
6. Keep Tenant ID as `common` unless your admin requires a specific tenant.

## Notes

- This is **not** an Outlook plugin/add-in; it is a standalone app.
- In Graph mode, sign-in uses device code flow and a browser confirmation.
