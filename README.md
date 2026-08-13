# PowerShell Graph Calendar Diagnostic

`Get-CalendarDiagnostic.ps1` uses delegated Microsoft Graph authentication to inspect
the signed-in user's calendar. It does not require an application registration.

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
.\Get-CalendarDiagnostic.ps1 -OutputPath .\CalendarDiagnostic
```

Sign in with the device-code prompt. The script detects recurring meetings with no
end, meetings ending more than one year in the future, duplicate iCalendar UIDs,
recurrence ranges above 100 occurrences, and series with more than 20 exceptions.
Use `-HighRecurrenceThreshold` and `-ExceptionThreshold` to change those limits.

The output directory contains an HTML summary with event details, `CalendarEvents.csv`,
and one CSV for each issue category.
