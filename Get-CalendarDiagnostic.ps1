[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PWD 'CalendarDiagnostic'),
    [ValidateRange(1, [int]::MaxValue)]
    [int]$HighRecurrenceThreshold = 100,
    [ValidateRange(1, [int]::MaxValue)]
    [int]$ExceptionThreshold = 20
)

function Get-GraphPagedCollection {
    param([Parameter(Mandatory)][string]$Uri)

    $items = [System.Collections.Generic.List[object]]::new()
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        foreach ($item in @($response.value)) {
            $items.Add($item)
        }
        $Uri = $response.'@odata.nextLink'
    } while ($Uri)

    return $items
}

function Get-CalendarDiagnostic {
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [object[]]$Events,
        [datetime]$AsOf = (Get-Date),
        [ValidateRange(1, [int]::MaxValue)]
        [int]$HighRecurrenceThreshold = 100,
        [ValidateRange(1, [int]::MaxValue)]
        [int]$ExceptionThreshold = 20
    )

    $issues = @{
        UnendingMeetings = [System.Collections.Generic.List[object]]::new()
        LongRunningMeetings = [System.Collections.Generic.List[object]]::new()
        DuplicateUids = [System.Collections.Generic.List[object]]::new()
        HighRecurrenceCounts = [System.Collections.Generic.List[object]]::new()
        ExcessiveExceptions = [System.Collections.Generic.List[object]]::new()
    }
    $issueIds = @{}
    $details = foreach ($calendarEvent in $Events) {
        [pscustomobject]@{
            Id = $calendarEvent.id
            Subject = $calendarEvent.subject
            Type = $calendarEvent.type
            ICalUId = $calendarEvent.iCalUId
            SeriesMasterId = $calendarEvent.seriesMasterId
            Start = $calendarEvent.start.dateTime
            End = $calendarEvent.end.dateTime
            RecurrenceRangeType = $calendarEvent.recurrence.range.type
            RecurrenceEndDate = $calendarEvent.recurrence.range.endDate
            NumberOfOccurrences = $calendarEvent.recurrence.range.numberOfOccurrences
            DiagnosticIssues = ''
        }
    }
    $detailById = @{}
    foreach ($detail in $details) {
        $detailById[$detail.Id] = $detail
    }

    $masters = @($Events | Where-Object { $_.type -in @('seriesMaster', 'singleInstance') })
    foreach ($calendarEvent in $masters | Where-Object { $_.recurrence.range.type -eq 'noEnd' }) {
        $issues.UnendingMeetings.Add($calendarEvent)
        $issueIds[$calendarEvent.id] = @($issueIds[$calendarEvent.id]) + 'Unending meeting'
    }

    $oneYearFromNow = $AsOf.Date.AddYears(1)
    foreach ($calendarEvent in $masters) {
        [datetime]$endDate = [datetime]::MinValue
        if ($calendarEvent.recurrence.range.endDate) {
            [datetime]::TryParse($calendarEvent.recurrence.range.endDate, [ref]$endDate) | Out-Null
        }
        if ($endDate -ne [datetime]::MinValue -and $endDate.Date -gt $oneYearFromNow) {
            $issues.LongRunningMeetings.Add($calendarEvent)
            $issueIds[$calendarEvent.id] = @($issueIds[$calendarEvent.id]) + 'Ends more than one year out'
        }
    }

    foreach ($group in $masters | Where-Object { -not [string]::IsNullOrWhiteSpace($_.iCalUId) } | Group-Object iCalUId | Where-Object Count -gt 1) {
        foreach ($calendarEvent in $group.Group) {
            $issues.DuplicateUids.Add($calendarEvent)
            $issueIds[$calendarEvent.id] = @($issueIds[$calendarEvent.id]) + 'Duplicate UID'
        }
    }

    foreach ($calendarEvent in $masters | Where-Object { $_.recurrence.range.numberOfOccurrences -gt $HighRecurrenceThreshold }) {
        $issues.HighRecurrenceCounts.Add($calendarEvent)
        $issueIds[$calendarEvent.id] = @($issueIds[$calendarEvent.id]) + "More than $HighRecurrenceThreshold occurrences"
    }

    foreach ($group in $Events | Where-Object { $_.type -eq 'exception' -and $_.seriesMasterId } | Group-Object seriesMasterId | Where-Object Count -gt $ExceptionThreshold) {
        $master = $Events | Where-Object id -eq $group.Name | Select-Object -First 1
        $record = if ($master) { $master } else { $group.Group | Select-Object -First 1 }
        $record | Add-Member -NotePropertyName ExceptionCount -NotePropertyValue $group.Count -Force
        $issues.ExcessiveExceptions.Add($record)
        $issueIds[$record.id] = @($issueIds[$record.id]) + "More than $ExceptionThreshold exceptions ($($group.Count))"
    }

    foreach ($id in $issueIds.Keys) {
        if ($detailById.ContainsKey($id)) {
            $detailById[$id].DiagnosticIssues = $issueIds[$id] -join '; '
        }
    }

    [pscustomobject]@{
        Details = @($details)
        Issues = $issues
        Summary = [pscustomobject]@{
            EventsScanned = @($Events).Count
            UnendingMeetings = $issues.UnendingMeetings.Count
            LongRunningMeetings = $issues.LongRunningMeetings.Count
            DuplicateUidEvents = $issues.DuplicateUids.Count
            HighRecurrenceCounts = $issues.HighRecurrenceCounts.Count
            ExcessiveExceptions = $issues.ExcessiveExceptions.Count
        }
    }
}

function Export-CalendarDiagnostic {
    param(
        [Parameter(Mandatory)][pscustomobject]$Diagnostic,
        [Parameter(Mandatory)][string]$Path
    )

    New-Item -ItemType Directory -Path $Path -Force | Out-Null
    $Diagnostic.Details | Export-Csv -Path (Join-Path $Path 'CalendarEvents.csv') -NoTypeInformation -Encoding utf8
    $fileNames = @{
        UnendingMeetings = 'UnendingMeetings.csv'
        LongRunningMeetings = 'LongRunningMeetings.csv'
        DuplicateUids = 'DuplicateUids.csv'
        HighRecurrenceCounts = 'HighRecurrenceCounts.csv'
        ExcessiveExceptions = 'ExcessiveExceptions.csv'
    }
    foreach ($issueName in $fileNames.Keys) {
        @($Diagnostic.Issues[$issueName]) | Select-Object id, subject, type, iCalUId, seriesMasterId, recurrence, ExceptionCount |
            Export-Csv -Path (Join-Path $Path $fileNames[$issueName]) -NoTypeInformation -Encoding utf8
    }

    $summaryHtml = $Diagnostic.Summary | ConvertTo-Html -Fragment
    $detailsHtml = $Diagnostic.Details | ConvertTo-Html -Fragment
    @"
<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Calendar Diagnostic Report</title></head>
<body><h1>Calendar Diagnostic Report</h1><h2>Summary</h2>$summaryHtml<h2>Event details</h2>$detailsHtml</body></html>
"@ | Set-Content -Path (Join-Path $Path 'CalendarDiagnosticReport.html') -Encoding utf8
}

if ($MyInvocation.InvocationName -ne '.') {
    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        throw 'Microsoft Graph PowerShell is required. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser'
    }

    Connect-MgGraph -Scopes @('User.Read', 'Calendars.Read') -UseDeviceAuthentication -NoWelcome
    try {
        $select = 'id,subject,type,iCalUId,seriesMasterId,start,end,recurrence'
        $events = Get-GraphPagedCollection -Uri "https://graph.microsoft.com/v1.0/me/events?`$select=$select&`$top=1000"
        $diagnostic = Get-CalendarDiagnostic -Events $events -HighRecurrenceThreshold $HighRecurrenceThreshold -ExceptionThreshold $ExceptionThreshold
        Export-CalendarDiagnostic -Diagnostic $diagnostic -Path $OutputPath
        $diagnostic.Summary | Format-List
        Write-Information "Report and CSV files written to $OutputPath" -InformationAction Continue
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}
