# PowerShell-Graph-Calendar-Diagnostic-Delegate.ps1 

<#
.SYNOPSIS
Reports calendar master records and recurrence exceptions in a specified time period.

.DESCRIPTION
Uses Microsoft Graph delegated OAuth to read a user's default calendar. The sign-in
page opens with the device code prefilled, so the user only completes interactive
sign-in. The calendarView endpoint expands recurring series for the requested period.
The script then groups occurrences and exceptions by their series master ID and treats
each non-recurring event as  its own master record.

You do not need to register an application or install any PowerShell modules. The script u
ses Microsoft's public client application for delegated sign-in.	

The text report contains per-master details and totals. Separate CSV reports list
recurring masters with more than 20 modified exceptions and masters longer than one year.

.PARAMETER UserGuid
User principal name or object GUID of the mailbox to inspect. Defaults to "me", which
uses the mailbox of the account that completes delegated sign-in.

.PARAMETER StartDateTime
Beginning of the calendar range as an ISO 8601 value with a UTC offset or Z suffix.

.PARAMETER EndDateTime
End of the calendar range as an ISO 8601 value with a UTC offset or Z suffix.

.PARAMETER OutputDirectory
Directory where the text and CSV reports are written. Defaults to the script directory.

.EXAMPLE
pwsh -File .\PowerShell-Graph-Calendar-Diagnostic-Delegate.ps1

.EXAMPLE
powershell.exe -NoProfile -File .\PowerShell-Graph-Calendar-Diagnostic-Delegate.ps1 -UserGuid user@contoso.com -StartDateTime '2025-01-01T00:00:00Z' -EndDateTime '2026-01-01T00:00:00Z' -OutputDirectory C:\Reports

.REQUIREMENTS
- The signed-in user must have access to the target mailbox calendar.
- Calendars.Read and Calendars.Read.Shared delegated consent must be permitted by the tenant.
- Windows PowerShell 5.1 or PowerShell 7+.

.NOTES
No PowerShell modules or user-created application registration are required. Authentication
uses Microsoft's existing Microsoft Graph Command Line Tools public client application.
An "exception" is a calendarView item whose Graph event type is "exception". Cancelled
items returned by Graph are counted separately.
#>

[CmdletBinding()]
param(
	[ValidateNotNullOrEmpty()]
	[string]$UserGuid = 'me',

	[string]$StartDateTime = '2023-01-01T00:00:00Z',

	[string]$EndDateTime = '2027-12-31T23:59:59Z',

	[string]$OutputDirectory = $PSScriptRoot
)

$ReportTextPath = Join-Path $OutputDirectory 'CalendarMasterDiagnostic.txt'
$CreateHighExceptionCsv = $true
$HighExceptionThreshold = 20
$HighExceptionCsvPath = Join-Path $OutputDirectory 'CalendarMastersOver20Exceptions.csv'
$LongMeetingCsvPath = Join-Path $OutputDirectory 'CalendarMastersOverOneYear.csv'
$MasterBreakdownCsvPath = Join-Path $OutputDirectory 'CalendarMasterBreakdown.csv'
$DuplicateIcalUidCsvPath = Join-Path $OutputDirectory 'CalendarMastersWithDuplicateIcalUid.csv'
$PageSize = 500
$MaximumCalendarViewDays = 1825
# Note: this will pull data in view windows of up to 5 years (1825 days) at a time, which is the maximum allowed by Microsoft Graph. This code will chunk
# for larger date ranges.
#  

$ErrorActionPreference = 'Stop'
$GraphBaseUri = 'https://graph.microsoft.com/v1.0'

function Get-GraphAccessToken {
	$clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
	$tenant = 'organizations'
	$deviceCodeUri = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/devicecode"
	$tokenUri = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token"
	$deviceCodeResponse = Invoke-RestMethod -Method Post -Uri $deviceCodeUri -ContentType 'application/x-www-form-urlencoded' -Body @{
		client_id = $clientId
		scope = 'https://graph.microsoft.com/Calendars.Read https://graph.microsoft.com/Calendars.Read.Shared'
	}

	$verificationUri = [string]$deviceCodeResponse.verification_uri_complete
	if ([string]::IsNullOrWhiteSpace($verificationUri)) {
		$verificationUri = '{0}?otc={1}' -f $deviceCodeResponse.verification_uri, [uri]::EscapeDataString($deviceCodeResponse.user_code)
	}
	Write-Information 'Opening the delegated sign-in page.' -InformationAction Continue
	Write-Information "If the browser requests a device code, enter: $($deviceCodeResponse.user_code)" -InformationAction Continue
	Start-Process $verificationUri

	$expiresAt = [datetime]::UtcNow.AddSeconds([int]$deviceCodeResponse.expires_in)
	$pollIntervalSeconds = [math]::Max([int]$deviceCodeResponse.interval, 1)
	do {
		Start-Sleep -Seconds $pollIntervalSeconds
		try {
			$tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUri -ContentType 'application/x-www-form-urlencoded' -Body @{
				grant_type = 'urn:ietf:params:oauth:grant-type:device_code'
				client_id = $clientId
				device_code = $deviceCodeResponse.device_code
			}
			if ($tokenResponse.access_token) {
				return $tokenResponse.access_token
			}
		}
		catch {
			if ([datetime]::UtcNow -ge $expiresAt) {
				throw 'Delegated sign-in timed out before authentication completed.'
			}
		}
	} while ([datetime]::UtcNow -lt $expiresAt)

	throw 'Delegated sign-in timed out before authentication completed.'
}

function Invoke-GraphGet {
	param(
		[Parameter(Mandatory = $true)][string]$Uri,
		[Parameter(Mandatory = $true)][string]$AccessToken,
		[int]$MaximumAttempts = 5
	)

	$headers = @{
		Authorization = "Bearer $AccessToken"
		Accept = 'application/json'
		Prefer = "odata.maxpagesize=$PageSize"
	}

	for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
		try {
			return Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers
		}
		catch {
			$statusCode = $null
			if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
				$statusCode = [int]$_.Exception.Response.StatusCode
			}

			$isTransient = $statusCode -in @(429, 500, 502, 503, 504)
			if (-not $isTransient -or $attempt -eq $MaximumAttempts) {
				throw
			}

			$delaySeconds = [math]::Min([math]::Pow(2, $attempt), 30)
			Write-Warning "Graph returned HTTP $statusCode. Retrying in $delaySeconds seconds (attempt $attempt of $MaximumAttempts)."
			Start-Sleep -Seconds $delaySeconds
		}
	}
}

function Get-AllGraphItem {
	param(
		[Parameter(Mandatory = $true)][string]$Uri,
		[Parameter(Mandatory = $true)][string]$AccessToken
	)

	$items = New-Object System.Collections.Generic.List[object]
	$nextUri = $Uri
	$pageNumber = 0

	while ($nextUri) {
		$pageNumber++
		Write-Information "Reading Graph page $pageNumber..." -InformationAction Continue
		$response = Invoke-GraphGet -Uri $nextUri -AccessToken $AccessToken

		foreach ($item in @($response.value)) {
			if ($null -ne $item) {
				$items.Add($item)
			}
		}

		$nextUri = $response.'@odata.nextLink'
	}

	return $items.ToArray()
}

function Get-EventDateTimeText {
	param([object]$DateTimeTimeZone)

	if ($null -eq $DateTimeTimeZone -or [string]::IsNullOrWhiteSpace($DateTimeTimeZone.dateTime)) {
		return ''
	}

	if ([string]::IsNullOrWhiteSpace($DateTimeTimeZone.timeZone)) {
		return [string]$DateTimeTimeZone.dateTime
	}

	return '{0} [{1}]' -f $DateTimeTimeZone.dateTime, $DateTimeTimeZone.timeZone
}

function Get-MasterRangeDate {
	param(
		[object]$Master,
		[ValidateSet('Start', 'End')][string]$Boundary
	)

	if ($Master.recurrence -and $Master.recurrence.range) {
		if ($Boundary -eq 'Start') {
			return [string]$Master.recurrence.range.startDate
		}
		$rangeEndDate = [datetime]::MinValue
		if (-not [datetime]::TryParse([string]$Master.recurrence.range.endDate, [ref]$rangeEndDate) -or $rangeEndDate -eq [datetime]::MinValue) {
			return ''
		}
		return [string]$Master.recurrence.range.endDate
	}

	if ($Boundary -eq 'Start') {
		return [string]$Master.start.dateTime
	}
	return [string]$Master.end.dateTime
}

function Test-RecurrenceLongerThanOneYear {
	param([object]$Master)

	if (-not $Master.recurrence -or -not $Master.recurrence.range) {
		return $false
	}

	$range = $Master.recurrence.range
	$startDate = [datetime]::MinValue
	if (-not [datetime]::TryParse([string]$range.startDate, [ref]$startDate)) {
		return $false
	}
	$endDate = [datetime]::MinValue
	if (-not [datetime]::TryParse([string]$range.endDate, [ref]$endDate)) {
		return $false
	}
	if ($endDate -eq [datetime]::MinValue) {
		return $false
	}

	return $endDate -gt $startDate.AddYears(1)
}

function Test-MeetingMasterLongerThanOneYear {
	param([object]$Master)

	if (-not $Master) {
		return $false
	}
	if ($Master.recurrence) {
		return Test-RecurrenceLongerThanOneYear -Master $Master
	}

	$startDate = [datetime]::MinValue
	$endDate = [datetime]::MinValue
	if (-not [datetime]::TryParse([string]$Master.start.dateTime, [ref]$startDate)) {
		return $false
	}
	if (-not [datetime]::TryParse([string]$Master.end.dateTime, [ref]$endDate)) {
		return $false
	}

	return $endDate -gt $startDate.AddYears(1)
}

function Test-RecurrenceHasNoEndDate {
	param([object]$Master)

	if (-not $Master.recurrence -or -not $Master.recurrence.range) {
		return $false
	}

	$endDate = [datetime]::MinValue
	return -not [datetime]::TryParse([string]$Master.recurrence.range.endDate, [ref]$endDate) -or $endDate -eq [datetime]::MinValue
}

function Add-ReportLine {
	param(
		[Parameter(Mandatory = $true)]
		[AllowEmptyCollection()]
		[AllowEmptyString()]
		[System.Collections.Generic.List[string]]$ReportLines,
		[AllowEmptyString()][string]$Text = ''
	)

	$ReportLines.Add($Text) | Out-Null  
}

$parsedStart = [datetimeoffset]::MinValue
$parsedEnd = [datetimeoffset]::MinValue
if (-not [datetimeoffset]::TryParse($StartDateTime, [ref]$parsedStart)) {
	throw "StartDateTime is not a valid ISO 8601 date/time: $StartDateTime"
}
if (-not [datetimeoffset]::TryParse($EndDateTime, [ref]$parsedEnd)) {
	throw "EndDateTime is not a valid ISO 8601 date/time: $EndDateTime"
}
if ($parsedStart -ge $parsedEnd) {
	throw 'StartDateTime must be earlier than EndDateTime.'
}
if ($PageSize -lt 1 -or $PageSize -gt 999) {
	throw 'PageSize must be between 1 and 999.'
}

foreach ($outputPath in @($ReportTextPath, $HighExceptionCsvPath, $LongMeetingCsvPath, $MasterBreakdownCsvPath, $DuplicateIcalUidCsvPath)) {
	$outputDirectory = Split-Path -Parent $outputPath
	if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
		New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
	}
}

Write-Information 'Acquiring delegated OAuth token...' -InformationAction Continue
$accessToken = Get-GraphAccessToken

$calendarOwnerPath = 'me'
if ($UserGuid -ne 'me') {
	$calendarOwnerPath = 'users/{0}' -f [uri]::EscapeDataString($UserGuid)
}
$eventSelect = 'id,subject,type,seriesMasterId,start,end,isCancelled,iCalUId'

Write-Information "Reading expanded calendar events from $($parsedStart.ToString('o')) through $($parsedEnd.ToString('o'))..." -InformationAction Continue
$calendarItemsById = @{}
$windowStart = $parsedStart
while ($windowStart -lt $parsedEnd) {
	$windowEnd = $windowStart.AddDays($MaximumCalendarViewDays)
	if ($windowEnd -gt $parsedEnd) {
		$windowEnd = $parsedEnd
	}

	$encodedStart = [uri]::EscapeDataString($windowStart.ToString('o'))
	$encodedEnd = [uri]::EscapeDataString($windowEnd.ToString('o'))
	$calendarViewUri = "$GraphBaseUri/$calendarOwnerPath/calendarView?startDateTime=$encodedStart&endDateTime=$encodedEnd&`$select=$eventSelect"
	Write-Information "Reading calendar window $($windowStart.ToString('o')) through $($windowEnd.ToString('o'))..." -InformationAction Continue

	foreach ($calendarItem in @(Get-AllGraphItem -Uri $calendarViewUri -AccessToken $accessToken)) {
		if ($null -ne $calendarItem -and -not [string]::IsNullOrWhiteSpace($calendarItem.id)) {
			$calendarItemsById[[string]$calendarItem.id] = $calendarItem
		}
	}

	$windowStart = $windowEnd
}
$calendarItems = @($calendarItemsById.Values)

$masterStatsById = @{}
foreach ($calendarItem in $calendarItems) {
	$isRecurringItem = $calendarItem.type -in @('occurrence', 'exception')
	if ($isRecurringItem -and -not [string]::IsNullOrWhiteSpace($calendarItem.seriesMasterId)) {
		$masterId = [string]$calendarItem.seriesMasterId
		$masterKind = 'Recurring series'
	}
	else {
		$masterId = [string]$calendarItem.id
		$masterKind = 'Single event'
	}

	if (-not $masterStatsById.ContainsKey($masterId)) {
		$masterStatsById[$masterId] = [pscustomobject]@{
			GraphId = $masterId
			IcalUid = [string]$calendarItem.iCalUId
			MasterKind = $masterKind
			Subject = [string]$calendarItem.subject
			Start = Get-EventDateTimeText -DateTimeTimeZone $calendarItem.start
			End = Get-EventDateTimeText -DateTimeTimeZone $calendarItem.end
			CalendarItems = 0
			Occurrences = 0
			Exceptions = 0
			CancelledItems = 0
			MasterObject = $calendarItem
		}
	}

	$stats = $masterStatsById[$masterId]
	$stats.CalendarItems++
	switch ($calendarItem.type) {
		'occurrence' { $stats.Occurrences++ }
		'exception' { $stats.Exceptions++ }
	}
	if ($calendarItem.isCancelled) {
		$stats.CancelledItems++
	}
}

$recurringMasterStats = @($masterStatsById.Values | Where-Object { $_.MasterKind -eq 'Recurring series' })
foreach ($stats in $recurringMasterStats) {
	$encodedMasterId = [uri]::EscapeDataString($stats.GraphId)
	$masterSelect = 'id,subject,type,start,end,recurrence,iCalUId'
	$masterUri = "$GraphBaseUri/$calendarOwnerPath/events/$encodedMasterId`?`$select=$masterSelect"
	$master = Invoke-GraphGet -Uri $masterUri -AccessToken $accessToken
	$stats.MasterObject = $master
	$stats.IcalUid = [string]$master.iCalUId
	$stats.Subject = [string]$master.subject
	$stats.Start = Get-EventDateTimeText -DateTimeTimeZone $master.start
	$stats.End = Get-EventDateTimeText -DateTimeTimeZone $master.end
}

$masterStats = @($masterStatsById.Values | Sort-Object Subject, GraphId)
$duplicateIcalUidGroups = @(
	$masterStats |
		Where-Object { -not [string]::IsNullOrWhiteSpace($_.IcalUid) } |
		Group-Object IcalUid |
		Where-Object { $_.Count -gt 1 }
)
$duplicateIcalUidMasterCounts = @{}
$duplicateIcalUidPeersByMasterId = @{}
foreach ($duplicateGroup in $duplicateIcalUidGroups) {
	foreach ($stats in $duplicateGroup.Group) {
		$duplicateIcalUidMasterCounts[$stats.GraphId] = $duplicateGroup.Count
		$duplicateIcalUidPeersByMasterId[$stats.GraphId] = @(
			$duplicateGroup.Group | Where-Object { $_.GraphId -ne $stats.GraphId }
		)
	}
}
$totalMasterRecords = $masterStats.Count
$totalCalendarItems = $calendarItems.Count
$totalOccurrences = @($calendarItems | Where-Object { $_.type -eq 'occurrence' }).Count
$totalExceptions = @($calendarItems | Where-Object { $_.type -eq 'exception' }).Count
$totalSingleEvents = @($calendarItems | Where-Object { $_.type -eq 'singleInstance' }).Count
$totalCancelledItems = @($calendarItems | Where-Object { $_.isCancelled }).Count
$highestRecurrencesForMeeting = 0
$highestExceptionsForMeeting = 0
if ($recurringMasterStats.Count -gt 0) {
	$highestRecurrencesForMeeting = ($recurringMasterStats | Measure-Object -Property CalendarItems -Maximum).Maximum
	$highestExceptionsForMeeting = ($recurringMasterStats | Measure-Object -Property Exceptions -Maximum).Maximum
}

$highExceptionMasterStats = @(
	$recurringMasterStats | Where-Object { $_.Exceptions -gt $HighExceptionThreshold }
)
$longerThanOneYearMasterStats = @(
	$recurringMasterStats | Where-Object { Test-RecurrenceLongerThanOneYear -Master $_.MasterObject }
)
$withoutEndDateMasterStats = @(
	$recurringMasterStats | Where-Object { Test-RecurrenceHasNoEndDate -Master $_.MasterObject }
)
$longMeetingReportMasterStats = @(
	$recurringMasterStats | Where-Object {
		Test-RecurrenceLongerThanOneYear -Master $_.MasterObject -or
		Test-RecurrenceHasNoEndDate -Master $_.MasterObject
	}
)
$reportLines = New-Object System.Collections.Generic.List[string]
Add-ReportLine -ReportLines $reportLines -Text 'MICROSOFT GRAPH CALENDAR MASTER DIAGNOSTIC'
Add-ReportLine -ReportLines $reportLines -Text ('Generated (UTC): {0}' -f [datetime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'))
Add-ReportLine -ReportLines $reportLines -Text ('User GUID: {0}' -f $UserGuid)
Add-ReportLine -ReportLines $reportLines -Text ('Period: {0} through {1}' -f $parsedStart.ToString('o'), $parsedEnd.ToString('o'))
Add-ReportLine -ReportLines $reportLines -Text ('')
Add-ReportLine -ReportLines $reportLines -Text ('TOTALS')
Add-ReportLine -ReportLines $reportLines -Text ('Master records with events in period: {0}' -f $totalMasterRecords)
Add-ReportLine -ReportLines $reportLines -Text ('  Recurring series masters: {0}' -f $recurringMasterStats.Count)
Add-ReportLine -ReportLines $reportLines -Text ('  Recurring meeting masters with more than {0} exceptions: {1}' -f $HighExceptionThreshold, $highExceptionMasterStats.Count)
Add-ReportLine -ReportLines $reportLines -Text ('  Recurring meeting masters longer than one year: {0}' -f $longerThanOneYearMasterStats.Count)
Add-ReportLine -ReportLines $reportLines -Text ('  Recurring meeting masters without an end date: {0}' -f $withoutEndDateMasterStats.Count)
Add-ReportLine -ReportLines $reportLines -Text ('  Single-event masters: {0}' -f ($totalMasterRecords - $recurringMasterStats.Count))
Add-ReportLine -ReportLines $reportLines -Text ('  Duplicate iCalUID values: {0}' -f $duplicateIcalUidGroups.Count)
Add-ReportLine -ReportLines $reportLines -Text ('  Master records sharing duplicate iCalUIDs: {0}' -f $duplicateIcalUidMasterCounts.Count)
Add-ReportLine -ReportLines $reportLines -Text ('Events in period, including recurring instances: {0}' -f $totalCalendarItems)
Add-ReportLine -ReportLines $reportLines -Text ('  Single-instance events: {0}' -f $totalSingleEvents)
Add-ReportLine -ReportLines $reportLines -Text ('  Normal recurring occurrences: {0}' -f $totalOccurrences)
Add-ReportLine -ReportLines $reportLines -Text ('  Modified recurring exceptions: {0}' -f $totalExceptions)
Add-ReportLine -ReportLines $reportLines -Text ('  Cancelled items returned by Graph: {0}' -f $totalCancelledItems)
Add-ReportLine -ReportLines $reportLines -Text ('Highest recurring instances for one meeting master in period: {0}' -f $highestRecurrencesForMeeting)
Add-ReportLine -ReportLines $reportLines -Text ('Highest modified exception count for one recurring meeting master in period: {0}' -f $highestExceptionsForMeeting)
Add-ReportLine -ReportLines $reportLines
Add-ReportLine -ReportLines $reportLines -Text 'MASTER RECORD BREAKDOWN'

if ($masterStats.Count -eq 0) {
	Add-ReportLine -ReportLines $reportLines -Text 'No calendar records were returned for the requested period.'
}
else {
	foreach ($stats in $masterStats) {
		Add-ReportLine -ReportLines $reportLines -Text ('-' * 100)
		Add-ReportLine -ReportLines $reportLines -Text ('Graph ID: {0}' -f $stats.GraphId)
		Add-ReportLine -ReportLines $reportLines -Text ('iCalUID: {0}' -f $stats.IcalUid)
		Add-ReportLine -ReportLines $reportLines -Text ('Subject: {0}' -f $stats.Subject)
		Add-ReportLine -ReportLines $reportLines -Text ('Master type: {0}' -f $stats.MasterKind)
		Add-ReportLine -ReportLines $reportLines -Text ('Master/first event start: {0}' -f $stats.Start)
		Add-ReportLine -ReportLines $reportLines -Text ('Master/first event end: {0}' -f $stats.End)
		Add-ReportLine -ReportLines $reportLines -Text ('Events in period: {0}' -f $stats.CalendarItems)
		Add-ReportLine -ReportLines $reportLines -Text ('Normal recurring occurrences in period: {0}' -f $stats.Occurrences)
		Add-ReportLine -ReportLines $reportLines -Text ('Modified recurring exceptions in period: {0}' -f $stats.Exceptions)
		Add-ReportLine -ReportLines $reportLines -Text ('Cancelled items returned by Graph in period: {0}' -f $stats.CancelledItems)
		if ($stats.Exceptions -gt $HighExceptionThreshold) {
			Add-ReportLine -ReportLines $reportLines -Text '*** Warning: More than 20 exceptions for this master event. ***'
		}
		if (Test-MeetingMasterLongerThanOneYear -Master $stats.MasterObject) {
			Add-ReportLine -ReportLines $reportLines -Text '*** Warning: The meeting is over a year long. ***'
		}
		if (Test-RecurrenceHasNoEndDate -Master $stats.MasterObject) {
			Add-ReportLine -ReportLines $reportLines -Text '*** Warning: The meeting is over has no end date. ***'
		}
		if ($duplicateIcalUidMasterCounts.ContainsKey($stats.GraphId)) {
			Add-ReportLine -ReportLines $reportLines -Text ('*** Warning: Duplicate iCalUID shared by {0} master records. ***' -f $duplicateIcalUidMasterCounts[$stats.GraphId])
			foreach ($duplicatePeer in $duplicateIcalUidPeersByMasterId[$stats.GraphId]) {
				Add-ReportLine -ReportLines $reportLines -Text ('    Duplicate Graph ID: {0}' -f $duplicatePeer.GraphId)
				Add-ReportLine -ReportLines $reportLines -Text ('    Duplicate Subject: {0}' -f $duplicatePeer.Subject)
			}
		}
	}
}

$reportLines | Set-Content -LiteralPath $ReportTextPath -Encoding UTF8

$masterBreakdownRows = @(
	$masterStats | ForEach-Object {
		$warningMessages = New-Object System.Collections.Generic.List[string]
		if ($_.Exceptions -gt $HighExceptionThreshold) {
			$warningMessages.Add('More than 20 exceptions for this master event.')
		}
		if (Test-MeetingMasterLongerThanOneYear -Master $_.MasterObject) {
			$warningMessages.Add('The meeting is over a year long.')
		}
		if (Test-RecurrenceHasNoEndDate -Master $_.MasterObject) {
			$warningMessages.Add('The meeting is over has no end date.')
		}
		if ($duplicateIcalUidMasterCounts.ContainsKey($_.GraphId)) {
			$warningMessages.Add(('Duplicate iCalUID shared by {0} master records.' -f $duplicateIcalUidMasterCounts[$_.GraphId]))
		}

		[pscustomobject]@{
			GraphId = $_.GraphId
			IcalUid = $_.IcalUid
			Subject = $_.Subject
			MasterType = $_.MasterKind
			MasterFirstEventStart = $_.Start
			MasterFirstEventEnd = $_.End
			EventsInPeriod = $_.CalendarItems
			NormalRecurringOccurrencesInPeriod = $_.Occurrences
			ModifiedRecurringExceptionsInPeriod = $_.Exceptions
			CancelledItemsReturnedByGraphInPeriod = $_.CancelledItems
			Warnings = $warningMessages -join ' | '
		}
	}
)

if ($masterBreakdownRows.Count -gt 0) {
	$masterBreakdownRows | Export-Csv -LiteralPath $MasterBreakdownCsvPath -NoTypeInformation -Encoding UTF8
}
else {
	'GraphId,IcalUid,Subject,MasterType,MasterFirstEventStart,MasterFirstEventEnd,EventsInPeriod,NormalRecurringOccurrencesInPeriod,ModifiedRecurringExceptionsInPeriod,CancelledItemsReturnedByGraphInPeriod,Warnings' |
		Set-Content -LiteralPath $MasterBreakdownCsvPath -Encoding UTF8
}

$duplicateIcalUidRows = @(
	foreach ($duplicateGroup in $duplicateIcalUidGroups) {
		foreach ($stats in $duplicateGroup.Group | Sort-Object Subject, GraphId) {
			$duplicatePeers = @($duplicateIcalUidPeersByMasterId[$stats.GraphId])
			[pscustomobject]@{
				IcalUid = $duplicateGroup.Name
				DuplicateMasterCount = $duplicateGroup.Count
				GraphId = $stats.GraphId
				Subject = $stats.Subject
				DuplicateGraphIds = ($duplicatePeers.GraphId -join ' | ')
				DuplicateSubjects = ($duplicatePeers.Subject -join ' | ')
				MasterType = $stats.MasterKind
				MasterFirstEventStart = $stats.Start
				MasterFirstEventEnd = $stats.End
			}
		}
	}
)

if ($duplicateIcalUidRows.Count -gt 0) {
	$duplicateIcalUidRows | Export-Csv -LiteralPath $DuplicateIcalUidCsvPath -NoTypeInformation -Encoding UTF8
}
else {
	'IcalUid,DuplicateMasterCount,GraphId,Subject,DuplicateGraphIds,DuplicateSubjects,MasterType,MasterFirstEventStart,MasterFirstEventEnd' |
		Set-Content -LiteralPath $DuplicateIcalUidCsvPath -Encoding UTF8
}

if ($CreateHighExceptionCsv) {
	$highExceptionRows = @(
		$highExceptionMasterStats |
			Sort-Object Exceptions -Descending |
			ForEach-Object {
				[pscustomobject]@{
					GraphId = $_.GraphId
					Subject = $_.Subject
					StartDate = Get-MasterRangeDate -Master $_.MasterObject -Boundary Start
					EndDate = Get-MasterRangeDate -Master $_.MasterObject -Boundary End
				}
			}
	)

	if ($highExceptionRows.Count -gt 0) {
		$highExceptionRows | Export-Csv -LiteralPath $HighExceptionCsvPath -NoTypeInformation -Encoding UTF8
	}
	else {
		'GraphId,Subject,StartDate,EndDate' | Set-Content -LiteralPath $HighExceptionCsvPath -Encoding UTF8
	}
}

$longMeetingRows = @(
	$longMeetingReportMasterStats |
		Sort-Object Subject, GraphId |
		ForEach-Object {
			[pscustomobject]@{
				GraphId = $_.GraphId
				Subject = $_.Subject
				StartDate = Get-MasterRangeDate -Master $_.MasterObject -Boundary Start
				EndDate = Get-MasterRangeDate -Master $_.MasterObject -Boundary End
			}
		}
)

if ($longMeetingRows.Count -gt 0) {
	$longMeetingRows | Export-Csv -LiteralPath $LongMeetingCsvPath -NoTypeInformation -Encoding UTF8
}
else {
	'GraphId,Subject,StartDate,EndDate' | Set-Content -LiteralPath $LongMeetingCsvPath -Encoding UTF8
}

Write-Information "Text report written to: $ReportTextPath" -InformationAction Continue
Write-Information "Master breakdown CSV written to: $MasterBreakdownCsvPath" -InformationAction Continue
Write-Information "Duplicate iCalUID CSV written to: $DuplicateIcalUidCsvPath" -InformationAction Continue
if ($CreateHighExceptionCsv) {
	Write-Information "High-exception CSV written to: $HighExceptionCsvPath" -InformationAction Continue
}
Write-Information "Over-one-year CSV written to: $LongMeetingCsvPath" -InformationAction Continue

