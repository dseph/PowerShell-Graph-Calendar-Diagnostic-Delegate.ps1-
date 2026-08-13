BeforeAll {
    . "$PSScriptRoot/../Get-CalendarDiagnostic.ps1"

    function Get-TestEvent {
        param(
            [string]$Id,
            [string]$Type = 'singleInstance',
            [string]$Uid,
            [string]$SeriesMasterId,
            [string]$RangeType,
            [string]$EndDate,
            [int]$NumberOfOccurrences
        )

        [pscustomobject]@{
            id = $Id
            subject = $Id
            type = $Type
            iCalUId = $Uid
            seriesMasterId = $SeriesMasterId
            start = [pscustomobject]@{ dateTime = '2026-01-01T09:00:00' }
            end = [pscustomobject]@{ dateTime = '2026-01-01T10:00:00' }
            recurrence = [pscustomobject]@{
                range = [pscustomobject]@{
                    type = $RangeType
                    endDate = $EndDate
                    numberOfOccurrences = $NumberOfOccurrences
                }
            }
        }
    }
}

Describe 'Get-CalendarDiagnostic' {
    It 'finds every requested calendar issue category' {
        $events = @(
            (Get-TestEvent -Id unending -Type seriesMaster -RangeType noEnd),
            (Get-TestEvent -Id long -Type seriesMaster -RangeType endDate -EndDate '2027-01-02'),
            (Get-TestEvent -Id duplicate-one -Uid duplicated),
            (Get-TestEvent -Id duplicate-two -Uid duplicated),
            (Get-TestEvent -Id recurring -Type seriesMaster -RangeType numbered -NumberOfOccurrences 101)
        )
        $events += 1..21 | ForEach-Object { Get-TestEvent -Id "exception-$_" -Type exception -SeriesMasterId recurring }

        $result = Get-CalendarDiagnostic -Events $events -AsOf ([datetime]'2026-01-01')

        $result.Summary.UnendingMeetings | Should -Be 1
        $result.Summary.LongRunningMeetings | Should -Be 1
        $result.Summary.DuplicateUidEvents | Should -Be 2
        $result.Summary.HighRecurrenceCounts | Should -Be 1
        $result.Summary.ExcessiveExceptions | Should -Be 1
        $result.Details.Where({ $_.Id -eq 'recurring' }).DiagnosticIssues | Should -Match 'More than 20 exceptions \(21\)'
    }

    It 'writes the report and every issue CSV' {
        $outputPath = Join-Path $TestDrive 'report'
        $result = Get-CalendarDiagnostic -Events @((Get-TestEvent -Id event -Uid uid))

        Export-CalendarDiagnostic -Diagnostic $result -Path $outputPath

        @(
            'CalendarDiagnosticReport.html',
            'CalendarEvents.csv',
            'UnendingMeetings.csv',
            'LongRunningMeetings.csv',
            'DuplicateUids.csv',
            'HighRecurrenceCounts.csv',
            'ExcessiveExceptions.csv'
        ) | ForEach-Object { Join-Path $outputPath $_ | Should -Exist }
    }
}
