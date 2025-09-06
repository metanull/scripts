[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [DateTime]$Date
)

try {
    $calendar = [System.Globalization.CultureInfo]::InvariantCulture.Calendar
    $week = $calendar.GetWeekOfYear(
        $Date,
        [System.Globalization.CalendarWeekRule]::FirstFourDayWeek,
        [System.DayOfWeek]::Monday
    )

    # If the week is 53, check if it belongs to next year
    if ($week -eq 53) {
        $nextYearJan1 = Get-Date -Year ($Date.Year + 1) -Month 1 -Day 1
        $nextWeek = $calendar.GetWeekOfYear(
            $nextYearJan1,
            [System.Globalization.CalendarWeekRule]::FirstFourDayWeek,
            [System.DayOfWeek]::Monday
        )

        if ($nextWeek -eq 1 -and $Date -ge $nextYearJan1.AddDays(-3)) {
            return 1
        }
    }
    return $week
}
catch {
    Write-Error "Failed to calculate ISO week number for date $($Date): $_"
    throw
}