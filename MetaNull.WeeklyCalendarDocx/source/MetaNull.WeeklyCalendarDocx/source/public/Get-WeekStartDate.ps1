[CmdletBinding(DefaultParameterSetName="ByYearWeek")]
param(
    [Parameter(Mandatory,ParameterSetName="ByYearWeek")]
    #[ValidateRange(1900, 2100)]
    [int]$Year,

    [Parameter(Mandatory,ParameterSetName="ByYearWeek")]
    [ValidateRange(1, 53)]
    [int]$Week,

    [Parameter(Mandatory,ParameterSetName="ByDate")]
    [ValidateNotNull()]
    [DateTime]$Date
)
try {
    if ($PSCmdlet.ParameterSetName -eq "ByDate") {
        # Calculate the Monday of the week containing $Date
        $dayOfWeek = ([int]$Date.DayOfWeek + 6) % 7
        return $Date.AddDays(-$dayOfWeek).Date
    } else {
        # Calculate the Monday of the specified ISO week and year
        # ISO weeks start with the week containing the first Thursday of the year
        $jan4 = Get-Date -Year $Year -Month 1 -Day 4
        $dayOfWeek = [int]$jan4.DayOfWeek
        $monday = $jan4.AddDays(-($dayOfWeek - 1))
        return $monday.AddDays(($Week - 1) * 7)
    }
}
catch {
    Write-Error "Failed to calculate week start date for Year $Year, Week $($Week): $_"
    throw
}