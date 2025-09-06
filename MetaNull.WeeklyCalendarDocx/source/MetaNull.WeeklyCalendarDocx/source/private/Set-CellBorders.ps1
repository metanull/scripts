[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    $Cell,

    [int]$LineStyle = (Get-WordConstants -ConstantName 'WD_LINE_STYLE_NONE')
)

try {
    $Cell.Borders.Item(1).LineStyle = $LineStyle
    $Cell.Borders.Item(2).LineStyle = $LineStyle
    $Cell.Borders.Item(3).LineStyle = $LineStyle
    $Cell.Borders.Item(4).LineStyle = $LineStyle
    $Cell.Borders.Item(5).LineStyle = $LineStyle
    $Cell.Borders.Item(6).LineStyle = $LineStyle
}
catch {
    Write-Warning "Failed to set cell borders: $_"
}