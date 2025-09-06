[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    $Paragraph,

    [int]$BottomLineStyle = (Get-WordConstants -ConstantName 'WD_LINE_STYLE_SINGLE'),
    [int]$LineWidth = (Get-WordConstants -ConstantName 'WD_LINE_WIDTH_075PT')
)

$Constants = Get-WordConstants -All

try {
    $Paragraph.Format.Borders.Enable = $true
    $Paragraph.Format.Borders.Item(1).LineStyle = $Constants.WD_LINE_STYLE_NONE  # Top
    $Paragraph.Format.Borders.Item(2).LineStyle = $Constants.WD_LINE_STYLE_NONE  # Left
    $Paragraph.Format.Borders.Item(4).LineStyle = $Constants.WD_LINE_STYLE_NONE  # Right
    $Paragraph.Format.Borders.Item(3).LineStyle = $BottomLineStyle     # Bottom
    $Paragraph.Format.Borders.Item(3).LineWidth = $LineWidth
}
catch {
    Write-Warning "Failed to set paragraph borders: $_"
}