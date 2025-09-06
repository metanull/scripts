[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    $Table,

    [Parameter(Mandatory)]
    [int]$RowNumber,

    [Parameter(Mandatory)]
    [AllowEmptyString()]
    [string]$Text,

    [string]$FontFamily,

    [int]$FontSize = 8,

    [double]$LineHeight = 1.0,

    [switch]$Bold,

    [System.Drawing.Color]$FontColor,

    [switch]$HasNoBorders
)

$Constants = Get-WordConstants -All

try {
    # Get the cell and clear it
    $cell = $Table.Cell($RowNumber, 1)
    $cell.Range.Text = "" # Required to avoid Word reusing the previous range!
    Set-CellBorders -Cell $cell -LineStyle $Constants.WD_LINE_STYLE_NONE

    # Set the Line Height
    $row = $Table.Rows.Item($RowNumber)
    $row.HeightRule = $Constants.WD_ROW_HEIGHT_EXACTLY
    $row.Height = $Table.Application.CentimetersToPoints($LineHeight)

    # Configure the paragraph
    $paragraph = $cell.Range.Paragraphs.Item(1)
    $paragraph.Range.Text = $Text
    $paragraph.Range.Font.Size = $FontSize

    if ($FontFamily) {
        $paragraph.Range.Font.Name = $FontFamily
    }

    if ($Bold) {
        $paragraph.Range.Font.Bold = $true
    }

    if ($FontColor) {
        $paragraph.Range.Font.Color = $FontColor
    }

    # Apply borders unless HasNoBorders is specified
    if (-not $HasNoBorders) {
        Set-ParagraphBorders -Paragraph $paragraph
    }
}
catch {
    Write-Warning "Failed to add table row $RowNumber with text '$Text': $_"
}