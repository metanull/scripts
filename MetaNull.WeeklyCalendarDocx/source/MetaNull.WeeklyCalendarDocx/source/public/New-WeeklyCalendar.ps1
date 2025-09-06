<#
.SYNOPSIS
    Generates a weekly calendar document using Microsoft Word.

.DESCRIPTION
    Creates a formatted weekly calendar showing work days (Monday-Friday) for a specified number of weeks.
    The calendar displays ISO week numbers and provides space for daily planning.

.PARAMETER Year
    The year for which to generate the calendar. Must be between 1900 and 2100.

.PARAMETER FromWeek
    The starting ISO week number (1-53). Defaults to 1.

.PARAMETER NumberOfWeeks
    The number of weeks to generate (1-52). Defaults to 6.

.PARAMETER BreakAfter
    Number of weeks after which to insert a page break. Defaults to 3. Set to 0 to disable page breaks.

.PARAMETER SaveAs
    Path where to save the document. If not specified, Word will be opened visibly for interactive use.

.PARAMETER Language
    Language for day names and date formatting. Supported: French, English.

.PARAMETER FontSize
    Font size. Defaults to 10.

.PARAMETER FontFamily
    Font family to use throughout the document. Defaults to 'Aptos'.

.PARAMETER Force
    When saving, overwrite existing files without prompting.

.EXAMPLE
    New-WeeklyCalendar -Year 2025 -FromWeek 10 -NumberOfWeeks 8
    Creates a calendar for weeks 10-17 of 2025 and opens it in Word.

.EXAMPLE
    New-WeeklyCalendar -SaveAs "C:\temp\calendar.docx" -Year 2025 -Force
    Creates a full year calendar for 2025 and saves it to the specified path.

.EXAMPLE
    New-WeeklyCalendar -Language English -FontSize 16
    Creates a calendar with English language settings and larger font.

.NOTES
    Requires Microsoft Word to be installed.
    Uses ISO 8601 week numbering (weeks start on Monday).

    Author: Pascal Havelange
    License: MIT License - https://opensource.org/licenses/MIT
             You are free to use, modify, and distribute this software without restriction.
#>
[CmdletBinding()]
param(
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'SaveAs')]
    #[ValidateRange(1900, 2100)]
    [int]$Year = (Get-Date).Year,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'SaveAs')]
    #[ValidateRange(1, 53)]
    [int]$FromWeek = 1,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'SaveAs')]
    #[ValidateRange(1, 52)]
    [int]$NumberOfWeeks = 6,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'SaveAs')]
    [ValidateRange(0, 53)]
    [int]$BreakAfter = 3,

    [Parameter(ParameterSetName = 'SaveAs', Mandatory = $true)]
    [ValidateScript({
        $directory = Split-Path $_ -Parent
        if (-not (Test-Path $directory -PathType Container)) {
            throw "Directory '$directory' does not exist."
        }
        $true
    })]
    [string]$SaveAs,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'SaveAs')]
    [ValidateScript({
        $availableLanguages = Get-LanguageConfiguration -ListAvailable
        if ($_ -in $availableLanguages) {
            $true
        } else {
            throw "Language '$_' is not supported. Available languages: $($availableLanguages -join ', ')"
        }
    })]
    [string]$Language = 'French',

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'SaveAs')]
    [ValidateRange(8, 24)]
    [int]$FontSize = 10,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'SaveAs')]
    [string]$FontFamily = 'Aptos',

    [Parameter(ParameterSetName = 'SaveAs')]
    [switch]$Force
)

$ParameterSetName = $PSCmdlet.ParameterSetName
$Constants = Get-WordConstants -All
$LanguageConfig = Get-LanguageConfiguration -Language $Language

# Validate Word application availability
if (-not (Test-WordApplication)) {
    throw "Microsoft Word is required but not available."
}

# Test file overwrite if saving
if ($ParameterSetName -eq 'SaveAs') {
    Test-FileOverwrite -FilePath $SaveAs -Force:$Force
}

# Initialize progress tracking
$progress = @{
    Activity = "Generating Weekly Calendar"
    Status = "Initializing..."
    PercentComplete = 0
}

Write-Progress @progress

# Create Word application with timeout handling
$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    if ($ParameterSetName -eq 'SaveAs') {
        $word.Visible = $false
    } else {
        $word.Visible = $true
    }

    # Create a new document
    $doc = $word.Documents.Add()
    $doc.PageSetup.TopMargin = $word.CentimetersToPoints(2.0)
    $doc.PageSetup.BottomMargin = $word.CentimetersToPoints(1.5)
    $doc.PageSetup.LeftMargin = $word.CentimetersToPoints(0.5)
    $doc.PageSetup.RightMargin = $word.CentimetersToPoints(0.5)

    $groupingCounter = 0
    $totalWeeks = $NumberOfWeeks + 1

    for ($week = $FromWeek; $week -lT ($FromWeek + $NumberOfWeeks); $week++) {
        $groupingCounter++
        $currentProgress = [math]::Round(($groupingCounter / $totalWeeks) * 100, 0)

        $startDate = Get-WeekStartDate -Year $Year -Week $week
        $endDate = $startDate.AddDays(4)
        $actualWeek = Get-ISOWeekNumber -Date $startDate

        $progress.Status = "Processing Year $($startDate.Year), week # $actualWeek ($groupingCounter of $totalWeeks)"
        $progress.PercentComplete = $currentProgress
        Write-Progress @progress

        if ($groupingCounter -ne 1 -and $week -gt $FromWeek -and $endDate.Year -gt $startDate.Year) {
            # Section break on new year - insert break at current position
            $doc.Range($doc.Content.End - 1, $doc.Content.End - 1).InsertBreak($Constants.WD_SECTION_BREAK_NEXT_PAGE)
            $groupingCounter = 1
        }

        # Add a table with 8 rows and 1 column
        $docParagraph = $doc.Paragraphs.Add()
        $docParagraph.Format.KeepTogether = $true
        $docParagraph.Format.KeepWithNext = $true

        $table = $doc.Tables.Add($docParagraph.Range, 8, 1)
        $table.AutoFitBehavior($Constants.WD_AUTOFIT_FIXED)
        $table.AllowAutoFit = $false
        $table.PreferredWidthType = $Constants.WD_PREFERRED_WIDTH_PERCENT
        $table.PreferredWidth = 100
        $table.Columns.Item(1).PreferredWidthType = $Constants.WD_PREFERRED_WIDTH_PERCENT
        $table.Columns.Item(1).PreferredWidth = 100

        # Row 1: Header
        $CurrentRow = 1

        if ($groupingCounter -eq 1) {
            # First week of the group - show ending year or starting year week number is 1
            if ($actualWeek -eq 1) {
                $HeaderText = "$($LanguageConfig.WeekPrefix)$actualWeek ($($endDate.Year))"
            } else {
                $HeaderText = "$($LanguageConfig.WeekPrefix)$actualWeek ($($startDate.Year))"
            }
        } elseif ($actualWeek -in @(52,53)) {
            # Year transition case for week 52/53
            $HeaderText = "$($LanguageConfig.WeekPrefix)$actualWeek ($($startDate.Year))"
        } elseif ($actualWeek -in @(1,2)) {
            # Year transition case for week 1/2
            $HeaderText = "$($LanguageConfig.WeekPrefix)$actualWeek ($($endDate.Year))"
        } else {
            $HeaderText = "$($LanguageConfig.WeekPrefix)$actualWeek"
        }

        Add-TableRow -Table $table -RowNumber $CurrentRow -Text $HeaderText -FontFamily $FontFamily -FontSize ($FontSize + [int]($FontSize * .2)) -Bold

        # Row 2: Subtitle
        $CurrentRow++
        if ($startDate.Month -ne $endDate.Month) {
            if ($startDate.Year -ne $endDate.Year) {
                $SubtitleText = $LanguageConfig.DateRangeFormat -f $startDate.ToString($LanguageConfig.DateFormat.DifferentYear.From), $endDate.ToString($LanguageConfig.DateFormat.DifferentYear.To)
            } else {
                $SubtitleText = $LanguageConfig.DateRangeFormat -f $startDate.ToString($LanguageConfig.DateFormat.DifferentMonth.From), $endDate.ToString($LanguageConfig.DateFormat.DifferentMonth.To)
            }
        } else {
            $SubtitleText = $LanguageConfig.DateRangeFormat -f $startDate.ToString($LanguageConfig.DateFormat.SameMonth.From), $endDate.ToString($LanguageConfig.DateFormat.SameMonth.To)
        }

        Add-TableRow -Table $table -RowNumber $CurrentRow -Text $SubtitleText -FontFamily $FontFamily -FontSize $FontSize -FontColor ([System.Drawing.Color]::FromArgb(128, 0, 0))

        # Row 3: Spacer
        $CurrentRow++
        Add-TableRow -Table $table -RowNumber $CurrentRow -Text "`t" -FontFamily $FontFamily -FontSize $FontSize

        # Rows 4 to 8: Days
        $days = $LanguageConfig.Days
        for ($i = 0; $i -lt $days.Count; $i++) {
            $CurrentRow++
            Add-TableRow -Table $table -RowNumber $CurrentRow -Text "$($days[$i])`t:" -FontFamily $FontFamily -FontSize $FontSize
        }

        if ($breakAfter -gt 0) {
            if ($groupingCounter -gt 0 -and $groupingCounter % $breakAfter -eq 0) {
                # Page break on every multiple of breakAfter (save for the very last week)
                if ($week -lt (($FromWeek + $NumberOfWeeks) - 1)) {
                    $doc.Range($doc.Content.End - 1, $doc.Content.End - 1).InsertBreak($Constants.WD_SECTION_BREAK_NEXT_PAGE)
                }
            } else {
                # Add controlled spacing after each table
                $newPara = $doc.Paragraphs.Add()
                $newPara.SpaceAfter = 0  # Remove after-paragraph spacing
                $newPara.SpaceBefore = 6  # Small before-paragraph spacing (6 points)
            }
        } else {
            # Add controlled spacing after each table
            $newPara = $doc.Paragraphs.Add()
            $newPara.SpaceAfter = 0  # Remove after-paragraph spacing
            $newPara.SpaceBefore = 6  # Small before-paragraph spacing (6 points)
        }
    }

    $progress.PercentComplete = 100
    $progress.Status = "Finalizing document..."
    Write-Progress @progress

    if ($ParameterSetName -eq 'SaveAs') {
        Write-Information "Saving Calendar to: $SaveAs"
        $doc.SaveAs([ref] $SaveAs)
        $doc.Close()
        $word.Quit()
    } else {
        Write-Information "Calendar created and opened in Word."
    }
}
finally {
    # Cleanup COM objects
    Write-Progress -Activity "Generating Weekly Calendar" -Completed

    if ($doc) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
        catch {
            Write-Warning "Failed to release document COM object: $_"
        }
    }

    if ($word) {
        try {
            if ($ParameterSetName -eq 'SaveAs') {
                $word.Quit()
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        catch {
            Write-Warning "Failed to release Word COM object: $_"
        }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}