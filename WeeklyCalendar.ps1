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
    .\WeeklyCalendar.ps1 -Year 2025 -FromWeek 10 -NumberOfWeeks 8
    Creates a calendar for weeks 10-17 of 2025 and opens it in Word.

.EXAMPLE
    .\WeeklyCalendar.ps1 -SaveAs "C:\temp\calendar.docx" -Year 2025 -Force
    Creates a full year calendar for 2025 and saves it to the specified path.

.EXAMPLE
    .\WeeklyCalendar.ps1 -Language English -FontSize 16
    Creates a calendar with English language settings and larger font.

.NOTES
    Requires Microsoft Word to be installed.
    Uses ISO 8601 week numbering (weeks start on Monday).

    Author: Pascal Havelange
    License: MIT License - https://opensource.org/licenses/MIT
             You are free to use, modify, and distribute this software without restriction.
    Contact: havelangep [at] hotmail.com
#>
[CmdletBinding(DefaultParameterSetName = 'Default')]
param (
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
    [switch]$Force,

    [Parameter(ParameterSetName = 'CurrentWeek')]
    [switch]$CurrentWeek
)
Process {
    # Main script execution
    if ($CurrentWeek.IsPresent -and $CurrentWeek) {
        $Week = Get-ISOWeekNumber -Date (Get-Date)
        Write-Host "Week number for (Get-Date) is $Week" -ForegroundColor Green
        return $Week
    } else {
        try {
            New-WeeklyCalendar -Year $Year -FromWeek $FromWeek -NumberOfWeeks $NumberOfWeeks -SaveAs $SaveAs -Language $Language -FontSize $FontSize -FontFamily $FontFamily -BreakAfter $BreakAfter -Force:$Force -ParameterSetName $PSCmdlet.ParameterSetName
        }
        catch {
            Write-Error "Failed to generate calendar: $_"
            throw
        }
    }
}
Begin {
    #region Calendar Generation
    function New-WeeklyCalendar {
        [CmdletBinding()]
        param(
            [int]$Year,
            [int]$FromWeek,
            [int]$NumberOfWeeks,
            [int]$BreakAfter,
            [string]$SaveAs,
            [string]$Language,
            [int]$FontSize,
            [string]$FontFamily,
            [switch]$Force,
            [string]$ParameterSetName
        )

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

                Add-TableRow -Table $table -RowNumber $CurrentRow -Text $HeaderText -FontFamily $FontFamily -FontSize ($FontSize + [int]($FontSize * .2)) -Bold -IsHeader

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
                $doc.SaveAs([ref] $SaveAs)
                $doc.Close()
                $word.Quit()
                Write-Host "Calendar saved to: $SaveAs" -ForegroundColor Green
            } else {
                Write-Host "Calendar created and opened in Word." -ForegroundColor Green
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
    }
    #endregion

    #region Helper Functions
    function Get-WordConstants {
        [CmdletBinding()]
        param(
            [Parameter(ParameterSetName = 'GetConstant')]
            [string]$ConstantName,

            [Parameter(ParameterSetName = 'GetAll')]
            [switch]$All,

            [Parameter(ParameterSetName = 'ListConstants')]
            [switch]$ListAvailable
        )

        $WordConstants = @{
            WD_LINE_STYLE_NONE = 0
            WD_LINE_STYLE_SINGLE = 1
            WD_SECTION_BREAK_NEXT_PAGE = 2
            WD_LINE_WIDTH_075PT = 6
            WD_AUTOFIT_FIXED = 0
            WD_PREFERRED_WIDTH_PERCENT = 1
            # Additional commonly used Word constants
            WD_PAGE_BREAK = 1
            WD_LINE_STYLE_DOUBLE = 7
            WD_LINE_WIDTH_150PT = 12
            WD_LINE_WIDTH_225PT = 18
            WD_LINE_WIDTH_300PT = 24
            WD_BORDER_TOP = 1
            WD_BORDER_LEFT = 2
            WD_BORDER_BOTTOM = 3
            WD_BORDER_RIGHT = 4
            WD_BORDER_HORIZONTAL = 5
            WD_BORDER_VERTICAL = 6
            WD_ROW_HEIGHT_AUTO = 0
            WD_ROW_HEIGHT_AT_LEAST = 1
            WD_ROW_HEIGHT_EXACTLY = 2
        }

        switch ($PSCmdlet.ParameterSetName) {
            'GetConstant' {
                if ($WordConstants.ContainsKey($ConstantName)) {
                    return $WordConstants[$ConstantName]
                } else {
                    throw "Constant '$ConstantName' not found. Use -ListAvailable to see available constants."
                }
            }
            'GetAll' {
                return $WordConstants
            }
            'ListConstants' {
                return $WordConstants.Keys | Sort-Object
            }
            default {
                # Default behavior - return all constants
                return $WordConstants
            }
        }
    }

    function Get-LanguageConfiguration {
        [CmdletBinding()]
        param(
            [Parameter(ParameterSetName = 'GetLanguage')]
            [string]$Language,

            [Parameter(ParameterSetName = 'ListLanguages')]
            [switch]$ListAvailable
        )

        $LanguageConfig = @{
            French = @{
                Days = @("LUN", "MAR", "MER", "JEU", "VEN")
                WeekPrefix = "SEM."
                DateRangeFormat = "({0} → {1})"
                DateFormat = @{
                    SameMonth = @{
                        From = "%d"
                        To = "d'/'MM"
                    }
                    DifferentMonth = @{
                        From = "d'/'MM"
                        To = "d'/'MM'/'yyyy"
                    }
                    DifferentYear = @{
                        From = "d'/'MM'/'yyyy"
                        To = "d'/'MM'/'yyyy"
                    }
                }
            }
            English = @{
                Days = @("MON", "TUE", "WED", "THU", "FRI")
                WeekPrefix = "WK. "
                DateFormat = @{
                    SameMonth = @{
                        From = "d"
                        To = "d'/'MM"
                    }
                    DifferentMonth = @{
                        From = "d'/'MM"
                        To = "d'/'MM'/'yyyy"
                    }
                    DifferentYear = @{
                        From = "d'/'MM'/'yyyy"
                        To = "d'/'MM'/'yyyy"
                    }
                }
            }
        }

        if ($ListAvailable) {
            return $LanguageConfig.Keys | Sort-Object
        }

        if ($Language -and $LanguageConfig.ContainsKey($Language)) {
            return $LanguageConfig[$Language]
        } elseif ($Language) {
            throw "Language '$Language' is not supported. Available languages: $($LanguageConfig.Keys -join ', ')"
        } else {
            throw "Language parameter is required when not listing available languages."
        }
    }

    function Get-WeekStartDate {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            #[ValidateRange(1900, 2100)]
            [int]$Year,

            [Parameter(Mandatory)]
            #[ValidateRange(1, 53)]
            [int]$Week
        )

        try {
            $jan4 = Get-Date -Year $Year -Month 1 -Day 4
            $dayOfWeek = [int]$jan4.DayOfWeek
            $monday = $jan4.AddDays(-($dayOfWeek - 1))
            return $monday.AddDays(($Week - 1) * 7)
        }
        catch {
            Write-Error "Failed to calculate week start date for Year $Year, Week $($Week): $_"
            throw
        }
    }

    function Get-ISOWeekNumber {
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
    }

    function Set-CellBorders {
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
    }

    function Set-ParagraphBorders {
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
    }

    function Add-TableRow {
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

            [switch]$IsHeader,

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
    }

    function Test-WordApplication {
        [CmdletBinding()]
        param()

        try {
            $testWord = New-Object -ComObject Word.Application -ErrorAction Stop
            $testWord.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testWord) | Out-Null
            return $true
        }
        catch {
            Write-Error "Microsoft Word is not available: $_"
            return $false
        }
    }

    function Test-FileOverwrite {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$FilePath,

            [switch]$Force
        )

        if (Test-Path $FilePath) {
            if (-not $Force) {
                $response = Read-Host "File '$FilePath' already exists. Overwrite? (Y/N)"
                if ($response -notmatch '^[Yy]') {
                    throw "Operation cancelled by user."
                }
            }

            # Test if file is locked
            try {
                [System.IO.File]::OpenWrite($FilePath).Close()
            }
            catch {
                throw "File '$FilePath' is locked or cannot be overwritten: $_"
            }
        }
    }
    #endregion
}