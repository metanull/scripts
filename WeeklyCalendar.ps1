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

.PARAMETER SaveAs
    Path where to save the document. If not specified, Word will be opened visibly for interactive use.

.PARAMETER Language
    Language for day names and date formatting. Supported: French, English, German, Spanish.

.PARAMETER HeaderFontSize
    Font size for week headers (8-24). Defaults to 16.

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
    .\WeeklyCalendar.ps1 -Language English -HeaderFontSize 18
    Creates a calendar with English language settings and larger headers.

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
    [int]$HeaderFontSize = 16,
    
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'SaveAs')]
    [string]$FontFamily = 'Aptos',
    
    [Parameter(ParameterSetName = 'SaveAs')]
    [switch]$Force
)
Process {
    # Main script execution
    try {
        New-WeeklyCalendar -Year $Year -FromWeek $FromWeek -NumberOfWeeks $NumberOfWeeks -SaveAs $SaveAs -Language $Language -HeaderFontSize $HeaderFontSize -FontFamily $FontFamily -Force:$Force -ParameterSetName $PSCmdlet.ParameterSetName
    }
    catch {
        Write-Error "Failed to generate calendar: $_"
        throw
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
            [string]$SaveAs,
            [string]$Language,
            [int]$HeaderFontSize,
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
            
            for ($week = $FromWeek; $week -le ($FromWeek + $NumberOfWeeks); $week++) {
                $groupingCounter++
                $currentProgress = [math]::Round(($groupingCounter / $totalWeeks) * 100, 0)

                $startDate = Get-WeekStartDate -Year $Year -Week $week
                $endDate = $startDate.AddDays(4)
                $actualWeek = Get-ISOWeekNumber -Date $startDate

                $progress.Status = "Processing Year $($startDate.Year), week # $actualWeek ($groupingCounter of $totalWeeks)"
                $progress.PercentComplete = $currentProgress
                Write-Progress @progress

                if ($week -gt $FromWeek -and $endDate.Year -gt $startDate.Year) {
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

                if ($actualWeek -in @(52,53)) {
                    # Year transition case for week 52/53
                    $HeaderText = "$($LanguageConfig.WeekPrefix)$actualWeek ($($startDate.Year))"
                } elseif ($actualWeek -in @(1,2)) {
                    # Year transition case for week 1/2
                    $HeaderText = "$($LanguageConfig.WeekPrefix)$actualWeek ($($endDate.Year))"
                } elseif ($groupingCounter -eq 1) {
                    # First week of the group - show ending year or starting year week number is 1
                    if ($actualWeek -eq 1) {
                        $HeaderText = "$($LanguageConfig.WeekPrefix)$actualWeek ($($startDate.Year))"
                    } else {
                        $HeaderText = "$($LanguageConfig.WeekPrefix)$actualWeek ($($endDate.Year))"
                    }
                } else {
                    $HeaderText = "$($LanguageConfig.WeekPrefix)$actualWeek"
                }
                
                Add-TableRow -Table $table -RowNumber $CurrentRow -Text $HeaderText -FontFamily $FontFamily -FontSize $HeaderFontSize -Bold -IsHeader

                # Row 2: Subtitle
                $CurrentRow++
                if ($startDate.Month -ne $endDate.Month) {
                    if ($startDate.Year -ne $endDate.Year) {
                        $SubtitleText = $LanguageConfig.DateFormat.DifferentYear -f $startDate.ToString("d MMMM yyyy"), $endDate.ToString("d MMMM yyyy")
                    } else {
                        $SubtitleText = $LanguageConfig.DateFormat.DifferentMonth -f $startDate.ToString("d MMM"), $endDate.ToString("d MMMM yyyy")
                    }
                } else {
                    $SubtitleText = $LanguageConfig.DateFormat.SameMonth -f $startDate.Day, $endDate.ToString("d MMMM")
                }
                
                Add-TableRow -Table $table -RowNumber $CurrentRow -Text $SubtitleText -FontFamily $FontFamily -FontSize 8 -Bold -FontColor ([System.Drawing.Color]::FromArgb(128, 0, 0))

                # Row 3: Spacer
                $CurrentRow++
                Add-TableRow -Table $table -RowNumber $CurrentRow -Text "" -FontFamily $FontFamily -FontSize 8

                # Rows 4 to 8: Days
                $days = $LanguageConfig.Days
                for ($i = 0; $i -lt $days.Count; $i++) {
                    $CurrentRow++
                    Add-TableRow -Table $table -RowNumber $CurrentRow -Text "$($days[$i])`t:" -FontFamily $FontFamily -FontSize 8
                }

                if ($groupingCounter -gt 0 -and $groupingCounter % 4 -eq 0) {
                    # Page break on every multiple of 4 - insert break at current position
                    $doc.Range($doc.Content.End - 1, $doc.Content.End - 1).InsertBreak($Constants.WD_SECTION_BREAK_NEXT_PAGE)
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
                DateFormat = @{
                    SameMonth = "Du {0} au {1}."
                    DifferentMonth = "Du {0} au {1}."
                    DifferentYear = "Du {0} au {1}."
                }
            }
            English = @{
                Days = @("MON", "TUE", "WED", "THU", "FRI")
                WeekPrefix = "WK. "
                DateFormat = @{
                    SameMonth = "From {0} to {1}."
                    DifferentMonth = "From {0} to {1}."
                    DifferentYear = "From {0} to {1}."
                }
            }
            German = @{
                Days = @("MON", "DIE", "MIT", "DON", "FRE")
                WeekPrefix = "KW. "
                DateFormat = @{
                    SameMonth = "Vom {0} bis {1}."
                    DifferentMonth = "Vom {0} bis {1}."
                    DifferentYear = "Vom {0} bis {1}."
                }
            }
            Spanish = @{
                Days = @("LUN", "MAR", "MIÉ", "JUE", "VIE")
                WeekPrefix = "SEM. "
                DateFormat = @{
                    SameMonth = "Del {0} al {1}."
                    DifferentMonth = "Del {0} al {1}."
                    DifferentYear = "Del {0} al {1}."
                }
            }
            Italian = @{
                Days = @("LUN", "MAR", "MER", "GIO", "VEN")
                WeekPrefix = "SETT. "
                DateFormat = @{
                    SameMonth = "Dal {0} al {1}."
                    DifferentMonth = "Dal {0} al {1}."
                    DifferentYear = "Dal {0} al {1}."
                }
            }
            Portuguese = @{
                Days = @("SEG", "TER", "QUA", "QUI", "SEX")
                WeekPrefix = "SEM. "
                DateFormat = @{
                    SameMonth = "De {0} a {1}."
                    DifferentMonth = "De {0} a {1}."
                    DifferentYear = "De {0} a {1}."
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
            
            [switch]$Bold,
            
            [System.Drawing.Color]$FontColor,
            
            [switch]$IsHeader,
            
            [switch]$HasBorders = $true
        )
        
        $Constants = Get-WordConstants -All
        
        try {
            # Get the cell and clear it
            $cell = $Table.Cell($RowNumber, 1)
            $cell.Range.Text = "" # Required to avoid Word reusing the previous range!
            Set-CellBorders -Cell $cell -LineStyle $Constants.WD_LINE_STYLE_NONE
            
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
            
            # Apply borders for non-header rows
            if (-not $IsHeader -and $HasBorders) {
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