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