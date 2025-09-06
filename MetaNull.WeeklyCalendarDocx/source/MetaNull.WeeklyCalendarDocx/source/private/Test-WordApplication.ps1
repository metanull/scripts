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