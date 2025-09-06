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