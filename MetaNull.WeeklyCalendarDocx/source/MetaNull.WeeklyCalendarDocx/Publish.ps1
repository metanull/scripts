#requires -module Microsoft.PowerShell.PSResourceGet
<#
.SYNOPSIS
    Publish the module to the repository
.DESCRIPTION
    Publish the module to the repository using the Microsoft.PowerShell.PSResourceGet module
.PARAMETER Path
    The path to the module to publish. If not specified (default), Finds and publishes the latest of all builds found in the same location as the script.
.PARAMETER RepositoryName
    The name of the repository to publish to. Default is PSGallery.
    Additional repositories can be configured using the PSRepository cmdlets (see: Register-PSResourceRepository).
.PARAMETER Credential
    The credential to use to publish the module. This should be a PSCredential object (see: Get-Credential or Get-Secret).
.OUTPUTS
    None
#>
[CmdletBinding(DefaultParameterSetName = 'psgallery')]
param(
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [AllowEmptyString()]
    [ValidateScript({
        if($null -eq $_ -or [string]::empty -eq $_) {
            return $true
        }
        return Test-Path -Path $_ -PathType Container
    })]
    [string] $Path,

    [Parameter(Mandatory = $false)]
    [string] $RepositoryName = 'PSGallery',

    [Parameter(Mandatory)]
    [System.Management.Automation.Credential()]
    [System.Management.Automation.PSCredential]
    $Credential
)
Process {
    # Set ErrorAction to Stop
    $BackupErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Stop'

        if(-not $Path) {
            "Loading the Build Settings" | Write-Verbose
            $Build = Import-PowerShellDataFile -Path (Join-Path -Path $PSScriptRoot -ChildPath '.\Blueprint.psd1' -Resolve)

            $ModuleVersion = Import-PowerShellDataFile -Path (Join-Path -Path $PSScriptRoot -ChildPath '.\Version.psd1' -Resolve)
            $ModuleVersion = [version]::new("$($ModuleVersion.Major).$($ModuleVersion.Minor).$($ModuleVersion.Build).$($ModuleVersion.Revision)")

            $Path = Join-Path -Resolve -Path $PSScriptRoot -ChildPath "$($Build.Destination)/$($Build.name)/$($ModuleVersion.ToString())"
        }

        Publish-PSResource -Path $Path -Repository $RepositoryName -ApiKey $Credential.GetNetworkCredential().Password -SkipDependenciesCheck
    } finally {
        # Restore ErrorAction
        $ErrorActionPreference = $BackupErrorActionPreference
    }
}