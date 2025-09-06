<#
.SYNOPSIS
    Build the module from the source code
.DESCRIPTION
    This script builds the module from the source code. It creates the module's manifest, the module's code file, the module's class file, and the module's init script file.
    It also increments the build number.
.PARAMETER IncrementMajor
    If set, increment the major version number (and reset the minor and revision numbers); the build number is always incremented.
    If set, IncrementMinor and IncrementRevision are ignored.
.PARAMETER IncrementMinor
    Increment the minor version number (and reset the revision number); the build number is always incremented.
    If set, IncrementRevision is ignored.
.PARAMETER IncrementRevision
    Increment the revision number; the build number is always incremented.
.OUTPUTS
    The path where the module has been built
#>
[CmdletBinding(DefaultParameterSetName = 'IncrementBuild')]
param(
    [switch] $IncrementMajor,

    [switch] $IncrementMinor,

    [switch] $IncrementRevision
)
Begin {
    $EPB = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Stop'

        $BuildDefinition = (Join-Path -Path $PSScriptRoot -ChildPath '.\Blueprint.psd1' -Resolve)

        # 1. Load the Build Settings
        Write-Verbose "Loading Build settings from $BuildDefinition"
        $Build = Import-PowerShellDataFile -Path $BuildDefinition

        Write-Verbose "Building into $($Build.Destination)"

        # 2. Populate build settings with calculated values
        Write-Verbose "Populating build settings"
        $Build.CurrentYMD = Get-Date -Format 'yyyy.MM.dd'

        # 2.1 Create the destination directory (where the module code is placed)
        if(-not (Test-Path -Path $Build.Destination)) {
            Write-Debug "Creating the destination directory $($Build.Destination)"
            New-Item -Path $Build.Destination -ItemType Container | Out-null
        }

        # 2.2 Resolve source and destination root directories (get their fully qualified path)
        $Build.Destination = Join-Path -Path $PSScriptRoot -ChildPath $Build.Destination -Resolve
        $Build.Source = Join-Path -Path $PSScriptRoot -ChildPath $Build.Source -Resolve
        Write-Debug "Source: $($Build.Source)"
        Write-Debug "Destination: $($Build.Destination)"

        # 2.3 Add directory structure if required
        @('init','tools','class','public','private') | Foreach-Object {
            if(-not (Test-Path (Join-Path -Path $Build.Source -ChildPath $_) )) {
                Write-Debug "Creating the directory $(Join-Path -Path $Build.Source -ChildPath $_)"
                New-Item -Path $Build.Source -Name $_ -ItemType Directory | Out-Null
            }
        }

        # 2.4 Load and manage the build number
        $Build.VersionPath = Join-Path -Path $PSScriptRoot -ChildPath "Version.psd1"
        if(-not (Test-Path -Path $Build.VersionPath)) {
            Write-Debug "Creating the version file $($Build.VersionPath)"
            $Item = New-Item -Path $Build.VersionPath -ItemType File
            '@{Major=0;Minor=0;Build=0;Revision=0}' | Set-Content $Item
        }

        $Version = Import-PowerShellDataFile -Path $Build.VersionPath
        if($IncrementMajor.IsPresent -and $IncrementMajor) {
            $Version.Major ++
            $Version.Minor = 0
            $Version.Revision = 0
        }
        elseif($IncrementMinor.IsPresent -and $IncrementMinor) {
            $Version.Minor ++
            $Version.Revision = 0
        }
        elseif($IncrementRevision.IsPresent -and $IncrementRevision) {
            $Version.Revision ++
        }
        $Version.Build ++
        $Build.Version = [version]::new("$($Version.Major).$($Version.Minor).$($Version.Build).$($Version.Revision)")
        Write-Verbose "Version: $($Build.Version)"

        # 2.5 Create the versionned destination directory
        $Destination = Join-Path (Join-Path $Build.Destination $Build.Name) $Build.Version
        if((Test-Path -Path $Destination)) {
            Write-Warning 'Destination exists!'
            Remove-Item -Force -Path $Destination -Confirm
        }
        Write-Debug "Creating the versionned destination directory $Destination"
        New-Item -Path $Destination -ItemType Container | Out-null
        $Build.Destination = Resolve-Path -Path $Destination

        # 2.5.2 Create the tools subdirectory in the module
        Write-Debug "Creating the tools subdirectory in the module"
        New-Item -Path (Join-Path $Destination 'tools') -ItemType Container | Out-null

        # 2.6 Calculate the fully qualified path of the Module's code file
        $Build.ManifestPath = Join-Path -Path $Build.Destination -ChildPath "$($Build.Name).psd1"
        Write-Debug "Creating the module's manifest file $($Build.ManifestPath)"
        New-Item -Path $Build.ManifestPath -ItemType File | Out-null

        # 2.7 Calculate the fully qualified path of the Module's manifest
        $Build.ModulePath = Join-Path -Path $Build.Destination -ChildPath "$($Build.Name).psm1"
        Write-Debug "Creating the module's code file $($Build.ModulePath)"
        New-Item -Path $Build.ModulePath -ItemType File | Out-null

        # 2.8 Calculate the fully qualified path of the Module's init script file (this script will be executed, in the client's scope while module is getting loaded)
        $Build.InitScriptPath = Join-Path -Path $Build.Destination -ChildPath "tools\init.ps1"
        Write-Debug "Creating the module's init script file $($Build.InitScriptPath)"
        New-Item -Path $Build.InitScriptPath -ItemType File | Out-null

        # 2.9 Calculate the fully qualified path of the Module's class script file (there are some limitation for classes exposed in this way, see https://stackoverflow.com/questions/31051103/how-to-export-a-class-in-a-powershell-v5-module)
        $Build.ClassScriptPath = Join-Path -Path $Build.Destination -ChildPath "tools\classes.ps1"
        Write-Debug "Creating the module's class script file $($Build.ClassScriptPath)"
        New-Item -Path $Build.ClassScriptPath -ItemType File | Out-null
    } catch {
        Write-Error $_
    } finally {
        $ErrorActionPreference = $EPB
    }
}
End {
    Write-Verbose "Incrementing the build number"
    # Increment the build number, and save
    Clear-Content -Path $Build.VersionPath
    '@{' | Set-Content -Path $Build.VersionPath
    "  Major = $($Build.Version.Major)" | Add-Content -Path $Build.VersionPath
    "  Minor = $($Build.Version.Minor)" | Add-Content -Path $Build.VersionPath
    "  Revision = $($Build.Version.Revision)" | Add-Content -Path $Build.VersionPath
    "  Build = $($Build.Version.Build)" | Add-Content -Path $Build.VersionPath
    '}' | Add-Content -Path $Build.VersionPath

    Write-Debug "New version number: $($Build.Version)"
}
Process {
    # Stores the list of public functions exposed by the module
    $FunctionsToExport = @()

    # Stores the list of TypeData (for module's classes) exposed by the module
    $TypesToExport = @()

    # Stores the list of FormatData (for module's classes) exposed by the module
    $FormatsToExport = @()

    # Load dependencies INTO THE BUILD process, if required
    Write-Debug "Processing .Net dependencies"
    $Build.AssemblyDependencies | ForEach-Object {
        $AssemblyName = $_ -replace '\.[^\.]+$'
        $ImportScript = @"
if (-not ("$($_)" -as [Type])) {
    Add-Type -Assembly $($AssemblyName)
}

"@
        $ImportScript | Add-Content -Path $Build.InitScriptPath -Encoding UTF8BOM
    }

    # Register module dependencies
    Write-Debug "Processing module dependencies"
    $Build.ModuleDependencies | ForEach-Object {
        "#Requires -Module $($_)" | Add-Content -Path $Build.ModulePath -Encoding UTF8BOM

        # Add the constraint also to the init script
        "#Requires -Module $($_)" | Add-Content -Path $Build.InitScriptPath -Encoding UTF8BOM
    }

    # 1. Build the Module's code file
    Write-Verbose "Building the module's code file"
    # 1.0 Copy the README file
    Get-ChildItem -Path $Build.Source -File -Filter '*.md' | Foreach-Object {
        if([System.IO.Path]::GetExtension($_) -eq '.md' ) {
            Write-Debug "Copying the README file $($_)"
            $_ | Copy-Item -Destination $Build.Destination
        }
    }

    # 1.1 Add code on top of the module, before any Function definition
    Get-ChildItem -Path (Join-Path -Path $Build.Source -ChildPath 'init') -File -Filter '*.ps1' | Foreach-Object {
        if([System.IO.Path]::GetExtension($_) -eq '.ps1' ) {
            Write-Debug "Adding code from $($_)"
            Get-Content -Path $_.FullName | Add-Content -Path $Build.ModulePath -Encoding UTF8BOM
        }
    }
    # 1.2 Add private functions (not exposed by the module)
    Get-ChildItem -Path (Join-Path -Path $Build.Source -ChildPath 'private') -File -Filter '*.ps1' | Foreach-Object {
        if([System.IO.Path]::GetExtension($_) -eq '.ps1' ) {
            $FunctionName  = $_.Name -replace '\.ps1$'
            Write-Debug "Adding private function $($FunctionName) from $($_)"
            "Function $($FunctionName) {" | Add-Content -Path $Build.ModulePath -Encoding UTF8BOM
            Get-Content -Path $_.FullName | Add-Content -Path $Build.ModulePath -Encoding UTF8BOM
            '}' | Add-Content -Path $Build.ModulePath -Encoding UTF8BOM
        }
    }
    # 1.3 Add public functions (exposed by the module)
    Get-ChildItem -Path (Join-Path -Path $Build.Source -ChildPath 'public') -File -Filter '*.ps1' | Foreach-Object {
        if([System.IO.Path]::GetExtension($_) -eq '.ps1' ) {
            $FunctionName  = $_.Name -replace '\.ps1$'
            Write-Debug "Adding public function $($FunctionName) from $($_)"
            "Function $($FunctionName) {" | Add-Content -Path $Build.ModulePath -Encoding UTF8BOM
            Get-Content -Path $_.FullName | Add-Content -Path $Build.ModulePath -Encoding UTF8BOM
            '}' | Add-Content -Path $Build.ModulePath -Encoding UTF8BOM
            $FunctionsToExport += $FunctionName
        }
    }
    # 1.4 Add class definitions
    # 1.4.a Add class' code to the module's class file
    Get-ChildItem -Path (Join-Path -Path $Build.Source -ChildPath 'class') -File -Filter '*.ps1' | Foreach-Object {
        if([System.IO.Path]::GetExtension($_) -eq '.ps1' ) {
            Write-Debug "Adding class code from $($_)"
            Get-Content -Path $_.FullName | Add-Content -Path $Build.ClassScriptPath -Encoding UTF8BOM
        }
    }
    # # 1.4.b Add Init code to make the class modules available to the calling process when they use Import-Module
    # if( Get-ChildItem -Path (Join-Path -Path $Build.Source -ChildPath 'class') -File -Filter '*.ps1' | Where-Object {[System.IO.Path]::GetExtension($_) -eq '.ps1' }) {
    #     # "# using module $($Build.ModulePath)" | Add-Content -Path $Build.InitScriptPath -Encoding UTF8BOM
    #     "# using module $($Build.Name)" | Add-Content -Path $Build.InitScriptPath -Encoding UTF8BOM
    # }

    # 1.4.b Copy "TypesData" file to the module's root directory
    Get-ChildItem -Path (Join-Path -Path $Build.Source -ChildPath 'class') -File -Filter '*.types.ps1xml' | Foreach-Object {
        if([System.IO.Path]::GetExtension($_) -eq '.ps1xml' ) {
            write-debug "Copying the TypesData file $($_)"
            $TypesToExport += $_.Name
            $_ | Copy-Item -Destination $Build.Destination
        }
    }
    # 1.4.b Copy "FormatData" file to the module's root directory
    Get-ChildItem -Path (Join-Path -Path $Build.Source -ChildPath 'class') -File -Filter '*.formats.ps1xml' | Foreach-Object {
        if([System.IO.Path]::GetExtension($_) -eq '.ps1xml' ) {
            write-debug "Copying the FormatData file $($_)"
            $FormatsToExport += $_.Name
            $_ | Copy-Item -Destination $Build.Destination
        }
    }
    # 1.5 Copy "resource" directory to the module's root directory
    if((Test-Path -Path (Join-Path -Path (Split-Path $Build.Source -Parent) -ChildPath 'resource') -PathType Container)) {
        Write-Debug "Copying the resource directory"
        Copy-Item -Recurse -LiteralPath (Join-Path -Path (Split-Path $Build.Source -Parent) -ChildPath 'resource') -Destination $Build.Destination
    }

    # 2. Build the Module's manifest
    Write-Verbose "Building the module's manifest"
    $ModuleManifest = $Build.ModuleSettings.Clone()
    $ModuleManifest.RootModule = "$($Build.Name).psm1"
    $ModuleManifest.ModuleVersion = $Build.Version
    $ModuleManifest.FunctionsToExport = $FunctionsToExport
    $ModuleManifest.ScriptsToProcess = @(
        "$($Build.InitScriptPath | Split-Path -Parent | Split-Path -Leaf)\$($Build.InitScriptPath | Split-Path -Leaf)"
        "$($Build.ClassScriptPath | Split-Path -Parent | Split-Path -Leaf)\$($Build.ClassScriptPath | Split-Path -Leaf)"
    )
    if($TypesToExport.Length) {
        $ModuleManifest.TypesToProcess = $TypesToExport
    }
    if($FormatsToExport.Length) {
        $ModuleManifest.FormatsToProcess = $FormatsToExport
    }
    # Add the requirement to the Manifest
    $ModuleManifest.RequiredModules = $Build.ModuleDependencies

    # 3. Save the Module's manifest
    Write-Verbose "Saving the module's manifest $($Build.ManifestPath)"
    New-ModuleManifest @ModuleManifest -Path $Build.ManifestPath

    # Return the path where the module has been built
    Get-Item ($Build.ManifestPath | Split-Path)
}