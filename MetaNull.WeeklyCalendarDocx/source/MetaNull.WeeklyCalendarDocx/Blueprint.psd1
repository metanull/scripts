﻿@{
    # Where the cod eis stored
    Source = './source'

    # Where to write the 'built' module
    Destination = './build'

    # Name used to identify this module
    Name = 'MetaNull.WeeklyCalendarDocx'

    # Types & Dependencies
    AssemblyDependencies = @(
        'System.Drawing.Color'
    )
    ModuleDependencies = @(

    )

    # Module Settings
    ModuleSettings = @{
        # (!Overwritten by the build)
        # Script module or binary module file associated with this manifest.
        # RootModule = ''

        # Supported PSEditions
        # CompatiblePSEditions = @()

        # (!Overwritten by the build)
        # Version number of this module.
        # ModuleVersion = '0.0.1'

        # ID used to uniquely identify this module
        GUID              = '26ba686c-8dc4-466a-a3d9-d0e9deaf9e98'

        # Project URI
        ProjectUri        = 'https://github.com/metanull/scripts'

        # Author of this module
        Author            = 'Pascal Havelange'

        # Company or vendor of this module
        CompanyName       = 'Unknown'

        # Copyright statement for this module
        Copyright         = '© 2025. You are free to use, modify, and distribute this software without restriction.'

        # Description of the functionality provided by this module
        Description       = 'Generate a Weekly Calendar Word document'

        # Minimum version of the Windows PowerShell engine required by this module
        PowerShellVersion = '5.1'

        # Name of the PowerShell host required by this module
        # PowerShellHostName = ''

        # Minimum version of the PowerShell host required by this module
        # PowerShellHostVersion = ''

        # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
        DotNetFrameworkVersion = '4.8.1'

        # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
        # CLRVersion = ''

        # Processor architecture (None, X86, Amd64) required by this module
        # ProcessorArchitecture = ''

        # Modules that must be imported into the global environment prior to importing this module
        # RequiredModules = @()

        # Assemblies that must be loaded prior to importing this module
        # RequiredAssemblies = @()

        # Script files (.ps1) that are run in the caller's environment prior to importing this module.
        # ScriptsToProcess = @()

        # Type files (.ps1xml) to be loaded when importing this module
        # TypesToProcess = @()

        # Format files (.ps1xml) to be loaded when importing this module
        # FormatsToProcess = @()

        # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
        # NestedModules = @()

        # (!Overwritten by the build)
        # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
        # FunctionsToExport = @()

        # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
        # CmdletsToExport = @()

        # Variables to export from this module
        # VariablesToExport = '*'

        # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
        # AliasesToExport = @()

        # DSC resources to export from this module
        # DscResourcesToExport = @()

        # List of all modules packaged with this module
        # ModuleList = @()

        # List of all files packaged with this module
        # FileList = @()

        # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
        # PrivateData = @{)

        # HelpInfo URI of this module
        # HelpInfoURI = ''

        # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
        # DefaultCommandPrefix = ''

    }
}
