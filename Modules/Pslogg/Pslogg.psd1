﻿#
# Module manifest for module 'Pslogg'
#
# Generated by: SimonE
#
# Generated on: 22/02/2018
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'Pslogg.psm1'

# Version number of this module.
ModuleVersion = '3.2.0'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = 'ab80676f-a281-44d0-b44e-feffad061630'

# Author of this module
Author = 'Simon Elms'

# Copyright statement for this module
Copyright = '(c) 2018 Simon Elms. All rights reserved.'

# Description of the functionality provided by this module
Description = 'A PowerShell module for logging messages to the host, to PowerShell streams, or to a log file.'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '5.1'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# DotNetFrameworkVersion = ''

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

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = @('Write-LogMessage', 'Get-LogConfiguration', 'Set-LogConfiguration', 'Reset-LogConfiguration')

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = @()

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = @()

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        # NOTE: Do not enclose tags in @() if this module needs to be compatible with PowerShell 3.
        Tags = 'Log', 'Logging', 'Logger'

        # A URL to the license for this module.
        LicenseUri = 'https://opensource.org/licenses/isc-license.txt'

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/AnotherSadGit/Pslogg_PowerShellLogger'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        ReleaseNotes = 'Allow log file to be overwritten if LogFile.Overwrite set during a session'

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}