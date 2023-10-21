# Pslogg
A PowerShell module for logging messages to the PowerShell host or to PowerShell streams, and optionally to a log file.

|                               |                                                                                                            |
------------------------------ | -----------------------------------------------------------------------------------------------------------
**Most recent build status**   | [![Build status](https://ci.appveyor.com/api/projects/status/4ewhdwapoynjnta9?svg=true)](https://ci.appveyor.com/project/AnotherSadGit/pslogg-powershelllogger)
**Master branch build status** | [![Build status](https://ci.appveyor.com/api/projects/status/4ewhdwapoynjnta9/branch/master?svg=true)](https://ci.appveyor.com/project/AnotherSadGit/pslogg-powershelllogger/branch/master)

## Getting Started
See the [Quick Start](https://github.com/AnotherSadGit/Pslogg_PowerShellLogger/wiki/Quick-Start) wiki page.

## Features

The Pslogg module exports four functions:

1. **[Write-LogMessage](https://github.com/AnotherSadGit/Pslogg_PowerShellLogger/wiki/Write‐LogMessage)**:  Writes log messages to the PowerShell host or to a PowerShell stream, and optionally to a log file;

2. **[Get-LogConfiguration](https://github.com/AnotherSadGit/Pslogg_PowerShellLogger/wiki/Get‐LogConfiguration)**:  Retrieves a hash table which is a copy of the current configuration settings of the Pslogg module;

3. **[Set-LogConfiguration](https://github.com/AnotherSadGit/Pslogg_PowerShellLogger/wiki/Set‐LogConfiguration)**:  Sets one or more configuration settings.  Use this function to 
set up the Pslogg module prior to writing any log messages;

4. **[Reset-LogConfiguration](https://github.com/AnotherSadGit/Pslogg_PowerShellLogger/wiki/Reset‐LogConfiguration)**:  Resets the configuration back to its default settings.

## Further Information
See the [wiki](https://github.com/AnotherSadGit/Pslogg_PowerShellLogger/wiki) pages for more information.