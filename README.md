# Prog
A PowerShell module for logging messages to the host, to PowerShell streams, or to a log file.

### master branch build status: [![build status](https://ci.appveyor.com/api/projects/status/0xu4p7bvxdxgkaql/branch/master?svg=true)](https://ci.appveyor.com/project/AnotherSadGit/prog-powershelllogger/branch/master)
### v0.9 branch build status: [![build status](https://ci.appveyor.com/api/projects/status/0xu4p7bvxdxgkaql/branch/v0.9?svg=true)](https://ci.appveyor.com/project/AnotherSadGit/prog-powershelllogger/branch/v0.9)

## Getting Started
Copy the Prog_PowerShellLoggingModule > Modules > Prog folder, with its contents, to one of the 
locations that PowerShell recognises for modules.  The two default locations are:

1. For all users:  **%ProgramFiles%\WindowsPowerShell\Modules** 
(usually resolves to C:\Program Files\WindowsPowerShell\Modules);

2. For the current user only:  **%UserProfile%\Documents\WindowsPowerShell\Modules** 
(usually resolves to C:\Users\\{user name}\Documents\WindowsPowerShell\Modules)

If the PowerShell console or the PowerShell ISE is open when you copy the Prog folder to a 
recognised module location you may need to close and reopen the console or ISE for it to 
recognise the new Prog module.

Once the Prog folder has been saved to a recognised module location you should be able to call 
the module's functions without explicitly importing the module.

## Features

The Prog module exports four functions:

1) **_Write-LogMessage_**:  Writes log messages to the host or to a PowerShell stream, and 
optionally to a log file;

2) **_Get-LogConfiguration_**:  Retrieves a hash table with the current configuration settings 
of the Prog module;

3) **_Set-LogConfiguration_**:  Sets one or more configuration settings.  Use this function to 
set up the Prog module prior to writing any log messages;

4) **_Reset-LogConfiguration_**:  Resets the configuration back to its default settings.

## Detailed Help for Exported Functions

Once the Prog module has been imported into the local PowerShell session comment-based help can 
be used to view details of each of the functions exported from the module.  For example, to see 
details of the Write-LogMessage function enter the following in the PowerShell console:

```
help Write-LogMessage -full
```

## Usage

### Configuration

Prior to writing log messages, use **_Set-LogConfiguration_** to configure the Prog module.  

_Set-LogConfiguration_ may be used to set the following log properties:

1) **The log level:**  This determines whether a message will be logged or not.  
	
   Possible log levels, in order from lowest to highest, are: 
   * Off
   * Error
   * Warning
   * Information
   * Debug
   * Verbose 

   Only log messages at a level the same as, or lower than, the LogLevel will be logged.  For 
   example, if the LogLevel is "Information" then only log messages at a level of Information, 
   Warning or Error will be logged.  Messages at a level of Debug or Verbose will not be logged, 
   as these log levels are higher than Information;

2) **The message destination:**  Messages may be written to the host or to PowerShell streams 
such as the Information stream or the Verbose stream.  In addition, if a log file name is set in 
the configuration, the messages will be written to the log file;

3) **The host text color:**  Messages written to the host, as opposed to PowerShell streams, may 
be written in any PowerShell console color.  Different colors may be specified for different 
message types, such as Error, Warning or Information;

4) **The message format:**  In addition to the specified message, the text written to the log may 
include additional fields that are automatically populated, such as a timestamp or the name of 
the function writing to the log.  A simple template can be defined to specify the format of the 
logged text, including the fields to be displayed and any field separators;

5) **Whether an existing log file will be overwritten or appended to:**  If a log file is specified 
in the configuration you can determine whether new log messages will overwrite an existing log 
file with the same file name or will be appended to the end of it.  If the option to overwrite an 
existing file is chosen it will only be overwritten by the first message written to the log in a 
given session.  Subsequent messages written in the same session will be appended to the log file.

The configuration can be updated by _Set-LogConfiguration_ in two different ways:

1) **Use parameter -LogConfiguration:**  Pass in a hash table representing all the configuration 
settings.  This can be used together with function _Get-LogConfiguration_:  Use 
_Get-LogConfiguration_ to read the current configuration out as a hash table, update that hash 
table with new settings, then finally write the updated hash table back using 
`Set-LogConfiguration -LogConfiguration <hash table>`;

2) **Use different parameters to update individual configuration settings:**  For example, 
parameter _-LogLevel_ can be used to set the configured log level, and parameter _-LogFileName_ 
can be used to set the configured log file name.  

When updating the host text colors you can either update them all at once, passing a hash table 
into parameter _-HostTextColorConfiguration_, or you can update individual colors with individual 
parameters such as _-ErrorTextColor_ and _-WarningTextColor_.

#### Examples

##### Set the LogLevel and MessageFormat using individual parameters:
```
Set-LogConfiguration -LogLevel Warning `
	-MessageFormat '{Timestamp:yyyy-MM-dd hh:mm:ss},{CallingObjectName},{Message}'
```

##### Set the text colors used by the host to display error and warning messages:
```
Set-LogConfiguration -ErrorTextColor Magenta -WarningTextColor DarkYellow
```

##### Set all text colors simultaneously:
```
$hostColors = @{
					Error = "DarkRed"
					Warning = "DarkYellow"
					Information = "DarkCyan"
					Debug = "Cyan"
					Verbose = "Gray"
					Success = "Green"
					Failure = "Red"
					PartialFailure = "Yellow"
				}
Set-LogConfiguration -HostTextColorConfiguration $hostColors
```

##### Use parameter `-LogConfiguration` to update the entire configuration at once:
```
$configuration = Get-LogConfiguration

$configuration.LogLevel = Verbose
$configuration.IncludeDateInFileName = $False
$configuration.OverwriteLogFile = $False
$configuration.MessageFormat = '{Timestamp:T}: {Message}'

Set-LogConfiguration -LogConfiguration $configuration
```

### Writing Log Messages
Messages are logged using the **Write-LogMessage** function.

#### Examples

##### Write error message to log:
```
try
{
	...
}
catch [System.IO.FileNotFoundException]
{
	Write-LogMessage -Message "Error while updating file: $_.Exception.Message" -IsError
}
```

##### Write debug message to log:
```
Write-LogMessage -Message "Updating user settings for $userName..." -IsDebug
```

### Overriding Configuration for a Single Message
Some log properties can be overridden when writing a single log message.  The changes apply only 
to that one message; subsequent messages will return to using the settings specified via 
_Set-LogConfiguration_.

#### Examples

##### Write message to host in a specified color:
```
Write-LogMessage -Message 'Updating user settings...' -WriteToHost -HostTextColor Cyan
```

##### Write message with a custom message format which only applies to this one message:
```
Write-LogMessage -Message "***** Running on server: $serverName *****" -MessageFormat '{message}'
```
The message written to the log will only include the specified message text.  It will not 
include other fields, such as {Timestamp} or {CallingObjectName}.

##### Write message to the Debug PowerShell stream, rather than to the host:
```
Write-LogMessage -Message "Updating user settings for $userName..." -WriteToStreams -IsDebug
```
