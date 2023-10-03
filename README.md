# Pslogg
A PowerShell module for logging messages to the host, to PowerShell streams, or to a log file.

|                               |                                                                                                            |
------------------------------ | -----------------------------------------------------------------------------------------------------------
**Most recent build status**   | [![Build status](https://ci.appveyor.com/api/projects/status/4ewhdwapoynjnta9?svg=true)](https://ci.appveyor.com/project/AnotherSadGit/pslogg-powershelllogger)
**Master branch build status** | [![Build status](https://ci.appveyor.com/api/projects/status/4ewhdwapoynjnta9/branch/master?svg=true)](https://ci.appveyor.com/project/AnotherSadGit/pslogg-powershelllogger/branch/master)

## Getting Started
There are two ways of installing the Pslogg module:  from the PowerShell Gallery via PowerShellGet 
or Manually:

### Installing from the PowerShell Gallery via PowerShellGet
You will need to run the following commands in a console or terminal with **Administrator privileges**.

##### If you have direct access to the internet:
```powershell
install-module -Name Pslogg -Repository 'PSGallery'
```
**NOTE:** If you get an error message similar to:<br/>
*WARNING: Source Location 'https://www.powershellgallery.com/api/v2/package/Pslogg/2.0.0' is not valid. 
PackageManagement\Install-Package : Package 'Pslogg' failed to download.*<br/>
then you are probably behind a proxy server.  See how to install the module from behind a proxy, below.

##### If you're behind a proxy server:
```powershell
$proxyCredential = Get-Credential -Message 'Please enter credentials for proxy server'
install-module -Name Pslogg -Repository 'PSGallery' `
    -Proxy 'http://...' -ProxyCredential $proxyCredential
```
(replace the 'http://...' with the correct URL for your proxy server)

##### To check if the module is installed:
```powershell
get-installedmodule -Name Pslogg
```

**NOTE:** You may get an error message along the lines of:<br/>
*"PowerShellGet requires NuGet provider version '2.8.5.201' or newer to interact with NuGet-based repositories."*<br/>
See the following document from Microsoft to resolve this issue:<br/>
*"Bootstrap the NuGet provider and NuGet.exe"*<br/>
at https://docs.microsoft.com/en-us/powershell/gallery/how-to/getting-support/bootstrapping-nuget

### Manually
Copy the Pslogg_PowerShellLogger > Modules > Pslogg folder, with its contents, to one of the 
locations that PowerShell recognizes for modules.  The two default locations are:

1. For all users:  **%ProgramFiles%\WindowsPowerShell\Modules** 
(usually resolves to C:\Program Files\WindowsPowerShell\Modules);

2. For the current user only:  **%UserProfile%\Documents\WindowsPowerShell\Modules** 
(usually resolves to C:\Users\\{user name}\Documents\WindowsPowerShell\Modules)

If the PowerShell console or the PowerShell ISE is open when you copy the Pslogg folder to a 
recognized module location you may need to close and reopen the console or ISE for it to 
recognize the new Pslogg module.

Once the Pslogg folder has been saved to a recognized module location you should be able to call 
the module's functions without explicitly importing the module.

## Features

The Pslogg module exports four functions:

1. **_Write-LogMessage_**:  Writes log messages to the host or to a PowerShell stream, and 
optionally to a log file;

2. **_Get-LogConfiguration_**:  Retrieves a hash table which is a copy of the current configuration 
settings of the Pslogg module;

3. **_Set-LogConfiguration_**:  Sets one or more configuration settings.  Use this function to 
set up the Pslogg module prior to writing any log messages;

4. **_Reset-LogConfiguration_**:  Resets the configuration back to its default settings.

## Detailed Help for Exported Functions

Once the Pslogg module has been imported into the local PowerShell session comment-based help can 
be used to view details of each of the functions exported from the module.  For example, to see 
details of the Write-LogMessage function enter the following in the PowerShell console:

```powershell
help Write-LogMessage -full
```

## Usage

### Configuration

Prior to writing log messages, use **_Set-LogConfiguration_** to configure the Pslogg module.  

_Set-LogConfiguration_ may be used to set the following log properties:

1. **The log level:**  This determines whether a message will be logged or not.  
	
   Possible log levels, in order from lowest to highest, are: 
   * OFF
   * ERROR
   * WARNING
   * INFORMATION
   * DEBUG
   * VERBOSE 

   Only log messages at a level the same as, or lower than, the LogLevel will be logged.  For 
   example, if the LogLevel is INFORMATION then only log messages at a level of INFORMATION, 
   WARNING or ERROR will be logged.  Messages at a level of DEBUG or VERBOSE will not be logged, 
   as these log levels are higher than INFORMATION;

2. **The message destination:**  Messages may be written to the host or to PowerShell streams 
such as the Information stream or the Verbose stream.  In addition, if a log file name is set in 
the configuration, the messages will be written to the log file;

3. **The host text color:**  Messages written to the host, as opposed to PowerShell streams, may 
be written in any PowerShell console color.  Different colors may be specified for different 
message types, such as Error, Warning or Information.  Different colors may also be specified 
for different message categories;

4. **The message format:**  In addition to the specified message, the text written to the log may 
include additional fields that are automatically populated, such as a timestamp or the name of 
the function writing to the log.  A simple template can be defined to specify the format of the 
logged text, including the fields to be displayed and any field separators;

5. **Whether an existing log file will be overwritten or appended to:**  If a log file is specified 
in the configuration you can determine whether new log messages will overwrite an existing log 
file with the same file name or will be appended to the end of it.  If the option to overwrite an 
existing file is chosen it will only be overwritten by the first message written to the log in a 
given session.  Subsequent messages written in the same session will be appended to the log file.

The configuration can be updated by _Set-LogConfiguration_ in two different ways:

1. **Use parameter -LogConfiguration:**  Pass in a hash table representing all the configuration 
settings.  This can be used together with function _Get-LogConfiguration_:  Use 
_Get-LogConfiguration_ to read the current configuration out as a hash table, update that hash 
table with new settings, then write the updated hash table back using 
`Set-LogConfiguration -LogConfiguration <hash table>`;

2. **Use different parameters to update individual configuration settings:**  For example, 
parameter _-LogLevel_ can be used to set the configured log level, and parameter _-LogFileName_ 
can be used to set the configured log file name.  

When updating the host text colors you can either update them all at once, passing a hash table 
into parameter _-HostTextColorConfiguration_, or you can update individual colors with individual 
parameters such as _-ErrorTextColor_ and _-WarningTextColor_.

#### Get-LogConfiguration
**_Get-LogConfiguration_** retrieves a copy of the Pslogg configuration hash table, NOT a reference to 
the live configuration.  As a result the Pslogg configuration can only be updated via 
_Set-LogConfiguration_.  This ensures that the Pslogg internal state is updated correctly.  

For example, if a user were able to use _Get-LogConfiguration_ to access the live configuration 
and modify it to set the configuration MessageFormat string directly, the modified MessageFormat 
would not be used when writing log messages.  That is because _Set-LogConfiguration_ parses the 
new MessageFormat string and updates Pslogg's internal state to indicate which fields are to be 
included in log messages.  If the configuration MessageFormat string were updated directly it 
would not be parsed and the list of fields to include in log messages would not be updated.

Although changes to the hash table retrieved by _Get-LogConfiguration_ will not be reflected in 
the Pslogg configuration, the updated hash table can be written back into the Pslogg configuration 
via _Set-LogConfiguration_.  

#### Examples

##### Get the text color for messages with category Success:
```powershell
PS C:\Users\Me> $config = Get-LogConfiguration
PS C:\Users\Me> $config.CategoryInfo.Success.Color 

    Green
```

##### Get the text colors for all message levels:
```powershell
PS C:\Users\Me> $config = Get-LogConfiguration
PS C:\Users\Me> $config.HostTextColor 

    Name                 Value
    ----                 -----
    Debug                White
    Error                Red
    Warning              Yellow
    Verbose              White
    Information          Cyan
```

##### Get the text color for messages of level ERROR:
```powershell
PS C:\Users\Me> $config = Get-LogConfiguration
PS C:\Users\Me> $config.HostTextColor.Error 

    Red
```	

##### Get the name of the file messages will be logged to:
```powershell
PS C:\Users\Me> $config = Get-LogConfiguration
PS C:\Users\Me> $config.LogFile.Name 

    Results.log
```
**NOTE:** The name returned is the "raw" file name.  It will not include a date, if Pslogg is 
configured to include dates in log file names.
(Pslogg is configured to include the date in the log file name by setting 
`$config.LogFile.IncludeDateInFileName = $true`, or via 
`Set-LogConfiguration -IncludeDateInFileName`)

##### Get the full path of the file messages will be logged to:
```powershell
PS C:\Users\Me> $config = Get-LogConfiguration
PS C:\Users\Me> $config.LogFile.FullPath 

    C:\Users\Me\Documents\PowerShell\MyTest\Results_20201027.log
```
In contrast to `$config.LogFile.Name`, `$config.LogFile.FullPath` is the absolute path to the 
log file.  It will include the date, if Pslogg is configured to include dates in log file names.

##### Get the format of log messages:
```powershell
PS C:\Users\Me> $config = Get-LogConfiguration
PS C:\Users\Me> $config.MessageFormat 

    {Timestamp:yyyy-MM-dd HH:mm:ss.fff} | {CallerName} | {Category} | {MessageLevel} | {Message}
```

##### Use _Get-LogConfiguration_ and _Set-LogConfiguration_ to update Pslogg's configuration:
```powershell
    $config = Get-LogConfiguration
    $config.LogLevel = 'ERROR'
    $config.LogFile.Name = 'Error.log'
    $config.CategoryInfo['FileCopy'] = @{Color = 'DarkYellow'}
    Set-LogConfiguration -LogConfiguration $config
 ```

##### Set the details of the log file using individual parameters:
```powershell
Set-LogConfiguration -LogFileName 'Debug.log' -ExcludeDateFromFileName -AppendToLogFile
```
	
##### Set the LogLevel and MessageFormat using individual parameters:
```powershell
Set-LogConfiguration -LogLevel Warning `
    -MessageFormat '{Timestamp:yyyy-MM-dd HH:mm:ss},{Category},{Message}'
```

##### Set the text color used by the host to display messages with category 'FileCopy':
```powershell
Set-LogConfiguration -CategoryInfoItem 'FileCopy', @{ Color = 'Blue' }
```	
	
##### Set the text colors used by the host to display error and warning messages:
```powershell
Set-LogConfiguration -ErrorTextColor DarkRed -WarningTextColor DarkYellow
```

##### Set all text colors simultaneously:
```powershell
$hostColors = @{
					Error = 'DarkRed'
					Warning = 'DarkYellow'
					Information = 'DarkCyan'
					Debug = 'Cyan'
					Verbose = 'Gray'
				}
Set-LogConfiguration -HostTextColorConfiguration $hostColors
```

### Writing Log Messages
Messages are logged using the **Write-LogMessage** function.

#### Examples

##### Write message to log:
```powershell
 Write-LogMessage -Message "Updating user settings for $userName..."
```

The `-Message` parameter is optional:

```powershell
 Write-LogMessage "Updating user settings for $userName..."
```

#### Message Level
The Message Level of a message is compared to the configuration LogLevel to determine whether the 
message gets logged or not.  

Possible values are the LogLevels: 
* ERROR
* WARNING
* INFORMATION
* DEBUG
* VERBOSE

The Message Level is compared to the LogLevel in the logger configuration.  If the Message Level 
is the same as or lower than the LogLevel the message will be logged.  If the Message Level is 
higher than the LogLevel the message will not be logged.

For example, if the LogLevel is INFORMATION then only messages with a Message Level of 
INFORMATION, WARNING or ERROR will be logged.  Messages with a Message Level of DEBUG or 
VERBOSE will not be logged, as those levels are higher than INFORMATION.

The Message Level may be specified as text, via the _-MessageLevel_ parameter, or via a switch 
parameter.  There is one Message Level switch parameter for each Message Level:
_-IsError, -IsWarning, -IsInformation, -IsDebug, -IsVerbose_.

#### Examples

##### Write error message to log:
```powershell
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
```powershell
 Write-LogMessage "Updating user settings for $userName..." -IsDebug
```

##### Specify the MessageLevel as text rather than using a Message Level switch:
```powershell
Write-LogMessage "Updating user settings for $userName..." -MessageLevel 'DEBUG'
```

### Message Category
The Message Category allows logged messages to be categorized which is useful for filtering or 
querying the log.  Any text can be used as a Message Category when writing to the log.  

The Message Category can have a color specified in the configuration CategoryInfo hash table.  If 
the message is being written to the host, as opposed to a PowerShell stream, the ForegroundColor 
will be set to the CategoryInfo color from the configuration.

For example, if the Message Category is 'Success' the text ForegroundColor will be set to the 
color specified by the configuration CategoryInfo.Success.Color, if it exists.

#### Example

##### Write message to the log with a certain category:
```powershell
    Write-LogMessage 'File copy completed successfully.' -Category 'Success' -IsInformation
```

### Overriding Configuration for a Single Message
Some log properties can be overridden when writing a single log message.  The changes apply only 
to that one message; subsequent messages will return to using the settings specified via 
_Set-LogConfiguration_.

#### Examples

##### Write message to host in a specified color:
```powershell
Write-LogMessage 'Updating user settings...' -WriteToHost -HostTextColor Cyan
```

##### Write message with a custom message format which only applies to this one message:
```powershell
Write-LogMessage "***** Running on server: $serverName *****" -MessageFormat '{message}'
```

The message written to the log will only include the specified message text.  It will not 
include other fields, such as {Timestamp} or {CallerName}.

##### Write message to the Debug PowerShell stream, rather than to the host:
```powershell
Write-LogMessage 'Updating user settings...' -WriteToStreams -IsDebug
```
