<#
.SYNOPSIS
Functions for logging messages to the host or to PowerShell streams and, optionally, to a log file.

.DESCRIPTION
A module for logging messages to the host or to PowerShell streams, such as the Error stream or 
the Information stream.  In addition, messages may optionally be logged to a log file.

Messages are logged using the Write-LogMessage function.

The logger may be configured prior to logging any messages via function Set-LogConfiguration.  
For example, the logger may be configured to write to the PowerShell host, or to PowerShell 
streams such as the Error stream or the Verbose stream.

Function Get-LogConfiguration will return a copy of the current logger configuration as a hash 
table.  The configuration can be reset back to its default values via function 
Reset-LogConfiguration.

.NOTES

#>

$_logLevels = @{
                    Off = 0
                    Error = 1
                    Warning = 2
                    Information = 3
                    Debug = 4
                    Verbose = 5
                }

$_defaultHostTextColor = @{
                                Error = 'Red'
                                Warning = 'Yellow'
                                Information = 'Cyan'
                                Debug = 'White'
                                Verbose = 'White'
                            }

$_defaultCategoryInfo = @{
                        Progress = @{ IsDefault = $True }
                        Success = @{ Color = 'Green' }
                        Failure = @{ Color = 'Red' }
                        PartialFailure = @{ Color = 'Yellow' }
                    }

$_defaultLogConfiguration = @{   
                                LogLevel = 'INFORMATION'
								MessageFormat = '{Timestamp:yyyy-MM-dd hh:mm:ss.fff} | {CallerName} | {Category} | {MessageLevel} | {Message}'
                                WriteToHost = $True
                                HostTextColor = $_defaultHostTextColor
                                LogFile = @{
                                                Name = 'Results.log'
                                                IncludeDateInFileName = $True
                                                Overwrite = $True
                                            }
                                CategoryInfo = $_defaultCategoryInfo
                            }

$_defaultTimestampFormat = 'yyyy-MM-dd hh:mm:ss.fff'	
						
$_logConfiguration = @{}
$_messageFormatInfo = @{}

$_logFilePath = ''
$_logFileOverwritten = $False

# Function naming conventions:
# ----------------------------
# Functions to be exported: Follow the standard PowerShell naming convention of 
#                           "<verb>-<singular noun>", eg "Write-LogMessage".
#
# Private functions that 
# shouldn't be exported:    Have a "Private_" prefix.  They must not include a dash, "-".
#
# These naming conventions simplify the exporting of public functions but not private ones 
# because we can then simply export all functions containing a dash, "-".


#region Write Log Messages ************************************************************************

<#
.SYNOPSIS
Writes a message to the host or a stream and, optionally, to a log file.

.DESCRIPTION
Writes a message to either the host or to PowerShell streams such as the Information stream or the 
Verbose stream, depending on the logging configuration.  In addition the message may be written to 
a log file, once again depending on the logging configuration.

.NOTES
The Prog logger can be configured via function Set-LogConfiguration with settings that persist 
between messages.  For example, it can be configured to write to the PowerShell host, or to 
PowerShell streams such as the Error stream or the Verbose stream.

The most important configuration setting is the LogLevel.  This determines which messages will be 
logged and which will not.  

Possible LogLevels, in order from lowest to highest, are:
    OFF
    ERROR
    WARNING
    INFORMATION
    DEBUG
    VERBOSE
        
Each message to be logged has a Message Level.  This may be set explicitly when calling 
Write-LogMessage or the default value of INFORMATION may be used.  The Message Level is compared 
to the LogLevel in the logger configuration.  Only messages with a Message Level the same as or 
lower than the configured LogLevel will be logged.  

For example, if the LogLevel is INFORMATION then only messages with a Message Level of 
INFORMATION, WARNING or ERROR will be logged.  Messages with a Message Level of DEBUG or 
VERBOSE will not be logged, as those levels are higher than INFORMATION.

When calling Write-LogMessage the Message Level can be set in two different ways:

    1) Via parameter -MessageLevel:  The Message Level is specified as text, for example:

        Write-LogMessage 'Hello world' -MessageLevel 'VERBOSE'

    2) Via Message Level switch parameters:  There are switch parameters for each possible 
        Message Level: -IsError, -IsWarning, -IsInformation, -IsDebug and -IsVerbose.  For 
        example:

        Write-LogMessage 'Hello world' -IsVerbose

        Only one Message Level switch may be set for a given message.  

Several configuration settings can be overridden for a single log message.  The changes apply 
only to that one message; subsequent messages will return to using the settings in the logger  
configuration.  Settings that can be overridden on a per-message basis are:

    1) The message destination:  The message can be logged to a different destination from the 
        one specified in the logger configuration by using the switch parameters -WriteToHost or 
        -WriteToStreams;

    2) The host text color:  If the message is being written to the host, as opposed to 
        PowerShell streams, its text color can be set via parameter -HostTextColor.  Any 
        PowerShell console color can be used;

    3) The message format:  The format of the message can be set via parameter -MessageFormat.

.PARAMETER Message
The message to be logged. 

.PARAMETER HostTextColor
The ForegroundColor the message will be written in, if the message is written to the host.

If the message is written to a PowerShell stream, such as the Information stream, this color is 
ignored.

Acceptable values are: 'Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 
'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White'.

.PARAMETER MessageFormat: 
A string that sets the format of the text that will be logged.  

Text enclosed in curly braces, {...}, represents the name of a field which will be included in 
the logged text.  The field names are not case sensitive.  
        
Any other text, not enclosed in curly braces, will be treated as a string literal and will appear 
in the logged text exactly as specified.	
		
Leading spaces in the MessageFormat string will be retained when the text is written to the 
log to allow log messages to be indented.  Trailing spaces in the MessageFormat string will be 
removed, and will not be written to the log.
		
Possible field names are:
	{Message}     : The supplied text message to write to the log;

	{Timestamp}	  : The date and time the log message is recorded.  

                    The Timestamp field may include an optional datetime format string, inside 
                    the curly braces, following the field name and separated from it by a 
                    colon, ':'.  For example, '{Timestamp:T}'.
                            
                    Any .NET datetime format string is valid.  For example, "{Timestamp:d}" will 
                    format the timestamp using the short date pattern, which is "MM/dd/yyyy" in 
                    the US.  
                            
                    While the field names in the MessageFormat string are NOT case sentive the 
                    datetime format string IS case sensitive.  This is because .NET datetime 
                    format strings are case sensitive.  For example, "d" is the short date 
                    pattern while "D" is the long date pattern.  
                            
                    The Timestamp field may be specified without any datetime format string.  For 
                    example, '{Timestamp}'.  In that case the default datetime format string,  
                    'yyyy-MM-dd hh:mm:ss.fff', will be used;

	{CallerName}  : The name of the function or script that is writing to the log.  

                    When determining the caller name all functions in this module will be ignored; 
                    the caller name will be the external function or script that calls into this 
                    module to write to the log.  
                            
                    If a function is writing to the log the function name will be displayed.  If 
                    the log is being written to from a script file, outside any function, the name 
                    of the script file will be displayed.  If the log is being written to manually 
                    from the Powershell console then '[CONSOLE]' will be displayed.

	{Category}    : The Message Category.  If no Message Category is explicitly specified when 
                    calling Write-LogMessage the default Category from the logger configuration 
                    will be used.

	{MessageLevel} : The Message Level at which the message is being recorded.  For example, the 
                    message may be an Error message or a Debug message.  The MessageLevel will 
                    always be displayed in upper case.

.PARAMETER MessageLevel
A string that specifies the Message Level of the message.  Possible values are the LogLevels:
    ERROR
    WARNING
    INFORMATION
    DEBUG
    VERBOSE

The Message Level is compared to the LogLevel in the logger configuration.  If the Message Level 
is the same as or lower than the LogLevel the message will be logged.  If the Message Level is 
higher than the LogLevel the message will not be logged.

For example, if the LogLevel is INFORMATION then only messages with a Message Level of 
INFORMATION, WARNING or ERROR will be logged.  Messages with a Message Level of DEBUG or 
VERBOSE will not be logged, as those levels are higher than INFORMATION.

-MessageLevel cannot be specified at the same time as one of the Message Level switch parameters: 
-IsError, -IsWarning, -IsInformation, -IsDebug or -IsVerbose.  Either -MessageLevel can be 
specified or one or the Message Level switches can be specified but not both.

In addition to determining whether the message will be logged or not, -MessageLevel has the 
following effects:

    1) If the message is set to be written to a PowerShell stream it determines which stream the 
        message will be written to: The Error stream, the Warning stream, the Information stream, 
        the Debug stream or the Verbose stream;

    2) If the message is set to be written to the host and the -HostTextColor parameter is not 
        specified -MessageLevel determines the ForegroundColor the message will be written in.  
        The appropriate color is read from the logger configuration HostTextColor hash table.  
        For example, if the -MessageLevel is ERROR the text ForegroundColor will be set to the 
        color specified by logger configuration HostTextColor.Error;

    3) The {MessageLevel} placeholder in the MessageFormat string, if present, will be replaced 
        by the -MessageLevel text.  For example, if -MessageLevel is ERROR the {MessageLevel} 
        placeholder will be replaced by the text 'ERROR'.

.PARAMETER IsError
Sets the Message Level to ERROR.  

-IsError is one of the Message Type switch parameters.  Only one Message Type switch may be set 
at the same time.  The Message Type switch parameters are:
    -IsError, -IsWarning, -IsInformation, -IsDebug, -IsVerbose.

.PARAMETER IsWarning
Sets the Message Level to WARNING.

-IsWarning is one of the Message Type switch parameters.  Only one Message Type switch may be set 
at the same time.  The Message Type switch parameters are:
    -IsError, -IsWarning, -IsInformation, -IsDebug, -IsVerbose.

.PARAMETER IsInformation
Sets the Message Level to INFORMATION.

-IsInformation is one of the Message Type switch parameters.  Only one Message Type switch may be 
set at the same time.  The Message Type switch parameters are:
    -IsError, -IsWarning, -IsInformation, -IsDebug, -IsVerbose.

.PARAMETER IsDebug
Sets the Message Level to DEBUG.

-IsDebug is one of the Message Type switch parameters.  Only one Message Type switch may be set 
at the same time.  The Message Type switch parameters are:
    -IsError, -IsWarning, -IsInformation, -IsDebug, -IsVerbose.

.PARAMETER IsVerbose
Sets the Message Level to VERBOSE.

-IsVerbose is one of the Message Type switch parameters.  Only one Message Type switch may be set 
at the same time.  The Message Type switch parameters are:
    -IsError, -IsWarning, -IsInformation, -IsDebug, -IsVerbose

.PARAMETER Category
A string that specifies the Message Category of the message.  Any string can be specified.  

The Message Category can have a color specified in the configuration CategoryInfo hash table.  If 
the -HostTextColor parameter is not specified and the message is being written to the host, 
the ForegroundColor will be set to the CategoryInfo color from the configuration.

For example, if the -Category is 'Success' the text ForegroundColor will be set to the logger 
configuration CategoryInfo.Success.Color.

-Category will default to the configuration CategoryInfo name which has IsDefault set.

.PARAMETER WriteToHost
A switch parameter that, if set, will write the message to the host, as opposed to one of the 
PowerShell streams such as Error or Warning, overriding the logger configuration setting 
WriteToHost.

-WriteToHost and -WriteToStreams cannot both be set at the same time.

.PARAMETER WriteToStreams
A switch parameter that complements -WriteToHost.  If set the message will be written to a 
PowerShell stream.  This overrides the logger configuration setting WriteToHost.

Which PowerShell stream is written to is determined by the Message Level, which may be set via 
the -MessageLevel parameter or by one of the Message Level switch parameters: 
-IsError, -IsWarning, -IsInformation, -IsDebug or -IsVerbose.

-WriteToHost and -WriteToStreams cannot both be set at the same time.

.EXAMPLE
Write error message to the log:

    try
    {
	    ...
    }
    catch [System.IO.FileNotFoundException]
    {
	    Write-LogMessage -Message "Error while updating file: $_.Exception.Message" -IsError
    }

.EXAMPLE
Write debug message to the log:

    Write-LogMessage "Updating user settings for $userName..." -IsDebug

The -Message parameter is optional.

.EXAMPLE
Write to the log, specifying the MessageLevel rather than using a Message Level switch:

    Write-LogMessage "Updating user settings for $userName..." -MessageLevel 'DEBUG'

.EXAMPLE
Write message to the log with a certain category:

    Write-LogMessage 'File copy completed successfully.' -Category 'Success' -IsInformation

.EXAMPLE
Write message to the PowerShell host in a specified color:

    Write-LogMessage 'Updating user settings...' -WriteToHost -HostTextColor Cyan

The MessageLevel wasn't specified so it will default to INFORMATION.

.EXAMPLE
Write message to the Debug PowerShell stream, rather than to the host:

    Write-LogMessage 'Updating user settings...' -WriteToStreams -IsDebug

.EXAMPLE
Write message with a custom message format which only applies to this one message:

    Write-LogMessage "***** Running on server: $serverName *****" -MessageFormat '{message}'

The message written to the log will only include the specified message text.  It will not 
include other fields, such as {Timestamp} or {CallerName}.

.LINK
Get-LogConfiguration

.LINK
Set-LogConfiguration

.LINK
Reset-LogConfiguration

#>
function Write-LogMessage     
{
    [CmdletBinding(DefaultParameterSetName='MessageLevelText')]
    Param
    (
        [Parameter(Mandatory=$False, 
                    Position=0)]
        [string]$Message,

        [Parameter(Mandatory=$False)]
        [ValidateScript({Private_ValidateHostColor $_})]
        [string]$HostTextColor,      

        [Parameter(Mandatory=$False)]
        [string]$MessageFormat,      

        [Parameter(Mandatory=$False,
                     ParameterSetName='MessageLevelText')]
        [ValidateScript({ Private_ValidateLogLevel -LevelToTest $_ -ExcludeOffLevel })]
        [string]$MessageLevel,

        [Parameter(Mandatory=$False,
                     ParameterSetName='MessageLevelSwitches')]
        [switch]$IsError, 

        [Parameter(Mandatory=$False,
                     ParameterSetName='MessageLevelSwitches')]
        [switch]$IsWarning,

        [Parameter(Mandatory=$False,
                     ParameterSetName='MessageLevelSwitches')]
        [switch]$IsInformation, 

        [Parameter(Mandatory=$False,
                     ParameterSetName='MessageLevelSwitches')]
        [switch]$IsDebug, 

        [Parameter(Mandatory=$False,
                     ParameterSetName='MessageLevelSwitches')]
        [switch]$IsVerbose, 

        [Parameter(Mandatory=$False)]
        [string]$Category,

        [Parameter(Mandatory=$False)]
        [switch]$WriteToHost,      

        [Parameter(Mandatory=$False)]
        [switch]$WriteToStreams
    )

    Private_ValidateSwitchParameterGroup -SwitchList $IsError,$IsWarning,$IsInformation,$IsDebug,$IsVerbose `
		-ErrorMessage 'Only one Message Level switch parameter may be set when calling the function. Message Level switch parameters: -IsError, -IsWarning, -IsInformation, -IsDebug, -IsVerbose'

    Private_ValidateSwitchParameterGroup -SwitchList $WriteToHost,$WriteToStreams `
		-ErrorMessage 'Only one Destination switch parameter may be set when calling the function. Destination switch parameters: -WriteToHost, -WriteToStreams'
	
    $Timestamp = Get-Date
    $CallerName = ''
    $TextColor = $Null

    $messageFormatInfo = $script:_messageFormatInfo
    if ($MessageFormat)
    {
        $messageFormatInfo = Private_GetMessageFormatInfo $MessageFormat
    }

    # Getting the calling object name is an expensive operation so only perform it if needed.
    if ($messageFormatInfo.FieldsPresent -contains 'CallerName')
    {
        $CallerName = Private_GetCallerName
    }

    # Parameter sets mean either $MessageLevel is supplied or a message level switch, such as 
    # -IsError, but not both.  Of course, they're all optional so none have to be specified, in 
    # which case we set the default values:

    if ($IsError.IsPresent)
    {
        $MessageLevel = 'ERROR'
    }
    elseif ($IsWarning.IsPresent)
    {
        $MessageLevel = 'WARNING'
    }
    elseif ($IsInformation.IsPresent)
    {
        $MessageLevel = 'INFORMATION'
    }
    elseif ($IsDebug.IsPresent)
    {
        $MessageLevel = 'DEBUG'
    }
    elseif ($IsVerbose.IsPresent)
    {
        $MessageLevel = 'VERBOSE'
    }

    # Default.
    if (-not $MessageLevel)
    {
        $MessageLevel = 'INFORMATION'
    }    

    $configuredLogLevelValue = $script:_logLevels[$script:_logConfiguration.LogLevel]
    $messageLogLevelValue = $script:_logLevels[$MessageLevel]
    if ($messageLogLevelValue -gt $configuredLogLevelValue)
    {
        return
    }

    # Long-winded logic because we want either of the local parameters to override the 
    # configuration setting: If either of the parameters is set ignore the configuration.
    $LogTarget = ''
    if ($WriteToHost.IsPresent)
    {
        $LogTarget = 'Host'
    }
    elseif ($WriteToStreams.IsPresent)
    {
        $LogTarget = 'Streams'
    }
    elseif ($script:_logConfiguration.WriteToHost)
    {
        $LogTarget = 'Host'
    }
    else
    {
        $LogTarget = 'Streams'
    }

    $configuredCategories = @{}
    if ($script:_logConfiguration.ContainsKey('CategoryInfo'))
    {
        $configuredCategories = $script:_logConfiguration.CategoryInfo
    }

    if ([string]::IsNullOrWhiteSpace($Category) -and $configuredCategories)
    {
        $Category = $configuredCategories.Keys.Where( 
            { $configuredCategories[$_] -is [hashtable] `
                -and $configuredCategories[$_].ContainsKey('IsDefault') `
                -and $configuredCategories[$_]['IsDefault'] -eq $True }, 
            'First', 1)
    }
    if ($Category)
    {
        $Category = $Category.Trim()
    }

    $textToLog = $ExecutionContext.InvokeCommand.ExpandString($messageFormatInfo.WorkingFormat)

    if ($LogTarget -eq 'Host')
    {
        if ($HostTextColor)
        {
            $TextColor = $HostTextColor
        }
        elseif ($Category -and $configuredCategories `
            -and $configuredCategories.ContainsKey($Category) `
            -and $configuredCategories[$Category].ContainsKey('Color'))
        {
            $TextColor = $configuredCategories[$Category].Color
        }
        else
        {
            switch ($MessageLevel)
            {
                ERROR	{ $TextColor = $script:_logConfiguration.HostTextColor.Error; break }
                WARNING	{ $TextColor = $script:_logConfiguration.HostTextColor.Warning; break }
                INFORMATION	{ $TextColor = $script:_logConfiguration.HostTextColor.Information; break }
                DEBUG	{ $TextColor = $script:_logConfiguration.HostTextColor.Debug; break }
                VERBOSE	{ $TextColor = $script:_logConfiguration.HostTextColor.Verbose; break }
            }
        }

        if ($TextColor)
        {
            Write-Host $textToLog -ForegroundColor $TextColor
        }
        else
        {
            Write-Host $textToLog
        }
    }
    elseif ($LogTarget -eq 'Streams')
    {
        switch ($MessageLevel)
        {
            ERROR           { Write-Error $textToLog; break }
            WARNING         { Write-Warning $textToLog; break }
            INFORMATION     { Write-Information $textToLog; break }
            DEBUG           { Write-Debug $textToLog; break }
            VERBOSE         { Write-Verbose $textToLog; break }                            
        }
    }

    if (-not $script:_logConfiguration.ContainsKey('LogFile') `
        -or -not $script:_logConfiguration.LogFile.ContainsKey('Name') `
        -or [string]::IsNullOrWhiteSpace($script:_logConfiguration.LogFile.Name))
    {
        return
    }

    if (-not (Test-Path $script:_logFilePath -IsValid))
    {
        # Fail silently so that every message output to the console doesn't include an error 
        # message.
        return
    }

    $overwriteLogFile = $False
    if ($script:_logConfiguration.LogFile.ContainsKey('Overwrite'))
    {
        $overwriteLogFile = $script:_logConfiguration.LogFile.Overwrite
    }
    if ($overwriteLogFile -and (-not $script:_logFileOverwritten))
    {
        Set-Content -Path $script:_logFilePath -Value $textToLog
        $script:_logFileOverwritten = $True
    }
    else
    {
        Add-Content -Path $script:_logFilePath -Value $textToLog
    }
}

<#
.SYNOPSIS
Gets the name of the function calling into this module.

.DESCRIPTION
Walks up the call stack until it finds a stack frame where the ScriptName is not the filename of 
this module.  

If the call stack cannot be read then the function returns '[UNKNOWN CALLER]'.  

If no stack frame is found with a different ScriptName then the function returns "----".

If the ScriptName of the first stack frame outside of this module is $Null then the module is 
being called from the PowerShell console.  In that case the function returns '[CONSOLE]'.  

If the ScriptName of the first stack frame outside of this module is NOT $Null then the module 
is being called from a script file.  In that case the function will return the the stack frame 
FunctionName, unless the FunctionName is "<ScriptBlock>".  

A FunctionName of "<ScriptBlock>" means the module is being called from the root of a script 
file, outside of any function.  In that case the function returns 
"Script <script short file name>".  The script short file name returned will include the file 
extension but not any path information.

.NOTES
This function is NOT intended to be exported from this module.

#>
function Private_GetCallerName()
{
	$callStack = Get-PSCallStack
	if ($callStack -eq $null -or $callStack.Count -eq 0)
	{
		return '[UNKNOWN CALLER]'
	}
	
	$thisFunctionStackFrame = $callStack[0]
	$thisModuleFileName = $thisFunctionStackFrame.ScriptName
	$stackFrameFileName = $thisModuleFileName
    # Skip this function in the call stack as we've already read it.  We also know there must 
    # be at least two stack frames in the call stack as this function will only be called from 
    # another function in this module, so it's safe to skip the first stack frame.
	$i = 1
	$stackFrameFunctionName = '----'
	while ($stackFrameFileName -eq $thisModuleFileName -and $i -lt $callStack.Count)
	{
		$stackFrame = $callStack[$i]
		$stackFrameFileName = $stackFrame.ScriptName
		$stackFrameFunctionName = $stackFrame.FunctionName
		$i++
	}
	
	if ($stackFrameFileName -eq $null)
	{
		return '[CONSOLE]'
	}
	if ($stackFrameFunctionName -eq '<ScriptBlock>')
	{
		$scriptFileNameWithoutPath = (Split-Path -Path $stackFrameFileName -Leaf)
		return "Script $scriptFileNameWithoutPath"
	}
	
	return $stackFrameFunctionName
}

#endregion

#region Configuration *****************************************************************************

<#
.SYNOPSIS
Gets the log configuration settings.

.DESCRIPTION
Gets the log configuration settings.

.OUTPUTS
A hash table with the following keys:

    LogLevel: A string that specifies the Log Level of the logger.  It determines whether a 
        message will be logged or not.  

        Possible values, in order from lowest to highest, are:
            OFF
            ERROR
            WARNING
            INFORMATION
            DEBUG
            VERBOSE

        Only messages with a Message Level the same as or lower than the LogLevel will be logged.

        For example, if the LogLevel is INFORMATION then only messages with a Message Level of 
        INFORMATION, WARNING or ERROR will be logged.  Messages with a Message Level of DEBUG or 
        VERBOSE will not be logged, as those levels are higher than INFORMATION;

    LogFile: A hash table with the configuration details of the log file that log messages will be 
        written to, in addition to the PowerShell host or PowerShell streams.  If you don't want 
        to write to a log file either set the LogFile value to $Null, or set LogFile.Name to 
        $Null.
        
        The hash table has the following keys:

            Name: The path to the log file.  If LogFile.Name is $Null, empty or blank log then 
                messages will be written to the PowerShell host or PowerShell streams but not 
                written to a log file.  
                
                If LogFile.Name is specified without a path, or with a relative path, it will be 
                relative to the directory of the calling script, not this module.  The default 
                value for Log.FileName is "Results.log";

            IncludeDateInFileName: If $True then the log file name will have a date, of the form 
                '_yyyyMMdd' appended to the end of the file name.  For example, 
                'Results_20171129.log'.  The default value is $True;

            Overwrite: If $True any existing log file with the same name as LogFile.Name, 
                including the date if LogFile.IncludeDateInFileName is set, will be overwritten 
                by the first message logged in a given session.  Subsequent messages written in 
                the same session will be appended to the end of the log file.  
        
                If $False new log messages will be appended to the end of the existing log file.  
        
                If no file with the same name exists it will be created, regardless of the value 
                of Log.OverwriteLogFile.  
        
                The default value is $True;

    WriteToHost: If $True then all log messages will be written to the host.  If $False then log 
        messages will be written to the appropriate stream.  For example, Error messages will be 
        written to the error stream, Warning messages will be written to the warning stream, etc. 

        The default value is $True;
		
	MessageFormat: A string that sets the format of log messages.  

        Text enclosed in curly braces, {...}, represents the name of a field which will be included 
        in the logged message.  The field names are not case sensitive.  
        
        Any other text, not enclosed in curly braces, will be treated as a string literal and will 
        appear in the logged message exactly as specified.	
		
		Leading spaces in the MessageFormat string will be retained when the message is written to 
		the logs to allow log messages to be indented.  Trailing spaces in the MessageFormat string 
		will not be included in the logged messages.
		
		Possible field names are:

			{Message}     : The supplied text message to write to the log;

			{Timestamp}	  : The date and time the log message is recorded.  The Timestamp field may 	
							include an optional datetime format string, following the field name 
                            and separated from it by a colon, ":".  
                            
                            Any .NET datetime format string is valid.  For example, "{Timestamp:d}" 
                            will format the timestamp using the short date pattern, which is 
                            "MM/dd/yyyy" in the US.  
                            
                            While the field names in the MessageFormat string are NOT case sentive 
                            the datetime format string IS case sensitive.  This is because .NET 
                            datetime format strings are case sensitive.  For example "d" is the 
                            short date pattern while "D" is the long date pattern.  
                            
                            The default datetime format string is "yyyy-MM-dd hh:mm:ss.fff".

			{CallerName} : The name of the function or script that is writing to the log.  

                            When determining the caller name all functions in this module will be 
                            ignored; the caller name will be the external function or script that 
                            calls into this module to write to the log.  
                            
                            If a function is writing to the log the function name will be 
                            displayed.  If the log is being written to from a script file, outside 
                            any function, the name of the script file will be displayed.  If the 
                            log is being written to manually from the Powershell console then 
                            '[CONSOLE]' will be displayed.

			{Category} : The Category of the message.  It will always be displayed in upper case.

			{MessageLevel}    : The Log Level at which the message is being recorded.  For example, the 
                            message may be an Error message or a Debug message.  The MessageLevel will 
                            always be displayed in upper case.
			
		The default MessageFormat is: 
		'{Timestamp:yyyy-MM-dd hh:mm:ss.fff} | {CallerName} | {Category} | {MessageLevel} | {Message}';

    HostTextColor: A hash table that specifies the different text colors that will be used for 
        different log levels, for log messages written to the host.  HostTextColor only applies 
        if WriteToHost is $True.  
        
        The hash table has the following keys:

            Error: The text color for messages of log level Error.  The default value is Red;

            Warning: The text color for messages of log level Warning.  The default value is 
                Yellow;

            Information: The text color for messages of log level Information.  The default 
                value is Cyan;

            Debug: The text color for messages of log level Debug.  The default value is White;

            Verbose: The text color for messages of log level Verbose.  The default value is 
                White.  

        Possible values for text colors are: 'Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 
        'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 
        'Red', 'Magenta', 'Yellow', 'White';

    CategoryInfo: A hash table that defines properties for Message Categories.  
    
        When writing a log message any string can be used for a Message Category.  However, to 
        provide special functionality for messages of a given category, that category should be 
        added to the configuration CategoryInfo hash table.
        
        The keys of the CategoryInfo hash table are the category names that will be used as 
        Message Categories.  The CategoryInfo values are nested hash tables that set the 
        properties of each category.  
        
        Currently two properties are supported:

            IsDefault: Indicates the category that will be used as the default, if no -Category 
                is specified in Write-LogMesssage;

            Color: The text color for messages of the specified category, if they are written to 
                the host. 
.NOTES
The hash table returned by Get-LogConfiguration is a copy of the Prog configuration, NOT a
reference to the live configuration.  This means any changes to the hash table retrieved by 
Get-LogConfiguration will NOT be reflected in Prog's configuration.

As a result the Prog configuration can only be updated via Set-LogConfiguration.  This ensures 
that the Prog internal state is updated correctly.  

For example, if a user were able to set the configuration MessageFormat string directly this 
modified MessageFormat would not be used when writing log messages.  That is because 
Set-LogConfiguration parses the new MessageFormat string and updates Prog's internal state to 
indicate which fields are to be included in log messages.  If the configuration MessageFormat 
string were updated directly it would not be parsed and the list of fields to include in 
log messages would not be updated.

Although updating the hash table retrieved by Get-LogConfiguration will not update the Prog 
configuration, the updated hash table can be written back as the Prog configuration via 
Set-LogConfiguration.  

.EXAMPLE
Get the text color for messages with category Success:

    PS C:\Users\Me> $config = Get-LogConfiguration
    PS C:\Users\Me> $config.CategoryInfo.Success.Color 

    Green

.EXAMPLE
Get the text colors for all message levels:

    PS C:\Users\Me> $config = Get-LogConfiguration
    PS C:\Users\Me> $config.HostTextColor 

    Name                 Value
    ----                 -----
    Debug                White
    Error                Red
    Warning              Yellow
    Verbose              White
    Information          Cyan

.EXAMPLE
Get the text color for messages of level ERROR:

    PS C:\Users\Me> $config = Get-LogConfiguration
    PS C:\Users\Me> $config.HostTextColor.Error 

    Red

.EXAMPLE
Get the name of the file messages will be logged to:

    PS C:\Users\Me> $config = Get-LogConfiguration
    PS C:\Users\Me> $config.LogFile.Name 

    Results.log

.EXAMPLE
Get the format of log messages:

    PS C:\Users\Me> $config = Get-LogConfiguration
    PS C:\Users\Me> $config.MessageFormat 

    {Timestamp:yyyy-MM-dd hh:mm:ss.fff} | {CallerName} | {Category} | {MessageLevel} | {Message}

.EXAMPLE
Use Get-LogConfiguration and Set-LogConfiguration to update Prog's configuration:

    $config = Get-LogConfiguration
    $config.LogLevel = 'ERROR'
    $config.LogFile.Name = 'Error.log'
    $config.CategoryInfo['FileCopy'] = @{Color = 'DarkYellow'}
    Set-LogConfiguration -LogConfiguration $config
    
.LINK
Write-LogMessage

.LINK
Set-LogConfiguration

.LINK
Reset-LogConfiguration
       
#>
function Get-LogConfiguration()
{
    if ($script:_logConfiguration -eq $Null -or $script:_logConfiguration.Keys.Count -eq 0)
    {
        $script:_logConfiguration = Private_DeepCopyHashTable $script:_defaultLogConfiguration
    }
    return Private_DeepCopyHashTable $script:_logConfiguration
}

<#
.SYNOPSIS
Sets one or more of the log configuration settings.    

.DESCRIPTION 
Sets one or more of the log configuration settings. 

.PARAMETER LogConfiguration
A hash table representing all configuration settings.  For the hash table format see the help 
topic for Get-LogConifguration.

.PARAMETER LogLevel
A string that specifies the Log Level of the logger.  It determines whether a message will be 
logged or not.  

Possible values, in order from lowest to highest, are:
    OFF
    ERROR
    WARNING
    INFORMATION
    DEBUG
    VERBOSE

Only messages with a Message Level the same as or lower than the LogLevel will be logged.

For example, if the LogLevel is INFORMATION then only messages with a Message Level of 
INFORMATION, WARNING or ERROR will be logged.  Messages with a Message Level of DEBUG or 
VERBOSE will not be logged, as those levels are higher than INFORMATION.

.PARAMETER LogFileName
The path to the log file.  If LogFile.Name is $Null, empty or blank log then messages will 
be written to the PowerShell host or PowerShell streams but not written to a log file.  
                
If LogFile.Name is specified without a path, or with a relative path, it will be relative to 
the directory of the calling script, not this module.  The default value for Log.FileName is 
'Results.log'.

.PARAMETER IncludeDateInFileName
A switch parameter that, if set, will include a date in the log file name.  The date will take 
the form '_yyyyMMdd' appended to the end of the file name.  For example, 'Results_20171129.log'.  

IncludeDateInFileName and ExcludeDateFromFileName cannot both be set at the same time.

.PARAMETER ExcludeDateFromFileName
A switch parameter that is the opposite of IncludeDateInFileName.  If set it will exclude the 
date from the log file name.  For example, 'Results.log'.  

IncludeDateInFileName and ExcludeDateFromFileName cannot both be set at the same time.

.PARAMETER OverwriteLogFile
A switch parameter that, if set, will overwrite any existing log file with the same name as 
LogFile.Name, including a date if LogFile.IncludeDateInFileName is set.  The log file will only 
be overwritten by the first message logged in a given session.  Subsequent messages written in 
the same session will be appended to the end of the log file.

OverwriteLogFile and AppendToLogFile cannot both be set at the same time.

.PARAMETER AppendToLogFile
A switch parameter that is the opposite of OverwriteLogFile.  If set new log messages will be 
appended to the end of an existing log file, if it has the same name as Log.FileName, including a 
date if IncludeDateInFileName is set.   

OverwriteLogFile and AppendToLogFile cannot both be set at the same time.

.PARAMETER WriteToHost
A switch parameter that, if set, will direct all output to the host, as opposed to one of the 
streams such as Error or Warning.  If the LogFileName parameter is set the output will also be 
written to a log file.  

WriteToHost and WriteToStreams cannot both be set at the same time.

.PARAMETER WriteToStreams
A switch parameter that complements WriteToHost.  If set all output will be directed to streams, 
such as Error or Warning, rather than the host.  If the LogFileName parameter is set the output 
will also be written to a log file.   

WriteToHost and WriteToStreams cannot both be set at the same time.

.PARAMETER MessageFormat: 
A string that sets the format of log messages.  

Text enclosed in curly braces, {...}, represents the name of a field which will be included in the 
logged message.  The field names are not case sensitive.  
        
Any other text, not enclosed in curly braces, will be treated as a string literal and will appear 
in the logged message exactly as specified.	
		
Leading spaces in the MessageFormat string will be retained when the message is written to the 
logs to allow log messages to be indented.  Trailing spaces in the MessageFormat string will not 
be included in the logged messages.
		
Possible field names are:
	{Message}     : The supplied text message to write to the log;

	{Timestamp}	  : The date and time the log message is recorded.  The Timestamp field may 
					include an optional datetime format string, following the field name and 
                    separated from it by a colon, ":".  
                            
                    Any .NET datetime format string is valid.  For example, "{Timestamp:d}" will 
                    format the timestamp using the short date pattern, which is "MM/dd/yyyy" in 
                    the US.  
                            
                    While the field names in the MessageFormat string are NOT case sentive the 
                    datetime format string IS case sensitive.  This is because .NET datetime 
                    format strings are case sensitive.  For example "d" is the short date pattern 
                    while "D" is the long date pattern.  
                            
                    The default datetime format string is "yyyy-MM-dd hh:mm:ss.fff".

	{CallerName} : The name of the function or script that is writing to the log.  

                    When determining the caller name all functions in this module will be ignored; 
                    the caller name will be the external function or script that calls into this 
                    module to write to the log.  
                            
                    If a function is writing to the log the function name will be displayed.  If 
                    the log is being written to from a script file, outside any function, the name 
                    of the script file will be displayed.  If the log is being written to manually 
                    from the Powershell console then '[CONSOLE]' will be displayed.

	{Category}    : The Category of the message.  It will always be displayed in upper case.

	{MessageLevel} : The Log Level at which the message is being recorded.  For example, the 
                    message may be an Error message or a Debug message.  The MessageLevel will 
                    always be displayed in upper case.

.PARAMETER CategoryInfoItem
Sets one or more items in the CategoryInfo hash table. 

CategoryInfoItem can take two different arguments:

    1) A hash table, of the form:
            @{
                <key1> = @{ <property1>=<value1>; <property2>=<value2>; ...n }
                <key2> = @{ <property1>=<value1>; <property2>=<value2>; ...n }
                ... 
            }

        The items of the hash table will be added to the CategoryInfo hash table, if the keys do 
        not already exist in the CategoryInfo hash table.  If the keys do exist in the 
        CategoryInfo hash table their values will be replaced.         
        
        The keys of the items (<key1>, <key2> in the hash table above) are Message Category 
        names.  The values of the items ( @{ <property1>=<value1>; <property2>=<value2>; ...n } 
        in the hash table above) are hash tables that attach properties to the associated Message 
        Categories.

        Currently two properties are supported for each CategoryInfoItem:

            IsDefault: Indicates the category that will be used as the default, if no -Category 
                is specified in Write-LogMesssage;

            Color: The text color for messages of the specified category, if they are written to 
                the host. 

    2) A two-element array, of the form:
            <key>, @{ <property1>=<value1>; <property2>=<value2>; ...n }

        This sets a single CategoryInfo item.  If the key already exists the value will be 
        replaced.  If the key does not already exist it will be created.

.PARAMETER CategoryInfoKeyToRemove
Removes one or more items from the CategoryInfo hash table.

CategoryInfoKeyToRemove can take two different arguments:

    1) An array of CategoryInfo keys;

    2) A single CategoryInfo key.

.PARAMETER HostTextColorConfiguration
A hash table specifying the different text colors that will be used for different log levels, 
for log messages written to the host.  

The hash table must have the following keys:

    Error: The text color for messages of log level Error.  The default value is Red;

    Warning: The text color for messages of log level Warning.  The default value is 
        Yellow;

    Information: The text color for messages of log level Information.  The default 
        value is Cyan;

    Debug: The text color for messages of log level Debug.  The default value is White;

    Verbose: The text color for messages of log level Verbose.  The default value is White.  

Possible values for text colors are: 'Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 
'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 
'Red', 'Magenta', 'Yellow', 'White'.

These colors are only used if WriteToHost is set.  If WriteToStreams is set these colors are 
ignored.

.PARAMETER ErrorTextColor
The text color for messages written to the host at message level Error.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are: 'Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 
'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 
'Red', 'Magenta', 'Yellow', 'White'.

.PARAMETER WarningTextColor
The text color for messages written to the host at message level Warning.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are as per ErrorTextColor.

.PARAMETER InformationTextColor
The text color for messages written to the host at message level Information.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are as per ErrorTextColor.

.PARAMETER DebugTextColor
The text color for messages written to the host at message level Debug.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are as per ErrorTextColor.

.PARAMETER VerboseTextColor
The text color for messages written to the host at message level Verbose.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are as per ErrorTextColor.

.EXAMPLE
Use parameter -LogConfiguration to update the entire configuration at once:

    $hostTextColor = @{
							Error = 'DarkRed'
							Warning = 'DarkYellow'
							Information = 'DarkCyan'
							Debug = 'Gray'
							Verbose = 'White'
						}

	$logConfiguration = @{   
							LogLevel = 'DEBUG'
							MessageFormat = '{CallerName} | {Category} | {Message}'
							WriteToHost = $True
							HostTextColor = $hostTextColor
							LogFile = @{
											Name = 'Debug.log'
											IncludeDateInFileName = $False
											Overwrite = $False
										}
							CategoryInfo = @{}
						}
						
	Set-LogConfiguration -LogConfiguration $logConfiguration

.EXAMPLE
Set the details of the log file using individual parameters:

    Set-LogConfiguration -LogFileName 'Debug.log' -ExcludeDateFromFileName -AppendToLogFile

.EXAMPLE
Set the LogLevel and MessageFormat using individual parameters:

    Set-LogConfiguration -LogLevel Warning `
	    -MessageFormat '{Timestamp:yyyy-MM-dd hh:mm:ss},{Category},{Message}'

.EXAMPLE
Add a single CategoryInfo item using the tuple (two-element array) syntax:

    Set-LogConfiguration -CategoryInfoItem FileCopy, @{ Color = 'Blue' }

.EXAMPLE
Change the default Category using the tuple (two-element array) syntax:

    Set-LogConfiguration -CategoryInfoItem FileCopy, @{ IsDefault = $True }

.EXAMPLE
Add multiple CategoryInfo items using the hash table syntax:

    Set-LogConfiguration -CategoryInfoItem @{ 
                                            FileCopy = @{ Color = 'Blue' } 
                                            FileAdd = @{ Color = 'Yellow' } 
                                            }

If the configuration CategoryInfo hash table already includes keys 'FileCopy' and 'FileAdd' 
the colors of those CategoryInfo items will be updated.  If the keys do not already exist 
they will be created.

.EXAMPLE
Remove a single CategoryInfo item:

    Set-LogConfiguration -CategoryInfoKeyToRemove PartialFailure

The CategoryInfo item with key 'PartialFailure' will be removed, if it exists.  No error is 
thrown if the key does not exist.

.EXAMPLE
Remove multiple CategoryInfo items:

    Set-LogConfiguration -CategoryInfoKeyToRemove Progress,PartialFailure

The CategoryInfo items with keys 'Progress' and 'PartialFailure' will be removed, if they exist.

.EXAMPLE
Set the text colors used by the host to display error and warning messages:

    Set-LogConfiguration -ErrorTextColor DarkRed -WarningTextColor DarkYellow

.EXAMPLE
Set all text colors simultaneously:

	$hostColors = @{
						Error = 'DarkRed'
						Warning = 'DarkYellow'
						Information = 'DarkCyan'
						Debug = 'Cyan'
						Verbose = 'Gray'
					}
					
	Set-LogConfiguration -HostTextColorConfiguration $hostColors

.EXAMPLE
Use Get-LogConfiguration and Set-LogConfiguration to update Prog's configuration:

    $config = Get-LogConfiguration
    $config.LogLevel = 'ERROR'
    $config.LogFile.Name = 'Error.log'
    $config.CategoryInfo['FileCopy'] = @{Color = 'DarkYellow'}
    Set-LogConfiguration -LogConfiguration $config
    
.LINK
Write-LogMessage

.LINK
Get-LogConfiguration

.LINK
Reset-LogConfiguration
#>
function Set-LogConfiguration
{
    # CmdletBinding attribute must be on first non-comment line of the function
    # and requires that the parameters be defined via the Param keyword rather 
    # than in parentheses outside the function body.
    [CmdletBinding(DefaultParameterSetName="IndividualSettings_IndividualColors")]
    Param 
    (
        [parameter(Mandatory=$True, 
                    ParameterSetName="AllSettings")]
        [ValidateNotNull()]
        [Hashtable]$LogConfiguration, 

        [parameter(ParameterSetName="IndividualSettings_AllColors")]
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]        
        [ValidateScript({ Private_ValidateLogLevel -LevelToTest $_ })]
        [string]$LogLevel, 
        
        [parameter(ParameterSetName="IndividualSettings_AllColors")]
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [string]$LogFileName,
        
        [parameter(ParameterSetName="IndividualSettings_AllColors")]
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [switch]$IncludeDateInFileName,
        
        [parameter(ParameterSetName="IndividualSettings_AllColors")]
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [switch]$ExcludeDateFromFileName,
        
        [parameter(ParameterSetName="IndividualSettings_AllColors")]
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [switch]$OverwriteLogFile,
        
        [parameter(ParameterSetName="IndividualSettings_AllColors")]
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [switch]$AppendToLogFile,
        
        [parameter(ParameterSetName="IndividualSettings_AllColors")]
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [switch]$WriteToHost,      
        
        [parameter(ParameterSetName="IndividualSettings_AllColors")]
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [switch]$WriteToStreams,    
        
        [parameter(ParameterSetName="IndividualSettings_AllColors")]
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [string]$MessageFormat,
        
        [parameter(ParameterSetName="IndividualSettings_AllColors")]
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        $CategoryInfoItem,
        
        [parameter(ParameterSetName="IndividualSettings_AllColors")]
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        $CategoryInfoKeyToRemove,
        
        [parameter(ParameterSetName="IndividualSettings_AllColors")]
        [ValidateNotNull()]
        [Hashtable]$HostTextColorConfiguration, 

        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [ValidateScript({Private_ValidateHostColor $_})]
        [string]$ErrorTextColor,            
        
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [ValidateScript({Private_ValidateHostColor $_})]
        [string]$WarningTextColor,      
        
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [ValidateScript({Private_ValidateHostColor $_})]
        [string]$InformationTextColor,    
        
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [ValidateScript({Private_ValidateHostColor $_})]
        [string]$DebugTextColor,          
        
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [ValidateScript({Private_ValidateHostColor $_})]
        [string]$VerboseTextColor
    )

    if ($LogConfiguration -ne $Null)
    {
        $script:_logConfiguration = Private_DeepCopyHashTable $LogConfiguration
        Private_SetMessageFormat $LogConfiguration.MessageFormat
        Private_SetLogFilePath
        return
    }

    # Ensure that mutually exclusive pairs of switch parameters are not both set:

    Private_ValidateSwitchParameterGroup -SwitchList $IncludeDateInFileName,$ExcludeDateFromFileName `
		-ErrorMessage "Only one FileName switch parameter may be set when calling the function. FileName switch parameters: -IncludeDateInFileName, -ExcludeDateFromFileName"

    Private_ValidateSwitchParameterGroup -SwitchList $OverwriteLogFile,$AppendToLogFile `
		-ErrorMessage "Only one LogFileWriteBehavior switch parameter may be set when calling the function. LogFileWriteBehavior switch parameters: -OverwriteLogFile, -AppendToLogFile"

    Private_ValidateSwitchParameterGroup -SwitchList $WriteToHost,$WriteToStreams `
		-ErrorMessage "Only one Destination switch parameter may be set when calling the function. Destination switch parameters: -WriteToHost, -WriteToStreams"

    if (![string]::IsNullOrWhiteSpace($LogLevel))
    {
        $script:_logConfiguration.LogLevel = $LogLevel
    }

    if (![string]::IsNullOrWhiteSpace($LogFileName))
    {
        $script:_logConfiguration.LogFile.Name = $LogFileName
        Private_SetLogFilePath
    }

    if ($ExcludeDateFromFileName.IsPresent)
    {
        $script:_logConfiguration.LogFile.IncludeDateInFileName = $False
        Private_SetLogFilePath
    }

    if ($IncludeDateInFileName.IsPresent)
    {
        $script:_logConfiguration.LogFile.IncludeDateInFileName = $True
        Private_SetLogFilePath
    }

    if ($AppendToLogFile.IsPresent)
    {
        $script:_logConfiguration.LogFile.Overwrite = $False
    }

    if ($OverwriteLogFile.IsPresent)
    {
        $script:_logConfiguration.LogFile.Overwrite = $True
    }

    if ($WriteToStreams.IsPresent)
    {
        $script:_logConfiguration.WriteToHost = $False
    }

    if ($WriteToHost.IsPresent)
    {
        $script:_logConfiguration.WriteToHost = $True
    }

    if (![string]::IsNullOrWhiteSpace($MessageFormat))
    {
        Private_SetMessageFormat $MessageFormat
    }

    if ($CategoryInfoItem)
    {
        if (-not $script:_logConfiguration.ContainsKey('CategoryInfo') `
        -or (-not $script:_logConfiguration.CategoryInfo))
        {
            $script:_logConfiguration.CategoryInfo = @{}
        }

        if ($CategoryInfoItem -is [array])
        {            
            $key = $CategoryInfoItem[0]
            $value = $CategoryInfoItem[1]

            Private_SetCategoryInfoItem `
                -CategoryInfoHashtable $script:_logConfiguration.CategoryInfo `
                -Key $key -Value $value
        }
        elseif ($CategoryInfoItem -is [hashtable])
        {
            foreach($key in $CategoryInfoItem.Keys)
            {
                 Private_SetCategoryInfoItem `
                    -CategoryInfoHashtable $script:_logConfiguration.CategoryInfo `
                    -Key $key -Value $CategoryInfoItem[$key]
            }
        }
    }

    if ($CategoryInfoKeyToRemove -and $script:_logConfiguration.CategoryInfo)
    {
        foreach($key in $CategoryInfoKeyToRemove)
        {
            $script:_logConfiguration.CategoryInfo.Remove($key)
        }
    }

    if ($HostTextColorConfiguration -ne $Null)
    {
        $script:_logConfiguration.HostTextColor = $HostTextColorConfiguration
    }

    if (![string]::IsNullOrWhiteSpace($ErrorTextColor))
    {
        Private_SetConfigTextColor -ConfigurationKey "Error" -ColorName $ErrorTextColor
    }

    if (![string]::IsNullOrWhiteSpace($WarningTextColor))
    {
        Private_SetConfigTextColor -ConfigurationKey "Warning" -ColorName $WarningTextColor
    }

    if (![string]::IsNullOrWhiteSpace($InformationTextColor))
    {
        Private_SetConfigTextColor -ConfigurationKey "Information" -ColorName $InformationTextColor
    }

    if (![string]::IsNullOrWhiteSpace($DebugTextColor))
    {
        Private_SetConfigTextColor -ConfigurationKey "Debug" -ColorName $DebugTextColor
    }

    if (![string]::IsNullOrWhiteSpace($VerboseTextColor))
    {
        Private_SetConfigTextColor -ConfigurationKey "Verbose" -ColorName $VerboseTextColor
    }
}

<#
.SYNOPSIS
Resets the log configuration to the default settings. 

.DESCRIPTION    
Resets the log configuration to the default settings. 
    
.LINK
Write-LogMessage

.LINK
Get-LogConfiguration

.LINK
Set-LogConfiguration
#>
function Reset-LogConfiguration()
{
    Set-LogConfiguration -LogConfiguration $script:_defaultLogConfiguration
}

<#
.SYNOPSIS
Gets the folder name of the top-most script or function calling into this module.

.DESCRIPTION
Returns the folder name from the first non-console stack frame at the top of the call stack.  If 
that stack frame represents this module the function will return the current location set in the 
console.

If the call stack cannot be read then the function returns $Null.


.NOTES
This function is NOT intended to be exported from this module.

This is an expensive function.  However, it will only be called while setting the logging 
configuration which shouldn't happen often.
#>
function Private_GetCallerDirectory()
{
	$callStack = Get-PSCallStack
	if ($callStack -eq $null -or $callStack.Count -eq 0)
	{
		return $Null
	}
    
    $thisFunctionStackFrame = $callStack[0]
	$thisModuleFileName = $thisFunctionStackFrame.ScriptName
    
    $i = $callStack.Count - 1
	$stackFrameFileName = $Null
	do
	{
		$stackFrame = $callStack[$i]
		$stackFrameFileName = $stackFrame.ScriptName
        $i--
	} while ($stackFrameFileName -eq $Null -and $stackFrameFileName -ne $thisModuleFileName -and $i -ge 0)
	
    $stackFrameDirectory = (Get-Location).Path
    # A stack frame representing a call from the console will have ScriptName equal to $Null.  A  
    # stack frame representing a call from a script file (whether from the root of the file or 
    # from a function) will have a non-null ScriptName.
    if ($stackFrameFileName -ne $Null -and $stackFrameFileName -ne $thisModuleFileName)
    {
        $stackFrameDirectory = Split-Path -Path $stackFrameFileName -Parent
    }

	return $stackFrameDirectory
}

<#
.SYNOPSIS
Gets the absolute path of the specified path.

.DESCRIPTION
Determines whether the path supplied is an absolute or a relative path.  If it is 
absolute it is returned unaltered.  If it is relative then the path to the directory the 
calling script is running in will be prepended to the specified path.

.NOTES
This function is NOT intended to be exported from this module.

#>
function Private_GetAbsolutePath (
    [Parameter(Mandatory=$True)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
    )
{
    if (-not (Test-Path $Path -IsValid))
    {
        throw [ArgumentException] "Invalid file path: '$Path'"
    }

    if ([System.IO.Path]::IsPathRooted($Path))
    {
        return $Path
    }

    $callingDirectoryPath = Private_GetCallerDirectory

    $Path = Join-Path $callingDirectoryPath $Path

    return $Path
}

<#
.SYNOPSIS
Sets the full path to the log file.

.DESCRIPTION
Sets module variable $_logFilePath.  If $_logFilePath has changed then $_logFileOverwritten 
will be cleared.

Determines whether the LogFile.Name specified in the configuration settings is an absolute or a 
relative path.  If it is relative then the path to the directory the calling script is running 
in will be prepended to the specified LogFile.Name.

If configuration setting LogFile.IncludeDateInFileName is $True then the date will be included in 
the log file name, in the form: "<log file name>_yyyyMMdd.<file extension>".  For example, 
"Results_20171129.log".

.NOTES
This function is NOT intended to be exported from this module.

#>
function Private_SetLogFilePath ()
{
    $oldLogFilePath = $script:_logFilePath

    if ([string]::IsNullOrWhiteSpace($script:_logConfiguration.LogFile.Name))
    {
        $script:_logFilePath = ''
        return
    }

    $logFilePath = Private_GetAbsolutePath $script:_logConfiguration.LogFile.Name

    if ($script:_logConfiguration.LogFile.IncludeDateInFileName)
    {
        $directory = [System.IO.Path]::GetDirectoryName($logFilePath)
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($logFilePath)
        $fileName += (Get-Date -Format "_yyyyMMdd")

        # Will include the leading ".":
        $fileExtension = [System.IO.Path]::GetExtension($logFilePath)

        $logFilePath = [System.IO.Path]::Combine($directory, $fileName + $fileExtension)
    }

    $script:_logFilePath = $logFilePath

    if ($script:_logFilePath -ne $oldLogFilePath)
    {
        $script:_logFileOverwritten = $False
    }
}

<#
.SYNOPSIS
Returns a deep copy of a hash table.

.DESCRIPTION
Returns a deep copy of a hash table.

.NOTES
Assumes the hash table values are either value types or nested hash tables.  This function 
will not deal properly with values that are reference types; it will make shallow copies of 
them.

This function is required because the Clone method will only perform a shallow copy of a 
hash table.  This would not be a problem if all values of the hash table were value types 
but that is not the case: HostTextColor is a nested hash table.

This function is NOT intended to be exported from this module.
#>
function Private_DeepCopyHashTable([Collections.Hashtable]$HashTable)
{
    if ($HashTable -eq $Null)
    {
        return $Null
    }

    if ($HashTable.Keys.Count -eq 0)
    {
        return @{}
    }

    $copy = @{}
    foreach($key in $HashTable.Keys)
    {
        if ($HashTable[$key] -is [Collections.Hashtable])
        {
            $copy[$key] = (Private_DeepCopyHashTable $HashTable[$key])
        }
        else
        {
            # Assumes the value of the hash table element is a value type, not a reference type.
			# Works also if the value is an array of values types (ie does a deep copy of the 
			# array).
            $copy[$key] = $HashTable[$key]
        }
    }

    return $copy
}

<#
.SYNOPSIS
Sets the message format in the log configuration settings.

.DESCRIPTION
Sets the message format in the log configuration settings.

.PARAMETER MessageFormat: 
A string that sets the format of log messages.  

Text enclosed in curly braces, {...}, represents the name of a field which will be included in 
the logged text.  The field names are not case sensitive.  
        
Any other text, not enclosed in curly braces, will be treated as a string literal and will appear 
in the logged text exactly as specified.	
		
Leading spaces in the MessageFormat string will be retained when the text is written to the 
log to allow log messages to be indented.  Trailing spaces in the MessageFormat string will be 
removed, and will not be written to the log.
		
Possible field names are:
	{Message}     : The supplied text message to write to the log;

	{Timestamp}	  : The date and time the log message is recorded.  

                    The Timestamp field may include an optional datetime format string, inside 
                    the curly braces, following the field name and separated from it by a 
                    colon, ':'.  For example, '{Timestamp:T}'.
                            
                    Any .NET datetime format string is valid.  For example, "{Timestamp:d}" will 
                    format the timestamp using the short date pattern, which is "MM/dd/yyyy" in 
                    the US.  
                            
                    While the field names in the MessageFormat string are NOT case sentive the 
                    datetime format string IS case sensitive.  This is because .NET datetime 
                    format strings are case sensitive.  For example, "d" is the short date 
                    pattern while "D" is the long date pattern.  
                            
                    The Timestamp field may be specified without any datetime format string.  For 
                    example, '{Timestamp}'.  In that case the default datetime format string,  
                    'yyyy-MM-dd hh:mm:ss.fff', will be used;

	{CallerName}  : The name of the function or script that is writing to the log.  

                    When determining the caller name all functions in this module will be ignored; 
                    the caller name will be the external function or script that calls into this 
                    module to write to the log.  
                            
                    If a function is writing to the log the function name will be displayed.  If 
                    the log is being written to from a script file, outside any function, the name 
                    of the script file will be displayed.  If the log is being written to manually 
                    from the Powershell console then '[CONSOLE]' will be displayed.

	{MessageLevel} : The Message Level at which the message is being recorded.  For example, the 
                    message may be an Error message or a Debug message.  The MessageLevel will 
                    always be displayed in upper case.

	{Category}    : The Message Category.  If no Message Category is explicitly specified when 
                    calling Write-LogMessage the default Message Category from the logger 
                    configuration will be used.

.NOTES
This function is NOT intended to be exported from this module.
#>
function Private_SetMessageFormat([string]$MessageFormat)
{
    $script:_logConfiguration["MessageFormat"] = $MessageFormat

    $script:_messageFormatInfo = Private_GetMessageFormatInfo $MessageFormat
}

<#
.SYNOPSIS
Sets one of the host text color values in the log configuration settings.

.DESCRIPTIONs
Sets one of the host text color values in the log configuration settings.

.NOTES
This function is NOT intended to be exported from this module.
#>
function Private_SetConfigTextColor([string]$ConfigurationKey, [string]$ColorName)
{
    if (-not $script:_logConfiguration.ContainsKey("HostTextColor"))
    {
        $script:_logConfiguration.HostTextColor = $script:_defaultHostTextColor
    }

    $script:_logConfiguration.HostTextColor[$ConfigurationKey] = $ColorName
}

<#
.SYNOPSIS
Function called by ValidateScript to check if the value passed to parameter -CategoryInfoItem 
is valid.

.DESCRIPTION
Checks whether the parameter is either a hash table or an array of two elements, the second of 
which is a hash table.  If the parameter meets these criteria this function returns $True.  If 
the parameter doesn't meet the criteria the function throws an exception rather than returning 
$False.  

.NOTES        
Throwing an exception allows us to specify a custom error message.  If the function simply 
returned $False PowerShell would generate a standard error message that would not indicate why 
the validation failed.
#>
function Private_ValidateCategoryInfoItem (
	[Parameter(Mandatory=$True)]
	$CategoryInfoItem
	)
{	
    if ($CategoryInfoItem -is [array])
    {
        if ($CategoryInfoItem.Count -ne 2)
        {
            throw [ArgumentException]::new( `
                "Expected an array of 2 elements but $($CategoryInfoItem.Count) supplied.", 
                'CategoryInfoItem')
        }
                        
        $key = $CategoryInfoItem[0]
        $value = $CategoryInfoItem[1]

        if (-not ($key -is [string]))
        {
            throw [ArgumentException]::new( `
                "Expected first element to be a string but it is $($key.GetType().FullName).", 
                'CategoryInfoItem')
        }

        if (-not ($value -is [hashtable]))
        {
            throw [ArgumentException]::new( `
                "Expected second element to be a hashtable but it is $($value.GetType().FullName).", 
                'CategoryInfoItem')
        }

        return $True
    }

    if ($CategoryInfoItem -is [hashtable]) 
    {
        foreach($key in $CategoryInfoItem.Keys)
        {
            if (-not ($key -is [string]))
            {
                throw [ArgumentException]::new( `
                    "Expected key to be a string but it is $($key.GetType().FullName).", 
                    'CategoryInfoItem')
            }

            $value = $CategoryInfoItem[$key]
            if (-not ($value -is [hashtable]))
            {
                throw [ArgumentException]::new( `
                    "Expected value to be a hashtable but it is $($value.GetType().FullName).", 
                    'CategoryInfoItem')
            }
        }

        return $True
    }

    throw [ArgumentException]::new( `
        "Expected argument to be either a hashtable or an array but it is $($CategoryInfoItem.GetType().FullName).",
        'CategoryInfoItem')
}

<#
.SYNOPSIS
Sets a configuration CategoryInfo item.

.DESCRIPTION
The item to set is specified via the -Key and -Value parameters.  If the value hashtable contains 
an IsDefault key then any existing value hashtable with an IsDefault key will have the key 
removed.
#>
function Private_SetCategoryInfoItem (
        [hashtable]$CategoryInfoHashtable,                
        [string]$Key,                
        [hashtable]$Value
    )
{
    if (-not $CategoryInfoHashtable)
    {
        $CategoryInfoHashtable = @{}
    }

    if ($Value.ContainsKey('IsDefault') -and $Value.IsDefault -eq $True)
    {
        foreach ($existingKey in $CategoryInfoHashtable.Keys)
        {
            $existingValue = $CategoryInfoHashtable[$existingKey]
            $existingValue.Remove('IsDefault')
        }
    }

    $CategoryInfoHashtable[$Key] = $Value
}

#endregion

#region Private Functions Shared by Logging and Configuration Functions ***************************

<#
.SYNOPSIS
Ensures at most one of the switch values passed as an argument is set.

.DESCRIPTION
Ensures at most one of the switch values passed as an argument is set.

.NOTES
This function is NOT intended to be exported from this module.
#>
function Private_ValidateSwitchParameterGroup (
	[parameter(Mandatory=$True)]
    [ValidateNotNull()]
	[switch[]]$SwitchList,
	
	[parameter(Mandatory=$True)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
	[string]$ErrorMessage
	)
{
	# Can't use "if ($SwitchList.Count -gt 1)..." because it will always be true, even if no 
	# switches are set when calling the parent function.  If one of the switch parameters is not 
	# set it will still be passed to this function but with value $False.	
	# Could use ".Where{$_}" but ".Where{$_ -eq $True}" is easier to understand.
	if ($SwitchList.Where{$_ -eq $True}.Count -gt 1)
	{
		throw [ArgumentException] $ErrorMessage
	}
}

<#
.SYNOPSIS
Function called by ValidateScript to check if the specified host color name is valid when 
passed as a parameter.

.DESCRIPTION
If the specified color name is valid this function returns $True.  If the specified color name 
is not valid the function throws an exception rather than returning $False.  

.NOTES        
Allows multiple parameters to be validated in a single place, so the validation code does not 
have to be repeated for each parameter.  

Throwing an exception when the color name is invalid allows us to specify a custom error message. 
If the function simply returned $False PowerShell would generate a standard error message that 
does not indicate why the validation failed.
#>
function Private_ValidateHostColor (
	[Parameter(Mandatory=$True)]
	[string]$ColorToTest
	)
{	
    $validColors = @('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 
        'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 
        'Red', 'Magenta', 'Yellow', 'White')
	
	if ($validColors -notcontains $ColorToTest)
	{
		throw [ArgumentException] "INVALID TEXT COLOR ERROR: '$ColorToTest' is not a valid text color for the PowerShell host."
	}
			
	return $True
}

<#
.SYNOPSIS
Function called by ValidateScript to check if the specified log level is valid when passed as a 
parameter.

.DESCRIPTION
If the specified log level is valid this function returns $True.  If the specified log level is 
not valid the function throws an exception rather than returning $False.  

.NOTES        
Allows multiple parameters to be validated in a single place, so the validation code does not 
have to be repeated for each parameter.  

Throwing an exception when the log level is invalid allows us to specify a custom error message. 
If the function simply returned $False PowerShell would generate a standard error message that 
does not indicate why the validation failed.
#>
function Private_ValidateLogLevel (
	[Parameter(Mandatory=$True)]
	[string]$LevelToTest, 

    [parameter(Mandatory=$False)]
	[switch]$ExcludeOffLevel
	)
{	
    $validLevels = @('OFF', 'ERROR', 'WARNING', 'INFORMATION', 'DEBUG', 'VERBOSE')
    if ($ExcludeOffLevel.IsPresent)
    {
        $validLevels[0] = $Null
    }
	
	if ($validLevels -notcontains $LevelToTest)
	{
		throw [ArgumentException] "INVALID LOG LEVEL ERROR: '$LevelToTest' is not a valid log level."
	}
			
	return $True
}

<#
.SYNOPSIS
Gets a Timestamp format string from a specified message format string.

.DESCRIPTION
Parses a message format string to find a {Timestamp} field placeholder.  If the Timestamp 
plceholder is found and it contains a datetime format string the datetime format string will be 
returned.  If there is no Timestamp plceholder, or if the Timestamp placeholder has no datetime 
format string, then $Null will be returned.

.NOTES
This function is NOT intended to be exported from this module.
#>
function Private_GetTimestampFormat ([string]$MessageFormat)
{
    # The regex can handle zero or more white spaces (spaces or tabs) between the curly braces 
    # and the placeholder name.  eg "{  Timestamp}", '{ Timestamp   }".  It can also handle 
    # zero or more white spaces before or after the colon that separates the placeholder name 
    # from the datetime format string.  Note that (?: ... ) is a non-capturing group so $Matches 
    # should contain at most two groups: 
    #    $Matches[0]    : The overall match.  Always present if the {Timestamp} placeholder is 
    #                       present;
    #    $Matches[1]    : The datetime format string.  Only present if the {Timestamp} 
    #                       placeholder has a datetime format string.
    # Note the first colon in the regex pattern is part of the non-capturing group specifier.  
    # The second colon in the regex pattern represents the separator between the placeholder name 
    # and the datetime format string, eg {Timestamp:d}
    $regexPattern = "{\s*Timestamp\s*(?::\s*(.+?)\s*)?\s*}"
	
    # -imatch is a case insensitive regex match.
    # No need to compile the regex as it won't be used often.
    $isMatch = $MessageFormat -imatch $regexPattern
	if ($isMatch -and $Matches.Count -ge 2)
	{
		return $Matches[1].Trim()
	}

    return $Null
}

<#
.SYNOPSIS
Gets message format info from a message format string.

.DESCRIPTION
Parses a message format string to determine which fields will appear in the log message and what 
their format strings are, if applicable.  The results are returned in a hash table.

.OUTPUTS
A hash table with the following keys:
    RawFormat: The format string passed into this function;

    WorkingFormat: A modified format string with the field placeholders replaced with variable 
        names.  The variable names that may be embedded in the WorkingFormat string are:

        $Message        :   Replaces field placeholder {Message};
                            
        $Timestamp      :   Unlike other fields, Timestamp must include a datetime format string.  
                            If no datetime format string is included in the Timestamp placeholder 
                            it will default to 'yyyy-MM-dd hh:mm:ss.fff'.  
                            
                            The Timestamp placeholder will be replaced with 
                            "$($Timestamp.ToString('<datetime format string>'))".  
                            
                            Examples:
                                1)  Field placeholder '{Timestamp:d}' will be replaced by 
                                    "$($Timestamp.ToString('d'))";

                                2) Field placeholder '{Timestamp}' will use the default datetime 
                                    format string so will be replaced by 
                                    "$($Timestamp.ToString('yyyy-MM-dd hh:mm:ss.fff'))";

        $CallerName :   Replaces field placeholder {CallerName};

        $MessageLevel   :   Replaces field placeholder {MessageLevel};

        $Category    :   Replaces field placeholder {Category};
        
    FieldsPresent: An array of strings representing the names of fields that will appear in the 
        log message.  Field names that may appear in the array are:

        "Message"       :   Included if the RawFormat string contains field placeholder {Message};
                            
        "Timestamp"     :   Included if the RawFormat string contains field placeholder 
                            {Timestamp};

        "CallerName" :   Included if the RawFormat string contains field placeholder 
                            {CallerName};

        "MessageLevel"  :   Included if the RawFormat string contains field placeholder 
                            {MessageLevel};

        "Category"   :   Included if the RawFormat string contains field placeholder 
                            {Category}.

.NOTES
This function is NOT intended to be exported from this module.
#>
function Private_GetMessageFormatInfo([string]$MessageFormat)
{
    $messageFormatInfo = @{
                            RawFormat = $MessageFormat
                            WorkingFormat = ""
                            FieldsPresent = @()
                        }

    $workingFormat = $MessageFormat

    # -ireplace is a case insensitive find and replace.
    # The regex can handle zero or more white spaces (spaces or tabs) between the curly braces 
    # and the placeholder name.  eg "{ Messages}", '{  Messages   }".
    # No need to compile the regex as it won't be used often.
    $modifiedText = $workingFormat -ireplace '{\s*Message\s*}', '${Message}'
    if ($modifiedText -ne $workingFormat)
    {
        $messageFormatInfo.FieldsPresent += "Message"
        $workingFormat = $modifiedText
    }

    $timestampFormat = Private_GetTimestampFormat $workingFormat
    if (-not $timestampFormat)
    {
        $timestampFormat = $script:_defaultTimestampFormat
    }
    # Escape the first two "$" because we want to retain them in the replacement text.  Do 
    # not escape the "$" in "$timestampFormat" because we want to expand that variable.
    $replacementText = "`$(`$Timestamp.ToString('$timestampFormat'))"

    $modifiedText = $workingFormat -ireplace '{\s*Timestamp\s*(?::\s*.+?\s*)?\s*}', $replacementText
    if ($modifiedText -ne $workingFormat)
    {
        $messageFormatInfo.FieldsPresent += "Timestamp"
        $workingFormat = $modifiedText
    }

    $modifiedText = $workingFormat -ireplace '{\s*CallerName\s*}', '${CallerName}'
    if ($modifiedText -ne $workingFormat)
    {
        $messageFormatInfo.FieldsPresent += "CallerName"
        $workingFormat = $modifiedText
    }

    $modifiedText = $workingFormat -ireplace '{\s*MessageLevel\s*}', '${MessageLevel}'
    if ($modifiedText -ne $workingFormat)
    {
        $messageFormatInfo.FieldsPresent += "MessageLevel"
        $workingFormat = $modifiedText
    }

    $modifiedText = $workingFormat -ireplace '{\s*Category\s*}', '${Category}'
    if ($modifiedText -ne $workingFormat)
    {
        $messageFormatInfo.FieldsPresent += "Category"
        $workingFormat = $modifiedText
    }

     $messageFormatInfo.WorkingFormat = $workingFormat

     return $messageFormatInfo
}

#endregion

# Set up initial conditions.
Reset-LogConfiguration

# Only export public functions.  To simplify the exporting of public functions but not private 
# ones public functions must follow the standard PowerShell naming convention, 
# "<verb>-<singular noun>", while private functions must not contain a dash, "-".
Export-ModuleMember -Function *-*
