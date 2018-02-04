<#
.SYNOPSIS
Functions for logging messages to the host or to PowerShell streams and, optionally, to a log file.

.DESCRIPTION
A module for logging messages to the host or to PowerShell streams, such as the Error stream or 
the Information stream.  In addition, messages may optionally be logged to a log file.

Messages are logged using the Write-LogMessage function.

The logging configuration is flexible and powerful, allowing the following properties of log 
messages to be changed:

    1) The log level:  This determines whether a log message will be logged or not.  
        
        Possible log levels, in order from lowest to highest, are:
            "Off"
            "Error"
            "Warning"
            "Information"
            "Debug"
            "Verbose" 

        Only log messages at a level the same as or lower than the LogLevel will be logged.  For 
        example, if the LogLevel is "Information" then only log messages at a level of 
        Information, Warning or Error will be logged.  Messages at a level of Debug or Verbose 
        will not be logged, as these log levels are higher than Information;

    2) The message destination:  Messages may be written to the host or to PowerShell streams such 
        as the Information stream or the Verbose stream.  In addition, if a log file name is 
        set in the logging configuration, the messages will be written to a log file;

    3) The host text color:  Messages written to the host, as opposed to PowerShell streams, may 
        be written in any PowerShell console color.  Different colors may be specified for 
        different message types, such as Error, Warning or Information;

    4) The message format:  In addition to the specified message, the text written to the log may 
        include additional fields that are automatically populated, such as a timestamp or the  
        name of the function writing to the log.  The format of the logged text, including the 
        fields to be displayed, may be specified in the logging configuration.

The logging configuration may be altered via function Set-LogConfiguration.  Function 
Get-LogConfiguration will return a copy of the current configuration as a hash table.  The 
configuration can be reset back to its default values via function Reset-LogConfiguration.

Several configuration settings can be overridden when writing a single log message.  The changes 
apply only to that one message; subsequent messages will return to using the settings in the 
logging configuration.

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

# Message types need the same values as log levels for the keys they have in common.
$_messageTypes = @{
                    Error = 1
                    Warning = 2
                    Information = 3
                    Debug = 4
                    Verbose = 5
                    Success = 6
                    Failure = 7
                    PartialFailure = 8
                }

$_defaultHostTextColor = @{
                                Error = "Red"
                                Warning = "Yellow"
                                Information = "Cyan"
                                Debug = "White"
                                Verbose = "White"
                                Success = "Green"
                                Failure = "Red"
                                PartialFailure = "Yellow"
                            }

$_defaultLogConfiguration = @{   
                                LogLevel = "Debug"
                                LogFileName = "Script.log"
                                IncludeDateInFileName = $True
                                OverwriteLogFile = $True
                                WriteToHost = $True
								MessageFormat = "{Timestamp:yyyy-MM-dd hh:mm:ss.fff} | {CallingObjectName} | {MessageType} | {Message}"
                                HostTextColor = $_defaultHostTextColor
                            }

$_defaultTimestampFormat = "yyyy-MM-dd hh:mm:ss.fff"	
						
$_logConfiguration = @{}
$_messageFormatInfo = @{}

$_logFilePath = ""
$_logFileOverwritten = $False

Reset-LogConfiguration

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
The logging configuration is flexible and powerful, allowing the following properties of log 
messages to be changed:

    1) The log level:  This determines whether a log message will be logged or not.  
        
        Possible log levels, in order from lowest to highest, are:
            "Off"
            "Error"
            "Warning"
            "Information"
            "Debug"
            "Verbose" 

        Only log messages at a level the same as or lower than the LogLevel will be logged.  For 
        example, if the LogLevel is "Information" then only log messages at a level of 
        Information, Warning or Error will be logged.  Messages at a level of Debug or Verbose 
        will not be logged, as these log levels are higher than Information;

    2) The message destination:  Messages may be written to the host or to PowerShell streams such 
        as the Information stream or the Verbose stream.  In addition, if a log file name is 
        set in the logging configuration, the messages will be written to a log file;

    3) The host text color:  Messages written to the host, as opposed to PowerShell streams, may 
        be written in any PowerShell console color.  Different colors may be specified for 
        different message types, such as Error, Warning or Information;

    4) The message format:  In addition to the specified message, the text written to the log may 
        include additional fields that are automatically populated, such as a timestamp or the  
        name of the function writing to the log.  The format of the logged text, including the 
        fields to be displayed, may be specified in the logging configuration.

The logging configuration may be altered via function Set-LogConfiguration.  Function 
Get-LogConfiguration will return a copy of the current configuration as a hash table.  The 
configuration can be reset back to its default values via function Reset-LogConfiguration.

For more details on the log configuration and how to set it view the help topics for 
Get-LogConfiguration and Set-LogConfiguration.

Several configuration settings can be overridden when writing a single log message.  The changes 
apply only to that one message; subsequent messages will return to using the settings in the 
logging configuration.  Settings that can be overridden on a per-message basis are:

    1) The message destination:  The message can be logged to a different destination from the 
        one specified in the logging configuration by using the switch parameters 
        -WriteToHost or -WriteToStreams;

    2) The host text color:  If the message is being written to the host, as opposed to 
        PowerShell streams, its text color can be set via parameter -HostTextColor.  Any 
        PowerShell console color can be used;

    3) The message format:  The format of the message can be set via parameter -MessageFormat.

.PARAMETER Message
The message to be logged. 

.PARAMETER HostTextColor
The name of the text color if the message is to be written to the host.  

This is only used if the message is going to be written to the host.  If the message is to be 
written to a PowerShell stream, such as the Information stream, this color is ignored.

Acceptable values are: "Black", "DarkBlue", "DarkGreen", "DarkCyan", "DarkRed", "DarkMagenta", 
"DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", "Red", "Magenta", "Yellow", "White".

.PARAMETER MessageFormat: 
A string that sets the format of the text that will be logged.  

Text enclosed in curly braces, {...}, represents the name of a field which will be included in 
the logged text.  The field names are not case sensitive.  
        
Any other text, not enclosed in curly braces, will be treated as a string literal and will appear 
in the logged text exactly as specified.	
		
Leading spaces in the MessageFormat string will be retained when the text is written to the 
log to allow log messages to be indented.  Trailing spaces in the MessageFormat string will not 
be included in the logged text.
		
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

	{CallingObjectName} : The name of the function or script that is writing to the log.  

                    When determining the caller name all functions in this module will be ignored; 
                    the caller name will be the external function or script that calls into this 
                    module to write to the log.  
                            
                    If a function is writing to the log the function name will be displayed.  If 
                    the log is being written to from a script file, outside any function, the name 
                    of the script file will be displayed.  If the log is being written to manually 
                    from the Powershell console then "[CONSOLE]" will be displayed.

	{LogLevel}    : The LogLevel at which the message is being recorded.  For example, the message 
                    may be an Error message or a Debug message.  The LogLevel will always be 
                    displayed in upper case.

	{Result}      : Used with result-related message types: Success, Failure and PartialFailure.  
                    The Result will always be displayed in upper case.

	{MessageType} : The MessageType of the message.  This combines LogLevel and Result: It includes 
                    the log levels Error, Warning, Information, Debug and Verbose as well as the 
                    results Success, Failure and PartialFailure.  The MessageType will always be 
                    displayed in upper case.

.PARAMETER IsError
When set this switch parameter has five effects:

    1) It specifies the message is being logged at level Error.  The message will not be 
        logged if the LogLevel is set to Off in the logging configuration;

    2) If the message is set to be written to a PowerShell stream it will be written to the 
        Error stream;

    3) If the message is set to be written to the host it will be written using the text 
        color specified in configuration setting HostTextColor.Error, unless parameter 
        HostTextColor is used to override the HostTextColor.Error setting;

    4) The {LogLevel} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "ERROR";

    5) The {MessageType} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "ERROR".

IsError is one of the Message Type switch parameters.  Only one Message Type switch may be set 
at the same time.  The Message Type switch parameters are:
    IsError, IsWarning, IsInformation, IsDebug, IsVerbose, IsSuccessResult, IsFailureResult, 
    IsPartialFailureResult.

.PARAMETER IsWarning
When set this switch parameter has five effects:

    1) It specifies the message is being logged at level Warning.  The message will not be 
        logged if the LogLevel is set to Off or Error in the logging configuration;

    2) If the message is set to be written to a PowerShell stream it will be written to the 
        Warning stream;

    3) If the message is set to be written to the host it will be written using the text 
        color specified in configuration setting HostTextColor.Warning, unless parameter 
        HostTextColor is used to override the HostTextColor.Warning setting;

    4) The {LogLevel} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "WARNING";

    5) The {MessageType} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "WARNING".

IsWarning is one of the Message Type switch parameters.  Only one Message Type switch may be set 
at the same time.  The Message Type switch parameters are:
    IsError, IsWarning, IsInformation, IsDebug, IsVerbose, IsSuccessResult, IsFailureResult, 
    IsPartialFailureResult.

.PARAMETER IsInformation
When set this switch parameter has five effects:

    1) It specifies the message is being logged at level Information.  The message will not be 
        logged if the LogLevel is set to Off, Error or Warning in the logging configuration;

    2) If the message is set to be written to a PowerShell stream it will be written to the 
        Information stream;

    3) If the message is set to be written to the host it will be written using the text 
        color specified in configuration setting HostTextColor.Information, unless parameter 
        HostTextColor is used to override the HostTextColor.Information setting;

    4) The {LogLevel} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "INFORMATION";

    5) The {MessageType} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "INFORMATION".

IsInformation is one of the Message Type switch parameters.  Only one Message Type switch may be 
set at the same time.  The Message Type switch parameters are:
    IsError, IsWarning, IsInformation, IsDebug, IsVerbose, IsSuccessResult, IsFailureResult, 
    IsPartialFailureResult.

.PARAMETER IsDebug
When set this switch parameter has five effects:

    1) It specifies the message is being logged at level Debug.  The message will not be 
        logged if the LogLevel is set to Off, Error, Warning or Information in the logging 
        configuration;

    2) If the message is set to be written to a PowerShell stream it will be written to the 
        Debug stream;

    3) If the message is set to be written to the host it will be written using the text 
        color specified in configuration setting HostTextColor.Debug, unless parameter 
        HostTextColor is used to override the HostTextColor.Debug setting;

    4) The {LogLevel} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "DEBUG";

    5) The {MessageType} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "DEBUG".

IsDebug is one of the Message Type switch parameters.  Only one Message Type switch may be set 
at the same time.  The Message Type switch parameters are:
    IsError, IsWarning, IsInformation, IsDebug, IsVerbose, IsSuccessResult, IsFailureResult, 
    IsPartialFailureResult.

.PARAMETER IsVerbose
When set this switch parameter has five effects:

    1) It specifies the message is being logged at level Verbose.  The message will not be 
        logged if the LogLevel is set to Off, Error, Warning, Information or Verbose in the 
        logging configuration;

    2) If the message is set to be written to a PowerShell stream it will be written to the 
        Verbose stream;

    3) If the message is set to be written to the host it will be written using the text 
        color specified in configuration setting HostTextColor.Verbose, unless parameter 
        HostTextColor is used to override the HostTextColor.Verbose setting;

    4) The {LogLevel} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "VERBOSE";

    5) The {MessageType} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "VERBOSE".

IsVerbose is one of the Message Type switch parameters.  Only one Message Type switch may be set 
at the same time.  The Message Type switch parameters are:
    IsError, IsWarning, IsInformation, IsDebug, IsVerbose, IsSuccessResult, IsFailureResult, 
    IsPartialFailureResult.

.PARAMETER IsSuccessResult
When set this switch parameter has five effects:

    1) It specifies the message is being logged at level Information.  The message will not be 
        logged if the LogLevel is set to Off, Error or Warning in the logging configuration;

    2) If the message is set to be written to a PowerShell stream it will be written to the 
        Information stream;

    3) If the message is set to be written to the host it will be written using the text 
        color specified in configuration setting HostTextColor.Success, unless parameter 
        HostTextColor is used to override the HostTextColor.Success setting;

    4) The {Result} placeholder in the MessageFormat string, if present, will be replaced by the 
        text "SUCCESS";

    5) The {MessageType} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "SUCCESS".

IsSuccessResult is one of the Message Type switch parameters.  Only one Message Type switch may 
be set at the same time.  The Message Type switch parameters are:
    IsError, IsWarning, IsInformation, IsDebug, IsVerbose, IsSuccessResult, IsFailureResult, 
    IsPartialFailureResult.

.PARAMETER IsFailureResult
When set this switch parameter has five effects:

    1) It specifies the message is being logged at level Information.  The message will not be 
        logged if the LogLevel is set to Off, Error or Warning in the logging configuration;

    2) If the message is set to be written to a PowerShell stream it will be written to the 
        Information stream;

    3) If the message is set to be written to the host it will be written using the text 
        color specified in configuration setting HostTextColor.Failure, unless parameter 
        HostTextColor is used to override the HostTextColor.Failure setting;

    4) The {Result} placeholder in the MessageFormat string, if present, will be replaced by the 
        text "FAILURE";

    5) The {MessageType} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "FAILURE".

IsFailureResult is one of the Message Type switch parameters.  Only one Message Type switch may 
be set at the same time.  The Message Type switch parameters are:
    IsError, IsWarning, IsInformation, IsDebug, IsVerbose, IsSuccessResult, IsFailureResult, 
    IsPartialFailureResult.

.PARAMETER IsPartialFailureResult
When set this switch parameter has five effects:

    1) It specifies the message is being logged at level Information.  The message will not be 
        logged if the LogLevel is set to Off, Error or Warning in the logging configuration;

    2) If the message is set to be written to a PowerShell stream it will be written to the 
        Information stream;

    3) If the message is set to be written to the host it will be written using the text 
        color specified in configuration setting HostTextColor.PartialFailure, unless parameter 
        HostTextColor is used to override the HostTextColor.PartialFailure setting;

    4) The {Result} placeholder in the MessageFormat string, if present, will be replaced by the 
        text "PARTIAL FAILURE";

    5) The {MessageType} placeholder in the MessageFormat string, if present, will be replaced by 
        the text "PARTIAL FAILURE".

IsPartialFailureResult is one of the Message Type switch parameters.  Only one Message Type 
switch may be set at the same time.  The Message Type switch parameters are:
    IsError, IsWarning, IsInformation, IsDebug, IsVerbose, IsSuccessResult, IsFailureResult, 
    IsPartialFailureResult.

.PARAMETER WriteToHost
A switch parameter that, if set, will write the message to the host, as opposed to one of the 
PowerShell streams such as Error or Warning, overriding the boolean configuration setting 
OverwriteLogFile, which may be set to $True or $False.

WriteToHost and WriteToStreams cannot both be set at the same time.

.PARAMETER WriteToStreams
A switch parameter that complements WriteToHost.  If set the message will be written to a 
PowerShell stream.  This overrides the boolean boolean configuration setting OverwriteLogFile, 
which may be set to $True or $False.

Which PowerShell stream is written to is determined by which Message Type switch parameter is 
set: IsError, IsWarning, IsInformation, IsDebug, IsVerbose, IsSuccessResult, IsFailureResult, 
or IsPartialFailureResult.

WriteToHost and WriteToStreams cannot both be set at the same time.

#>
function Write-LogMessage (
    [Parameter(Mandatory=$True)]
    [AllowEmptyString()]
    [string]$Message,

    [Parameter(Mandatory=$False)]
    [string]$HostTextColor,      

    [Parameter(Mandatory=$False)]
    [switch]$MessageFormat,

    [Parameter(Mandatory=$False)]
    [switch]$IsError, 

    [Parameter(Mandatory=$False)]
    [switch]$IsWarning,

    [Parameter(Mandatory=$False)]
    [switch]$IsInformation, 

    [Parameter(Mandatory=$False)]
    [switch]$IsDebug, 

    [Parameter(Mandatory=$False)]
    [switch]$IsVerbose, 

    [Parameter(Mandatory=$False)]
    [switch]$IsSuccessResult, 

    [Parameter(Mandatory=$False)]
    [switch]$IsFailureResult, 

    [Parameter(Mandatory=$False)]
    [switch]$IsPartialFailureResult, 

    [Parameter(Mandatory=$False)]
    [switch]$WriteToHost,      

    [Parameter(Mandatory=$False)]
    [switch]$WriteToStreams
    )
{
    Private_ValidateSwitchParameterGroup -SwitchList $IsError,$IsWarning,$IsInformation,$IsDebug,$IsVerbose,$IsSuccessResult,$IsFailureResult,$IsPartialFailureResult `
		-ErrorMessage "Only one Message Type switch parameter may be set when calling the function. Message Type switch parameters: -IsError, -IsWarning, -IsInformation, -IsDebug, -IsVerbose, -IsSuccessResult, -IsFailureResult, -IsPartialFailureResult"

    Private_ValidateSwitchParameterGroup -SwitchList $WriteToHost,$WriteToStreams
		-ErrorMessage "Only one Destination switch parameter may be set when calling the function. Destination switch parameters: -WriteToHost, -WriteToStreams"
	
    $Timestamp = Get-Date
    $CallingObjectName = ""
    $LogLevel = $Null
    $Result = "UNKNOWN"
    $MessageType = $Null
    $TextColor = $Null

    $messageFormatInfo = $script:_messageFormatInfo
    if ($MessageFormat)
    {
        $messageFormatInfo = Private_GetMessageFormatInfo $MessageFormat
    }

    # Getting the calling object name is an expensive operation so only perform it if needed.
    if ($messageFormatInfo.FieldsPresent -contains "CallingObjectName")
    {
        $CallingObjectName = Private_GetCallingFunctionName
    }

    if ($IsError.IsPresent)
    {
        $LogLevel = "ERROR"
        $Result = ""
        $MessageType = "ERROR"
        $TextColor = $script:_logConfiguration.HostTextColor.Error
    }
    elseif ($IsWarning.IsPresent)
    {
        $LogLevel = "WARNING"
        $Result = ""
        $MessageType = "WARNING"
        $TextColor = $script:_logConfiguration.HostTextColor.Warning
    }
    elseif ($IsInformation.IsPresent)
    {
        $LogLevel = "INFORMATION"
        $Result = ""
        $MessageType = "INFORMATION"
        $TextColor = $script:_logConfiguration.HostTextColor.Information
    }
    elseif ($IsDebug.IsPresent)
    {
        $LogLevel = "DEBUG"
        $Result = ""
        $MessageType = "DEBUG"
        $TextColor = $script:_logConfiguration.HostTextColor.Debug
    }
    elseif ($IsVerbose.IsPresent)
    {
        $LogLevel = "VERBOSE"
        $Result = ""
        $MessageType = "VERBOSE"
        $TextColor = $script:_logConfiguration.HostTextColor.Verbose
    }
    elseif ($IsSuccessResult.IsPresent)
    {
        $LogLevel = "INFORMATION"
        $Result = "SUCCESS"
        $MessageType = "SUCCESS"
        $TextColor = $script:_logConfiguration.HostTextColor.Success
    }
    elseif ($IsFailureResult.IsPresent)
    {
        $LogLevel = "INFORMATION"
        $Result = "FAILURE"
        $MessageType = "FAILURE"
        $TextColor = $script:_logConfiguration.HostTextColor.Failure
    }
    elseif ($IsPartialFailureResult.IsPresent)
    {
        $LogLevel = "INFORMATION"
        $Result = "PARTIAL FAILURE"
        $MessageType = "PARTIAL FAILURE"
        $TextColor = $script:_logConfiguration.HostTextColor.PartialFailure
    }

    # Defaults.
    if (-not $MessageType)
    {
        $LogLevel = "INFORMATION"
        $Result = ""
        $MessageType = "INFORMATION"
        # For text color default to the current console text color.
    }

    if ($HostTextColor)
    {
        $TextColor = $HostTextColor
    }

    $textToLog = $ExecutionContext.InvokeCommand.ExpandString($messageFormatInfo.WorkingFormat)

    if ($WriteToHost.IsPresent)
    {
        if ($TextColor)
        {
            Write-Host $textToLog -ForegroundColor $TextColor
        }
        else
        {
            Write-Host $textToLog
        }
    }
    elseif ($WriteToStreams.IsPresent)
    {
        switch ($MessageType)
        {
            "ERROR"             { Write-Error $textToLog; break }
            "WARNING"           { Write-Warning $textToLog; break }
            "INFORMATION"       { Write-Information $textToLog; break }
            "DEBUG"             { Write-Debug $textToLog; break }
            "VERBOSE"           { Write-Verbose $textToLog; break }
            "SUCCESS"           { Write-Information $textToLog; break }
            "FAILURE"           { Write-Information $textToLog; break }
            "PARTIAL FAILURE"   { Write-Information $textToLog; break }
            default             { Write-Information $textToLog; break }
        }
    }

    if ([string]::IsNullOrWhiteSpace($script:_logConfiguration.LogFileName))
    {
        return
    }

    if (-not (Test-Path $script:_logFilePath -IsValid))
    {
        # Fail silently so that every message output to the console doesn't include an error 
        # message.
        return
    }

    if (($script:_logConfiguration.OverwriteLogFile -and (-not $script:_logFileOverwritten)) -or 
        (-not (Test-Path $script:_logFilePath)))
    {
        $textToLog | Set-Content $script:_logFilePath
        $script:_logFileOverwritten = $True
    }
    else
    {
        $textToLog | Add-Content $script:_logFilePath
    }
}

<#
.SYNOPSIS
Gets the name of the function calling into this module.

.DESCRIPTION
Walks up the call stack until it finds a stack frame where the ScriptName is not the filename of 
this module.  

If the call stack cannot be read then the function returns "[UNKNOWN CALLER]".  

If no stack frame is found with a different ScriptName then the function returns "----".

If the ScriptName of the first stack frame outside of this module is $Null then the module is 
being called from the PowerShell console.  In that case the function returns "[CONSOLE]".  

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
function Private_GetCallingFunctionName()
{
	$callStack = Get-PSCallStack
	if ($callStack -eq $null -or $callStack.Count -eq 0)
	{
		return "[UNKNOWN CALLER]"
	}
	
	$thisFunctionStackFrame = $callStack[0]
	$thisModuleFileName = $thisFunctionStackFrame.ScriptName
	$stackFrameFileName = $thisModuleFileName
	$i = 1
	$stackFrameFunctionName = "----"
	while ($stackFrameFileName -eq $thisModuleFileName -and $i -lt $callStack.Count)
	{
		$stackFrame = $callStack[$i]
		$stackFrameFileName = $stackFrame.ScriptName
		$stackFrameFunctionName = $stackFrame.FunctionName
		$i++
	}
	
	if ($stackFrameFileName -eq $null)
	{
		return "[CONSOLE]"
	}
	if ($stackFrameFunctionName -eq "<ScriptBlock>")
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
    LogLevel: The name of a log level.  This determines whether a log message will be logged or 
        not.  
        
        Possible log levels, in order from lowest to highest, are:
            "Off"
            "Error"
            "Warning"
            "Information"
            "Debug"
            "Verbose" 

        Only log messages at a level the same as or lower than the LogLevel will be logged.  For 
        example, if the LogLevel is "Information" then only log messages at a level of 
        Information, Warning or Error will be logged.  Messages at a level of Debug or Verbose 
        will not be logged, as these log levels are higher than Information;

    LogFileName: The path to the log file.  If LogFileName is $Null, empty or blank log messages 
        will be displayed on screen but not written to a log file.  If LogFileName is specified 
        without a path, or with a relative path, it will be relative to the directory of the 
        calling script, not this module.  The default value for LogFileName is "Script.log";

    IncludeDateInFileName: If $True then the log file name will have a date, of the form 
        "_yyyyMMdd" appended to the end of the file name.  For example, "Script_20171129.log".  
        The default value is $True;

    OverwriteLogFile: If $True any existing log file with the same name as LogFileName, including 
        a date if IncludeDateInFileName is set, will be overwritten.  If $False new log messages 
        will be appended to the end of the existing log file.  If no file with the same name 
        exists it will be created, regardless of the value of OverwriteLogFile.  The default value 
        is $True;

    WriteToHost: If $True then all log messages wiill be written to the host.  If $False then log 
        messages will be written to the appropriate stream.  For example, Error messages will be 
        written to the error stream, Warning messages will be written to the warning stream, etc.  
        The stream for Success, Failure and PartialFailure messages is the information stream.  
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
			{Message}     : The supplied text message to write to the log.

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
                            
                            The default datetime format string is "yyyy-MM-dd hh:mm:ss.fff"

			{CallingObjectName} : The name of the function or script that is writing to the log.  

                            When determining the caller name all functions in this module will be 
                            ignored; the caller name will be the external function or script that 
                            calls into this module to write to the log.  
                            
                            If a function is writing to the log the function name will be 
                            displayed.  If the log is being written to from a script file, outside 
                            any function, the name of the script file will be displayed.  If the 
                            log is being written to manually from the Powershell console then 
                            "[CONSOLE]" will be displayed.

			{LogLevel}    : The LogLevel at which the message is being recorded.  For example, the 
                            message may be an Error message or a Debug message.  The LogLevel will 
                            always be displayed in upper case.

			{Result}      : Used with result-related message types: Success, Failure and 
                            PartialFailure.  The Result will always be displayed in upper case.

			{MessageType} : The MessageType of the message.  This combines LogLevel and Result: 
                            It includes the log levels Error, Warning, Information, Debug and 
                            Verbose as well as the results Success, Failure and PartialFailure.  
                            If does not include the log level Off because in that case the 
                            message would not be logged.  The MessageType will always be 
                            displayed in upper case.
			
		The default MessageFormat is: 
		"{Timestamp:yyyy-MM-dd hh:mm:ss.fff} | {CallingObjectName} | {MessageType} | {Message}";

    HostTextColor: A hash table that specifies the different text colors that will be used for 
        different log levels, for log messages written to the host.  HostTextColor only applies 
        if WriteToHost is $True.  The hash table has the following keys:
            Error: The text color for messages of log level Error.  The default value is Red;

            Warning: The text color for messages of log level Warning.  The default value is 
            Yellow;

            Information: The text color for messages of log level Information.  The default 
            value is Cyan;

            Debug: The text color for messages of log level Debug.  The default value is White;

            Verbose: The text color for messages of log level Verbose.  The default value is White;

            Success: The text color for messages representing a result of Success.  The default 
                value is Green;

            Failure: The text color for messages representing a result of Failure.  The default 
                value is Red;

            PartialFailure: The text color for messages representing a result of 
                Partial Failure.  Partial Failure may be used where, for example, multiple items 
                are updated and some are updated successfully while some are not.  The default 
                value is Yellow.  

        Possible values for text colors are: "Black", "DarkBlue", "DarkGreen", "DarkCyan", 
        "DarkRed", "DarkMagenta", "DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", 
        "Red", "Magenta", "Yellow", "White".

.NOTES        
All result messages, Success, Failure and PartialFailure, are written at log level Information.    
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
Resets the log configuration to the default settings. 

.DESCRIPTION    
Resets the log configuration to the default settings. 
#>
function Reset-LogConfiguration()
{
    Set-LogConfiguration -LogConfiguration $script:_defaultLogConfiguration
}

<#
.SYNOPSIS
Sets one or more of the log configuration settings.    

.DESCRIPTION 
Sets one or more of the log configuration settings. 

.PARAMETER LogConfiguration
A hash table representing all configuration settings.  It must have the following keys:
    LogLevel: The name of a log level.  This determines whether a log message will be logged or 
        not.  
        
        Possible log levels, in order from lowest to highest, are:
            "Off"
            "Error"
            "Warning"
            "Information"
            "Debug"
            "Verbose" 

        Only log messages at a level the same as or lower than the LogLevel will be logged.  For 
        example, if the LogLevel is "Information" then only log messages at a level of 
        Information, Warning or Error will be logged.  Messages at a level of Debug or Verbose 
        will not be logged, as these log levels are higher than Information;

    LogFileName: The path to the log file.  If LogFileName is $Null, empty or blank log messages 
        will be displayed on screen but not written to a log file.  If LogFileName is specified 
        without a path, or with a relative path, it will be relative to the directory of the 
        calling script, not this module.  The default value for LogFileName is "Script.log";

    IncludeDateInFileName: If $True then the log file name will have a date, of the form 
        "_yyyyMMdd" appended to the end of the file name.  For example, "Script_20171129.log".  
        The default value is $True;

    OverwriteLogFile: If $True any existing log file with the same name as LogFileName, including 
        a date if IncludeDateInFileName is set, will be overwritten.  If $False new log messages 
        will be appended to the end of the existing log file.  If no file with the same name 
        exists it will be created, regardless of the value of OverwriteLogFile.  The default value 
        is $True;

    WriteToHost: If $True then all log messages wiill be written to the host.  If $False then log 
        messages will be written to the appropriate stream.  For example, Error messages will be 
        written to the error stream, Warning messages will be written to the warning stream, etc.  
        The stream for Success, Failure and PartialFailure messages is the information stream.  
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

			{CallingObjectName} : The name of the function or script that is writing to the log.  

                            When determining the caller name all functions in this module will be 
                            ignored; the caller name will be the external function or script that 
                            calls into this module to write to the log.  
                            
                            If a function is writing to the log the function name will be 
                            displayed.  If the log is being written to from a script file, outside 
                            any function, the name of the script file will be displayed.  If the 
                            log is being written to manually from the Powershell console then 
                            "[CONSOLE]" will be displayed.

			{LogLevel}    : The LogLevel at which the message is being recorded.  For example, the 
                            message may be an Error message or a Debug message.  The LogLevel will 
                            always be displayed in upper case.

			{Result}      : Used with result-related message types: Success, Failure and 
                            PartialFailure.  The Result will always be displayed in upper case.

			{MessageType} : The MessageType of the message.  This combines LogLevel and Result: 
                            It includes the log levels Error, Warning, Information, Debug and 
                            Verbose as well as the results Success, Failure and PartialFailure.  
                            If does not include the log level Off because in that case the 
                            message would not be logged.  The MessageType will always be 
                            displayed in upper case.
			
		The default MessageFormat is: 
		"{Timestamp:yyyy-MM-dd hh:mm:ss.fff} | {CallingObjectName} | {MessageType} | {Message}";

    HostTextColor: A hash table that specifies the different text colors that will be used for 
        different log levels, for log messages written to the host.  HostTextColor only applies 
        if WriteToHost is $True.  The hash table has the following keys:
            Error: The text color for messages of log level Error.  The default value is Red;

            Warning: The text color for messages of log level Warning.  The default value is 
            Yellow;

            Information: The text color for messages of log level Information.  The default 
            value is Cyan;

            Debug: The text color for messages of log level Debug.  The default value is White;

            Verbose: The text color for messages of log level Verbose.  The default value is White;

            Success: The text color for messages representing a result of Success.  The default 
                value is Green;

            Failure: The text color for messages representing a result of Failure.  The default 
                value is Red;

            PartialFailure: The text color for messages representing a result of 
                Partial Failure.  Partial Failure may be used where, for example, multiple items 
                are updated and some are updated successfully while some are not.  The default 
                value is Yellow.  

        Possible values for text colors are: "Black", "DarkBlue", "DarkGreen", "DarkCyan", 
        "DarkRed", "DarkMagenta", "DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", 
        "Red", "Magenta", "Yellow", "White". 

.PARAMETER LogLevel
The name of a log level, which determines whether a log message will be logged or not. 

Possible log levels, in order from lowest to highest, are:
    "Off"
    "Error"
    "Warning"
    "Information"
    "Debug"
    "Verbose" 

Only log messages at a level the same as or lower than the LogLevel will be logged.  For example, 
if the LogLevel is "Information" then only log messages at a level of Information, Warning or 
Error will be logged.  Messages at a level of Debug or Verbose will not be logged, as these log 
levels are higher than Information. 

.PARAMETER LogFileName
The path to the log file.  If LogFileName is $Null, empty or blank log messages will not be 
written to a log file, although they will be written to the host or to streams.  If LogFileName is 
specified without a path, or with a relative path, it will be relative to the directory of the 
calling script, not this module.

.PARAMETER IncludeDateInFileName
A switch parameter that, if set, will include a date in the log file name.  The date will take the 
form "_yyyyMMdd" appended to the end of the file name.  For example, "Script_20171129.log".  

IncludeDateInFileName and ExcludeDateFromFileName cannot both be set at the same time.

.PARAMETER ExcludeDateFromFileName
A switch parameter that is the opposite of IncludeDateInFileName.  If set it will exclude the date 
from the log file name.  For example, "Script.log".  

IncludeDateInFileName and ExcludeDateFromFileName cannot both be set at the same time.

.PARAMETER OverwriteLogFile
A switch parameter that, if set, will overwrite any existing log file with the same name as 
LogFileName, including a date if IncludeDateInFileName is set.  

OverwriteLogFile and AppendToLogFile cannot both be set at the same time.

.PARAMETER AppendToLogFile
A switch parameter that is the opposite of OverwriteLogFile.  If set new log messages will be 
appended to the end of an existing log file, if it has the same name as LogFileName, including a 
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

	{CallingObjectName} : The name of the function or script that is writing to the log.  

                    When determining the caller name all functions in this module will be ignored; 
                    the caller name will be the external function or script that calls into this 
                    module to write to the log.  
                            
                    If a function is writing to the log the function name will be displayed.  If 
                    the log is being written to from a script file, outside any function, the name 
                    of the script file will be displayed.  If the log is being written to manually 
                    from the Powershell console then "[CONSOLE]" will be displayed.

	{LogLevel}    : The LogLevel at which the message is being recorded.  For example, the message 
                    may be an Error message or a Debug message.  The LogLevel will always be 
                    displayed in upper case.

	{Result}      : Used with result-related message types: Success, Failure and PartialFailure.  
                    The Result will always be displayed in upper case.

	{MessageType} : The MessageType of the message.  This combines LogLevel and Result: It includes 
                    the log levels Error, Warning, Information, Debug and Verbose as well as the 
                    results Success, Failure and PartialFailure.  The MessageType will always be 
                    displayed in upper case.

.PARAMETER HostTextColorConfiguration
A hash table specifying the different text colors that will be used for different log levels, 
for log messages written to the host.  The hash table must have the following keys:
    Error: The text color for messages of log level Error.  The default value is Red;

    Warning: The text color for messages of log level Warning.  The default value is 
    Yellow;

    Information: The text color for messages of log level Information.  The default 
    value is Cyan;

    Debug: The text color for messages of log level Debug.  The default value is White;

    Verbose: The text color for messages of log level Verbose.  The default value is White;

    Success: The text color for messages representing a result of Success.  The default 
        value is Green;

    Failure: The text color for messages representing a result of Failure.  The default 
        value is Red;

    PartialFailure: The text color for messages representing a result of 
        Partial Failure.  Partial Failure may be used where, for example, multiple items 
        are updated and some are updated successfully while some are not.  The default 
        value is Yellow.  

Possible values for text colors are: "Black", "DarkBlue", "DarkGreen", "DarkCyan", 
"DarkRed", "DarkMagenta", "DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", 
"Red", "Magenta", "Yellow", "White".

.PARAMETER ErrorTextColor
The name of the text color for messages written to the host at log level Error.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are: "Black", "DarkBlue", "DarkGreen", "DarkCyan", "DarkRed", "DarkMagenta", 
"DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", "Red", "Magenta", "Yellow", "White".

.PARAMETER WarningTextColor
The name of the text color for messages written to the host at log level Warning.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are as per ErrorTextColor.

.PARAMETER InformationTextColor
The name of the text color for messages written to the host at log level Information.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are as per ErrorTextColor.

.PARAMETER DebugTextColor
The name of the text color for messages written to the host at log level Debug.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are as per ErrorTextColor.

.PARAMETER VerboseTextColor
The name of the text color for messages written to the host at log level Verbose.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are as per ErrorTextColor.

.PARAMETER SuccessTextColor
The name of the text color for Success messages written to the host.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are as per ErrorTextColor.

.PARAMETER FailureTextColor
The name of the text color for Failure messages written to the host.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are as per ErrorTextColor.

.PARAMETER PartialFailureTextColor
The name of the text color for PartialFailure messages written to the host.  

This is only used if WriteToHost is set.  If WriteToStreams is set this color is ignored.

Acceptable values are as per ErrorTextColor.
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
        [ValidateSet("Off", "Error", "Warning", "Information", "Debug", "Verbose")]
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
        [string]$VerboseTextColor,        
        
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [ValidateScript({Private_ValidateHostColor $_})]
        [string]$SuccessTextColor,         
        
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [ValidateScript({Private_ValidateHostColor $_})]
        [string]$FailureTextColor,         
        
        [parameter(ParameterSetName="IndividualSettings_IndividualColors")]
        [ValidateScript({Private_ValidateHostColor $_})]
        [string]$PartialFailureTextColor
    )

    if ($LogConfiguration -ne $Null)
    {
        $script:_logConfiguration = Private_DeepCopyHashTable $LogConfiguration
        Private_SetMessageFormat $LogConfiguration.MessageFormat
        Private_SetLogFilePath
        return
    }

    # Ensure that mutually exclusive pairs of switch parameters are not both set:

    Private_ValidateSwitchParameterGroup -SwitchList $IncludeDateInFileName,$ExcludeDateFromFileName
		-ErrorMessage "Only one FileName switch parameter may be set when calling the function. FileName switch parameters: -IncludeDateInFileName, -ExcludeDateFromFileName"

    Private_ValidateSwitchParameterGroup -SwitchList $OverwriteLogFile,$AppendToLogFile
		-ErrorMessage "Only one LogFileWriteBehavior switch parameter may be set when calling the function. LogFileWriteBehavior switch parameters: -OverwriteLogFile, -AppendToLogFile"

    Private_ValidateSwitchParameterGroup -SwitchList $WriteToHost,$WriteToStreams
		-ErrorMessage "Only one Destination switch parameter may be set when calling the function. Destination switch parameters: -WriteToHost, -WriteToStreams"

    if (![string]::IsNullOrWhiteSpace($LogLevel))
    {
        $script:_logConfiguration.LogLevel = $LogLevel
    }

    if (![string]::IsNullOrWhiteSpace($LogFileName))
    {
        $script:_logConfiguration.LogFileName = $LogFileName
        Private_SetLogFilePath
    }

    if ($ExcludeDateFromFileName.IsPresent)
    {
        $script:_logConfiguration.IncludeDateInFileName = $False
        Private_SetLogFilePath
    }

    if ($IncludeDateInFileName.IsPresent)
    {
        $script:_logConfiguration.IncludeDateInFileName = $True
        Private_SetLogFilePath
    }

    if ($AppendToLogFile.IsPresent)
    {
        $script:_logConfiguration.OverwriteLogFile = $False
    }

    if ($OverwriteLogFile.IsPresent)
    {
        $script:_logConfiguration.OverwriteLogFile = $True
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

    if (![string]::IsNullOrWhiteSpace($SuccessTextColor))
    {
        Private_SetConfigTextColor -ConfigurationKey "Success" -ColorName $SuccessTextColor
    }

    if (![string]::IsNullOrWhiteSpace($FailureTextColor))
    {
        Private_SetConfigTextColor -ConfigurationKey "Failure" -ColorName $FailureTextColor
    }

    if (![string]::IsNullOrWhiteSpace($PartialFailureTextColor))
    {
        Private_SetConfigTextColor -ConfigurationKey "PartialFailure" -ColorName $PartialFailureTextColor
    }
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
    if ([System.IO.Path]::IsPathRooted($Path))
    {
        return $Path
    }

    $callingDirectoryPath = $MyInvocation.PSScriptRoot

    $Path = Join-Path $callingDirectoryPath $Path

    return $Path
}

<#
.SYNOPSIS
Sets the full path to the log file.

.DESCRIPTION
Sets module variable $_logFilePath.  If $_logFilePath has changed then $_logFileOverwritten 
will be cleared.

Determines whether the LogFileName specified in the configuration settings is an absolute or a 
relative path.  If it is relative then the path to the directory the calling script is running 
in will be prepended to the specified LogFileName.

If configuration setting OverwriteLogFile is $True then the date will be included in the log 
file name, in the form: "<log file name>_yyyyMMdd.<file extension>".  For example, 
"Script_20171129.log".

.NOTES
This function is NOT intended to be exported from this module.

#>
function Private_SetLogFilePath ()
{
    $oldLogFilePath = $script:_logFilePath

    $logFilePath = Private_GetAbsolutePath $script:_logConfiguration.LogFileName

    if ($script:_logConfiguration.IncludeDateInFileName)
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

	{CallingObjectName} : The name of the function or script that is writing to the log.  

                    When determining the caller name all functions in this module will be ignored; 
                    the caller name will be the external function or script that calls into this 
                    module to write to the log.  
                            
                    If a function is writing to the log the function name will be displayed.  If 
                    the log is being written to from a script file, outside any function, the name 
                    of the script file will be displayed.  If the log is being written to manually 
                    from the Powershell console then "[CONSOLE]" will be displayed.

	{LogLevel}    : The LogLevel at which the message is being recorded.  For example, the message 
                    may be an Error message or a Debug message.  The LogLevel will always be 
                    displayed in upper case.

	{Result}      : Used with result-related message types: Success, Failure and PartialFailure.  
                    The Result will always be displayed in upper case.

	{MessageType} : The MessageType of the message.  This combines LogLevel and Result: It includes 
                    the log levels Error, Warning, Information, Debug and Verbose as well as the 
                    results Success, Failure and PartialFailure.  The MessageType will always be 
                    displayed in upper case.

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

.DESCRIPTION
Sets one of the host text color values in the log configuration settings.

.NOTES
This function is NOT intended to be exported from this module.
#>
function Private_SetConfigTextColor([string]$ConfigurationKey, [string]$ColorName)
{
    if (!$script:_logConfiguration.ContainsKey("HostTextColor"))
    {
        $script:_logConfiguration.HostTextColor = $script:_defaultHostTextColor
    }

    $script:_logConfiguration.HostTextColor[$ConfigurationKey] = $ColorName
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
		throw [System.ArgumentException] $ErrorMessage
	}
}

<#
.SYNOPSIS
Function called by ValidateScript to check if the specified host color name is valid when 
passed as a parameter.

.DESCRIPTION
If the specified color name is valid this function returns $True.  If the specified color name 
is not valid the function throws an exception rather than returns $False.  

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
    $validColors = @("Black", "DarkBlue", "DarkGreen", "DarkCyan", "DarkRed", "DarkMagenta", 
            "DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", "Red", "Magenta", 
            "Yellow", "White")
	
	if ($validColors -notcontains $ColorToTest)
	{
		throw [System.ArgumentException] "INVALID TEXT COLOR ERROR: '$ColorToTest' is not a valid text color for the PowerShell host."
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
    # The second colon in the regex pattern represents the separator between the placholder name 
    # and the datetime format string, eg {Timestamp:d}
    $regexPattern = "{\s*Timestamp\s*(?::\s*(.+)\s*)?\s*}"
	
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
                            it will default to "yyyy-MM-dd hh:mm:ss.fff".  
                            
                            The Timestamp placeholder will be replaced with 
                            "$($Timestamp.ToString('<datetime format string>'))".  
                            
                            Examples:
                                1)  Field placeholder "{Timestamp:d}" will be replaced by 
                                    "$($Timestamp.ToString('d'))";

                                2) Field placeholder "{Timestamp}" will use the default datetime 
                                    format string so will be replaced by 
                                    "$($Timestamp.ToString('yyyy-MM-dd hh:mm:ss.fff'))";

        $CallingObjectName :   Replaces field placeholder {CallingObjectName};

        $LogLevel       :   Replaces field placeholder {LogLevel};

        $Result         :   Replaces field placeholder {Result};

        $MessageType    :   Replaces field placeholder {MessageType};
        
    FieldsPresent: An array of strings representing the names of fields that will appear in the 
        log message.  Field names that may appear in the array are:

        "Message"       :   Included if the RawFormat string contains field placeholder {Message};
                            
        "Timestamp"     :   Included if the RawFormat string contains field placeholder 
                            {Timestamp};

        "CallingObjectName" :   Included if the RawFormat string contains field placeholder 
                            {CallingObjectName};

        "LogLevel"      :   Included if the RawFormat string contains field placeholder 
                            {LogLevel};

        "Result"        :   Included if the RawFormat string contains field placeholder 
                            {Result};

        "MessageType"   :   Included if the RawFormat string contains field placeholder 
                            {MessageType}.

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

    $modifiedText = $workingFormat -ireplace '{\s*Timestamp\s*(?::\s*.+\s*)?\s*}', $replacementText
    if ($modifiedText -ne $workingFormat)
    {
        $messageFormatInfo.FieldsPresent += "Timestamp"
        $workingFormat = $modifiedText
    }

    $modifiedText = $workingFormat -ireplace '{\s*CallingObjectName\s*}', '${CallingObjectName}'
    if ($modifiedText -ne $workingFormat)
    {
        $messageFormatInfo.FieldsPresent += "CallingObjectName"
        $workingFormat = $modifiedText
    }

    $modifiedText = $workingFormat -ireplace '{\s*LogLevel\s*}', '${LogLevel}'
    if ($modifiedText -ne $workingFormat)
    {
        $messageFormatInfo.FieldsPresent += "LogLevel"
        $workingFormat = $modifiedText
    }

    $modifiedText = $workingFormat -ireplace '{\s*Result\s*}', '${Result}'
    if ($modifiedText -ne $workingFormat)
    {
        $messageFormatInfo.FieldsPresent += "Result"
        $workingFormat = $modifiedText
    }

    $modifiedText = $workingFormat -ireplace '{\s*MessageType\s*}', '${MessageType}'
    if ($modifiedText -ne $workingFormat)
    {
        $messageFormatInfo.FieldsPresent += "MessageType"
        $workingFormat = $modifiedText
    }

     $messageFormatInfo.WorkingFormat = $workingFormat

     return $messageFormatInfo
}

#endregion

# Only export public functions.  To simplify the exporting of public functions but not private 
# ones public functions must follow the standard PowerShell naming convention, 
# "<verb>-<singular noun>", while private functions must not contain a dash, "-".
Export-ModuleMember -Function *-*