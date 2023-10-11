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

. $PSScriptRoot\Configuration.ps1

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
Writes a message to the PowerShell host or a PowerShell stream and, optionally, to a log file.

.DESCRIPTION
Writes a message to either the PowerShell host or to a PowerShell stream, such as the Information 
stream or the Verbose stream, depending on the logging configuration.  In addition, the message may 
be written to a log file, once again depending on the logging configuration.

.NOTES
The Pslogg logger can be configured via function Set-LogConfiguration with settings that persist 
between messages.  For example, it can be configured to write to the PowerShell host, or to 
PowerShell streams such as the Error stream or the Verbose stream.

The most important configuration setting is the LogLevel, the session logging level.  This 
determines which messages will be logged and which will not.  

Possible LogLevels, in order from highest to lowest, are:
    VERBOSE
    DEBUG
    INFORMATION
    WARNING
    ERROR
    OFF
        
Each message to be logged has a Message Level.  This may be set explicitly when calling 
Write-LogMessage or the default value of INFORMATION may be used.  The Message Level is compared 
to the LogLevel in the logger configuration.  Only messages with a Message Level the same as or 
lower than the configured LogLevel will be logged.  

For example, if the LogLevel is INFORMATION then only messages with a Message Level of 
INFORMATION, WARNING or ERROR will be logged.  Messages with a Message Level of DEBUG or 
VERBOSE will not be logged, as those levels are higher than INFORMATION.

When calling Write-LogMessage the Message Level can be set in two different ways:

    1) Via parameter -MessageLevel:  The Message Level is specified as text.  For example:

        Write-LogMessage 'Hello world' -MessageLevel 'VERBOSE'

    2) Via Message Level switch parameters:  There are switch parameters for each possible 
        Message Level: -IsVerbose, -IsDebug, -IsInformation, -IsWarning and -IsError.  For 
        example:

        Write-LogMessage 'Hello world' -IsVerbose

        Only one Message Level switch may be set for a given message.  

Several configuration settings can be overridden for a single log message.  The changes apply 
only to that one message; subsequent messages will return to using the settings in the logger  
configuration.  Settings that can be overridden on a per-message basis are:

    1) The message destination:  The message can be logged to a different destination from the 
        one specified in the logger configuration by using the switch parameters -WriteToHost or 
        -WriteToStreams;

    2) The host text color:  If the message is being written to the PowerShell host, as opposed to 
        PowerShell streams, its text color can be set via parameter -HostTextColor.  Any valid 
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
A string that sets the format of the text that will be logged.    If MessageFormat is not 
specified explicitly when calling Write-LogMessage the MessageFormat from the logger configuration 
will be used.

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

                    The Timestamp field may include an optional datetime format string inside 
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
                    'yyyy-MM-dd HH:mm:ss.fff', will be used;

	{CallerName}  : The name of the function or script that is writing to the log.  

                    When determining the caller name all functions in the Pslogg module will be 
                    ignored; the caller name will be the external function or script that calls 
                    the Pslogg module to write to the log.  
                            
                    If a function is writing to the log the function name will be displayed.  If 
                    the log is being written to from a script file, outside any function, the name 
                    of the script file will be displayed.  If the log is being written to manually 
                    from the Powershell console then '[CONSOLE]' will be displayed;

	{Category}    : The Message Category.  If no Message Category is explicitly specified when 
                    calling Write-LogMessage the default Category from the logger configuration 
                    will be used.  The Category will always be displayed in upper case;

	{MessageLevel} : The Message Level at which the message is being recorded.  For example, the 
                    message may be an Error message or a Debug message.  The MessageLevel will 
                    always be displayed in upper case.

.PARAMETER MessageLevel
A string that specifies the Message Level of the message.  Possible values are the LogLevels:
    VERBOSE
    DEBUG
    INFORMATION
    WARNING
    ERROR

The Message Level is compared to the LogLevel in the logger configuration.  If the Message Level 
is the same as or lower than the LogLevel the message will be logged.  If the Message Level is 
higher than the LogLevel the message will not be logged.

For example, if the LogLevel is INFORMATION then only messages with a Message Level of 
INFORMATION, WARNING or ERROR will be logged.  Messages with a Message Level of DEBUG or 
VERBOSE will not be logged, as those levels are higher than INFORMATION.

-MessageLevel cannot be specified at the same time as one of the Message Level switch parameters: 
-IsVerbose, -IsDebug, -IsInformation, -IsWarning and -IsError.  Either -MessageLevel can be 
specified or one or the Message Level switches can be specified but not both.

In addition to determining whether the message will be logged or not, -MessageLevel has the 
following effects:

    1) If the message is set to be written to a PowerShell stream it determines which stream the 
        message will be written to: The Verbose stream, the Debug stream, the Information stream, 
	the Warning stream or the Error stream;

    2) If the message is set to be written to the host and the -HostTextColor parameter is not 
        specified -MessageLevel determines the ForegroundColor the message will be written in.  
        The appropriate color is read from the logger configuration HostTextColor hash table.  
        For example, if the -MessageLevel is ERROR the text ForegroundColor will be set to the 
        color specified by logger configuration HostTextColor.Error;

    3) The {MessageLevel} placeholder in the MessageFormat string, if present, will be replaced 
        by the -MessageLevel text.  For example, if -MessageLevel is ERROR the {MessageLevel} 
        placeholder will be replaced by the text 'ERROR'.

Using the Message Level switches, for example -IsDebug, to set the Message Level is more concise 
than using -MessageLevel and specifying the Message Level via text.  However, -MessageLevel is 
useful for setting the  Message Level programmatically, based on the outcome of some process.

.PARAMETER IsVerbose
Sets the Message Level to VERBOSE.

-IsVerbose is one of the Message Level switch parameters.  Only one Message Level switch may be set 
at the same time.  The Message Level switch parameters are:
    -IsVerbose, -IsDebug, -IsInformation, -IsWarning, -IsError.

.PARAMETER IsDebug
Sets the Message Level to DEBUG.

-IsDebug is one of the Message Level switch parameters.  Only one Message Level switch may be set 
at the same time.  The Message Level switch parameters are:
    -IsVerbose, -IsDebug, -IsInformation, -IsWarning, -IsError.

.PARAMETER IsInformation
Sets the Message Level to INFORMATION.

-IsInformation is one of the Message Level switch parameters.  Only one Message Level switch may be 
set at the same time.  The Message Level switch parameters are:
    -IsVerbose, -IsDebug, -IsInformation, -IsWarning, -IsError.

.PARAMETER IsWarning
Sets the Message Level to WARNING.

-IsWarning is one of the Message Level switch parameters.  Only one Message Level switch may be set 
at the same time.  The Message Level switch parameters are:
    -IsVerbose, -IsDebug, -IsInformation, -IsWarning, -IsError.

.PARAMETER IsError
Sets the Message Level to ERROR.  

-IsError is one of the Message Level switch parameters.  Only one Message Level switch may be set 
at the same time.  The Message Level switch parameters are:
    -IsVerbose, -IsDebug, -IsInformation, -IsWarning, -IsError.

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
PowerShell streams such as Error or Warning.

-WriteToHost overrides the logger configuration setting WriteToHost.

-WriteToHost and -WriteToStreams cannot both be set at the same time.

.PARAMETER WriteToStreams
A switch parameter that complements -WriteToHost.  If set the message will be written to a 
PowerShell stream.  

-WriteToStreams overrides the logger configuration setting WriteToHost.

Which PowerShell stream is written to is determined by the Message Level, which may be set via 
the -MessageLevel parameter or by one of the Message Level switch parameters: 
-IsVerbose, -IsDebug, -IsInformation, -IsWarning or -IsError.

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
        [switch]$IsVerbose, 

        [Parameter(Mandatory=$False,
                    ParameterSetName='MessageLevelSwitches')]
        [switch]$IsDebug, 

        [Parameter(Mandatory=$False,
                    ParameterSetName='MessageLevelSwitches')]
        [switch]$IsInformation, 

        [Parameter(Mandatory=$False,
                    ParameterSetName='MessageLevelSwitches')]
        [switch]$IsWarning,

        [Parameter(Mandatory=$False,
                    ParameterSetName='MessageLevelSwitches')]
        [switch]$IsError, 

        [Parameter(Mandatory=$False)]
        [string]$Category,

        [Parameter(Mandatory=$False)]
        [switch]$WriteToHost,      

        [Parameter(Mandatory=$False)]
        [switch]$WriteToStreams
    )

    Private_ValidateSwitchParameterGroup -SwitchList $IsVerbose,$IsDebug,$IsInformation,$IsWarning,$IsError `
		-ErrorMessage 'Only one Message Level switch parameter may be set when calling the function. Message Level switch parameters: -IsVerbose, -IsDebug, -IsInformation, -IsWarning, -IsError'

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

    if (-not (Test-Path $script:_logConfiguration.LogFile.FullPath -IsValid))
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
        Set-Content -Path $script:_logConfiguration.LogFile.FullPathReadOnly -Value $textToLog
        $script:_logFileOverwritten = $True
    }
    else
    {
        Add-Content -Path $script:_logConfiguration.LogFile.FullPathReadOnly -Value $textToLog
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

# Set up initial conditions.
Reset-LogConfiguration

# Only export public functions.  To simplify the exporting of public functions but not private 
# ones public functions must follow the standard PowerShell naming convention, 
# "<verb>-<singular noun>", while private functions must not contain a dash, "-".
Export-ModuleMember -Function *-*
