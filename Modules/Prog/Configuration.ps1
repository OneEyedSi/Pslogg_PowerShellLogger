# Functions for reading and setting the Prog logger configuration.

. $PSScriptRoot\Private\SharedFunctions.ps1

<#
.SYNOPSIS
Gets a copy of the log configuration settings.

.DESCRIPTION
Gets a copy of the log configuration settings.  

The hash table returned by Get-LogConfiguration is a copy of the Prog configuration, NOT a
reference to the live configuration.  This means any changes to the hash table retrieved by 
Get-LogConfiguration will NOT be reflected in Prog's configuration.

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
Get-LogConfiguration returns a copy of the Prog configuration, NOT a reference to the live 
configuration.  As a result the Prog configuration can only be updated via Set-LogConfiguration.  
This ensures that the Prog internal state is updated correctly.  

For example, if a user were able to use Get-LogConfiguration to access the live configuration 
and modify it to set the configuration MessageFormat string directly, the modified MessageFormat 
would not be used when writing log messages.  That is because Set-LogConfiguration parses the 
new MessageFormat string and updates Prog's internal state to indicate which fields are to be 
included in log messages.  If the configuration MessageFormat string were updated directly it 
would not be parsed and the list of fields to include in log messages would not be updated.

Although changes to the hash table retrieved by Get-LogConfiguration will not be reflected in 
the Prog configuration, the updated hash table can be written back into the Prog configuration 
via Set-LogConfiguration.  

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

Only one CategoryInfo item can have the IsDefault property set.  If one of the supplied items has 
IsDefault set then the IsDefault property will be removed from all existing items.  If multiple 
supplied items have IsDefault set then only the last one processed will end up with IsDefault.  
The last item processed will depend on the sort order of the CategoryInfo hash table Keys 
collection.

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
Use Get-LogConfiguration and Set-LogConfiguration to update the configuration:

    $config = Get-LogConfiguration
    $config.LogLevel = 'ERROR'
    $config.LogFile.Name = 'Error.log'
    $config.CategoryInfo['FileCopy'] = @{Color = 'DarkYellow'}
    Set-LogConfiguration -LogConfiguration $config

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
                "Expected second element to be a hash table but it is $($value.GetType().FullName).", 
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
                    "Expected value to be a hash table but it is $($value.GetType().FullName).", 
                    'CategoryInfoItem')
            }
        }

        return $True
    }

    throw [ArgumentException]::new( `
        "Expected argument to be either a hash table or an array but it is $($CategoryInfoItem.GetType().FullName).",
        'CategoryInfoItem')
}

<#
.SYNOPSIS
Sets a configuration CategoryInfo item.

.DESCRIPTION
The item to set is specified via the -Key and -Value parameters.  If the value hash table contains 
an IsDefault key then any existing value hash table with an IsDefault key will have the key 
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