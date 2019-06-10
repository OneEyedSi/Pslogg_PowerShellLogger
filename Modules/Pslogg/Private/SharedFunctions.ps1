# Private functions shared by Configuration and Logging scripts.

. $PSScriptRoot\ModuleState.ps1

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