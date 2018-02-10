<#
.SYNOPSIS
Tests of the exported configuration functions in the Logging module.

.DESCRIPTION
Pester tests of the configuration-related functions exported from the Logging module.
#>

# PowerShell allows multiple modules of the same name to be imported from different locations.  
# This would confuse Pester.  So, to be sure there are not multiple Logging modules imported, 
# remove all Logging modules and re-import only one.
Get-Module Logging | Remove-Module -Force
# Use $PSScriptRoot so this script will always import the Logging module in the Modules folder 
# adjacent to the folder containing this script, regardless of the location that Pester is 
# invoked from:
#                                     {parent folder}
#                                             |
#                   -----------------------------------------------------
#                   |                                                   |
#     {folder containing this script}                                Modules folder
#                   \                                                   |
#                    ------------------> imports                     Logging module folder
#                                                \                      |
#                                                 -----------------> Logging.psm1 module script
Import-Module (Join-Path $PSScriptRoot ..\Modules\Logging\Logging.psm1 -Resolve) -Force

InModuleScope Logging {

    <#
    .SYNOPSIS
    Compares two hashtables and returns an array of error messages describing the differences.

    .DESCRIPTION
    Compares two hash tables and returns an array of error messages describing the differences.  If 
    there are no differences the array will be empty.

    .NOTES
    The function only deals with hashtable values of the following data types:
        Value types, such as integers;
        Strings;
        Arrays;
        Hashtables

    It specifically cannot deal with values that are reference types, such as objects.  While it can 
    deal with values that are arrays, it assumes those arrays do not contain reference types or 
    nested hashtables.

    .INPUTS
    Two hashtables.

    .OUTPUTS
    An array of strings.
    #>
    function GetHashTableDifferences (
        [hashtable]$HashTable1,
        [hashtable]$HashTable2, 
        [int]$IndentLevel = 0
    )
    {
        $spacesPerIndent = 4
        $indentSpaces = ' ' * $spacesPerIndent * $IndentLevel

        if ($HashTable1 -eq $Null)
        {
            if ($HashTable2 -eq $Null)
            {
                return @()
            }
            return @($indentSpaces + 'Hashtable 1 is $Null')
        }
        # HashTable1 must be non-null...
        if ($HashTable2 -eq $Null)
        {
            return @($indentSpaces + 'Hashtable 2 is $Null')
        }
        # Both hashtables are non-null...

        # Reference equality: Both hashtables reference the same hashtable object.
        if ($HashTable1 -eq $HashTable2)
        {
            return @()
        }

        # The two hashtables are not pointing to the same object...

        $returnArray = @()

        # Compare-Object doesn't work on the hashtable.Keys collections.  It assumes all keys in 
        # hashtable 1 are missing from 2 and vice versa.  Compare-Object works properly if the 
        # Keys collections are converted to arrays first.
        # CopyTo will only work if the keys array is created with the right length first.
        $keys1 = @($Null) * $HashTable1.Keys.Count
        $HashTable1.Keys.CopyTo($keys1, 0)
        $keys2 = @($Null) * $HashTable2.Keys.Count
        $HashTable2.Keys.CopyTo($keys2, 0)
        # 
        # Result will be a list of the keys that exist in one hashtable but not the other.  If all 
        # keys match an empty array will be returned.
        $result = Compare-Object -ReferenceObject $keys1 -DifferenceObject $keys2
        if ($result)
        {
            $keysMissingFrom2 = $result | 
                Where-Object {$_.SideIndicator -eq '<='} | 
                Select-Object InputObject -ExpandProperty InputObject
            if ($keysMissingFrom2)
            {            
                $returnArray += "${indentSpaces}Keys missing from hashtable 2: $($keysMissingFrom2 -join ', ')"
            }

            $keysAddedTo2 = $result | 
                Where-Object {$_.SideIndicator -eq '=>'} | 
                Select-Object InputObject -ExpandProperty InputObject
            if ($keysAddedTo2)
            {            
                $returnArray += "${indentSpaces}Keys added to hashtable 2: $($keysAddedTo2 -join ', ')"
            }
        }

        foreach ($key in $HashTable1.Keys)
        {
            $value1 = $hashTable1[$key]
            $typeName1 = $value1.GetType().FullName

            if (-not $HashTable2.ContainsKey($key))
            {
                continue
            }

            $value2 = $HashTable2[$key]
            $typeName2 = $value2.GetType().FullName

            if ($typeName1 -ne $typeName2)
            {
                $returnArray += "${indentSpaces}The data types of key [${key}] differ in the hashtables:  Hashtable 1 data type: $typeName1; Hashtable 2 data type: $typeName2" 
                continue
            }

            # $typeName1 and ...2 are identical, ie the values for the matching keys are of the same 
            # data type in the two hashtables...

            # Compare-Object, at the parent hashtable level, will always assume nested hashtables are 
            # identical, even if they aren't.  So treat nested hashtables as a special case.
            if ($typeName1 -eq 'System.Collections.Hashtable')
            {            
                $nestedHashTableDifferences = GetHashTableDifferences `
                    -HashTable1 $value1 -HashTable2 $value2 -IndentLevel ($IndentLevel + 1)
                if ($nestedHashTableDifferences)
                {
                    $returnArray += "${indentSpaces}The nested hashtables at key [${key}] differ:"
                    $returnArray += $nestedHashTableDifferences
                    continue
                }
            }

            # Arrays, strings and value types can be compared via Compare-Object.
            # ASSUMPTION: That no values are reference types and any arrays do not contain 
            # reference types or hashtables.

            # SyncWindow = 0 ensures arrays will be compared in element order.  If one array is 
            # @(1, 2, 3) and the other is @(3, 2, 1) with SyncWindow = 0 these would be seen as 
            # different.  Leaving out the SyncWindow parameter or setting it to a larger number the 
            # two arrays would be seen as identical.
            $result = Compare-Object -ReferenceObject $value1 -DifferenceObject $value2 -SyncWindow 0
            if ($result)
            {
                if ($typeName1 -eq 'System.String')
                {
                    $value1 = "'$value1'"
                    $value2 = "'$value2'"
                }
                if ($typeName1 -eq 'System.Object[]')
                {
                    $value1 = "@($($value1 -join ', '))"
                    $value2 = "@($($value2 -join ', '))"
                }
                $returnArray += "${indentSpaces}The values at key [${key}] differ:  Hashtable 1 value: $value1; Hashtable 2 value: $value2"
            }
        }

        return $returnArray
    }

    function GetNewConfigurationColour()
    {
        $hostTextColor = @{
                            Error = "DarkYellow"
                            Warning = "DarkYellow"
                            Information = "DarkYellow"
                            Debug = "DarkYellow"
                            Verbose = "DarkYellow"
                            Success = "DarkYellow"
                            Failure = "DarkYellow"
                            PartialFailure = "DarkYellow"
                        }
        return $hostTextColor
    }

    # Gets a configuration hashtable where every setting is different from the defaults.
    function GetNewConfiguration ()
    {
        $logConfiguration = Private_DeepCopyHashTable $script:_defaultLogConfiguration
        $logConfiguration.LogLevel = "Verbose"
        $logConfiguration.LogFileName = "C:\Test\Test.txt"
        $logConfiguration.IncludeDateInFileName = $False
        $logConfiguration.OverwriteLogFile = $False
        $logConfiguration.WriteToHost = $False
        $logConfiguration.MessageFormat = "{LogLevel} | {Message}"
        
        $hostTextColor = GetNewConfigurationColour
        $logConfiguration.HostTextColor = $hostTextColor

        return $logConfiguration
    }

    # Gets the new MessageFormat that matches the configuration returned from MessageFormatInfo.
    function GetNewMessageFormat ()
    {
        $newConfiguration = GetNewConfiguration
        return $newConfiguration.MessageFormat
    }

    # Gets the MessageFormatInfo hashtable that matches the configuration returned from 
    # GetNewConfiguration.
    function GetNewMessageFormatInfo ()
    {
        $messageFormatInfo = @{
                                RawFormat = '{LogLevel} | {Message}'
                                WorkingFormat = '${LogLevel} | ${Message}'
                                FieldsPresent = @('Message', 'LogLevel')
                            }
        return $messageFormatInfo
    }

    # Gets the MessageFormatInfo hashtable that matches the default configuration.
    function GetDefaultMessageFormatInfo ()
    {
        $messageFormatInfo = @{
                                RawFormat = '{Timestamp:yyyy-MM-dd hh:mm:ss.fff} | {CallingObjectName} | {MessageType} | {Message}'
                                WorkingFormat = '$($Timestamp.ToString(''yyyy-MM-dd hh:mm:ss.fff'')) | ${CallingObjectName} | ${MessageType} | ${Message}'
                                FieldsPresent = @('Message', 'Timestamp', 'CallingObjectName', 'MessageType')
                            }
        return $messageFormatInfo
    }

    # Gets the LogFilePath that matches the configuration returned from GetNewConfiguration.
    function GetNewLogFilePath ([switch]$IncludeDateInFileName)
    {
        if ($IncludeDateInFileName)
        {
            $dateString = Get-Date -Format "_yyyyMMdd"
            return "C:\Test\Test${dateString}.txt"
        }

        return 'C:\Test\Test.txt'
    }

    function GetCallingDirectoryPath ()
    {
        $callStack = Get-PSCallStack
        $stackFrame = $callStack[0]
        # Skip this function in the call stack as we've already read it.
	    $i = 1
	    while ($stackFrame.ScriptName -ne $Null -and $i -lt $callStack.Count)
	    {
		    $stackFrame = $callStack[$i]
            if ($stackFrame.ScriptName -ne $Null)
		    {
                $stackFrameFileName = $stackFrame.ScriptName
            }
		    $i++
	    }
        $pathOfCallerDirectory = Split-Path -Path $stackFrameFileName -Parent
        return $pathOfCallerDirectory
    }

    function GetDefaultLogFilePath ()
    {
        $dateString = Get-Date -Format "_yyyyMMdd"
        $fileName = "Script${dateString}.log"
        # Can't use $PSScriptRoot because it will return the folder containing this file, while 
        # the Logging module will see the ultimate caller as the Pester module running this 
        # test script.
        $callingDirectoryPath = GetCallingDirectoryPath
        $path = Join-Path $callingDirectoryPath $fileName
        return $path
    }

    # Sets the Logging configuration to its defaults.
    function SetConfigurationToDefault ()
    {
        $script:_logConfiguration = Private_DeepCopyHashTable $script:_defaultLogConfiguration
        $script:_messageFormatInfo = GetDefaultMessageFormatInfo
        $script:_logFilePath = GetDefaultLogFilePath
        $script:_logFileOverwritten = $False
    }

    # Sets the Logging configuration so that every settings differs from the defaults.
    function SetNewConfiguration ()
    {
        $script:_logConfiguration = GetNewConfiguration
        $script:_messageFormatInfo = GetNewMessageFormatInfo
        $script:_logFilePath = GetNewLogFilePath
        $script:_logFileOverwritten = $True
    }

    # Verifies the specified hashtable is identical to the reference hashtable.
    function AssertHashTablesMatch 
        (
            [hashtable]$ReferenceHashTable, 
            [hashtable]$HashTableToTest, 
            [switch]$ShouldBeNotEqual
        ) 
    {
        $differences = GetHashTableDifferences `
            -HashTable1 $ReferenceHashTable `
            -HashTable2 $HashTableToTest
        if ($ShouldBeNotEqual)
        {
            $differences | Should -Not -Be @()
        }
        else
        {
            $differences | Should -Be @()
        }
    }

    function AssertLogConfigurationMatchesReference 
        (
            [hashtable]$ReferenceHashTable, 
            [switch]$ShouldBeNotEqual
        ) 
    {
        AssertHashTablesMatch -ReferenceHashTable $ReferenceHashTable `
            -HashTableToTest $script:_logConfiguration -ShouldBeNotEqual:$ShouldBeNotEqual
    }

    function AssertMessageFormatInfoMatchesReference 
        (
            [hashtable]$ReferenceHashTable, 
            [switch]$ShouldBeNotEqual
        ) 
    {
        AssertHashTablesMatch -ReferenceHashTable $ReferenceHashTable `
            -HashTableToTest $script:_messageFormatInfo -ShouldBeNotEqual:$ShouldBeNotEqual
    }

    Describe 'Get-LogConfiguration' {     
        BeforeEach {
            SetConfigurationToDefault
        }

        It 'returns a copy of default configuration if current configuration is $Null' {
            $script:_logConfiguration = $Null

            $logConfiguration = Get-LogConfiguration
            
            AssertHashTablesMatch -ReferenceHashTable $script:_defaultLogConfiguration `
                -HashTableToTest $logConfiguration
        } 

        It 'returns a copy of default configuration if current configuration is empty hashtable' {
            $script:_logConfiguration = @{}

            $logConfiguration = Get-LogConfiguration

            AssertHashTablesMatch -ReferenceHashTable $script:_defaultLogConfiguration `
                -HashTableToTest $logConfiguration
        } 

        It 'returns a copy of current configuration if current configuration hashtable is populated' {
            $script:_logConfiguration.Keys.Count | Should -BeGreaterThan 6

            $logConfiguration = Get-LogConfiguration
            
            AssertHashTablesMatch -ReferenceHashTable $script:_logConfiguration `
                -HashTableToTest $logConfiguration
        }  

        It 'returns a static copy of current configuration, which does not reflect subsequent changes to configuration' {
            $script:_logConfiguration.HostTextColor.Error = 'Blue'
            $script:_logConfiguration.HostTextColor.Error | Should -Be 'Blue'

            $logConfiguration = Get-LogConfiguration
            
            $script:_logConfiguration.HostTextColor.Error = 'White'
            $script:_logConfiguration.HostTextColor.Error | Should -Be 'White'

            $logConfiguration.HostTextColor.Error | Should -Be 'Blue'
        }  

        It 'returns a static copy of current configuration, where subsequent changes to the copy are not reflected in the configuration' {
            $script:_logConfiguration.HostTextColor.Error = 'Red'

            $logConfiguration = Get-LogConfiguration
            $logConfiguration.HostTextColor.Error = 'Blue'
            $logConfiguration.HostTextColor.Error | Should -Be 'Blue'

            $script:_logConfiguration.HostTextColor.Error | Should -Be 'Red'
        } 
    }

    Describe 'Set-LogConfiguration' {     

        BeforeEach {
            SetConfigurationToDefault

            AssertLogConfigurationMatchesReference `
                -ReferenceHashTable $script:_defaultLogConfiguration

            $defaultMessageFormatInfo = GetDefaultMessageFormatInfo
            AssertMessageFormatInfoMatchesReference -ReferenceHashTable $defaultMessageFormatInfo
        }

        Context 'Parameter set "AllSettings"' {
            It 'sets configuration to specified hashtable via LogConfiguration parameter' {
                
                $newConfiguration = GetNewConfiguration
                Set-LogConfiguration -LogConfiguration $newConfiguration

                AssertLogConfigurationMatchesReference -ReferenceHashTable $newConfiguration
            }

            It 'updates MessageFormatInfo from the MessageFormat details' {                

                $newConfiguration = GetNewConfiguration
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $newMessageFormatInfo = GetNewMessageFormatInfo
                AssertMessageFormatInfoMatchesReference -ReferenceHashTable $newMessageFormatInfo
            }

            It 'updates LogFilePath' {

                $newConfiguration = GetNewConfiguration
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $newLogFilePath = GetNewLogFilePath
                $script:_logFilePath | Should -Be $newLogFilePath
            }

            It 'includes date stamp in LogFilePath if new configuration IncludeDateInFileName set' {

                $newConfiguration = GetNewConfiguration
                $newConfiguration.IncludeDateInFileName = $True
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $newLogFilePath = GetNewLogFilePath -IncludeDateInFileName
                $script:_logFilePath | Should -Be $newLogFilePath
            }

            It 'clears LogFileOverwritten if configuration LogFileName changed' {
                $script:_logFileOverwritten = $True

                $newConfiguration = GetNewConfiguration
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $script:_logFileOverwritten | Should -Be $False
            }

            It 'does not clear LogFileOverwritten if configuration LogFileName unchanged' {
                $script:_logFileOverwritten = $True
                $script:_logConfiguration.LogFileName = GetNewLogFilePath
                $script:_logFilePath = GetNewLogFilePath

                $newConfiguration = GetNewConfiguration
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $script:_logFileOverwritten | Should -Be $True
            }
        }

        Context 'Parameter LogLevel' {
            BeforeEach {                      
                $script:_logConfiguration.LogLevel = 'Debug'              
            }

            It 'sets configuration LogLevel as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -LogLevel Information `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogLevel | Should -Be Information
            }

            It 'sets configuration LogLevel as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -LogLevel Information -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogLevel | Should -Be Information
            }

            It 'sets configuration LogLevel as a member of the default parameter set' {
                                
                Set-LogConfiguration -LogLevel Information 
                    
                $script:_logConfiguration.LogLevel | Should -Be Information
            }
        }

        Context 'Parameter LogFileName' {
            BeforeEach {     
                $script:_logConfiguration.LogFileName = 'Text.log'
            }

            It 'sets configuration LogFileName as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -LogFileName New.test `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogFileName | Should -Be New.test
            }

            It 'sets configuration LogFileName as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -LogFileName New.test -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogFileName | Should -Be New.test
            }

            It 'sets configuration LogFileName as a member of the default parameter set' {
                                
                Set-LogConfiguration -LogFileName New.test 
                    
                $script:_logConfiguration.LogFileName | Should -Be New.test
            }

            It 'updates LogFilePath' {

                $newConfiguration = GetNewConfiguration
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $newLogFilePath = GetNewLogFilePath
                $script:_logFilePath | Should -Be $newLogFilePath
            }

            It 'includes date stamp in LogFilePath if IncludeDateInFileName set' {

                $newConfiguration = GetNewConfiguration
                $newConfiguration.IncludeDateInFileName = $True
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $newLogFilePath = GetNewLogFilePath -IncludeDateInFileName
                $script:_logFilePath | Should -Be $newLogFilePath
            }

            It 'does not include date stamp in LogFilePath if IncludeDateInFileName cleared' {

                $newConfiguration = GetNewConfiguration
                $newConfiguration.IncludeDateInFileName = $False
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $newLogFilePath = GetNewLogFilePath
                $script:_logFilePath | Should -Be $newLogFilePath
            }

            It 'clears LogFileOverwritten if configuration LogFileName changed' {
                $script:_logFileOverwritten = $True

                $newConfiguration = GetNewConfiguration
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $script:_logFileOverwritten | Should -Be $False
            }

            It 'does not clear LogFileOverwritten if configuration LogFileName unchanged' {
                $script:_logFileOverwritten = $True
                $script:_logConfiguration.LogFileName = GetNewLogFilePath
                $script:_logFilePath = GetNewLogFilePath

                $newConfiguration = GetNewConfiguration
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $script:_logFileOverwritten | Should -Be $True
            }
        }

        Context 'Parameter IncludeDateInFileName' {
            BeforeEach {   
                $script:_logConfiguration.IncludeDateInFileName = $False
            }

            It 'sets configuration IncludeDateInFileName as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -IncludeDateInFileName `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.IncludeDateInFileName | Should -Be $True
            }

            It 'sets configuration IncludeDateInFileName as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -IncludeDateInFileName -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.IncludeDateInFileName | Should -Be $True
            }

            It 'sets configuration IncludeDateInFileName as a member of the default parameter set' {
                                
                Set-LogConfiguration -IncludeDateInFileName
                    
                $script:_logConfiguration.IncludeDateInFileName | Should -Be $True
            }

            It 'clears LogFileOverwritten if calculated LogFilePath changed' {
                $script:_logFileOverwritten = $True
                $script:_logConfiguration.LogFileName = GetNewLogFilePath
                $script:_logFilePath = GetNewLogFilePath
                $script:_logConfiguration.IncludeDateInFileName = $False

                Set-LogConfiguration -IncludeDateInFileName

                $script:_logFileOverwritten | Should -Be $False
            }

            It 'does not clear LogFileOverwritten if calculated LogFilePath unchanged' {
                $script:_logFileOverwritten = $True
                $script:_logConfiguration.LogFileName = GetNewLogFilePath
                $script:_logFilePath = GetNewLogFilePath -IncludeDateInFileName
                $script:_logConfiguration.IncludeDateInFileName = $True

                Set-LogConfiguration -IncludeDateInFileName

                $script:_logFileOverwritten | Should -Be $True
            }
        }

        Context 'Parameter ExcludeDateFromFileName' {
            BeforeEach {              
                $script:_logConfiguration.IncludeDateInFileName = $True
            }

            It 'sets configuration ExcludeDateFromFileName as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -ExcludeDateFromFileName `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.IncludeDateInFileName | Should -Be $False
            }

            It 'sets configuration ExcludeDateFromFileName as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -ExcludeDateFromFileName -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.IncludeDateInFileName | Should -Be $False
            }

            It 'sets configuration ExcludeDateFromFileName as a member of the default parameter set' {
                                
                Set-LogConfiguration -ExcludeDateFromFileName
                    
                $script:_logConfiguration.IncludeDateInFileName | Should -Be $False
            }

            It 'clears LogFileOverwritten if calculated LogFilePath changed' {
                $script:_logFileOverwritten = $True
                $script:_logConfiguration.LogFileName = GetNewLogFilePath
                $script:_logFilePath = GetNewLogFilePath -IncludeDateInFileName
                $script:_logConfiguration.IncludeDateInFileName = $True

                Set-LogConfiguration -ExcludeDateFromFileName

                $script:_logFileOverwritten | Should -Be $False
            }

            It 'does not clear LogFileOverwritten if calculated LogFilePath unchanged' {
                $script:_logFileOverwritten = $True
                $script:_logConfiguration.LogFileName = GetNewLogFilePath
                $script:_logFilePath = GetNewLogFilePath
                $script:_logConfiguration.IncludeDateInFileName = $False

                Set-LogConfiguration -ExcludeDateFromFileName

                $script:_logFileOverwritten | Should -Be $True
            }
        }

        Context 'Parameter OverwriteLogFile' {
            BeforeEach {    
                $script:_logConfiguration.OverwriteLogFile = $False
            }

            It 'sets configuration OverwriteLogFile as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -OverwriteLogFile `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.OverwriteLogFile | Should -Be $True
            }

            It 'sets configuration OverwriteLogFile as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -OverwriteLogFile -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.OverwriteLogFile | Should -Be $True
            }

            It 'sets configuration OverwriteLogFile as a member of the default parameter set' {
                                
                Set-LogConfiguration -OverwriteLogFile
                    
                $script:_logConfiguration.OverwriteLogFile | Should -Be $True
            }
        }

        Context 'Parameter AppendToLogFile' {
            BeforeEach {                
                $script:_logConfiguration.OverwriteLogFile = $True
            }

            It 'sets configuration AppendToLogFile as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -AppendToLogFile `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.OverwriteLogFile | Should -Be $False
            }

            It 'sets configuration AppendToLogFile as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -AppendToLogFile -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.OverwriteLogFile | Should -Be $False
            }

            It 'sets configuration AppendToLogFile as a member of the default parameter set' {
                                
                Set-LogConfiguration -AppendToLogFile
                    
                $script:_logConfiguration.OverwriteLogFile | Should -Be $False
            }
        }

        Context 'Parameter WriteToHost' {
            BeforeEach {   
                $script:_logConfiguration.WriteToHost = $False
            }

            It 'sets configuration WriteToHost as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -WriteToHost `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.WriteToHost | Should -Be $True
            }

            It 'sets configuration WriteToHost as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -WriteToHost -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.WriteToHost | Should -Be $True
            }

            It 'sets configuration WriteToHost as a member of the default parameter set' {
                                
                Set-LogConfiguration -WriteToHost
                    
                $script:_logConfiguration.WriteToHost | Should -Be $True
            }
        }

        Context 'Parameter WriteToStreams' {
            BeforeEach {                
                $script:_logConfiguration.WriteToHost = $True
            }

            It 'sets configuration WriteToStreams as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -WriteToStreams `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.WriteToHost | Should -Be $False
            }

            It 'sets configuration WriteToStreams as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -WriteToStreams -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.WriteToHost | Should -Be $False
            }

            It 'sets configuration WriteToStreams as a member of the default parameter set' {
                                
                Set-LogConfiguration -WriteToStreams
                    
                $script:_logConfiguration.WriteToHost | Should -Be $False
            }
        }

        Context 'Parameter MessageFormat' {
            BeforeEach {                
                $script:_logConfiguration.MessageFormat = 'original format'
                $newFormat = GetNewMessageFormat
            }

            It 'sets configuration MessageFormat as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -MessageFormat $newFormat `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.MessageFormat | Should -Be $newFormat
            }

            It 'sets configuration MessageFormat as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -MessageFormat $newFormat -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.MessageFormat | Should -Be $newFormat
            }

            It 'sets configuration MessageFormat as a member of the default parameter set' {
                                
                Set-LogConfiguration -MessageFormat $newFormat
                    
                $script:_logConfiguration.MessageFormat | Should -Be $newFormat
            }

            It 'updates MessageFormatInfo from the MessageFormat details' {                

                Set-LogConfiguration -MessageFormat $newFormat

                $newMessageFormatInfo = GetNewMessageFormatInfo
                AssertMessageFormatInfoMatchesReference -ReferenceHashTable $newMessageFormatInfo
            }
        }

        Context 'Parameter set "IndividualSettings_AllColors"' {

            It 'sets configuration HostTextColor hashtable via HostTextColorConfiguration parameter' {
                $newHostTextColours = GetNewConfigurationColour
                $differences = GetHashTableDifferences `
                    -HashTable1 $script:_logConfiguration.HostTextColor `
                    -HashTable2 $newHostTextColours
                $differences | Should -Not -Be @()

                Set-LogConfiguration -HostTextColorConfiguration $newHostTextColours
                    
                $differences = GetHashTableDifferences `
                    -HashTable1 $script:_logConfiguration.HostTextColor `
                    -HashTable2 $newHostTextColours
                $differences | Should -Be @()


            }
        }

        Context 'Parameter set "IndividualSettings_IndividualColors"' {

            It 'sets HostTextColor Error via parameter ErrorTextColor' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'

                Set-LogConfiguration -ErrorTextColor Blue
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be Blue
            }

            It 'sets HostTextColor Warning via parameter WarningTextColor' {
                $script:_logConfiguration.HostTextColor.Warning = 'Red'

                Set-LogConfiguration -WarningTextColor Blue
                    
                $script:_logConfiguration.HostTextColor.Warning | Should -Be Blue
            }

            It 'sets HostTextColor Information via parameter InformationTextColor' {
                $script:_logConfiguration.HostTextColor.Information = 'Red'

                Set-LogConfiguration -InformationTextColor Blue
                    
                $script:_logConfiguration.HostTextColor.Information | Should -Be Blue
            }

            It 'sets HostTextColor Debug via parameter DebugTextColor' {
                $script:_logConfiguration.HostTextColor.Debug = 'Red'

                Set-LogConfiguration -DebugTextColor Blue
                    
                $script:_logConfiguration.HostTextColor.Debug | Should -Be Blue
            }

            It 'sets HostTextColor Verbose via parameter VerboseTextColor' {
                $script:_logConfiguration.HostTextColor.Verbose = 'Red'

                Set-LogConfiguration -VerboseTextColor Blue
                    
                $script:_logConfiguration.HostTextColor.Verbose | Should -Be Blue
            }

            It 'sets HostTextColor Success via parameter SuccessTextColor' {
                $script:_logConfiguration.HostTextColor.Success = 'Red'

                Set-LogConfiguration -SuccessTextColor Blue
                    
                $script:_logConfiguration.HostTextColor.Success | Should -Be Blue
            }

            It 'sets HostTextColor Failure via parameter FailureTextColor' {
                $script:_logConfiguration.HostTextColor.Failure = 'Red'

                Set-LogConfiguration -FailureTextColor Blue
                    
                $script:_logConfiguration.HostTextColor.Failure | Should -Be Blue
            }

            It 'sets HostTextColor PartialFailure via parameter PartialFailureTextColor' {
                $script:_logConfiguration.HostTextColor.PartialFailure = 'Red'

                Set-LogConfiguration -PartialFailureTextColor Blue
                    
                $script:_logConfiguration.HostTextColor.PartialFailure | Should -Be Blue
            }
        }

        Context 'Multiple parameters set simultaneously' {
            It 'sets multiple configuration properties if multiple parameters are set' {
                $script:_logFileOverwritten = $True

                Set-LogConfiguration -LogLevel Verbose -LogFileName 'C:\Test\Test.txt' `
                    -ExcludeDateFromFileName -AppendToLogFile -WriteToStreams `
                    -MessageFormat '{LogLevel} | {Message}' -ErrorTextColor DarkYellow `
                    -WarningTextColor DarkYellow -InformationTextColor DarkYellow `
                    -DebugTextColor DarkYellow -VerboseTextColor DarkYellow `
                    -SuccessTextColor DarkYellow -FailureTextColor DarkYellow `
                    -PartialFailureTextColor DarkYellow

                $referenceLogConfiguration = GetNewConfiguration
                $referenceMessageFormatInfo = GetNewMessageFormatInfo
                $newLogFilePath = GetNewLogFilePath
                $newLogFileOverwritten = $False

                AssertLogConfigurationMatchesReference `
                    -ReferenceHashTable $referenceLogConfiguration
                
                AssertMessageFormatInfoMatchesReference `
                    -ReferenceHashTable $referenceMessageFormatInfo

                $script:_logFilePath | Should -Be $newLogFilePath
                $script:_logFileOverwritten | Should -Be $False
            }
        }

        Context 'Mutually exclusive switch parameter validation' {

            It 'throws exception if switches IncludeDateInFileName and ExcludeDateFromFileName are both set' {
                $newHostTextColours = GetNewConfigurationColour

                { Set-LogConfiguration -IncludeDateInFileName -ExcludeDateFromFileName } | 
                    Should -Throw 'Only one FileName switch parameter may be set'
            }

            It 'throws exception if switches OverwriteLogFile and AppendToLogFile are both set' {
                $newHostTextColours = GetNewConfigurationColour

                { Set-LogConfiguration -OverwriteLogFile -AppendToLogFile } | 
                    Should -Throw 'Only one LogFileWriteBehavior switch parameter may be set'
            }

            It 'throws exception if switches WriteToHost and WriteToStreams are both set' {
                $newHostTextColours = GetNewConfigurationColour

                { Set-LogConfiguration -WriteToHost -WriteToStreams } | 
                    Should -Throw 'Only one Destination switch parameter may be set'
            }
        }
    }

    Describe 'Reset-LogConfiguration' {             

        BeforeEach {
            SetNewConfiguration

            $newConfiguration = GetNewConfiguration
            AssertLogConfigurationMatchesReference `
                -ReferenceHashTable $newConfiguration

            $newMessageFormatInfo = GetNewMessageFormatInfo
            AssertMessageFormatInfoMatchesReference -ReferenceHashTable $newMessageFormatInfo

            $newLogFilePath = GetNewLogFilePath
            $script:_logFilePath | Should -Be $newLogFilePath

            $script:_logFileOverwritten | Should -Be $True
        }

        It 'resets configuration to default hashtable' {
                
            Reset-LogConfiguration 

            AssertLogConfigurationMatchesReference `
                -ReferenceHashTable $script:_defaultLogConfiguration
        }

        It 'resets MessageFormatInfo to default values' {                

            Reset-LogConfiguration 

            $defaultMessageFormatInfo = GetDefaultMessageFormatInfo
            AssertMessageFormatInfoMatchesReference -ReferenceHashTable $defaultMessageFormatInfo
        }

        It 'resets LogFilePath to default value' {

            Reset-LogConfiguration

            $defaultLogFilePath = GetDefaultLogFilePath
            $script:_logFilePath | Should -Be $defaultLogFilePath
        }

        It 'clears LogFileOverwritten if configuration LogFileName changed' {
            $script:_logConfiguration.LogFileName = GetNewLogFilePath
            $script:_logFileOverwritten = $True

            Reset-LogConfiguration

            $script:_logFileOverwritten | Should -Be $False
        }

        It 'does not clear LogFileOverwritten if configuration LogFileName unchanged' {
            $script:_logConfiguration.LogFileName = $script:_defaultLogConfiguration.LogFileName
            $script:_logFilePath = GetDefaultLogFilePath
            $script:_logFileOverwritten = $True

            Reset-LogConfiguration

            $script:_logFileOverwritten | Should -Be $True
        }
    }
}
