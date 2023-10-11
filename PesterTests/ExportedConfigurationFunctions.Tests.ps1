<#
.SYNOPSIS
Tests of the exported configuration functions in the Pslogg module.

.DESCRIPTION
Pester tests of the configuration-related functions exported from the Pslogg module.
#>

BeforeDiscovery {
    # NOTE: The module under test has to be imported in a BeforeDiscovery block, not a 
    # BeforeAll block.  If placed in a BeforeAll block the tests will fail with the following 
    # message:
    #   Discovery in ... failed with:
    #   System.Management.Automation.RuntimeException: No modules named 'Pslogg' are currently 
    #   loaded.

    # PowerShell allows multiple modules of the same name to be imported from different locations.  
    # This would confuse Pester.  So, to be sure there are not multiple Pslogg modules imported, 
    # remove all Pslogg modules and re-import only one.
    Get-Module Pslogg | Remove-Module -Force

    # Use $PSScriptRoot so this script will always import the Pslogg module in the Modules folder 
    # adjacent to the folder containing this script, regardless of the location that Pester is 
    # invoked from:
    #                                     {parent folder}
    #                                             |
    #                   -----------------------------------------------------
    #                   |                                                   |
    #     {folder containing this script}                                Modules folder
    #                   |                                                   |
    #                   |                                                Pslogg module folder
    #                   |                                                   |
    #               This script -------------> imports --------------->  Pslogg.psd1 module script
    Import-Module (Join-Path $PSScriptRoot ..\Modules\Pslogg\Pslogg.psd1 -Resolve) -Force
}

InModuleScope Pslogg {
    BeforeAll {
        
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
            [int]$IndentLevel = 0,
            [switch]$ValueTypesOnly
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
                if ($typeName1 -eq 'System.Collections.Hashtable' -and -not $ValueTypesOnly)
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
                            }
            return $hostTextColor
        }

        function GetDefaultCategoryInfo()
        {
            $categoryInfo = Private_DeepCopyHashTable $script:_defaultLogConfiguration.CategoryInfo
            return $categoryInfo
        }

        function GetNewCategoryInfo()
        {
            $categoryInfo = GetDefaultCategoryInfo
            $categoryInfo.Remove('Progress')
            $categoryInfo.FileCopy = @{Color = 'DarkCyan'}
            return $categoryInfo                     
        }

        function GetNewLogFileInfo()
        {
            $logFileInfo = @{
                                Name = 'C:\Test\Test.txt'
                                IncludeDateInFileName = $False
                                Overwrite = $False
                            }

            return $logFileInfo
        }

        # Gets a configuration hashtable where every setting is different from the defaults.
        function GetNewConfiguration ()
        {
            $logConfiguration = Private_DeepCopyHashTable $script:_defaultLogConfiguration
            $logConfiguration.LogLevel = "Verbose"
            $logConfiguration.WriteToHost = $False
            $logConfiguration.MessageFormat = "{MessageLevel} | {Message}"
            
            $logFileInfo = GetNewLogFileInfo
            $logConfiguration.LogFile = $logFileInfo

            $hostTextColor = GetNewConfigurationColour
            $logConfiguration.HostTextColor = $hostTextColor

            $categoryInfo = GetNewCategoryInfo
            $logConfiguration.CategoryInfo = $categoryInfo

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
                                    RawFormat = '{MessageLevel} | {Message}'
                                    WorkingFormat = '${MessageLevel} | ${Message}'
                                    FieldsPresent = @('Message', 'MessageLevel')
                                }
            return $messageFormatInfo
        }

        # Gets the MessageFormatInfo hashtable that matches the default configuration.
        function GetDefaultMessageFormatInfo ()
        {
            $messageFormatInfo = @{
                                    RawFormat = '{Timestamp:yyyy-MM-dd HH:mm:ss.fff} | {CallerName} | {Category} | {MessageLevel} | {Message}'
                                    WorkingFormat = '$($Timestamp.ToString(''yyyy-MM-dd HH:mm:ss.fff'')) | ${CallerName} | ${Category} | ${MessageLevel} | ${Message}'
                                    FieldsPresent = @('Message', 'Timestamp', 'CallerName', 'MessageLevel', 'Category')
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
            $fileName = "Results${dateString}.log"
            # Can't use $PSScriptRoot because it will return the folder containing this file, while 
            # the Pslogg module will see the ultimate caller as the Pester module running this 
            # test script.
            $callingDirectoryPath = GetCallingDirectoryPath
            $path = Join-Path $callingDirectoryPath $fileName
            return $path
        }

        # Sets the Pslogg configuration to its defaults.
        function SetConfigurationToDefault ()
        {
            $script:_logConfiguration = Private_DeepCopyHashTable $script:_defaultLogConfiguration
            $script:_messageFormatInfo = GetDefaultMessageFormatInfo
            $script:_logConfiguration.LogFile.FullPathReadOnly = GetDefaultLogFilePath
            $script:_logFileOverwritten = $False
        }

        # Sets the Pslogg configuration so that every settings differs from the defaults.
        function SetNewConfiguration ()
        {
            $script:_logConfiguration = GetNewConfiguration
            $script:_messageFormatInfo = GetNewMessageFormatInfo
            $script:_logConfiguration.LogFile.FullPathReadOnly = GetNewLogFilePath
            $script:_logFileOverwritten = $True
        }

        # Verifies the specified hashtable is identical to the reference hashtable.
        function AssertHashTablesMatch 
            (
                [hashtable]$ReferenceHashTable, 
                [hashtable]$HashTableToTest, 
                [switch]$ShouldBeNotEqual,
                [switch]$ValueTypesOnly
            ) 
        {
            $differences = GetHashTableDifferences `
                -HashTable1 $ReferenceHashTable `
                -HashTable2 $HashTableToTest `
                -ValueTypesOnly:$ValueTypesOnly

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
                -HashTableToTest $script:_logConfiguration -ShouldBeNotEqual:$ShouldBeNotEqual `
                -ValueTypesOnly
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
            $script:_logConfiguration.Keys.Count | Should -Be 6

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
                $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Be $newLogFilePath
            }

            It 'includes date stamp in LogFilePath if new configuration LogFile.IncludeDateInFileName set' {

                $newConfiguration = GetNewConfiguration
                $newConfiguration.LogFile.IncludeDateInFileName = $True
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $newLogFilePath = GetNewLogFilePath -IncludeDateInFileName
                $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Be $newLogFilePath
            }

            It 'clears LogFileOverwritten if configuration LogFile.Name changed' {
                $script:_logFileOverwritten = $True

                $newConfiguration = GetNewConfiguration
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $script:_logFileOverwritten | Should -Be $False
            }

            It 'does not clear LogFileOverwritten if configuration LogFile.Name unchanged' {
                $script:_logFileOverwritten = $True
                $script:_logConfiguration.LogFile.Name = GetNewLogFilePath
                $script:_logConfiguration.LogFile.FullPathReadOnly = GetNewLogFilePath

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
                $script:_logConfiguration.LogFile.Name = 'Text.log'
            }

            It 'sets configuration LogFile.Name as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -LogFileName New.test `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogFile.Name | Should -Be New.test
            }

            It 'sets configuration LogFile.Name as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -LogFileName New.test -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogFile.Name | Should -Be New.test
            }

            It 'sets configuration LogFile.Name as a member of the default parameter set' {
                                
                Set-LogConfiguration -LogFileName New.test 
                    
                $script:_logConfiguration.LogFile.Name | Should -Be New.test
            }

            It 'throws ArgumentException if attempt to set LogFileName to an invalid path' {
                { Set-LogConfiguration -LogFileName 'CC:\Test\Test.log' } |
                    Should -Throw -ExceptionType ([ArgumentException]) `
                                    -ExpectedMessage '*Invalid file path*'
            }

            It 'updates LogFilePath' {

                $newConfiguration = GetNewConfiguration
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $newLogFilePath = GetNewLogFilePath
                $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Be $newLogFilePath
            }

            It 'includes date stamp in LogFilePath if LogFile.IncludeDateInFileName set' {

                $newConfiguration = GetNewConfiguration
                $newConfiguration.LogFile.IncludeDateInFileName = $True
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $newLogFilePath = GetNewLogFilePath -IncludeDateInFileName
                $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Be $newLogFilePath
            }

            It 'does not include date stamp in LogFilePath if LogFile.IncludeDateInFileName cleared' {

                $newConfiguration = GetNewConfiguration
                $newConfiguration.LogFile.IncludeDateInFileName = $False
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $newLogFilePath = GetNewLogFilePath
                $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Be $newLogFilePath
            }

            It 'clears LogFileOverwritten if configuration LogFile.Name changed' {
                $script:_logFileOverwritten = $True

                $newConfiguration = GetNewConfiguration
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $script:_logFileOverwritten | Should -Be $False
            }

            It 'does not clear LogFileOverwritten if configuration LogFile.Name unchanged' {
                $script:_logFileOverwritten = $True
                $script:_logConfiguration.LogFile.Name = GetNewLogFilePath
                $script:_logConfiguration.LogFile.FullPathReadOnly = GetNewLogFilePath

                $newConfiguration = GetNewConfiguration
                Set-LogConfiguration -LogConfiguration $newConfiguration

                $script:_logFileOverwritten | Should -Be $True
            }
        }

        Context 'Parameter IncludeDateInFileName' {
            BeforeEach {   
                $script:_logConfiguration.LogFile.IncludeDateInFileName = $False
            }

            It 'sets configuration LogFile.IncludeDateInFileName as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -IncludeDateInFileName `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogFile.IncludeDateInFileName | Should -Be $True
            }

            It 'sets configuration LogFile.IncludeDateInFileName as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -IncludeDateInFileName -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogFile.IncludeDateInFileName | Should -Be $True
            }

            It 'sets configuration LogFile.IncludeDateInFileName as a member of the default parameter set' {
                                
                Set-LogConfiguration -IncludeDateInFileName
                    
                $script:_logConfiguration.LogFile.IncludeDateInFileName | Should -Be $True
            }

            It 'clears LogFileOverwritten if calculated LogFilePath changed' {
                $script:_logFileOverwritten = $True
                $script:_logConfiguration.LogFile.Name = GetNewLogFilePath
                $script:_logConfiguration.LogFile.FullPathReadOnly = GetNewLogFilePath
                $script:_logConfiguration.LogFile.IncludeDateInFileName = $False

                Set-LogConfiguration -IncludeDateInFileName

                $script:_logFileOverwritten | Should -Be $False
            }

            It 'does not clear LogFileOverwritten if calculated LogFilePath unchanged' {
                $script:_logFileOverwritten = $True
                $script:_logConfiguration.LogFile.Name = GetNewLogFilePath
                $script:_logConfiguration.LogFile.FullPathReadOnly = GetNewLogFilePath -IncludeDateInFileName
                $script:_logConfiguration.LogFile.IncludeDateInFileName = $True

                Set-LogConfiguration -IncludeDateInFileName

                $script:_logFileOverwritten | Should -Be $True
            }
        }

        Context 'Parameter ExcludeDateFromFileName' {
            BeforeEach {              
                $script:_logConfiguration.LogFile.IncludeDateInFileName = $True
            }

            It 'sets configuration ExcludeDateFromFileName as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -ExcludeDateFromFileName `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogFile.IncludeDateInFileName | Should -Be $False
            }

            It 'sets configuration ExcludeDateFromFileName as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -ExcludeDateFromFileName -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogFile.IncludeDateInFileName | Should -Be $False
            }

            It 'sets configuration ExcludeDateFromFileName as a member of the default parameter set' {
                                
                Set-LogConfiguration -ExcludeDateFromFileName
                    
                $script:_logConfiguration.LogFile.IncludeDateInFileName | Should -Be $False
            }

            It 'clears LogFileOverwritten if calculated LogFilePath changed' {
                $script:_logFileOverwritten = $True
                $script:_logConfiguration.LogFile.Name = GetNewLogFilePath
                $script:_logConfiguration.LogFile.FullPathReadOnly = GetNewLogFilePath -IncludeDateInFileName
                $script:_logConfiguration.LogFile.IncludeDateInFileName = $True

                Set-LogConfiguration -ExcludeDateFromFileName

                $script:_logFileOverwritten | Should -Be $False
            }

            It 'does not clear LogFileOverwritten if calculated LogFilePath unchanged' {
                $script:_logFileOverwritten = $True
                $script:_logConfiguration.LogFile.Name = GetNewLogFilePath
                $script:_logConfiguration.LogFile.FullPathReadOnly = GetNewLogFilePath
                $script:_logConfiguration.LogFile.IncludeDateInFileName = $False

                Set-LogConfiguration -ExcludeDateFromFileName

                $script:_logFileOverwritten | Should -Be $True
            }
        }

        Context 'Parameter LogFileOverwrite' {
            BeforeEach {    
                $script:_logConfiguration.LogFile.Overwrite = $False
            }

            It 'sets configuration LogFile.Overwrite as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -OverwriteLogFile `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogFile.Overwrite | Should -Be $True
            }

            It 'sets configuration LogFile.Overwrite as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -OverwriteLogFile -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogFile.Overwrite | Should -Be $True
            }

            It 'sets configuration LogFile.Overwrite as a member of the default parameter set' {
                                
                Set-LogConfiguration -OverwriteLogFile
                    
                $script:_logConfiguration.LogFile.Overwrite | Should -Be $True
            }
        }

        Context 'Parameter AppendToLogFile' {
            BeforeEach {                
                $script:_logConfiguration.LogFile.Overwrite = $True
            }

            It 'sets configuration AppendToLogFile as a member of parameter set "IndividualSettings_AllColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                $newHostTextColours = GetNewConfigurationColour

                Set-LogConfiguration -AppendToLogFile `
                    -HostTextColorConfiguration $newHostTextColours
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogFile.Overwrite | Should -Be $False
            }

            It 'sets configuration AppendToLogFile as a member of parameter set "IndividualSettings_IndividualColors"' {
                $script:_logConfiguration.HostTextColor.Error = 'Red'
                                
                Set-LogConfiguration -AppendToLogFile -ErrorTextColor DarkYellow
                    
                $script:_logConfiguration.HostTextColor.Error | Should -Be DarkYellow
                $script:_logConfiguration.LogFile.Overwrite | Should -Be $False
            }

            It 'sets configuration AppendToLogFile as a member of the default parameter set' {
                                
                Set-LogConfiguration -AppendToLogFile
                    
                $script:_logConfiguration.LogFile.Overwrite | Should -Be $False
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

        Context 'Parameter CategoryInfoItem' {
            BeforeEach {                
                $script:_logConfiguration.CategoryInfo = @{
                                                            Progress = @{ IsDefault = $True }
                                                            Success = @{ Color = 'Green' }
                                                            Failure = @{ Color = 'Red' }
                                                            PartialFailure = @{ Color = 'Yellow' }
                                                        }
            }

            It 'updates the value of an existing CategoryInfo item when a tuple is supplied' {
                
                Set-LogConfiguration -CategoryInfoItem Success, @{ Color = 'DarkGreen' }

                $script:_logConfiguration.CategoryInfo.Success.Color | Should -Be 'DarkGreen'
            }

            It 'adds a new CategoryInfo item when a tuple is supplied which does not match any existing item' {
                
                Set-LogConfiguration -CategoryInfoItem Blue, @{ Color = 'Blue' }

                $script:_logConfiguration.CategoryInfo.ContainsKey('Blue') | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Blue.Color | Should -Be 'Blue'
            }

            It 'removes existing CategoryInfo item IsDefault property when a tuple is supplied which has an IsDefault property' {
                
                Set-LogConfiguration -CategoryInfoItem Blue, @{ IsDefault = $True }

                $script:_logConfiguration.CategoryInfo.ContainsKey('Blue') | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Blue.IsDefault | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Progress.ContainsKey('IsDefault') | Should -Be $False
            }

            It 'creates the CategoryInfo hashtable when CategoryInfo does not exist and a tuple is supplied' {
                $script:_logConfiguration.CategoryInfo = $Null

                Set-LogConfiguration -CategoryInfoItem Blue, @{ Color = 'Blue' }

                ($script:_logConfiguration.CategoryInfo -is [hashtable]) | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Keys.Count | Should -Be 1
                $script:_logConfiguration.CategoryInfo.ContainsKey('Blue') | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Blue.Color | Should -Be 'Blue'
            }

            It 'updates the value of an existing CategoryInfo item when a single-item hashtable is supplied' {
                
                Set-LogConfiguration -CategoryInfoItem @{ Success = @{ Color = 'DarkGreen' } }

                $script:_logConfiguration.CategoryInfo.Success.Color | Should -Be 'DarkGreen'
            }

            It 'adds a new CategoryInfo item when a single-item hashtable is supplied which does not match any existing item' {
                
                Set-LogConfiguration -CategoryInfoItem @{ Blue = @{ Color = 'Blue' } }

                $script:_logConfiguration.CategoryInfo.ContainsKey('Blue') | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Blue.Color | Should -Be 'Blue'
            }

            It 'removes existing CategoryInfo item IsDefault property when a single-item hashtable is supplied which has an IsDefault property' {
                
                Set-LogConfiguration -CategoryInfoItem @{ Blue = @{ IsDefault = $True } }

                $script:_logConfiguration.CategoryInfo.ContainsKey('Blue') | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Blue.IsDefault | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Progress.ContainsKey('IsDefault') | Should -Be $False
            }

            It 'creates the CategoryInfo hashtable when CategoryInfo does not exist and a single-item hashtable is supplied' {
                $script:_logConfiguration.CategoryInfo = $Null

                Set-LogConfiguration -CategoryInfoItem @{ Blue = @{ Color = 'Blue' } }

                ($script:_logConfiguration.CategoryInfo -is [hashtable]) | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Keys.Count | Should -Be 1
                $script:_logConfiguration.CategoryInfo.ContainsKey('Blue') | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Blue.Color | Should -Be 'Blue'
            }

            It 'updates the value of multiple existing CategoryInfo items when a multi-item hashtable is supplied' {
                
                Set-LogConfiguration -CategoryInfoItem @{ 
                                                        Success = @{ Color = 'DarkGreen' }                                                         
                                                        Failure = @{ Color = 'DarkRed' }
                                                        }

                $script:_logConfiguration.CategoryInfo.Success.Color | Should -Be 'DarkGreen'
                $script:_logConfiguration.CategoryInfo.Failure.Color | Should -Be 'DarkRed'
            }

            It 'adds multiple new CategoryInfo items when a multi-item hashtable is supplied which does not match any existing item' {
                
                Set-LogConfiguration -CategoryInfoItem @{ 
                                                        Blue = @{ Color = 'Blue' } 
                                                        Yellow = @{ Color = 'Yellow' } 
                                                        }

                $script:_logConfiguration.CategoryInfo.ContainsKey('Blue') | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Blue.Color | Should -Be 'Blue'
                $script:_logConfiguration.CategoryInfo.ContainsKey('Yellow') | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Yellow.Color | Should -Be 'Yellow'
            }

            It 'creates the CategoryInfo hashtable when CategoryInfo does not exist and a multi-item hashtable is supplied' {
                $script:_logConfiguration.CategoryInfo = $Null

                Set-LogConfiguration -CategoryInfoItem @{ 
                                                        Blue = @{ Color = 'Blue' } 
                                                        Yellow = @{ Color = 'Yellow' } 
                                                        }

                ($script:_logConfiguration.CategoryInfo -is [hashtable]) | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Keys.Count | Should -Be 2
                $script:_logConfiguration.CategoryInfo.ContainsKey('Blue') | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Blue.Color | Should -Be 'Blue'
                $script:_logConfiguration.CategoryInfo.ContainsKey('Yellow') | Should -Be $True
                $script:_logConfiguration.CategoryInfo.Yellow.Color | Should -Be 'Yellow'
            }
        }

        Context 'Parameter CategoryInfoKeyToRemove' {
            BeforeEach {                
                $script:_logConfiguration.CategoryInfo = @{
                                                            Progress = @{ IsDefault = $True }
                                                            Success = @{ Color = 'Green' }
                                                            Failure = @{ Color = 'Red' }
                                                            PartialFailure = @{ Color = 'Yellow' }
                                                        }
            }

            It 'has no effect if the supplied keys do not exist in the CategoryInfo hashtable' {
                
                Set-LogConfiguration -CategoryInfoKeyToRemove Blue,Yellow

                $script:_logConfiguration.CategoryInfo.Keys.Count | Should -Be 4
                $referenceHashtable = @{
                                            Progress = @{ IsDefault = $True }
                                            Success = @{ Color = 'Green' }
                                            Failure = @{ Color = 'Red' }
                                            PartialFailure = @{ Color = 'Yellow' }
                                        }
                AssertHashTablesMatch -ReferenceHashTable $referenceHashtable `
                    -HashtableToTest $script:_logConfiguration.CategoryInfo
            }

            It 'removes the key from the CategoryInfo hashtable if a single key is supplied' {
                
                Set-LogConfiguration -CategoryInfoKeyToRemove PartialFailure

                $script:_logConfiguration.CategoryInfo.Keys.Count | Should -Be 3
                $script:_logConfiguration.CategoryInfo.ContainsKey('PartialFailure') | 
                    Should -Be $False
            }

            It 'removes multiple keys from the CategoryInfo hashtable if multiple keys are supplied' {
                
                Set-LogConfiguration -CategoryInfoKeyToRemove Progress,PartialFailure

                $script:_logConfiguration.CategoryInfo.Keys.Count | Should -Be 2
                $script:_logConfiguration.CategoryInfo.ContainsKey('Progress') | 
                    Should -Be $False
                $script:_logConfiguration.CategoryInfo.ContainsKey('PartialFailure') | 
                    Should -Be $False
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
        }

        Context 'Multiple parameters set simultaneously' {
            It 'sets multiple configuration properties if multiple parameters are set' {
                $script:_logFileOverwritten = $True

                Set-LogConfiguration -LogLevel Verbose -LogFileName 'C:\Test\Test.txt' `
                    -ExcludeDateFromFileName -AppendToLogFile -WriteToStreams `
                    -MessageFormat '{MessageLevel} | {Message}' `
                    -CategoryInfoItem 'FileCopy', @{Color = 'DarkCyan'} `
                    -CategoryInfoKeyToRemove 'Progress' -ErrorTextColor DarkYellow `
                    -WarningTextColor DarkYellow -InformationTextColor DarkYellow `
                    -DebugTextColor DarkYellow -VerboseTextColor DarkYellow

                $referenceLogConfiguration = GetNewConfiguration
                $referenceMessageFormatInfo = GetNewMessageFormatInfo
                $newLogFilePath = GetNewLogFilePath
                $newLogFileOverwritten = $False

                AssertLogConfigurationMatchesReference `
                    -ReferenceHashTable $referenceLogConfiguration
                
                AssertMessageFormatInfoMatchesReference `
                    -ReferenceHashTable $referenceMessageFormatInfo

                $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Be $newLogFilePath
                $script:_logFileOverwritten | Should -Be $False
            }
        }

        Context 'Mutually exclusive switch parameter validation' {

            It 'throws exception if switches IncludeDateInFileName and ExcludeDateFromFileName are both set' {
                $newHostTextColours = GetNewConfigurationColour

                { Set-LogConfiguration -IncludeDateInFileName -ExcludeDateFromFileName } | 
                    Should -Throw 'Only one FileName switch parameter may be set*'
            }

            It 'throws exception if switches OverwriteLogFile and AppendToLogFile are both set' {
                $newHostTextColours = GetNewConfigurationColour

                { Set-LogConfiguration -OverwriteLogFile -AppendToLogFile } | 
                    Should -Throw 'Only one LogFileWriteBehavior switch parameter may be set*'
            }

            It 'throws exception if switches WriteToHost and WriteToStreams are both set' {
                $newHostTextColours = GetNewConfigurationColour

                { Set-LogConfiguration -WriteToHost -WriteToStreams } | 
                    Should -Throw 'Only one Destination switch parameter may be set*'
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
            $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Be $newLogFilePath

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

        It 'resets LogFile.FullPathReadOnly to default value' {

            Reset-LogConfiguration

            $defaultLogFilePath = GetDefaultLogFilePath
            $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Be $defaultLogFilePath
        }

        It 'clears LogFileOverwritten if configuration LogFile.Name changed' {
            $script:_logConfiguration.LogFile.Name = GetNewLogFilePath
            $script:_logFileOverwritten = $True

            Reset-LogConfiguration

            $script:_logFileOverwritten | Should -Be $False
        }

        It 'does not clear LogFileOverwritten if configuration LogFile.Name unchanged' {
            $script:_logConfiguration.LogFile.Name = $script:_defaultLogConfiguration.LogFile.Name
            $script:_logConfiguration.LogFile.FullPathReadOnly = GetDefaultLogFilePath
            $script:_logFileOverwritten = $True

            Reset-LogConfiguration

            $script:_logFileOverwritten | Should -Be $True
        }
    }
}