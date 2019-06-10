<#
.SYNOPSIS
Tests of the shared private functions in the Pslogg module.

.DESCRIPTION
Pester tests of the private functions in the Pslogg module that are called by both the 
configuration and the logging functions.
#>

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
#                   \                                                   |
#                    ------------------> imports                     Pslogg module folder
#                                                \                      |
#                                                 -----------------> Pslogg.psd1 module script
Import-Module (Join-Path $PSScriptRoot ..\Modules\Pslogg\Pslogg.psd1 -Resolve) -Force

InModuleScope Pslogg {

    # Need to dot source the helper file within the InModuleScope block to be able to use its 
    # functions within a test.
    . (Join-Path $PSScriptRoot .\AssertExceptionThrown.ps1 -Resolve)

    Describe "ValidateSwitchParameterGroup" {
    
        It 'throws ParameterBindingValidationException when no switches supplied' {
            [switch[]]$switchList = @()
            try
            {
                Private_ValidateSwitchParameterGroup -SwitchList $switchList -ErrorMessage "Throwing validation exception"
            }
            catch
            {
                $_.Exception.GetType().Name | Should -Be 'ParameterBindingValidationException'
            }            
        }
    
        It 'throws ParameterBindingValidationException when no error message supplied' {
            [switch]$firstSwitch = $True
            [switch[]]$switchList = @($firstSwitch)
            try
            {
                Private_ValidateSwitchParameterGroup -SwitchList $switchList -ErrorMessage ''
            }
            catch
            {
                $_.Exception.GetType().Name | Should -Be 'ParameterBindingValidationException'
            }            
        }

        It 'does not throw when one switch defined and not set' {
            [switch]$firstSwitch = $True
            [switch[]]$switchList = @($firstSwitch)
            { Private_ValidateSwitchParameterGroup -SwitchList $switchList -ErrorMessage "Should not throw" } | 
                Should -Not -Throw
        }

        It 'does not throw when one switch defined and is set' {
            [switch]$firstSwitch = $True
            [switch[]]$switchList = @($firstSwitch)
            { Private_ValidateSwitchParameterGroup -SwitchList $switchList -ErrorMessage "Should not throw" } | 
                Should -Not -Throw
        }

        It 'does not throw on zero switch set out or two' {
            $errorMessage = "Should not throw"
            [switch]$firstSwitch = $False
            [switch]$secondSwitch = $False
            [switch[]]$switchList = @($firstSwitch, $secondSwitch)
            { Private_ValidateSwitchParameterGroup -SwitchList $switchList -ErrorMessage $errorMessage } | 
                Should -Not -Throw
        }

        It 'does not throw on one switch set out of two' {
            $errorMessage = "Should not throw"
            [switch]$firstSwitch = $True
            [switch]$secondSwitch = $False
            [switch[]]$switchList = @($firstSwitch, $secondSwitch)
            { Private_ValidateSwitchParameterGroup -SwitchList $switchList -ErrorMessage $errorMessage } | 
                Should -Not -Throw
        }

        It 'throws ArgumentException on two switches set out of two' {
            $errorMessage = "Should throw ArgumentException"
            [switch]$firstSwitch = $True
            [switch]$secondSwitch = $True
            [switch[]]$switchList = @($firstSwitch, $secondSwitch)
            try
            {
                Private_ValidateSwitchParameterGroup -SwitchList $switchList -ErrorMessage $errorMessage
            }
            catch
            {
                # One way of checking exception type.
                $_.Exception.GetType().Name | Should Be 'ArgumentException'
            }
        }

        It 'throws ArgumentException on two switches set out of three' {
            $errorMessage = "Should throw ArgumentException"
            $exception = $Null
            [switch]$firstSwitch = $True
            [switch]$secondSwitch = $True
            [switch]$thirdSwitch = $False
            [switch[]]$switchList = @($firstSwitch, $secondSwitch, $thirdSwitch)
            try
            {
                Private_ValidateSwitchParameterGroup -SwitchList $switchList -ErrorMessage $errorMessage
            }
            catch
            {
                $exception = $_.Exception
            }
            # A second way of checking exception type.
            $exception | Should -BeOfType [ArgumentException]
        }
    }

    Describe "ValidateHostColor" {

        It 'does not throw when color name is valid' {
            [string]$colourName = 'Yellow'
            { Private_ValidateHostColor -ColorToTest $colourName } | Should -Not -Throw
        }

        It 'returns $True when color name is valid' {
            [string]$colourName = 'Yellow'
            $result = Private_ValidateHostColor -ColorToTest $colourName
            $result | Should -Be $True
        }

        It 'throws ArgumentException when color name is invalid' {
            [string]$colourName = 'Turquoise'
            
            { Private_ValidateHostColor -ColorToTest $colourName} | 
                Assert-ExceptionThrown -WithTypeName ArgumentException
        }

        It 'exception error message includes "INVALID TEXT COLOR ERROR: $colourName" when color name is invalid' {
            [string]$colourName = 'Turquoise'
            
            { Private_ValidateHostColor -ColorToTest $colourName} | 
                Assert-ExceptionThrown -WithMessage "INVALID TEXT COLOR ERROR: '$colourName'"
        }
    }

    Describe "ValidateLogLevel" {

        It 'does not throw when log level is valid' {
            [string]$logLevel = 'ERROR'

            { Private_ValidateLogLevel -LevelToTest $logLevel } | Should -Not -Throw
        }

        It 'returns $True when log level is valid' {
            [string]$logLevel = 'ERROR'

            $result = Private_ValidateLogLevel -LevelToTest $logLevel

            $result | Should -Be $True
        }

        It 'throws ArgumentException when log level is invalid' {
            [string]$logLevel = 'INVALID'

            { Private_ValidateLogLevel -LevelToTest $logLevel } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException
        }

        It 'exception error message includes "INVALID LOG LEVEL ERROR: $logLevel" when log level is invalid' {
            [string]$logLevel = 'INVALID'
            
             { Private_ValidateLogLevel -LevelToTest $logLevel } | 
                Assert-ExceptionThrown -WithMessage "INVALID LOG LEVEL ERROR: '$logLevel'"
        }

        It 'throws ArgumentException when log level is OFF and -ExcludeOffLevel switch is set' {
            [string]$logLevel = 'OFF'

            { Private_ValidateLogLevel -LevelToTest $logLevel -ExcludeOffLevel } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException
        }

        It 'does not throw when log level is OFF and -ExcludeOffLevel switch is not set' {
            [string]$logLevel = 'OFF'

            { Private_ValidateLogLevel -LevelToTest $logLevel } | Should -Not -Throw
        }

        It 'returns $True when log level is OFF and -ExcludeOffLevel switch is not set' {
            [string]$logLevel = 'OFF'

            $result = Private_ValidateLogLevel -LevelToTest $logLevel

            $result | Should -Be $True
        }
    }

    Describe "GetTimestampFormat" {

        It 'returns $Null when not a Timestamp placeholder' {
            [string]$textToSearch = 'xxx {other} xxx'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be $Null
        }

        It 'returns $Null when Timestamp placeholder without format string' {
            [string]$textToSearch = 'xxx {Timestamp} xxx'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be $Null
        }

        It 'returns $Null when Timestamp placeholder is missing colon' {
            [string]$textToSearch = 'xxx {Timestamp yyyy-MM-dd hh:mm:ss} xxx'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be $Null
        }

        It 'returns simple format string' {
            [string]$textToSearch = 'xxx {Timestamp:d} xxx'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be d
        }

        It 'performs case-insensitive search for Timestamp placeholder' {
            [string]$textToSearch = 'xxx {TIMESTAMP:d} xxx'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be d
        }

        It 'returns simple format string from Timestamp placeholder with leading and trailing spaces' {
            [string]$textToSearch = 'xxx { Timestamp :d} xxx'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be d
        }

        It 'returns simple format string from Timestamp placeholder with leading and trailing tabs' {
            [string]$textToSearch = 'xxx {	Timestamp	:d} xxx'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be d
        }

        It 'strips leading and trailing spaces from simple format string' {
            [string]$textToSearch = 'xxx {Timestamp:  d  } xxx'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be d
        }

        It 'returns simple format string when Timestamp placeholder is at start of text' {
            [string]$textToSearch = '{Timestamp:d} xxx'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be d
        }

        It 'returns simple format string when Timestamp placeholder is at end of text' {
            [string]$textToSearch = 'xxx {Timestamp:d}'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be d
        }

        It 'returns simple format string when Timestamp placeholder is the only text' {
            [string]$textToSearch = '{Timestamp:d}'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be d
        }

        It 'returns format string containing colon' {
            [string]$textToSearch = 'xxx {Timestamp:hh:mm:ss.fff} xxx'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be hh:mm:ss.fff
        }

        It 'returns format string containing space' {
            [string]$textToSearch = 'xxx {Timestamp : yyyy-MM-dd hh:mm:ss.fff } xxx'
            Private_GetTimestampFormat -MessageFormat $textToSearch | Should -Be 'yyyy-MM-dd hh:mm:ss.fff'
        }
    }

    Describe "GetMessageFormatInfo" {    

        It 'returns hashtable' {
            $messageFormat = 'xxx {Message} xxx'
            Private_GetMessageFormatInfo -MessageFormat $messageFormat | Should -BeOfType [Hashtable]
        }   

        It 'hashtable has key "RawFormat"' {
            $messageFormat = 'xxx {Message} xxx'
            $hashTable = Private_GetMessageFormatInfo -MessageFormat $messageFormat
            $hashTable.ContainsKey('RawFormat') | Should -Be $True
        }  

        It 'hashtable has key "WorkingFormat"' {
            $messageFormat = 'xxx {Message} xxx'
            $hashTable = Private_GetMessageFormatInfo -MessageFormat $messageFormat
            $hashTable.ContainsKey('WorkingFormat') | Should -Be $True
        }  

        It 'hashtable has key "FieldsPresent"' {
            $messageFormat = 'xxx {Message} xxx'
            $hashTable = Private_GetMessageFormatInfo -MessageFormat $messageFormat
            $hashTable.ContainsKey('FieldsPresent') | Should -Be $True
        } 

        It 'FieldsPresent is an array' {
            $messageFormat = 'xxx {Message} xxx'
            $hashTable = Private_GetMessageFormatInfo -MessageFormat $messageFormat
            $hashTable.FieldsPresent.GetType().FullName | Should -Be 'System.Object[]'
        }

        It 'RawFormat is a copy of original message format text' {
            $messageFormat = 'xxx {Message} xxx'
            $hashTable = Private_GetMessageFormatInfo -MessageFormat $messageFormat
            $hashTable.RawFormat | Should -BeExactly $messageFormat
        }

        It 'WorkingFormat is a copy of original message format text when it contains no field placeholders' {
            $messageFormat = 'xxx xxx'
            $hashTable = Private_GetMessageFormatInfo -MessageFormat $messageFormat
            $hashTable.WorkingFormat | Should -BeExactly $messageFormat
        }

        function TestWorkingFormatFieldSurroundedBySpaces ([string]$FieldName)
        {
            $formatTemplate = 'xxx [FIELD PLACEHOLDER] xxx' 
            TestWorkingFormatField -FieldName $FieldName -FormatTemplate $formatTemplate
        }

        function TestWorkingFormatField ([string]$FieldName, [string]$FormatTemplate)
        {
            $messageFormatTemplate = $FormatTemplate -replace "\[FIELD PLACEHOLDER\]", '{$FieldName}'
            $workingFormatTemplate = $FormatTemplate -replace "\[FIELD PLACEHOLDER\]", '`${$${FieldName}}'
            $messageFormat = $ExecutionContext.InvokeCommand.ExpandString($messageFormatTemplate)
            $workingFormat = $ExecutionContext.InvokeCommand.ExpandString($workingFormatTemplate)
            TestWorkingFormat -InputMessageFormat $messageFormat `
                -ExpectedWorkingFormat $workingFormat
        }

        function TestWorkingFormat 
            (
                [string]$InputMessageFormat, 
                [string]$ExpectedWorkingFormat, 
                [switch]$DoRegexMatch
            )
        {
            $hashTable = Private_GetMessageFormatInfo -MessageFormat $InputMessageFormat
            if ($DoRegexMatch)
            {
                $hashTable.WorkingFormat | Should -Match $ExpectedWorkingFormat
            }
            else
            {
                $hashTable.WorkingFormat | Should -Be $ExpectedWorkingFormat
            }
        }

        It 'replaces {Message} placeholder with "${Message}" in WorkingFormat' {
            TestWorkingFormatFieldSurroundedBySpaces -FieldName Message
        }

        It 'replaces {CallerName} placeholder with "${CallerName}" in WorkingFormat' {
            TestWorkingFormatFieldSurroundedBySpaces -FieldName CallerName
        }

        It 'replaces {MessageLevel} placeholder with "${MessageLevel}" in WorkingFormat' {
            TestWorkingFormatFieldSurroundedBySpaces -FieldName MessageLevel
        }

        It 'replaces {Category} placeholder with "${Category}" in WorkingFormat' {
            TestWorkingFormatFieldSurroundedBySpaces -FieldName Category
        }

        It 'replaces {TimeStamp} placeholder with "`$(`$Timestamp.ToString(<timestamp format>))" in WorkingFormat' {
            # Double the single quotes to escape them.
            TestWorkingFormat -InputMessageFormat 'xxx {Timestamp} xxx' `
                -ExpectedWorkingFormat 'xxx \$\(\$Timestamp\.ToString\(''.*''\)\) xxx' `
                -DoRegexMatch
        }

        It 'uses default <timestamp format> in WorkingFormat when none specified' {
            $messageFormat = 'xxx {Timestamp} xxx'
            $workingFormat = 'xxx $($Timestamp.ToString(''yyyy-MM-dd hh:mm:ss.fff'')) xxx'
            TestWorkingFormat -InputMessageFormat $messageFormat `
                -ExpectedWorkingFormat $workingFormat
        }

        It 'includes specified <timestamp format> in WorkingFormat when supplied' {
            $messageFormat = 'xxx {Timestamp : d} xxx'
            $workingFormat = 'xxx $($Timestamp.ToString(''d'')) xxx'
            TestWorkingFormat -InputMessageFormat $messageFormat `
                -ExpectedWorkingFormat $workingFormat
        }

        It 'replaces all other placeholders that follow a {TimeStamp} placeholder correctly' {
            $messageFormat = '{Timestamp:yyyy-MM-dd hh:mm:ss.fff} | {CallerName} | {Category} | {Message}'
            $workingFormat = '$($Timestamp.ToString(''yyyy-MM-dd hh:mm:ss.fff'')) | ${CallerName} | ${Category} | ${Message}'
            TestWorkingFormat -InputMessageFormat $messageFormat `
                -ExpectedWorkingFormat $workingFormat
        }

        It 'generates correct WorkingFormat when placeholder embedded in text without surrounding spaces' {
            $formatTemplate = 'xxx[FIELD PLACEHOLDER]xxx'
            TestWorkingFormatField -FieldName Message -FormatTemplate $formatTemplate
        }

        It 'generates correct Timestamp field in WorkingFormat when placeholder embedded in text without surrounding spaces' {
            $messageFormat = 'xxx{Timestamp : d}xxx'
            $workingFormat = 'xxx$($Timestamp.ToString(''d''))xxx'
            TestWorkingFormat -InputMessageFormat $messageFormat `
                -ExpectedWorkingFormat $workingFormat
        }

        It 'leaves leading and trailing spaces in WorkingFormat' {
            $formatTemplate = '   xxx [FIELD PLACEHOLDER] xxx   '
            TestWorkingFormatField -FieldName Message -FormatTemplate $formatTemplate
        }

        It 'replaces multiple placeholders in WorkingFormat' {
            $messageFormat = 'xxx {MessageLevel} {CallerName} {Message} xxx'
            $workingFormat = 'xxx ${MessageLevel} ${CallerName} ${Message} xxx'
            TestWorkingFormat -InputMessageFormat $messageFormat `
                -ExpectedWorkingFormat $workingFormat
        }

        It 'FieldsPresent is empty when no placeholders are present in message format text' {
            $messageFormat = 'xxx xxx'
            $hashTable = Private_GetMessageFormatInfo -MessageFormat $messageFormat
            $hashTable.FieldsPresent.Count | Should -Be 0
        }

        function TestFieldsPresent([string]$FieldName)
        {
            It "adds '$FieldName' to FieldsPresent array when {$FieldName} placeholder present in message format text" {
                $messageFormat = "xxx {$FieldName} xxx"
                $hashTable = Private_GetMessageFormatInfo -MessageFormat $messageFormat
                $hashTable.FieldsPresent.Count | Should -Be 1
                $hashTable.FieldsPresent[0] | Should -Be $FieldName
            }
        }

        TestFieldsPresent "Message"
        TestFieldsPresent "CallerName"
        TestFieldsPresent "MessageLevel"
        TestFieldsPresent "Category"
        TestFieldsPresent "Timestamp"

        It 'adds multiple field names to FieldsPresent when multiple placeholders in message format text' {
            $messageFormat = 'xxx {MessageLevel} {CallerName} {Message} xxx'
            $hashTable = Private_GetMessageFormatInfo -MessageFormat $messageFormat
            $hashTable.FieldsPresent.Count | Should -Be 3
            $hashTable.FieldsPresent.Contains("MessageLevel")
            $hashTable.FieldsPresent.Contains("CallerName")
            $hashTable.FieldsPresent.Contains("Message")
        }
    }
}
