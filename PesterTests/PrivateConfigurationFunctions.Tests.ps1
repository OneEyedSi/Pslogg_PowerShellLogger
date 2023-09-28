<#
.SYNOPSIS
Tests of the private configuration functions in the Pslogg module.

.DESCRIPTION
Pester tests of the private functions in the Pslogg module that are called by the public 
configuration functions.
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
        # Need to dot source the helper file within the InModuleScope block to be able to use its 
        # functions within a test.
        . (Join-Path $PSScriptRoot .\AssertExceptionThrown.ps1 -Resolve)
    }

    Describe "GetAbsolutePath" {     

        It 'throws ParameterBindingValidationException if empty path supplied' {
            { Private_GetAbsolutePath -Path '' } | 
                Assert-ExceptionThrown -WithTypeName ParameterBindingValidationException
        }  

        It 'throws ArgumentException if invalid path supplied' {
            { Private_GetAbsolutePath -Path 'CC:\Test\Test.log' } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException
        }

        It 'returns rooted path unchanged' {
            $originalPath = 'C:\Test\test.log'
            Private_GetAbsolutePath -Path $originalPath | Should -Be $originalPath
        }   

        It 'relative path is rooted on directory of caller' {
            $originalPath = 'SubDirectory\test.log'
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
            $expectedPath = Join-Path $pathOfCallerDirectory $originalPath

            $absolutePath = Private_GetAbsolutePath -Path $originalPath 

            $absolutePath | Should -Be $expectedPath
        } 
    }

    Describe "SetLogFilePath" {    
        BeforeAll { 
            function GetFileNameFromTemplate ([string]$FileNameTemplate, [switch]$IncludeDateInFileName)
            {
                $dateText = ""
                if ($IncludeDateInFileName)
                {
                    $dateText = (Get-Date -Format "_yyyyMMdd")
                }
                return $ExecutionContext.InvokeCommand.ExpandString($FileNameTemplate)
            }

            function TestLogFileConfiguration(
                [string]$ExistingFileName,
                [string]$NewFileNameTemplate,
                [switch]$IncludeDateInFileName
                )
            {
                $fileNameFromConfiguration = GetFileNameFromTemplate -FileNameTemplate $NewFileNameTemplate

                $script:_logConfiguration.LogFile.FullPath = $ExistingFileName
                $script:_logConfiguration.LogFile.Name = $fileNameFromConfiguration
                $script:_logConfiguration.LogFile.IncludeDateInFileName = $IncludeDateInFileName
                $script:_logFileOverwritten = $True

                Private_SetLogFilePath -OldLogFilePath $ExistingFileName
            }
        }

        It 'sets $script:_logConfiguration.LogFile.FullPath to configuration LogFile.Name' {
            $existingFileName = 'C:\Original\old.log'
            $newFileNameTemplate = 'C:\New\New${dateText}.log'
            $expectedFileName = GetFileNameFromTemplate -FileNameTemplate $newFileNameTemplate

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $newFileNameTemplate

            $script:_logConfiguration.LogFile.FullPath | Should -Be $expectedFileName
        }     

        It 'adds date to $script:_logConfiguration.LogFile.FullPath if LogFile.IncludeDateInFileName configuration value is set' {
            $existingFileName = 'C:\Original\old.log'
            $newFileNameTemplate = 'C:\New\New${dateText}.log'
            $expectedFileName = GetFileNameFromTemplate -FileNameTemplate $newFileNameTemplate `
                -IncludeDateInFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $newFileNameTemplate -IncludeDateInFileName

            $script:_logConfiguration.LogFile.FullPath | Should -Be $expectedFileName
        }

        It 'leaves $script:_logConfiguration.LogFile.FullPath unchanged if identical to configuration LogFile.Name' {
            $existingFileName = 'C:\Original\old.log'
            $newFileNameTemplate = 'C:\Original\old.log'
            $expectedFileName = $existingFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $newFileNameTemplate

            $script:_logConfiguration.LogFile.FullPath | Should -Be $existingFileName
        }

        It 'leaves $script:_logConfiguration.LogFile.FullPath unchanged if existing file name includes today''s date' {
            $fileNameTemplate = 'C:\Test\Test${dateText}.log'
            $existingFileName = GetFileNameFromTemplate -FileNameTemplate $fileNameTemplate `
                -IncludeDateInFileName
            $expectedFileName = $existingFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $fileNameTemplate -IncludeDateInFileName

            $script:_logConfiguration.LogFile.FullPath | Should -Be $existingFileName
        } 

        It 'updates $script:_logConfiguration.LogFile.FullPath if existing file name includes old date' {
            $fileNameTemplate = 'C:\Test\Test${dateText}.log'
            $oldDate = Get-Date -Year 2018 -Month 1 -Day 15
            $existingFileName = $ExecutionContext.InvokeCommand.ExpandString($fileNameTemplate)
            $expectedFileName = GetFileNameFromTemplate -FileNameTemplate $fileNameTemplate `
                -IncludeDateInFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $fileNameTemplate -IncludeDateInFileName

            $script:_logConfiguration.LogFile.FullPath | Should -Be $expectedFileName
        } 

        It 'leaves $script:_logFileOverwritten unchanged if log file path not updated' {
            $existingFileName = 'C:\Original\old.log'
            $newFileNameTemplate = 'C:\Original\old.log'
            $expectedFileName = $existingFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $newFileNameTemplate

            $script:_logConfiguration.LogFile.FullPath | Should -Be $existingFileName
            $script:_logFileOverwritten | Should -Be $True
        }

        It 'leaves $script:_logFileOverwritten unchanged if log file path with date not updated' {
            $fileNameTemplate = 'C:\Test\Test${dateText}.log'
            $existingFileName = GetFileNameFromTemplate -FileNameTemplate $fileNameTemplate `
                -IncludeDateInFileName
            $expectedFileName = $existingFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $fileNameTemplate -IncludeDateInFileName

            $script:_logConfiguration.LogFile.FullPath | Should -Be $existingFileName
            $script:_logFileOverwritten | Should -Be $True
        } 

        It 'clears $script:_logFileOverwritten if log file path updated' {
            $existingFileName = 'C:\Original\old.log'
            $newFileNameTemplate = 'C:\New\New${dateText}.log'
            $expectedFileName = GetFileNameFromTemplate -FileNameTemplate $newFileNameTemplate `
                -IncludeDateInFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $newFileNameTemplate -IncludeDateInFileName

            $script:_logConfiguration.LogFile.FullPath | Should -Be $expectedFileName
            $script:_logFileOverwritten | Should -Be $False
        }

        It 'clears $script:_logFileOverwritten if log file containing date updated' {
            $fileNameTemplate = 'C:\Test\Test${dateText}.log'
            $oldDate = Get-Date -Year 2018 -Month 1 -Day 15
            $existingFileName = $ExecutionContext.InvokeCommand.ExpandString($fileNameTemplate)
            $expectedFileName = GetFileNameFromTemplate -FileNameTemplate $fileNameTemplate `
                -IncludeDateInFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $fileNameTemplate -IncludeDateInFileName

            $script:_logConfiguration.LogFile.FullPath | Should -Be $expectedFileName
            $script:_logFileOverwritten | Should -Be $False
        }
    }

    Describe "DeepCopyHashTable" {     

        It 'returns $Null if source is $Null' {
            $sourceHashTable = $Null
            Private_DeepCopyHashTable -HashTable $sourceHashTable | Should -Be $Null       
        }     

        It 'returns empty hash table if source is empty hash table' {
            $sourceHashTable = @{}
            $copiedHashTable = Private_DeepCopyHashTable -HashTable $sourceHashTable 
            $copiedHashTable | Should -BeOfType [Hashtable]
            $copiedHashTable.Keys.Count | Should -Be 0
        }      

        It 'copies value type element which is independent of the source element' {
            $sourceHashTable = @{ Key = 10 }
            $copiedHashTable = Private_DeepCopyHashTable -HashTable $sourceHashTable 
            $copiedHashTable | Should -BeOfType [Hashtable]
            $sourceHashTable.Key = -1
            $copiedHashTable.Key | Should -Be 10
        }       

        It 'copies string element which is independent of the source element' {
            $elementValue = "Hello"
            $sourceHashTable = @{ Key = $elementValue }
            $copiedHashTable = Private_DeepCopyHashTable -HashTable $sourceHashTable 
            $copiedHashTable | Should -BeOfType [Hashtable]
            $sourceHashTable.Key = "Goodbye"
            $copiedHashTable.Key | Should -Be $elementValue
        }        

        It 'copies array element which is independent of the source element' {
            $elementValue = @(1, 2, 3)
            $sourceHashTable = @{ Key = $elementValue }
            $copiedHashTable = Private_DeepCopyHashTable -HashTable $sourceHashTable 
            $sourceHashTable.Key += @(4, 5)
            $copiedHashTable | Should -BeOfType [Hashtable]
            $copiedHashTable.Key | Should -Be @(1, 2, 3)
        }        

        It 'copies hashtable element which is independent of the source element' {
            $elementValue = @{ NestedKey = 10 }
            $sourceHashTable = @{ Key = $elementValue }
            $copiedHashTable = Private_DeepCopyHashTable -HashTable $sourceHashTable 
            $sourceHashTable.Key["SecondKey"] = "Twenty"
            $copiedHashTable | Should -BeOfType [Hashtable]
            $copiedNestedHashTable = $copiedHashTable.Key
            $copiedNestedHashTable | Should -BeOfType [Hashtable]
            $copiedNestedHashTable.Keys.Count | Should -Be 1
            $copiedNestedHashTable.NestedKey | Should -Be 10
        } 
    }

    Describe "SetMessageFormat" {     

        It 'sets LogConfiguration.MessageFormat equal to the specified string' {
            $newMessageFormat = '{MessageLevel}: {Message}'
            $originalMessageFormat = $script:_logConfiguration.MessageFormat
            Private_SetMessageFormat -MessageFormat $newMessageFormat 
            $script:_logConfiguration.MessageFormat | Should -Be $newMessageFormat
            $script:_logConfiguration.MessageFormat | Should -Not -Be $originalMessageFormat
        }    

        It 'sets MessageFormatInfo to be a hashtable' {
            $script:_messageFormatInfo = $Null
            $newMessageFormat = '{MessageLevel}: {Message}'
            Private_SetMessageFormat -MessageFormat $newMessageFormat  
            $script:_messageFormatInfo | Should -BeOfType [Hashtable]
        }  

        It "creates element '<_>' in MessageFormatInfo hashtable" -ForEach (
            'RawFormat', 'WorkingFormat', 'FieldsPresent'
        ) {
            $script:_messageFormatInfo = $Null
            $newMessageFormat = '{MessageLevel}: {Message}'
            Private_SetMessageFormat -MessageFormat $newMessageFormat
            $script:_messageFormatInfo.ContainsKey($_) | Should -Be $True
        }  

        It 'sets MessageFormatInfo.RawFormat equal to the specified string' {
            $script:_messageFormatInfo = $Null
            $newMessageFormat = '{MessageLevel}: {Message}'
            Private_SetMessageFormat -MessageFormat $newMessageFormat  
            $script:_messageFormatInfo.RawFormat | Should -Be $newMessageFormat
        }

        It 'sets MessageFormatInfo.WorkingFormat equal to the specified string with placeholders replaced by variable names' {
            $script:_messageFormatInfo = $Null
            $newMessageFormat = '{MessageLevel}: {Message}'
            Private_SetMessageFormat -MessageFormat $newMessageFormat  
            $script:_messageFormatInfo.WorkingFormat | Should -Be '${MessageLevel}: ${Message}'
        }

        It 'sets MessageFormatInfo.FieldsPresent to be an array' {
            $script:_messageFormatInfo = $Null
            $newMessageFormat = '{MessageLevel}: {Message}'
            Private_SetMessageFormat -MessageFormat $newMessageFormat  
            $script:_messageFormatInfo.FieldsPresent.GetType().FullName | Should -Be "System.Object[]"
        }

        It 'adds placeholder names from specified string to MessageFormatInfo.FieldsPresent' {
            $script:_messageFormatInfo = $Null
            $newMessageFormat = '{MessageLevel}: {Message}'
            Private_SetMessageFormat -MessageFormat $newMessageFormat  
            $script:_messageFormatInfo.FieldsPresent.Contains('MessageLevel') | Should -Be $True 
            $script:_messageFormatInfo.FieldsPresent.Contains('Message') | Should -Be $True
        }
    }

    Describe "SetConfigTextColor" { 

        It "sets LogConfiguration.HostTextColor to DefaultHostTextColor if HostTextColor doesn't exist" {
            $script:_logConfiguration.Remove("HostTextColor")
            $script:_logConfiguration.ContainsKey("HostTextColor") | Should -Not -Be $True
            Private_SetConfigTextColor -ConfigurationKey Error -ColorName Green
            $script:_logConfiguration.HostTextColor | Should -Be $script:_defaultHostTextColor
        }

        It "sets specified LogConfiguration.HostTextColor element to specified colour" {
            Private_SetConfigTextColor -ConfigurationKey Error -ColorName Yellow
            $script:_logConfiguration.HostTextColor.Error | Should -Be Yellow
        }
    }

    Describe 'ValidateCategoryInfoItem' {

        It 'throws ArgumentException if CategoryInfoItem is not a hashtable or an array' {
            $testValue = 'Hello world'

            { Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException  `
                    -WithMessage 'Expected argument to be either a hash table or an array but it is System.String'
        }

        It 'throws ArgumentException if CategoryInfoItem is empty array' {
            $testValue = @()

            { Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException  `
                    -WithMessage 'Expected an array of 2 elements but 0 supplied'
        }

        It 'throws ArgumentException if CategoryInfoItem is array with one element' {
            $testValue = @('text')

            { Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException  `
                    -WithMessage 'Expected an array of 2 elements but 1 supplied'
        }

        It 'throws ArgumentException if CategoryInfoItem is array with three elements' {
            $testValue = @('text1', 'text2', 'text3')

            { Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException  `
                    -WithMessage 'Expected an array of 2 elements but 3 supplied'
        }

        It 'throws ArgumentException if CategoryInfoItem is two-element array where first element is not a string' {
            $testValue = @(1, 'text')

            { Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException  `
                    -WithMessage 'Expected first element to be a string but it is System.Int32'
        }

        It 'throws ArgumentException if CategoryInfoItem is two-element array where second element is not a hashtable' {
            $testValue = @('Key', 'text')

            { Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException  `
                    -WithMessage 'Expected second element to be a hash table but it is System.String'
        }

        It 'returns $True if CategoryInfoItem is two-element array with types @([string], [hashtable])' {
            $testValue = @('Key', @{})

            $result = Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue
            $result | Should -Be $True
        }

        It 'returns $True if CategoryInfoItem is empty hashtable' {
            $testValue = @{}

            $result = Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue
            $result | Should -Be $True
        }

        It 'throws ArgumentException if CategoryInfoItem is hashtable where first key is not a string' {
            $testValue = @{
                            10=@{}
                            Key2=@{}
                        }

            { Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException  `
                    -WithMessage 'Expected key to be a string but it is System.Int32'
        }

        It 'throws ArgumentException if CategoryInfoItem is hashtable where second key is not a string' {
            $testValue = @{
                            Key1=@{}
                            10=@{}
                        }

            { Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException  `
                    -WithMessage 'Expected key to be a string but it is System.Int32'
        }

        It 'throws ArgumentException if CategoryInfoItem is hashtable where first value is not a hashtable' {
            $testValue = @{
                            Key1='Value'
                            Key2=@{}
                        }

            { Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException  `
                    -WithMessage 'Expected value to be a hash table but it is System.String'
        }

        It 'throws ArgumentException if CategoryInfoItem is hashtable where second value is not a hashtable' {
            $testValue = @{
                            Key1=@{}
                            Key2='Value'
                        }

            { Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue } | 
                Assert-ExceptionThrown -WithTypeName ArgumentException  `
                    -WithMessage 'Expected value to be a hash table but it is System.String'
        }

        It 'returns $True if CategoryInfoItem is hashtable with single item of types [string]=[hashtable]' {
            $testValue = @{
                            Key1=@{}
                        }

            $result = Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue
            $result | Should -Be $True
        }

        It 'returns $True if CategoryInfoItem is hashtable with two items of types [string]=[hashtable]' {
            $testValue = @{
                            Key1=@{}
                            Key2=@{}
                        }

            $result = Private_ValidateCategoryInfoItem -CategoryInfoItem $testValue
            $result | Should -Be $True
        }
    }
}
