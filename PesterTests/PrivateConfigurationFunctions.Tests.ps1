<#
.SYNOPSIS
Tests of the private configuration functions in the Prog module.

.DESCRIPTION
Pester tests of the private functions in the Prog module that are called by the public 
configuration functions.
#>

# PowerShell allows multiple modules of the same name to be imported from different locations.  
# This would confuse Pester.  So, to be sure there are not multiple Prog modules imported, 
# remove all Prog modules and re-import only one.
Get-Module Prog | Remove-Module -Force
# Use $PSScriptRoot so this script will always import the Prog module in the Modules folder 
# adjacent to the folder containing this script, regardless of the location that Pester is 
# invoked from:
#                                     {parent folder}
#                                             |
#                   -----------------------------------------------------
#                   |                                                   |
#     {folder containing this script}                                Modules folder
#                   \                                                   |
#                    ------------------> imports                     Prog module folder
#                                                \                      |
#                                                 -----------------> Prog.psm1 module script
Import-Module (Join-Path $PSScriptRoot ..\Modules\Prog\Prog.psm1 -Resolve) -Force

InModuleScope Prog {

    Describe "GetAbsolutePath" {     

        It 'throws ParameterBindingValidationException if empty path supplied' {
            $originalPath = ''
            try
            {
                Private_GetAbsolutePath -Path ''
            }
            catch
            {
                $_.Exception.GetType().Name | Should -Be 'ParameterBindingValidationException'
            }            
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

            $script:_logFilePath = $ExistingFileName
            $script:_logConfiguration.LogFileName = $fileNameFromConfiguration
            $script:_logConfiguration.IncludeDateInFileName = $IncludeDateInFileName
            $script:_logFileOverwritten = $True

            Private_SetLogFilePath
        }

        It 'sets $script:_logFilePath to log file name from configuration' {
            $existingFileName = 'C:\Original\old.log'
            $newFileNameTemplate = 'C:\New\New${dateText}.log'
            $expectedFileName = GetFileNameFromTemplate -FileNameTemplate $newFileNameTemplate

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $newFileNameTemplate

            $script:_logFilePath | Should -Be $expectedFileName
        }     

        It 'adds date to $script:_logFilePath if IncludeDateInFileName configuration value is set' {
            $existingFileName = 'C:\Original\old.log'
            $newFileNameTemplate = 'C:\New\New${dateText}.log'
            $expectedFileName = GetFileNameFromTemplate -FileNameTemplate $newFileNameTemplate `
                -IncludeDateInFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $newFileNameTemplate -IncludeDateInFileName

            $script:_logFilePath | Should -Be $expectedFileName
        }

        It 'leaves $script:_logFilePath unchanged if identical to log file name from configuration' {
            $existingFileName = 'C:\Original\old.log'
            $newFileNameTemplate = 'C:\Original\old.log'
            $expectedFileName = $existingFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $newFileNameTemplate

            $script:_logFilePath | Should -Be $existingFileName
        }

        It 'leaves $script:_logFilePath unchanged if existing file name includes today''s date' {
            $fileNameTemplate = 'C:\Test\Test${dateText}.log'
            $existingFileName = GetFileNameFromTemplate -FileNameTemplate $fileNameTemplate `
                -IncludeDateInFileName
            $expectedFileName = $existingFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $fileNameTemplate -IncludeDateInFileName

            $script:_logFilePath | Should -Be $existingFileName
        } 

        It 'updates $script:_logFilePath if existing file name includes old date' {
            $fileNameTemplate = 'C:\Test\Test${dateText}.log'
            $oldDate = Get-Date -Year 2018 -Month 1 -Day 15
            $existingFileName = $ExecutionContext.InvokeCommand.ExpandString($fileNameTemplate)
            $expectedFileName = GetFileNameFromTemplate -FileNameTemplate $fileNameTemplate `
                -IncludeDateInFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $fileNameTemplate -IncludeDateInFileName

            $script:_logFilePath | Should -Be $expectedFileName
        } 

        It 'leaves $script:_logFileOverwritten unchanged if log file path not updated' {
            $existingFileName = 'C:\Original\old.log'
            $newFileNameTemplate = 'C:\Original\old.log'
            $expectedFileName = $existingFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $newFileNameTemplate

            $script:_logFilePath | Should -Be $existingFileName
            $script:_logFileOverwritten | Should -Be $True
        }

        It 'leaves $script:_logFileOverwritten unchanged if log file path with date not updated' {
            $fileNameTemplate = 'C:\Test\Test${dateText}.log'
            $existingFileName = GetFileNameFromTemplate -FileNameTemplate $fileNameTemplate `
                -IncludeDateInFileName
            $expectedFileName = $existingFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $fileNameTemplate -IncludeDateInFileName

            $script:_logFilePath | Should -Be $existingFileName
            $script:_logFileOverwritten | Should -Be $True
        } 

        It 'clears $script:_logFileOverwritten if log file path updated' {
            $existingFileName = 'C:\Original\old.log'
            $newFileNameTemplate = 'C:\New\New${dateText}.log'
            $expectedFileName = GetFileNameFromTemplate -FileNameTemplate $newFileNameTemplate `
                -IncludeDateInFileName

            TestLogFileConfiguration -ExistingFileName $existingFileName `
                -NewFileNameTemplate $newFileNameTemplate -IncludeDateInFileName

            $script:_logFilePath | Should -Be $expectedFileName
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

            $script:_logFilePath | Should -Be $expectedFileName
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
            $newMessageFormat = '{LogLevel}: {Message}'
            $originalMessageFormat = $script:_logConfiguration.MessageFormat
            Private_SetMessageFormat -MessageFormat $newMessageFormat 
            $script:_logConfiguration.MessageFormat | Should -Be $newMessageFormat
            $script:_logConfiguration.MessageFormat | Should -Not -Be $originalMessageFormat
        }    

        It 'sets MessageFormatInfo to be a hashtable' {
            $script:_messageFormatInfo = $Null
            $newMessageFormat = '{LogLevel}: {Message}'
            Private_SetMessageFormat -MessageFormat $newMessageFormat  
            $script:_messageFormatInfo | Should -BeOfType [Hashtable]
        }    

        function TestMessageFormatInfoHasKey([string]$Key)
        {
            It "creates element '$key' in MessageFormatInfo hashtable" {
                $script:_messageFormatInfo = $Null
                $newMessageFormat = '{LogLevel}: {Message}'
                Private_SetMessageFormat -MessageFormat $newMessageFormat
                $script:_messageFormatInfo.ContainsKey($key) | Should -Be $True
            }  
        }

        TestMessageFormatInfoHasKey RawFormat
        TestMessageFormatInfoHasKey WorkingFormat
        TestMessageFormatInfoHasKey FieldsPresent

        It 'sets MessageFormatInfo.RawFormat equal to the specified string' {
            $script:_messageFormatInfo = $Null
            $newMessageFormat = '{LogLevel}: {Message}'
            Private_SetMessageFormat -MessageFormat $newMessageFormat  
            $script:_messageFormatInfo.RawFormat | Should -Be $newMessageFormat
        }

        It 'sets MessageFormatInfo.WorkingFormat equal to the specified string with placeholders replaced by variable names' {
            $script:_messageFormatInfo = $Null
            $newMessageFormat = '{LogLevel}: {Message}'
            Private_SetMessageFormat -MessageFormat $newMessageFormat  
            $script:_messageFormatInfo.WorkingFormat | Should -Be '${LogLevel}: ${Message}'
        }

        It 'sets MessageFormatInfo.FieldsPresent to be an array' {
            $script:_messageFormatInfo = $Null
            $newMessageFormat = '{LogLevel}: {Message}'
            Private_SetMessageFormat -MessageFormat $newMessageFormat  
            $script:_messageFormatInfo.FieldsPresent.GetType().FullName | Should -Be "System.Object[]"
        }

        It 'adds placeholder names from specified string to MessageFormatInfo.FieldsPresent' {
            $script:_messageFormatInfo = $Null
            $newMessageFormat = '{LogLevel}: {Message}'
            Private_SetMessageFormat -MessageFormat $newMessageFormat  
            $script:_messageFormatInfo.FieldsPresent.Contains("LogLevel") | Should -Be $True 
            $script:_messageFormatInfo.FieldsPresent.Contains("Message") | Should -Be $True
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
}
