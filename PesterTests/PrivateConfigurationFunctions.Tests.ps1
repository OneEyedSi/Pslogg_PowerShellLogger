<#
.SYNOPSIS
Tests of the private configuration functions in the Logging module.

.DESCRIPTION
Pester tests of the private functions in the Logging module that are called by the public 
configuration functions.
#>

# PowerShell allows multiple modules of the same name to be imported from different locations.  
# This would confuse Pester.  So, to be sure there are not multiple Logging modules imported, 
# remove all Logging modules and re-import only one.
Get-Module Logging | Remove-Module -Force
Import-Module ..\Modules\Logging\Logging.psm1 -Force

InModuleScope Logging {

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
            $pathOfFolderContainingThisScript = Convert-Path -Path .
            $expectedPath = Join-Path $pathOfFolderContainingThisScript $originalPath
            $absolutePath = Private_GetAbsolutePath -Path $originalPath 
            Private_GetAbsolutePath -Path $originalPath | Should -Be $expectedPath
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
}
