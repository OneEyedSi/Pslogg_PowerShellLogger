<#
.SYNOPSIS
Tests of the exported logging functions in the Logging module.

.DESCRIPTION
Pester tests of the logging functions exported from the Logging module.
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

<#
.SYNOPSIS
Test function with no arguments that throws an exception.
#>
function NoArgsException
{
    throw [ArgumentException] "This is the message"
}

InModuleScope Logging {

    # Need to dot source the helper file within the InModuleScope block to be able to use its 
    # functions within a test.
    . (Join-Path $PSScriptRoot .\AssertExceptionThrown.ps1 -Resolve)

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

    function GetResetLogFilePath ()
    {
        $dateString = Get-Date -Format "_yyyyMMdd"
        $path = "TestDrive:\Results${dateString}.log"
        return $path
    }

    # Sets the Logging configuration to its defaults, apart from LogFileName and LogFilePath.
    function ResetConfiguration ()
    {
        $script:_logConfiguration = Private_DeepCopyHashTable $script:_defaultLogConfiguration
        $script:_logConfiguration.LogFileName = "TestDrive:\Results.log"
        $script:_messageFormatInfo = GetDefaultMessageFormatInfo
        $script:_logFilePath = GetResetLogFilePath
        $script:_logFileOverwritten = $False
    }

    Describe 'Write-LogMessage' {             

        BeforeEach {
            ResetConfiguration
        }

        It 'throws no error when no Message supplied' {
            { Write-LogMessage } | Should -Not -Throw
        }

        It 'writes to log when no Message supplied' {
            Mock Write-Host

            Write-LogMessage

            Assert-MockCalled -CommandName Write-Host -Scope It -Times 1
        }

        It 'throws exception if both -IsError and -IsWarning switches set' {
            { Write-LogMessage -Message 'hello world' -IsError -IsWarning } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsError and -IsInformation switches set' {
            { Write-LogMessage -Message 'hello world' -IsError -IsInformation } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsError and -IsDebug switches set' {
            { Write-LogMessage -Message 'hello world' -IsError -IsDebug } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsError and -IsVerbose switches set' {
            { Write-LogMessage -Message 'hello world' -IsError -IsVerbose } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsError and -IsSuccessResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsError -IsSuccessResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsError and -IsFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsError -IsFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsError and -IsPartialFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsError -IsPartialFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsWarning and -IsInformation switches set' {
            { Write-LogMessage -Message 'hello world' -IsWarning -IsInformation } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsWarning and -IsDebug switches set' {
            { Write-LogMessage -Message 'hello world' -IsWarning -IsDebug } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsWarning and -IsVerbose switches set' {
            { Write-LogMessage -Message 'hello world' -IsWarning -IsVerbose } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsWarning and -IsSuccessResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsWarning -IsSuccessResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsWarning and -IsFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsWarning -IsFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsWarning and -IsPartialFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsWarning -IsPartialFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsInformation and -IsDebug switches set' {
            { Write-LogMessage -Message 'hello world' -IsInformation -IsDebug } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsInformation and -IsVerbose switches set' {
            { Write-LogMessage -Message 'hello world' -IsInformation -IsVerbose } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsInformation and -IsSuccessResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsInformation -IsSuccessResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsInformation and -IsFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsInformation -IsFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsInformation and -IsPartialFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsInformation -IsPartialFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsDebug and -IsVerbose switches set' {
            { Write-LogMessage -Message 'hello world' -IsDebug -IsVerbose } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsDebug and -IsSuccessResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsDebug -IsSuccessResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsDebug and -IsFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsDebug -IsFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsDebug and -IsPartialFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsDebug -IsPartialFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsVerbose and -IsSuccessResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsVerbose -IsSuccessResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsVerbose and -IsFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsVerbose -IsFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsVerbose and -IsPartialFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsVerbose -IsPartialFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsSuccessResult and -IsFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsSuccessResult -IsFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsSuccessResult and -IsPartialFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsSuccessResult -IsPartialFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -IsFailureResult and -IsPartialFailureResult switches set' {
            { Write-LogMessage -Message 'hello world' -IsFailureResult -IsPartialFailureResult } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Message Type switch parameter may be set'                                        
        }

        It 'throws exception if both -WriteToHost and -WriteToStreams switches set' {
            { Write-LogMessage -Message 'hello world' -WriteToHost -WriteToStreams } | 
                Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException `
                    -ExpectedExceptionMessage 'Only one Destination switch parameter may be set'                                        
        }
    }
}
