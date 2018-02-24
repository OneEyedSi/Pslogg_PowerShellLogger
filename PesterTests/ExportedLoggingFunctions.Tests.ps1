<#
.SYNOPSIS
Tests of the exported Prog functions in the Prog module.

.DESCRIPTION
Pester tests of the logging functions exported from the Prog module.
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
#                                                 -----------------> Prog.psd1 module script
Import-Module (Join-Path $PSScriptRoot ..\Modules\Prog\Prog.psd1 -Resolve) -Force

<#
.SYNOPSIS
Test function with no arguments that throws an exception.
#>
function NoArgsException
{
    throw [ArgumentException] "This is the message"
}

InModuleScope Prog {

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
        $path = "$TestDrive\Results${dateString}.log"
        return $path
    }

    # Sets the Prog configuration to its defaults, apart from LogFileName and LogFilePath.
    function ResetConfiguration ()
    {
        $script:_logConfiguration = Private_DeepCopyHashTable $script:_defaultLogConfiguration
        $script:_logConfiguration.LogFileName = "$TestDrive\Results.log"
        $script:_messageFormatInfo = GetDefaultMessageFormatInfo
        $script:_logFilePath = GetResetLogFilePath
        $script:_logFileOverwritten = $False
    }

    function MockHostAndStreamWriters ()
    {
        Mock Write-Host
        Mock Write-Error
        Mock Write-Warning
        Mock Write-Information
        Mock Write-Debug
        Mock Write-Verbose
    }

    function AssertCorrectWriteCommandCalled (
        [switch]$WriteToHost, 
        [switch]$WriteToStreams,
        [string]$messageType
        )
    {
        $timesWriteHostCalled = 0
        $timesWriteErrorCalled = 0
        $timesWriteWarningCalled = 0
        $timesWriteInformationCalled = 0
        $timesWriteDebugCalled = 0
        $timesWriteVerboseCalled = 0

        if ($WriteToHost.IsPresent)
        {
            $timesWriteHostCalled = 1
        }
        elseif ($WriteToStreams.IsPresent -or 
                (-not $WriteToHost.IsPresent -and -not $WriteToStreams.IsPresent))
        {
            switch ($messageType)
            {
                'ERROR'				{ $timesWriteErrorCalled = 1; break }
                'WARNING'   		{ $timesWriteWarningCalled = 1; break }
                'INFORMATION'		{ $timesWriteInformationCalled = 1; break }
                'DEBUG'				{ $timesWriteDebugCalled = 1; break }
                'VERBOSE'			{ $timesWriteVerboseCalled = 1; break } 
                'SUCCESS'			{ $timesWriteInformationCalled = 1; break }
                'FAILURE'			{ $timesWriteInformationCalled = 1; break }
                'PARTIAL FAILURE'	{ $timesWriteInformationCalled = 1; break } 
                default				{ $timesWriteInformationCalled = 1 }                                                             
            } 
        }
        
        Assert-MockCalled -CommandName Write-Host -Scope It -Times $timesWriteHostCalled
        Assert-MockCalled -CommandName Write-Error -Scope It -Times $timesWriteErrorCalled
        Assert-MockCalled -CommandName Write-Warning -Scope It -Times $timesWriteWarningCalled
        Assert-MockCalled -CommandName Write-Information -Scope It -Times $timesWriteInformationCalled
        Assert-MockCalled -CommandName Write-Debug -Scope It -Times $timesWriteDebugCalled
        Assert-MockCalled -CommandName Write-Verbose -Scope It -Times $timesWriteVerboseCalled             
    }

    function AssertWriteHostCalled ([string]$WithTextColor)
    {
        Assert-MockCalled -CommandName Write-Host `
            -ParameterFilter { $ForegroundColor -eq $WithTextColor } -Scope It -Times 1
    }

    function MockFileWriter ()
    {
        Mock Set-Content
        Mock Add-Content
    }

    function AssertFileWriterCalled 
    (
        [string]$LogFilePath,
        [bool]$OverwriteLogFile, 
        [bool]$LogFileOverwritten
    )
    {
        $commandName = 'Set-Content'

        if (-not $OverwriteLogFile -or $LogFileOverwritten)
        {
            $commandName = 'Add-Content'
        }
        
        Assert-MockCalled -CommandName $commandName -Scope It -Times 1 `
            -ParameterFilter { $Path -eq $LogFilePath }
    }

    function AssertFileWriterNotCalled ()
    {
        Assert-MockCalled -CommandName Add-Content -Scope It -Times 0 
        Assert-MockCalled -CommandName Set-Content -Scope It -Times 0
    }

    function RemoveLogFile ([string]$Path)
    {
        if (Test-Path $Path)
        {
            Remove-Item $Path
        }
    }

    function NewLogFile ([string]$Path)
    {
        'Original content line 1', 'Original content line 2' | Set-Content -Path $Path

        if (-not (Test-Path $Path))
        {
            throw [Exception] "Test log file '$Path' was not created during test set up."
        }

        $originalContent = (Get-Content -Path $Path)
        if ($originalContent.Count -lt 2)
        {
            throw [Exception] `
                "Contents of test log file '$Path' were not created during test set up."
        }

        return $originalContent
    }

    Describe 'Write-LogMessage' {             

        BeforeEach {
            ResetConfiguration
        }

        Context 'No message supplied' {

            It 'throws no error when no Message supplied' {
                # Only want the message field logged, so Prog isn't logging any text at all.
                Private_SetMessageFormat '{Message}'

                { Write-LogMessage } | Should -Not -Throw
            }

            It 'writes to log when no Message supplied' {
                # Only want the message field logged, so Prog isn't logging any text at all.
                Private_SetMessageFormat '{Message}'
                Mock Write-Host

                Write-LogMessage

                Assert-MockCalled -CommandName Write-Host -Scope It -Times 1
            }
        }
            
        Context 'Parameter validation' {
            # Any mocks declared in a Context block, even if they are declared inside an It block, 
            # remain until the end of the Context block, affecting subsequent It blocks within 
            # the same Context.
            Mock Write-Host

            It 'throws exception if an invalid -HostTextColor is specified' {
                { Write-LogMessage -Message 'hello world' -HostTextColor DeepPurple } | 
                    Assert-ExceptionThrown -WithTypeName ParameterBindingValidationException `
                        -WithMessage "'DeepPurple' is not a valid text color"                                       
            }

            It 'does not throw exception if no -HostTextColor is specified' {
                
                { Write-LogMessage -Message 'hello world' } | 
                    Assert-ExceptionThrown -Not
            }

            It 'does not throw exception if valid -HostTextColor is specified' {
                
                { Write-LogMessage -Message 'hello world' -HostTextColor Cyan } | 
                    Assert-ExceptionThrown -Not
            }

            It 'throws exception if both -IsError and -IsWarning switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsWarning } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsError and -IsInformation switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsInformation } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsError and -IsDebug switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsDebug } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsError and -IsVerbose switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsVerbose } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsError and -IsSuccessResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsSuccessResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsError and -IsFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsError and -IsPartialFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsPartialFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsWarning and -IsInformation switches set' {
                { Write-LogMessage -Message 'hello world' -IsWarning -IsInformation } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsWarning and -IsDebug switches set' {
                { Write-LogMessage -Message 'hello world' -IsWarning -IsDebug } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsWarning and -IsVerbose switches set' {
                { Write-LogMessage -Message 'hello world' -IsWarning -IsVerbose } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsWarning and -IsSuccessResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsWarning -IsSuccessResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsWarning and -IsFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsWarning -IsFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsWarning and -IsPartialFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsWarning -IsPartialFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsInformation and -IsDebug switches set' {
                { Write-LogMessage -Message 'hello world' -IsInformation -IsDebug } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsInformation and -IsVerbose switches set' {
                { Write-LogMessage -Message 'hello world' -IsInformation -IsVerbose } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsInformation and -IsSuccessResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsInformation -IsSuccessResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsInformation and -IsFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsInformation -IsFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsInformation and -IsPartialFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsInformation -IsPartialFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsDebug and -IsVerbose switches set' {
                { Write-LogMessage -Message 'hello world' -IsDebug -IsVerbose } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsDebug and -IsSuccessResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsDebug -IsSuccessResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsDebug and -IsFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsDebug -IsFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsDebug and -IsPartialFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsDebug -IsPartialFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsVerbose and -IsSuccessResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsVerbose -IsSuccessResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsVerbose and -IsFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsVerbose -IsFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsVerbose and -IsPartialFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsVerbose -IsPartialFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsSuccessResult and -IsFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsSuccessResult -IsFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsSuccessResult and -IsPartialFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsSuccessResult -IsPartialFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -IsFailureResult and -IsPartialFailureResult switches set' {
                { Write-LogMessage -Message 'hello world' -IsFailureResult -IsPartialFailureResult } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Type switch parameter may be set'                                        
            }

            It 'throws exception if both -WriteToHost and -WriteToStreams switches set' {
                { Write-LogMessage -Message 'hello world' -WriteToHost -WriteToStreams } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Destination switch parameter may be set'                                        
            }
        }

        Context 'Logging to host and PowerShell streams' {
            MockHostAndStreamWriters

            It 'writes to host if -WriteToHost switch set' {
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                AssertCorrectWriteCommandCalled -WriteToHost                          
            }

            It 'writes to stream if -WriteToStreams switch set' {
                
                Write-LogMessage -Message 'hello world' -WriteToStreams

                AssertCorrectWriteCommandCalled -WriteToStreams -MessageType 'INFORMATION'                          
            }

            It 'writes to host if neither -WriteToHost nor -WriteToStreams switches set and configuration.WriteToHost set' {
                $script:_logConfiguration.WriteToHost = $True
                
                Write-LogMessage -Message 'hello world'

                AssertCorrectWriteCommandCalled -WriteToHost
            }

            It 'writes to stream if neither -WriteToHost nor -WriteToStreams switches set and configuration.WriteToHost cleared' {
                $script:_logConfiguration.WriteToHost = $False
                
                Write-LogMessage -Message 'hello world'

                AssertCorrectWriteCommandCalled -WriteToStreams -MessageType 'INFORMATION'
            }

            It 'writes to host in colour specified via -HostTextColor' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Information = 'White'
                
                Write-LogMessage -Message 'hello world' -WriteToHost -HostTextColor $textColour

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It 'writes to host in Error text colour if -WriteToHost and -IsError switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Error = $textColour
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsError 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It 'writes to host in Warning text colour if -WriteToHost and -IsWarning switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Warning = $textColour
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsWarning 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It 'writes to host in Debug text colour if -WriteToHost and -IsDebug switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Debug = $textColour
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsDebug 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It 'writes to host in Information text colour if -WriteToHost and -IsInformation switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Information = $textColour
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsInformation 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It 'writes to host in Verbose text colour if -WriteToHost and -IsVerbose switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Verbose = $textColour
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsVerbose 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It 'writes to host in Success text colour if -WriteToHost and -IsSuccessResult switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Success = $textColour
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsSuccessResult 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It 'writes to host in Failure text colour if -WriteToHost and -IsFailureResult switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Failure = $textColour
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsFailureResult 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It 'writes to host in PartialFailure text colour if -WriteToHost and -IsPartialFailureResult switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.PartialFailure = $textColour
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsPartialFailureResult 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It '-HostTextColor overrides the colour determined by a Message Type switch' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Error = 'DarkMagenta'
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsError -HostTextColor $textColour

                AssertWriteHostCalled -WithTextColor $textColour
            }
        }

        Context 'Logging to file with mocked file writer' {
            Mock Write-Host
            MockFileWriter

            It 'does not attempt to write to a log file when configuration LogFileName blank' {
                $script:_logConfiguration.LogFileName = ''
                Private_SetLogFilePath

                Write-LogMessage -Message 'hello world' -WriteToHost

                AssertFileWriterNotCalled
            }

            It 'does not attempt to write to a log file when configuration LogFileName empty string' {
                $script:_logConfiguration.LogFileName = ' '
                Private_SetLogFilePath

                Write-LogMessage -Message 'hello world' -WriteToHost

                AssertFileWriterNotCalled
            }

            It 'does not attempt to write to a log file when configuration LogFileName $Null' {
                $script:_logConfiguration.LogFileName = $Null
                Private_SetLogFilePath

                Write-LogMessage -Message 'hello world' -WriteToHost

                AssertFileWriterNotCalled
            }

            It 'does not attempt to write to a log file when configuration LogFileName not a valid path' {
                $logFileName = 'CC:\Test\Test.log'
                # This scenario should never occur.  Prog should throw an exception when setting 
                # LogFileName to an invalid path via Set-LogConfiguration.
                $script:_logConfiguration.LogFileName = $logFileName
                $script:_logFilePath = $logFileName
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                AssertFileWriterNotCalled
            }

            It 'does not add date to log file name when IncludeDateInFileName cleared' {
                $script:_logConfiguration.IncludeDateInFileName = $False
                Private_SetLogFilePath
                $logFileAlreadyOverwritten = $script:_logFileOverwritten
                
                Write-LogMessage -Message 'hello world' -WriteToHost
                
                $logFileName = $script:_logConfiguration.LogFileName
                AssertFileWriterCalled -LogFilePath $logFileName `
                    -OverwriteLogFile $script:_logConfiguration.OverwriteLogFile `
                    -LogFileOverwritten $logFileAlreadyOverwritten
            }

            It 'adds date to log file name when IncludeDateInFileName set' {
                $script:_logConfiguration.IncludeDateInFileName = $True
                Private_SetLogFilePath
                $logFileAlreadyOverwritten = $script:_logFileOverwritten
                
                Write-LogMessage -Message 'hello world' -WriteToHost
                
                $logFileName = GetResetLogFilePath
                AssertFileWriterCalled -LogFilePath $logFileName `
                    -OverwriteLogFile $script:_logConfiguration.OverwriteLogFile `
                    -LogFileOverwritten $logFileAlreadyOverwritten
            }
        }

        Context 'Logging to file with actual file writer' {
            
            BeforeEach {
                RemoveLogFile -Path $script:_logFilePath
            }

            Mock Write-Host

            It 'creates a log file when configuration LogFileName set and OverwriteLogFile cleared, and log file does not exist' {
                $script:_logConfiguration.OverwriteLogFile = $False
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logFilePath | Should -Exist
            }

            It 'creates a log file when configuration LogFileName and OverwriteLogFile set, and log file does not exist' {
                $script:_logConfiguration.OverwriteLogFile = $True                
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logFilePath | Should -Exist
            }

            It 'appends to existing log file when configuration OverwriteLogFile cleared' {
                $script:_logConfiguration.OverwriteLogFile = $False

                $originalContent = NewLogFile -Path $script:_logFilePath                
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logFilePath | Should -Exist
                $newContent = (Get-Content -Path $script:_logFilePath)
                $newContent.Count | Should -Be 3
                $newContent[2] | Should -BeLike '*hello world*'               
                
                Write-LogMessage -Message 'second message' -WriteToHost

                $newContent = (Get-Content -Path $script:_logFilePath)
                $newContent.Count | Should -Be 4
                $newContent[3] | Should -BeLike '*second message*'
            }

            It 'overwrites an existing log file with first logged message when configuration OverwriteLogFile set' {
                $script:_logConfiguration.OverwriteLogFile = $True

                $originalContent = NewLogFile -Path $script:_logFilePath                
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logFilePath | Should -Exist
                $newContent = ,(Get-Content -Path $script:_logFilePath)
                $newContent.Count | Should -Be 1
                $newContent[0] | Should -BeLike '*hello world*'
            }

            It 'appends subsequent messages to log file when configuration OverwriteLogFile set' {
                $script:_logConfiguration.OverwriteLogFile = $True

                $originalContent = NewLogFile -Path $script:_logFilePath                
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logFilePath | Should -Exist
                $newContent = ,(Get-Content -Path $script:_logFilePath)
                $newContent.Count | Should -Be 1
                $newContent[0] | Should -BeLike '*hello world*'               
                
                Write-LogMessage -Message 'second message' -WriteToHost

                $newContent = (Get-Content -Path $script:_logFilePath)
                $newContent.Count | Should -Be 2
                $newContent[1] | Should -BeLike '*second message*'               
                
                Write-LogMessage -Message 'third message' -WriteToHost

                $newContent = (Get-Content -Path $script:_logFilePath)
                $newContent.Count | Should -Be 3
                $newContent[2] | Should -BeLike '*third message*'
            }
        }
    }
}
