<#
.SYNOPSIS
Tests of the exported Pslogg functions in the Pslogg module.

.DESCRIPTION
Pester tests of the logging functions exported from the Pslogg module.
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

<#
.SYNOPSIS
Test function with no arguments that throws an exception.
#>
function NoArgsException
{
    throw [ArgumentException] "This is the message"
}

InModuleScope Pslogg {

    # Need to dot source the helper file within the InModuleScope block to be able to use its 
    # functions within a test.
    . (Join-Path $PSScriptRoot .\AssertExceptionThrown.ps1 -Resolve)

    # Gets the MessageFormatInfo hashtable that matches the default configuration.
    function GetDefaultMessageFormatInfo ()
    {
        $messageFormatInfo = @{
                                RawFormat = '{Timestamp:yyyy-MM-dd hh:mm:ss.fff} | {CallerName} | {MessageLevel} | {Message}'
                                WorkingFormat = '$($Timestamp.ToString(''yyyy-MM-dd hh:mm:ss.fff'')) | ${CallerName} | ${MessageLevel} | ${Message}'
                                FieldsPresent = @('Message', 'Timestamp', 'CallerName', 'MessageLevel')
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
        $dateString = Get-Date -Format '_yyyyMMdd'
        $path = "$TestDrive\Results${dateString}.log"
        return $path
    }

    # Sets the Pslogg configuration to its defaults, apart from LogFileName and LogFilePath.
    function ResetConfiguration ()
    {
        $script:_logConfiguration = Private_DeepCopyHashTable $script:_defaultLogConfiguration
        $script:_logConfiguration.LogFile.Name = "$TestDrive\Results.log"
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
        [string]$MessageLevel
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
            switch ($MessageLevel)
            {
                ERROR				{ $timesWriteErrorCalled = 1; break }
                WARNING   			{ $timesWriteWarningCalled = 1; break }
                INFORMATION			{ $timesWriteInformationCalled = 1; break }
                DEBUG				{ $timesWriteDebugCalled = 1; break }
                VERBOSE				{ $timesWriteVerboseCalled = 1; break } 
                SUCCESS				{ $timesWriteInformationCalled = 1; break }
                FAILURE				{ $timesWriteInformationCalled = 1; break }
                PARTIAL_FAILURE		{ $timesWriteInformationCalled = 1; break } 
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

    function TestMessageLevelColor([string]$MessageLevel)
    {
        It "writes to host in $MessageLevel text colour when -WriteToHost switch set, -MessageLevel set to $MessageLevel and -HostTextColor not specified" {
            $textColour = 'DarkRed'
            $script:_logConfiguration.HostTextColor[$MessageLevel] = $textColour
            $script:_logConfiguration.LogLevel = 'VERBOSE'
                
            Write-LogMessage -Message 'hello world' -WriteToHost -MessageLevel $MessageLevel

            AssertWriteHostCalled -WithTextColor $textColour
        }
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

    function TestLogLevelWithMessageLevelSwitch
    (
        [string]$LogLevel, 
        [string]$MessageLevel, 
        [switch]$ShouldWrite
    )
    {
        $writeText = 'does not write'
        $numberTimesShouldWrite = 0
        if ($ShouldWrite.IsPresent)
        {
            $writeText = 'writes'
            $numberTimesShouldWrite = 1
        }

        $isError = $False
        $isWarning = $False
        $isInformation = $False
        $isDebug = $False
        $isVerbose = $False

        $switchName = ''

        switch ($MessageLevel)
        {
            ERROR { 
                $isError = $True 
                $switchName = '-IsError'
                break 
            }
            WARNING	{ 
                $isWarning = $True 
                $switchName = '-IsWarning'
                break 
            }
            INFORMATION	{ 
                $isInformation = $True 
                $switchName = '-IsInformation'
                break 
            }
            DEBUG { 
                $isDebug = $True 
                $switchName = '-IsDebug'
                break 
            }
            VERBOSE { 
                $isVerbose = $True 
                $switchName = '-IsVerbose'
                break 
            }
        }

        It "$writeText message when configuration LogLevel is $LogLevel and $switchName switch is set" {

                Write-LogMessage -Message 'hello world' -WriteToHost `
                    -IsError:$isError -IsWarning:$isWarning -IsInformation:$isInformation `
                    -IsDebug:$isDebug -IsVerbose:$isVerbose 

                Assert-MockCalled -CommandName Write-Host -Scope It -Times $numberTimesShouldWrite
            }
    }

    function TestLogLevelWithMessageLevelText
    (
        [string]$LogLevel, 
        [string]$MessageLevel, 
        [switch]$ShouldWrite
    )
    {
        $writeText = 'does not write'
        $numberTimesShouldWrite = 0
        if ($ShouldWrite.IsPresent)
        {
            $writeText = 'writes'
            $numberTimesShouldWrite = 1
        }

        It "$writeText message when configuration LogLevel is $LogLevel and Message Level is $MessageLevel" {

                Write-LogMessage -Message 'hello world' -WriteToHost -MessageLevel $MessageLevel

                Assert-MockCalled -CommandName Write-Host -Scope It -Times $numberTimesShouldWrite
            }
    }

    function TestMessageFormat
    (
        [scriptblock]$FunctionUnderTest, 
        [string]$ExpectedLoggedText,
        [switch]$DoRegexMatch
    )
    {
        Invoke-Command $FunctionUnderTest

        Assert-MockCalled -CommandName Write-Host -Scope It -Times 1 `
            -ParameterFilter {
                # It's messy throwing exceptions inside a parameter filter but we get 
                # more informative error messages which include expected and actual text, 
                # rather than unhelpful
                #   "Expected Write-Host in module Pslogg to be called at least 1 times but was called 0 times"
                if ($DoRegexMatch.IsPresent)
                {
                    $Object | Should -Match $ExpectedLoggedText
                    # Should never reach here if $Object doesn't match the regex expression. 
                    # Should throw an exception instead.
                    $True
                }
                else
                {
                    $Object | Should -Be $ExpectedLoggedText
                    # Should never reach here if $Object doesn't equal the expected text. 
                    # Should throw an exception instead.
                    $True
                }
            }
    }

    Describe 'Write-LogMessage' {             

        BeforeEach {
            ResetConfiguration
        }

        Context 'No message supplied' {

            It 'throws no error when no Message supplied' {
                # Only want the message field logged, so Pslogg isn't logging any text at all.
                Private_SetMessageFormat '{Message}'

                { Write-LogMessage } | Should -Not -Throw
            }

            It 'writes to log when no Message supplied' {
                # Only want the message field logged, so Pslogg isn't logging any text at all.
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

            It 'does not throw exception when -Message is $Null' {
                
                { Write-LogMessage -Message $Null } | 
                    Assert-ExceptionThrown -Not
            }

            It 'writes to log when -Message is $Null' {
                # Only want the message field logged, so Pslogg isn't logging any text at all.
                Private_SetMessageFormat '{Message}'

                Write-LogMessage -Message $Null

                Assert-MockCalled -CommandName Write-Host -Scope It -Times 1
            }

            It 'does not throw exception when -Message is empty string' {
                
                { Write-LogMessage -Message '' } | 
                    Assert-ExceptionThrown -Not
            }

            It 'writes to log when -Message is empty string' {
                # Only want the message field logged, so Pslogg isn't logging any text at all.
                Private_SetMessageFormat '{Message}'

                Write-LogMessage -Message ''

                Assert-MockCalled -CommandName Write-Host -Scope It -Times 1
            }

            It 'does not throw exception when Message passed by position' {
                { Write-LogMessage 'hello world' } | 
                    Assert-ExceptionThrown -Not                                     
            }

            It 'writes to log when Message passed by position' {
                # Only want the message field logged, so Pslogg isn't logging any text at all.
                Private_SetMessageFormat '{Message}'

                Write-LogMessage 'hello world'

                Assert-MockCalled -CommandName Write-Host -Scope It -Times 1
            }

            It 'throws exception when an invalid -HostTextColor is specified' {
                { Write-LogMessage -Message 'hello world' -HostTextColor DeepPurple } | 
                    Assert-ExceptionThrown -WithTypeName ParameterBindingValidationException `
                        -WithMessage "'DeepPurple' is not a valid text color"                                       
            }

            It 'does not throw exception when no -HostTextColor is specified' {
                
                { Write-LogMessage -Message 'hello world' } | 
                    Assert-ExceptionThrown -Not
            }

            It 'does not throw exception when valid -HostTextColor is specified' {
                
                { Write-LogMessage -Message 'hello world' -HostTextColor Cyan } | 
                    Assert-ExceptionThrown -Not
            }

            It 'does not throw exception when neither -MessageLevel nor a message level switch are specified' {
                { Write-LogMessage 'hello world' } | 
                    Assert-ExceptionThrown -Not
            }

            It 'throws exception when an invalid -MessageLevel is specified' {
                { Write-LogMessage 'hello world' -MessageLevel INVALID } | 
                    Assert-ExceptionThrown -WithTypeName ParameterBindingValidationException `
                        -WithMessage "Cannot validate argument on parameter 'MessageLevel'"                                       
            }

            It 'does not throw exception when valid -MessageLevel is specified' {
                { Write-LogMessage 'hello world' -MessageLevel ERROR } | 
                    Assert-ExceptionThrown -Not
            }

            It 'throws exception when both -MessageLevel and -IsError are specified' {
                { Write-LogMessage 'hello world' -MessageLevel ERROR -IsError } | 
                    Assert-ExceptionThrown -WithTypeName ParameterBindingException `
                        -WithMessage 'Parameter set cannot be resolved'
            }

            It 'throws exception when both -MessageLevel and -IsWarning are specified' {
                { Write-LogMessage 'hello world' -MessageLevel ERROR -IsWarning } | 
                    Assert-ExceptionThrown -WithTypeName ParameterBindingException `
                        -WithMessage 'Parameter set cannot be resolved'
            }

            It 'throws exception when both -MessageLevel and -IsInformation are specified' {
                { Write-LogMessage 'hello world' -MessageLevel ERROR -IsInformation } | 
                    Assert-ExceptionThrown -WithTypeName ParameterBindingException `
                        -WithMessage 'Parameter set cannot be resolved'
            }

            It 'throws exception when both -MessageLevel and -IsDebug are specified' {
                { Write-LogMessage 'hello world' -MessageLevel ERROR -IsDebug } | 
                    Assert-ExceptionThrown -WithTypeName ParameterBindingException `
                        -WithMessage 'Parameter set cannot be resolved'
            }

            It 'throws exception when both -MessageLevel and -IsVerbose are specified' {
                { Write-LogMessage 'hello world' -MessageLevel ERROR -IsVerbose } | 
                    Assert-ExceptionThrown -WithTypeName ParameterBindingException `
                        -WithMessage 'Parameter set cannot be resolved'
            }

            It 'throws exception when both -IsError and -IsWarning switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsWarning } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Level switch parameter may be set'                                        
            }

            It 'throws exception when both -IsError and -IsInformation switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsInformation } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Level switch parameter may be set'                                        
            }

            It 'throws exception when both -IsError and -IsDebug switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsDebug } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Level switch parameter may be set'                                        
            }

            It 'throws exception when both -IsError and -IsVerbose switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsVerbose } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Level switch parameter may be set'                                        
            }

            It 'throws exception when both -IsWarning and -IsInformation switches set' {
                { Write-LogMessage -Message 'hello world' -IsWarning -IsInformation } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Level switch parameter may be set'                                        
            }

            It 'throws exception when both -IsWarning and -IsDebug switches set' {
                { Write-LogMessage -Message 'hello world' -IsWarning -IsDebug } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Level switch parameter may be set'                                        
            }

            It 'throws exception when both -IsWarning and -IsVerbose switches set' {
                { Write-LogMessage -Message 'hello world' -IsWarning -IsVerbose } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Level switch parameter may be set'                                        
            }

            It 'throws exception when both -IsInformation and -IsDebug switches set' {
                { Write-LogMessage -Message 'hello world' -IsInformation -IsDebug } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Level switch parameter may be set'                                        
            }

            It 'throws exception when both -IsInformation and -IsVerbose switches set' {
                { Write-LogMessage -Message 'hello world' -IsInformation -IsVerbose } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Level switch parameter may be set'                                        
            }

            It 'throws exception when both -IsDebug and -IsVerbose switches set' {
                { Write-LogMessage -Message 'hello world' -IsDebug -IsVerbose } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Message Level switch parameter may be set'                                        
            }

            It 'throws exception when both -WriteToHost and -WriteToStreams switches set' {
                { Write-LogMessage -Message 'hello world' -WriteToHost -WriteToStreams } | 
                    Assert-ExceptionThrown -WithTypeName ArgumentException `
                        -WithMessage 'Only one Destination switch parameter may be set'                                        
            }
        }

        Context 'Logging to host and PowerShell streams' {
            MockHostAndStreamWriters

            It 'writes to host when -WriteToHost switch set' {
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                AssertCorrectWriteCommandCalled -WriteToHost                          
            }

            It 'writes to stream when -WriteToStreams switch set' {
                
                Write-LogMessage -Message 'hello world' -WriteToStreams

                AssertCorrectWriteCommandCalled -WriteToStreams -MessageLevel 'INFORMATION'                          
            }

            It 'writes to host when neither -WriteToHost nor -WriteToStreams switches set and configuration.WriteToHost set' {
                $script:_logConfiguration.WriteToHost = $True
                
                Write-LogMessage -Message 'hello world'

                AssertCorrectWriteCommandCalled -WriteToHost
            }

            It 'writes to stream when neither -WriteToHost nor -WriteToStreams switches set and configuration.WriteToHost cleared' {
                $script:_logConfiguration.WriteToHost = $False
                
                Write-LogMessage -Message 'hello world'

                AssertCorrectWriteCommandCalled -WriteToStreams -MessageLevel 'INFORMATION'
            }

            It 'writes to host in colour specified via -HostTextColor' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Information = 'White'
                
                Write-LogMessage -Message 'hello world' -WriteToHost -HostTextColor $textColour

                AssertWriteHostCalled -WithTextColor $textColour
            }

            TestMessageLevelColor -MessageLevel ERROR
            TestMessageLevelColor -MessageLevel WARNING
            TestMessageLevelColor -MessageLevel INFORMATION
            TestMessageLevelColor -MessageLevel DEBUG
            TestMessageLevelColor -MessageLevel VERBOSE

            It 'writes to host in Error text colour when -WriteToHost and -IsError switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Error = $textColour
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsError 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It 'writes to host in Warning text colour when -WriteToHost and -IsWarning switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Warning = $textColour
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsWarning 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It 'writes to host in Information text colour when -WriteToHost and -IsInformation switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Information = $textColour
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsInformation 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It 'writes to host in Debug text colour when -WriteToHost and -IsDebug switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Debug = $textColour
                $script:_logConfiguration.LogLevel = 'Verbose'
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsDebug 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It 'writes to host in Verbose text colour when -WriteToHost and -IsVerbose switches set, and -HostTextColor not specified' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Verbose = $textColour
                $script:_logConfiguration.LogLevel = 'Verbose'
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsVerbose 

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It '-Category colour overrides the colour determined by -MessageLevel' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.CategoryInfo.Success.Color = $textColour
                $script:_logConfiguration.HostTextColor.Information = 'DarkCyan'
                
                Write-LogMessage -Message 'hello world' -WriteToHost -MessageLevel 'INFORMATION' `
                    -Category Success

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It '-Category colour overrides the colour determined by a Message Level switch' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.CategoryInfo.Success.Color = $textColour
                $script:_logConfiguration.HostTextColor.Information = 'DarkCyan'
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsInformation `
                    -Category Success

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It '-HostTextColor overrides the colour determined by -MessageLevel' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Error = 'DarkMagenta'
                
                Write-LogMessage -Message 'hello world' -WriteToHost -MessageLevel 'ERROR' `
                    -HostTextColor $textColour

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It '-HostTextColor overrides the colour determined by a Message Level switch' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Error = 'DarkMagenta'
                
                Write-LogMessage -Message 'hello world' -WriteToHost -IsError -HostTextColor $textColour

                AssertWriteHostCalled -WithTextColor $textColour
            }

            It '-HostTextColor overrides the Category colour' {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor.Error = 'DarkMagenta'
                
                Write-LogMessage -Message 'hello world' -WriteToHost -HostTextColor $textColour `
                    -Category Success

                AssertWriteHostCalled -WithTextColor $textColour
            }
        }

        Context 'Logging to file with mocked file writer' {
            Mock Write-Host
            MockFileWriter

            It 'does not attempt to write to a log file when configuration LogFile.Name blank' {
                $script:_logConfiguration.LogFile.Name = ''
                Private_SetLogFilePath

                Write-LogMessage -Message 'hello world' -WriteToHost

                AssertFileWriterNotCalled
            }

            It 'does not attempt to write to a log file when configuration LogFile.Name empty string' {
                $script:_logConfiguration.LogFile.Name = ' '
                Private_SetLogFilePath

                Write-LogMessage -Message 'hello world' -WriteToHost

                AssertFileWriterNotCalled
            }

            It 'does not attempt to write to a log file when configuration LogFile.Name $Null' {
                $script:_logConfiguration.LogFile.Name = $Null
                Private_SetLogFilePath

                Write-LogMessage -Message 'hello world' -WriteToHost

                AssertFileWriterNotCalled
            }

            It 'does not attempt to write to a log file when configuration LogFile.Name not a valid path' {
                $logFileName = 'CC:\Test\Test.log'
                # This scenario should never occur.  Pslogg should throw an exception when setting 
                # LogFileName to an invalid path via Set-LogConfiguration.
                $script:_logConfiguration.LogFile.Name = $logFileName
                $script:_logFilePath = $logFileName
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                AssertFileWriterNotCalled
            }

            It 'does not add date to log file name when LogFile.IncludeDateInFileName cleared' {
                $script:_logConfiguration.LogFile.IncludeDateInFileName = $False
                Private_SetLogFilePath
                $logFileAlreadyOverwritten = $script:_logFileOverwritten
                
                Write-LogMessage -Message 'hello world' -WriteToHost
                
                $logFileName = $script:_logConfiguration.LogFile.Name
                AssertFileWriterCalled -LogFilePath $logFileName `
                    -OverwriteLogFile $script:_logConfiguration.LogFile.Overwrite `
                    -LogFileOverwritten $logFileAlreadyOverwritten
            }

            It 'adds date to log file name when LogFile.IncludeDateInFileName set' {
                $script:_logConfiguration.LogFile.IncludeDateInFileName = $True
                Private_SetLogFilePath
                $logFileAlreadyOverwritten = $script:_logFileOverwritten
                
                Write-LogMessage -Message 'hello world' -WriteToHost
                
                $logFileName = GetResetLogFilePath
                AssertFileWriterCalled -LogFilePath $logFileName `
                    -OverwriteLogFile $script:_logConfiguration.LogFile.Overwrite `
                    -LogFileOverwritten $logFileAlreadyOverwritten
            }
        }

        Context 'Logging to file with actual file writer' {
            
            BeforeEach {
                RemoveLogFile -Path $script:_logFilePath
            }

            Mock Write-Host

            It 'creates a log file when configuration LogFile.Name set and LogFile.Overwrite cleared, and log file does not exist' {
                $script:_logConfiguration.LogFile.Overwrite = $False
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logFilePath | Should -Exist
            }

            It 'creates a log file when configuration LogFile.Name and LogFile.Overwrite set, and log file does not exist' {
                $script:_logConfiguration.LogFile.Overwrite = $True                
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logFilePath | Should -Exist
            }

            It 'appends to existing log file when configuration LogFile.Overwrite cleared' {
                $script:_logConfiguration.LogFile.Overwrite = $False

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

            It 'overwrites an existing log file with first logged message when configuration LogFile.Overwrite set' {
                $script:_logConfiguration.LogFile.Overwrite = $True

                $originalContent = NewLogFile -Path $script:_logFilePath                
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logFilePath | Should -Exist
                $newContent = ,(Get-Content -Path $script:_logFilePath)
                $newContent.Count | Should -Be 1
                $newContent[0] | Should -BeLike '*hello world*'
            }

            It 'appends subsequent messages to log file when configuration LogFile.Overwrite set' {
                $script:_logConfiguration.LogFile.Overwrite = $True

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

        Context 'Log level Off' {
            BeforeEach {
                $script:_logConfiguration.LogLevel = 'Off'
            }

            Mock Write-Host
            $logLevel = 'Off'

            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel ERROR
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel WARNING
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel INFORMATION
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel DEBUG
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel VERBOSE
            
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel ERROR
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel WARNING
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel INFORMATION
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel DEBUG
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel VERBOSE
        }

        Context 'Log level Error' {
            BeforeEach {
                $script:_logConfiguration.LogLevel = 'Error'
            }

            Mock Write-Host
            $logLevel = 'Error'
            
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel ERROR -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel WARNING
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel INFORMATION
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel DEBUG
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel VERBOSE

            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel ERROR -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel WARNING
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel INFORMATION
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel DEBUG
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel VERBOSE
        }

        Context 'Log level Warning' {
            BeforeEach {
                $script:_logConfiguration.LogLevel = 'Warning'
            }

            Mock Write-Host
            $logLevel = 'Warning'
            
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel ERROR -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel WARNING -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel INFORMATION
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel DEBUG
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel VERBOSE

            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel ERROR -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel WARNING -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel INFORMATION
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel DEBUG
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel VERBOSE
        }

        Context 'Log level Information' {
            BeforeEach {
                $script:_logConfiguration.LogLevel = 'Information'
            }

            Mock Write-Host
            $logLevel = 'Information'
            
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel ERROR -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel WARNING -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel INFORMATION -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel DEBUG
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel VERBOSE

            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel ERROR -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel WARNING -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel INFORMATION -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel DEBUG
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel VERBOSE
        }

        Context 'Log level Debug' {
            BeforeEach {
                $script:_logConfiguration.LogLevel = 'Debug'
            }

            Mock Write-Host
            $logLevel = 'Debug'
            
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel ERROR -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel WARNING -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel INFORMATION -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel DEBUG -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel VERBOSE

            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel ERROR -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel WARNING -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel INFORMATION -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel DEBUG -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel VERBOSE
        }

        Context 'Log level Verbose' {
            BeforeEach {
                $script:_logConfiguration.LogLevel = 'Verbose'
            }

            Mock Write-Host
            $logLevel = 'Verbose'
            
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel ERROR -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel WARNING -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel INFORMATION -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel DEBUG -ShouldWrite
            TestLogLevelWithMessageLevelSwitch -LogLevel $logLevel -MessageLevel VERBOSE -ShouldWrite
            
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel ERROR -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel WARNING -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel INFORMATION -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel DEBUG -ShouldWrite
            TestLogLevelWithMessageLevelText -LogLevel $logLevel -MessageLevel VERBOSE -ShouldWrite
        }

        Context 'Message Format' {
            
            Mock Write-Host
            
            It 'writes only message text when -MessageFormat contains only {Message} field' {
                TestMessageFormat `
                    -ExpectedLoggedText 'hello world' `
                    -FunctionUnderTest `
                        { 
                            Write-LogMessage -Message 'hello world' `
                                            -MessageFormat '{message}' `
                                            -WriteToHost 
                        }
            }
            
            It 'writes only timestamp with default formatting when -MessageFormat contains only {Timestamp} field' {
                TestMessageFormat `
                    -ExpectedLoggedText '^\d\d\d\d-\d\d-\d\d \d\d:\d\d:\d\d\.\d\d\d$' -DoRegexMatch `
                    -FunctionUnderTest `
                        { 
                            Write-LogMessage -Message 'hello world' `
                                            -MessageFormat '{Timestamp}' `
                                            -WriteToHost 
                        }
            }
            
            It 'formats timestamp with specified format string when -MessageFormat contains formatted {Timestamp} field' {
                TestMessageFormat `
                    -ExpectedLoggedText '^\d\d:\d\d:\d\d$' -DoRegexMatch `
                    -FunctionUnderTest `
                        { 
                            Write-LogMessage -Message 'hello world' `
                                            -MessageFormat '{Timestamp:hh:mm:ss}' `
                                            -WriteToHost 
                        }
            }
            
            It 'writes only caller name when -MessageFormat contains only {CallerName} field' {
                TestMessageFormat `
                    -ExpectedLoggedText 'Script ExportedLoggingFunctions.Tests.ps1' `
                    -FunctionUnderTest `
                        { 
                            Write-LogMessage -Message 'hello world' `
                                            -MessageFormat '{CallerName}' `
                                            -WriteToHost 
                        }
            }
            
            It 'writes only message level when -MessageFormat contains only {MessageLevel} field' {
                TestMessageFormat `
                    -ExpectedLoggedText 'INFORMATION' `
                    -FunctionUnderTest `
                        { 
                            Write-LogMessage -Message 'hello world' `
                                            -MessageFormat '{MessageLevel}' `
                                            -WriteToHost 
                        }
            }
            
            It 'writes only category when -MessageFormat contains only {Category} field and -Category specified' {
                TestMessageFormat `
                    -ExpectedLoggedText 'SUCCESS' `
                    -FunctionUnderTest `
                        { 
                            Write-LogMessage -Message 'hello world' `
                                            -MessageFormat '{Category}' `
                                            -WriteToHost `
                                            -Category 'Success'
                        }
            }
            
            It 'writes default category when -MessageFormat contains only {Category} field and -Category not specified' {
                TestMessageFormat `
                    -ExpectedLoggedText 'PROGRESS' `
                    -FunctionUnderTest `
                        { 
                            Write-LogMessage -Message 'hello world' `
                                            -MessageFormat '{Category}' `
                                            -WriteToHost 
                        }
            }
            
            It 'writes empty string when -MessageFormat contains only {Category} field and -Category not specified and no default category' {
                $script:_logConfiguration.CategoryInfo.Progress.IsDefault = $False
                TestMessageFormat `
                    -ExpectedLoggedText '' `
                    -FunctionUnderTest `
                        { 
                            Write-LogMessage -Message 'hello world' `
                                            -MessageFormat '{Category}' `
                                            -WriteToHost 
                        }
            }        
            
            It 'writes text to match all fields when -MessageFormat contains multiple fields' {
                TestMessageFormat `
                    -ExpectedLoggedText '^\d\d\d\d-\d\d-\d\d \d\d:\d\d:\d\d\.\d\d\d | INFORMATION | hello world$' `
                    -DoRegexMatch `
                    -FunctionUnderTest `
                        { 
                            Write-LogMessage -Message 'hello world' `
                                            -MessageFormat '{Timestamp} | {MessageLevel} | {Message}' `
                                            -WriteToHost 
                        }
            }           
            
            It 'uses configuration MessageFormat when -MessageFormat parameter not specified' {
                Private_SetMessageFormat '{Timestamp:hh:mm:ss.fff} | {Message}'
                TestMessageFormat `
                    -ExpectedLoggedText '^\d\d:\d\d:\d\d\.\d\d\d | hello world$' `
                    -DoRegexMatch `
                    -FunctionUnderTest `
                        { 
                            Write-LogMessage -Message 'hello world' `
                                            -WriteToHost 
                        }
            }
        }
    }
}
