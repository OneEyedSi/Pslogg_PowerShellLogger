<#
.SYNOPSIS
Tests of the exported Pslogg functions in the Pslogg module.

.DESCRIPTION
Pester tests of the logging functions exported from the Pslogg module.
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

        # Gets the MessageFormatInfo hashtable that matches the default configuration.
        function GetDefaultMessageFormatInfo ()
        {
            $messageFormatInfo = @{
                                    RawFormat = '{Timestamp:yyyy-MM-dd HH:mm:ss.fff} | {CallerName} | {MessageLevel} | {Message}'
                                    WorkingFormat = '$($Timestamp.ToString(''yyyy-MM-dd HH:mm:ss.fff'')) | ${CallerName} | ${MessageLevel} | ${Message}'
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
            $script:_logConfiguration.LogFile.FullPathReadOnly = GetResetLogFilePath
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
            
            Should -Invoke -CommandName Write-Host -Scope It -Times $timesWriteHostCalled
            Should -Invoke -CommandName Write-Error -Scope It -Times $timesWriteErrorCalled
            Should -Invoke -CommandName Write-Warning -Scope It -Times $timesWriteWarningCalled
            Should -Invoke -CommandName Write-Information -Scope It -Times $timesWriteInformationCalled
            Should -Invoke -CommandName Write-Debug -Scope It -Times $timesWriteDebugCalled
            Should -Invoke -CommandName Write-Verbose -Scope It -Times $timesWriteVerboseCalled             
        }

        function AssertWriteHostCalled ([string]$WithTextColor)
        {
            Should -Invoke -CommandName Write-Host `
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
            
            Should -Invoke -CommandName $commandName -Scope It -Times 1 `
                -ParameterFilter { $Path -eq $LogFilePath }
        }

        function AssertFileWriterNotCalled ()
        {
            Should -Invoke -CommandName Add-Content -Scope It -Times 0 
            Should -Invoke -CommandName Set-Content -Scope It -Times 0
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

        function TestMessageFormat
        (
            [scriptblock]$FunctionUnderTest, 
            [string]$ExpectedLoggedText,
            [switch]$DoRegexMatch
        )
        {
            Invoke-Command $FunctionUnderTest

            Should -Invoke -CommandName Write-Host -Scope It -Times 1 `
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
    }

    Describe 'Write-LogMessage' {  
        BeforeDiscovery {
            $logLevels = @{
                Off = 0
                Error = 1
                Warning = 2
                Information = 3
                Debug = 4
                Verbose = 5
            }
        }           

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

                Should -Invoke -CommandName Write-Host -Scope It -Times 1
            }
        }
            
        Context 'Parameter validation' {
            BeforeAll {
                Mock Write-Host
            }            

            It 'does not throw exception when -Message is $Null' {
                
                { Write-LogMessage -Message $Null } | 
                    Should -Not -Throw
            }

            It 'writes to log when -Message is $Null' {
                # Only want the message field logged, so Pslogg isn't logging any text at all.
                Private_SetMessageFormat '{Message}'

                Write-LogMessage -Message $Null

                Should -Invoke -CommandName Write-Host -Scope It -Times 1
            }

            It 'does not throw exception when -Message is empty string' {
                
                { Write-LogMessage -Message '' } | 
                    Should -Not -Throw
            }

            It 'writes to log when -Message is empty string' {
                # Only want the message field logged, so Pslogg isn't logging any text at all.
                Private_SetMessageFormat '{Message}'

                Write-LogMessage -Message ''

                Should -Invoke -CommandName Write-Host -Scope It -Times 1
            }

            It 'does not throw exception when Message passed by position' {
                { Write-LogMessage 'hello world' } | 
                    Should -Not -Throw                                     
            }

            It 'writes to log when Message passed by position' {
                # Only want the message field logged, so Pslogg isn't logging any text at all.
                Private_SetMessageFormat '{Message}'

                Write-LogMessage 'hello world'

                Should -Invoke -CommandName Write-Host -Scope It -Times 1
            }

            It 'throws exception when an invalid -HostTextColor is specified' {
                { Write-LogMessage -Message 'hello world' -HostTextColor DeepPurple } | 
                    Should -Throw -ExceptionType ([System.Management.Automation.ParameterBindingException]) `
                        -ExpectedMessage "*'DeepPurple' is not a valid text color*"                                  
            }

            It 'does not throw exception when no -HostTextColor is specified' {
                
                { Write-LogMessage -Message 'hello world' } | 
                    Should -Not -Throw
            }

            It 'does not throw exception when valid -HostTextColor is specified' {
                
                { Write-LogMessage -Message 'hello world' -HostTextColor Cyan } | 
                    Should -Not -Throw
            }

            It 'does not throw exception when neither -MessageLevel nor a message level switch are specified' {
                { Write-LogMessage 'hello world' } | 
                    Should -Not -Throw
            }

            It 'throws exception when an invalid -MessageLevel is specified' {
                { Write-LogMessage 'hello world' -MessageLevel INVALID } | 
                Should -Throw -ExceptionType ([System.Management.Automation.ParameterBindingException]) `
                    -ExpectedMessage "*Cannot validate argument on parameter 'MessageLevel'*"                                   
            }

            It 'does not throw exception when valid -MessageLevel is specified' {
                { Write-LogMessage 'hello world' -MessageLevel ERROR } | 
                    Should -Not -Throw
            }

            It 'throws exception when both -MessageLevel and -IsError are specified' {
                { Write-LogMessage 'hello world' -MessageLevel ERROR -IsError } | 
                Should -Throw -ExceptionType ([System.Management.Automation.ParameterBindingException]) `
                    -ExpectedMessage "*Parameter set cannot be resolved*" 
            }

            It 'throws exception when both -MessageLevel and -IsWarning are specified' {
                { Write-LogMessage 'hello world' -MessageLevel ERROR -IsWarning } | 
                Should -Throw -ExceptionType ([System.Management.Automation.ParameterBindingException]) `
                    -ExpectedMessage "*Parameter set cannot be resolved*"  
            }

            It 'throws exception when both -MessageLevel and -IsInformation are specified' {
                { Write-LogMessage 'hello world' -MessageLevel ERROR -IsInformation } | 
                Should -Throw -ExceptionType ([System.Management.Automation.ParameterBindingException]) `
                    -ExpectedMessage "*Parameter set cannot be resolved*"  
            }

            It 'throws exception when both -MessageLevel and -IsDebug are specified' {
                { Write-LogMessage 'hello world' -MessageLevel ERROR -IsDebug } | 
                Should -Throw -ExceptionType ([System.Management.Automation.ParameterBindingException]) `
                    -ExpectedMessage "*Parameter set cannot be resolved*"   
            }

            It 'throws exception when both -MessageLevel and -IsVerbose are specified' {
                { Write-LogMessage 'hello world' -MessageLevel ERROR -IsVerbose } | 
                Should -Throw -ExceptionType ([System.Management.Automation.ParameterBindingException]) `
                    -ExpectedMessage "*Parameter set cannot be resolved*"  
            }

            It 'throws exception when both -IsError and -IsWarning switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsWarning } | 
                Should -Throw -ExceptionType ([ArgumentException]) `
                    -ExpectedMessage "*Only one Message Level switch parameter may be set*"                                      
            }

            It 'throws exception when both -IsError and -IsInformation switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsInformation } | 
                Should -Throw -ExceptionType ([ArgumentException]) `
                    -ExpectedMessage "*Only one Message Level switch parameter may be set*"                                     
            }

            It 'throws exception when both -IsError and -IsDebug switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsDebug } | 
                Should -Throw -ExceptionType ([ArgumentException]) `
                    -ExpectedMessage "*Only one Message Level switch parameter may be set*"                                      
            }

            It 'throws exception when both -IsError and -IsVerbose switches set' {
                { Write-LogMessage -Message 'hello world' -IsError -IsVerbose } | 
                Should -Throw -ExceptionType ([ArgumentException]) `
                    -ExpectedMessage "*Only one Message Level switch parameter may be set*"                                     
            }

            It 'throws exception when both -IsWarning and -IsInformation switches set' {
                { Write-LogMessage -Message 'hello world' -IsWarning -IsInformation } | 
                Should -Throw -ExceptionType ([ArgumentException]) `
                    -ExpectedMessage "*Only one Message Level switch parameter may be set*"                                     
            }

            It 'throws exception when both -IsWarning and -IsDebug switches set' {
                { Write-LogMessage -Message 'hello world' -IsWarning -IsDebug } | 
                Should -Throw -ExceptionType ([ArgumentException]) `
                    -ExpectedMessage "*Only one Message Level switch parameter may be set*"                                       
            }

            It 'throws exception when both -IsWarning and -IsVerbose switches set' {
                { Write-LogMessage -Message 'hello world' -IsWarning -IsVerbose } | 
                Should -Throw -ExceptionType ([ArgumentException]) `
                    -ExpectedMessage "*Only one Message Level switch parameter may be set*"                                        
            }

            It 'throws exception when both -IsInformation and -IsDebug switches set' {
                { Write-LogMessage -Message 'hello world' -IsInformation -IsDebug } | 
                Should -Throw -ExceptionType ([ArgumentException]) `
                    -ExpectedMessage "*Only one Message Level switch parameter may be set*"                                      
            }

            It 'throws exception when both -IsInformation and -IsVerbose switches set' {
                { Write-LogMessage -Message 'hello world' -IsInformation -IsVerbose } | 
                Should -Throw -ExceptionType ([ArgumentException]) `
                    -ExpectedMessage "*Only one Message Level switch parameter may be set*"                                       
            }

            It 'throws exception when both -IsDebug and -IsVerbose switches set' {
                { Write-LogMessage -Message 'hello world' -IsDebug -IsVerbose } | 
                Should -Throw -ExceptionType ([ArgumentException]) `
                    -ExpectedMessage "*Only one Message Level switch parameter may be set*"                                      
            }

            It 'throws exception when both -WriteToHost and -WriteToStreams switches set' {
                { Write-LogMessage -Message 'hello world' -WriteToHost -WriteToStreams } | 
                Should -Throw -ExceptionType ([ArgumentException]) `
                    -ExpectedMessage "*Only one Destination switch parameter may be set*"                                       
            }
        }

        Context 'Logging to host and PowerShell streams' {
            BeforeAll {
                MockHostAndStreamWriters
            }

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

            It "writes to host in <_> text colour when -WriteToHost switch set, -MessageLevel set to <_> and -HostTextColor not specified" -ForEach @(
                'Error', 'Warning', 'Information', 'Debug', 'Verbose'
            ) {
                $textColour = 'DarkRed'
                $script:_logConfiguration.HostTextColor[$_] = $textColour
                $script:_logConfiguration.LogLevel = 'Verbose'
                    
                Write-LogMessage -Message 'hello world' -WriteToHost -MessageLevel $_

                AssertWriteHostCalled -WithTextColor $textColour
            } 

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
                
                Write-LogMessage -Message 'hello world' -WriteToHost -MessageLevel 'Information' `
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
                
                Write-LogMessage -Message 'hello world' -WriteToHost -MessageLevel 'Error' `
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
            BeforeAll {
                Mock Write-Host
                MockFileWriter
            }
            
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
                $script:_logConfiguration.LogFile.FullPathReadOnly = $logFileName
                
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
            BeforeAll {
                Mock Write-Host
            }
            
            BeforeEach {
                RemoveLogFile -Path $script:_logConfiguration.LogFile.FullPathReadOnly
            }

            It 'creates a log file when configuration LogFile.Name set and LogFile.Overwrite cleared, and log file does not exist' {
                $script:_logConfiguration.LogFile.Overwrite = $False
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Exist
            }

            It 'creates a log file when configuration LogFile.Name and LogFile.Overwrite set, and log file does not exist' {
                $script:_logConfiguration.LogFile.Overwrite = $True                
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Exist
            }

            It 'appends to existing log file when configuration LogFile.Overwrite cleared' {
                $script:_logConfiguration.LogFile.Overwrite = $False

                $originalContent = NewLogFile -Path $script:_logConfiguration.LogFile.FullPathReadOnly                
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Exist
                $newContent = (Get-Content -Path $script:_logConfiguration.LogFile.FullPathReadOnly)
                $newContent.Count | Should -Be 3
                $newContent[2] | Should -BeLike '*hello world*'               
                
                Write-LogMessage -Message 'second message' -WriteToHost

                $newContent = (Get-Content -Path $script:_logConfiguration.LogFile.FullPathReadOnly)
                $newContent.Count | Should -Be 4
                $newContent[3] | Should -BeLike '*second message*'
            }

            It 'overwrites an existing log file with first logged message when configuration LogFile.Overwrite set' {
                $script:_logConfiguration.LogFile.Overwrite = $True

                $originalContent = NewLogFile -Path $script:_logConfiguration.LogFile.FullPathReadOnly                
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Exist
                $newContent = ,(Get-Content -Path $script:_logConfiguration.LogFile.FullPathReadOnly)
                $newContent.Count | Should -Be 1
                $newContent[0] | Should -BeLike '*hello world*'
            }

            It 'appends subsequent messages to log file when configuration LogFile.Overwrite set' {
                $script:_logConfiguration.LogFile.Overwrite = $True

                $originalContent = NewLogFile -Path $script:_logConfiguration.LogFile.FullPathReadOnly                
                
                Write-LogMessage -Message 'hello world' -WriteToHost

                $script:_logConfiguration.LogFile.FullPathReadOnly | Should -Exist
                $newContent = ,(Get-Content -Path $script:_logConfiguration.LogFile.FullPathReadOnly)
                $newContent.Count | Should -Be 1
                $newContent[0] | Should -BeLike '*hello world*'               
                
                Write-LogMessage -Message 'second message' -WriteToHost

                $newContent = (Get-Content -Path $script:_logConfiguration.LogFile.FullPathReadOnly)
                $newContent.Count | Should -Be 2
                $newContent[1] | Should -BeLike '*second message*'               
                
                Write-LogMessage -Message 'third message' -WriteToHost

                $newContent = (Get-Content -Path $script:_logConfiguration.LogFile.FullPathReadOnly)
                $newContent.Count | Should -Be 3
                $newContent[2] | Should -BeLike '*third message*'
            }
        }  
        
        Context "Configured LogLevel is <_>" -ForEach (
            'Off', 'Error', 'Warning', 'Information', 'Debug', 'Verbose'
        ) {
            BeforeDiscovery {
                $logLevelConfigured = $_
            }
            BeforeAll {
                Mock Write-Host
            }

            It "does <not>write message when -Is<messageLevel> switch is set" -ForEach (
                @{ 
                    MessageLevel = 'Error'                  
                    MessageLevelValue = $logLevels['Error']
                    ConfiguredLogLevel = $logLevelConfigured
                    ConfiguredLevelValue = $logLevels[$logLevelConfigured]
                    Not = if ($logLevels['Error'] -gt $logLevels[$logLevelConfigured]) { 'not ' }
                 },
                @{ 
                    MessageLevel = 'Warning'                  
                    MessageLevelValue = $logLevels['Warning']
                    ConfiguredLogLevel = $logLevelConfigured
                    ConfiguredLevelValue = $logLevels[$logLevelConfigured]
                    Not = if ($logLevels['Warning'] -gt $logLevels[$logLevelConfigured]) { 'not ' }
                },
                @{
                    MessageLevel = 'Information'                  
                    MessageLevelValue = $logLevels['Information']
                    ConfiguredLogLevel = $logLevelConfigured
                    ConfiguredLevelValue = $logLevels[$logLevelConfigured]
                    Not = if ($logLevels['Information'] -gt $logLevels[$logLevelConfigured]) { 'not ' }
                },
                @{ 
                    MessageLevel = 'Debug'                  
                    MessageLevelValue = $logLevels['Debug']
                    ConfiguredLogLevel = $logLevelConfigured
                    ConfiguredLevelValue = $logLevels[$logLevelConfigured]
                    Not = if ($logLevels['Debug'] -gt $logLevels[$logLevelConfigured]) { 'not ' }
                },
                @{ 
                    MessageLevel = 'Verbose'                  
                    MessageLevelValue = $logLevels['Verbose']
                    ConfiguredLogLevel = $logLevelConfigured
                    ConfiguredLevelValue = $logLevels[$logLevelConfigured]
                    Not = if ($logLevels['Verbose'] -gt $logLevels[$logLevelConfigured]) { 'not ' }
                }
            ) {
                $script:_logConfiguration.LogLevel = $ConfiguredLogLevel

                $numberTimesShouldWrite = 1
                if ($MessageLevelValue -gt $ConfiguredLevelValue)
                {
                    $numberTimesShouldWrite = 0
                }
                
                $isError = $False
                $isWarning = $False
                $isInformation = $False
                $isDebug = $False
                $isVerbose = $False

                switch ($MessageLevel)
                {
                    'Error' { 
                        $isError = $True 
                        break 
                    }
                    'Warning'	{ 
                        $isWarning = $True 
                        break 
                    }
                    'Information'	{ 
                        $isInformation = $True 
                        break 
                    }
                    'Debug' { 
                        $isDebug = $True 
                        break 
                    }
                    'Verbose' { 
                        $isVerbose = $True 
                        break 
                    }
                }
                Write-LogMessage -Message 'hello world' -WriteToHost `
                    -IsError:$isError -IsWarning:$isWarning -IsInformation:$isInformation `
                    -IsDebug:$isDebug -IsVerbose:$isVerbose 

                Should -Invoke -CommandName Write-Host -Times $numberTimesShouldWrite
            }

            It "does <not>write message when MessageLevel is <messageLevel>" -ForEach (
                @{ 
                    MessageLevel = 'Error'                  
                    MessageLevelValue = $logLevels['Error']
                    ConfiguredLogLevel = $logLevelConfigured
                    ConfiguredLevelValue = $logLevels[$logLevelConfigured]
                    Not = if ($logLevels['Error'] -gt $logLevels[$logLevelConfigured]) { 'not ' }
                 },
                @{ 
                    MessageLevel = 'Warning'                  
                    MessageLevelValue = $logLevels['Warning']
                    ConfiguredLogLevel = $logLevelConfigured
                    ConfiguredLevelValue = $logLevels[$logLevelConfigured]
                    Not = if ($logLevels['Warning'] -gt $logLevels[$logLevelConfigured]) { 'not ' }
                },
                @{
                    MessageLevel = 'Information'                  
                    MessageLevelValue = $logLevels['Information']
                    ConfiguredLogLevel = $logLevelConfigured
                    ConfiguredLevelValue = $logLevels[$logLevelConfigured]
                    Not = if ($logLevels['Information'] -gt $logLevels[$logLevelConfigured]) { 'not ' }
                },
                @{ 
                    MessageLevel = 'Debug'                  
                    MessageLevelValue = $logLevels['Debug']
                    ConfiguredLogLevel = $logLevelConfigured
                    ConfiguredLevelValue = $logLevels[$logLevelConfigured]
                    Not = if ($logLevels['Debug'] -gt $logLevels[$logLevelConfigured]) { 'not ' }
                },
                @{ 
                    MessageLevel = 'Verbose'                  
                    MessageLevelValue = $logLevels['Verbose']
                    ConfiguredLogLevel = $logLevelConfigured
                    ConfiguredLevelValue = $logLevels[$logLevelConfigured]
                    Not = if ($logLevels['Verbose'] -gt $logLevels[$logLevelConfigured]) { 'not ' }
                }
            ) {
                $script:_logConfiguration.LogLevel = $ConfiguredLogLevel

                $numberTimesShouldWrite = 1
                if ($MessageLevelValue -gt $ConfiguredLevelValue)
                {
                    $numberTimesShouldWrite = 0
                }

                Write-LogMessage -Message 'hello world' -WriteToHost `
                    -MessageLevel $MessageLevel 

                Should -Invoke -CommandName Write-Host -Times $numberTimesShouldWrite
            }
        }

        Context 'Message Format' {
            BeforeAll {
                Mock Write-Host
            }
            
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
                                            -MessageFormat '{Timestamp:HH:mm:ss}' `
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
                Private_SetMessageFormat '{Timestamp:HH:mm:ss.fff} | {Message}'
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
