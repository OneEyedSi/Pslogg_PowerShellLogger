<#
.SYNOPSIS
Tests of the private logging functions in the Pslogg module.

.DESCRIPTION
Pester tests of the private functions in the Pslogg module that are called by the public 
logging functions.
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

    Describe 'GetCallerName' {     

        Context 'Get-PSCallStack mocked to return $Null' {
            BeforeAll {
                Mock Get-PSCallStack { return $Null }
            }

            It 'returns "[UNKNOWN CALLER]"' {
                Private_GetCallerName | Should -Be "[UNKNOWN CALLER]"          
            } 
        }    

        Context 'Get-PSCallStack mocked to return empty collection' {
            BeforeAll {
                Mock Get-PSCallStack { return @() }
            }

            It 'returns "[UNKNOWN CALLER]"' {
                Private_GetCallerName | Should -Be "[UNKNOWN CALLER]"          
            } 
        }  

        Context 'Get-PSCallStack mocked to return single stack frame' {
            BeforeAll {
                Mock Get-PSCallStack { 
                    $callStack = @()
                    $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Pslogg.psd1'; FunctionName='Private_GetCallerName' }
                    $callStack += $stackFrame
                    return $callStack
                }
            }

            It 'returns "----"' {
                Private_GetCallerName | Should -Be "----"
            } 
        }  

        Context 'Get-PSCallStack mocked so second stack frame has no file name' {
            BeforeAll {
                Mock Get-PSCallStack { 
                    $callStack = @()
                    $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Pslogg.psd1'; FunctionName='Private_GetCallerName' }
                    $callStack += $stackFrame
                    $stackFrame = New-Object PSObject -Property @{ ScriptName=$Null; FunctionName=$Null }
                    $callStack += $stackFrame
                    return $callStack
                }
            }

            It 'returns "[CONSOLE]"' {
                Private_GetCallerName | Should -Be "[CONSOLE]"          
            } 
        }   

        Context 'Get-PSCallStack mocked so second stack frame has no script name' {
            BeforeAll {
                Mock Get-PSCallStack { 
                    $callStack = @()
                    $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Pslogg.psd1'; FunctionName='Private_GetCallerName' }
                    $callStack += $stackFrame
                    $stackFrame = New-Object PSObject -Property @{ ScriptName=$Null; FunctionName='<ScriptBlock>' }
                    $callStack += $stackFrame
                    return $callStack
                }
            }

            It 'returns "[CONSOLE]"' {
                Private_GetCallerName | Should -Be "[CONSOLE]"          
            } 
        }  

        Context 'Get-PSCallStack mocked so second stack frame function name is `<ScriptBlock`>' {
            BeforeAll {
                Mock Get-PSCallStack { 
                    $callStack = @()
                    $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Pslogg.psd1'; FunctionName='Private_GetCallerName' }
                    $callStack += $stackFrame
                    $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Test.ps1'; FunctionName='<ScriptBlock>' }
                    $callStack += $stackFrame
                    return $callStack
                }
            }

            It 'returns "Script [script name]"' {
                Private_GetCallerName | Should -Be "Script Test.ps1"
            } 
        } 

        Context 'Get-PSCallStack mocked so second stack frame has both script name and function name' {
            BeforeAll {
                    Mock Get-PSCallStack { 
                    $callStack = @()
                    $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Pslogg.psd1'; FunctionName='Private_GetCallerName' }
                    $callStack += $stackFrame
                    $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Test.ps1'; FunctionName='TestFunction' }
                    $callStack += $stackFrame
                    return $callStack
                }
            }

            It 'returns "[function name]"' {
                Private_GetCallerName | Should -Be "TestFunction"
            } 
        }
    }

    Describe 'ShouldWriteToFile' {
        BeforeEach {
            Reset-LogConfiguration
        }

        Context 'LogFile configuration partially or completely missing' {

            It 'returns $False when no configuration hashtable exists' {
                $script:_logConfiguration = $null
                Private_ShouldWriteToFile -CallerName 'Somescript.ps1' | Should -BeFalse
            }

            It 'returns $False when configuration is missing LogFile settings' {
                $script:_logConfiguration.LogFile = $null
                Private_ShouldWriteToFile -CallerName 'Somescript.ps1' | Should -BeFalse
            }

            It 'returns $False when LogFile.Name configuration is missing' {
                $script:_logConfiguration.LogFile.Remove('Name')
                Private_ShouldWriteToFile -CallerName 'Somescript.ps1' | Should -BeFalse
            }

            It 'returns $False when called from script and LogFile.WriteFromScript configuration is missing' {
                $script:_logConfiguration.LogFile.Remove('WriteFromScript')
                Private_ShouldWriteToFile -CallerName 'Somescript.ps1' | Should -BeFalse
            }

            It 'returns $False when called from PowerShell console and LogFile.WriteFromHost configuration is missing' {
                $script:_logConfiguration.LogFile.Remove('WriteFromHost')
                Private_ShouldWriteToFile -CallerName $script:_constCallerConsole | Should -BeFalse
            }

            It 'returns $False when called from unknown source and LogFile.WriteFromHost configuration is missing' {
                $script:_logConfiguration.LogFile.Remove('WriteFromHost')
                Private_ShouldWriteToFile -CallerName $script:_constCallerUnknown | Should -BeFalse
            }

            It 'returns $False when called from unknown source function and LogFile.WriteFromHost configuration is missing' {
                $script:_logConfiguration.LogFile.Remove('WriteFromHost')
                Private_ShouldWriteToFile -CallerName $script:_constCallerFunctionUnknown | Should -BeFalse
            }
        }

        Context 'Called from <CallerDescription>' -ForEach @(
            @{ CallerDescription = 'PowerShell console'; CallSource = $script:_constCallerConsole }
            @{ CallerDescription = 'Unknown caller'; CallSource = $script:_constCallerUnknown }
            @{ CallerDescription = 'Unknown function'; CallSource = $script:_constCallerFunctionUnknown }
        ) {
            BeforeEach {                
                $script:_logConfiguration.LogFile.WriteFromHost = $True
            }
            
            It 'returns $False when LogFile.WriteFromHost configuration is $False' {
                $script:_logConfiguration.LogFile.WriteFromHost = $False
                Private_ShouldWriteToFile -CallerName $CallSource | Should -BeFalse                
            }
            
            It 'returns $False when LogFile.Name configuration is $Null' {
                $script:_logConfiguration.LogFile.Name = $Null
                Private_ShouldWriteToFile -CallerName $CallSource | Should -BeFalse                
            }
            
            It 'returns $False when LogFile.Name configuration is empty string' {
                $script:_logConfiguration.LogFile.Name = ''
                Private_ShouldWriteToFile -CallerName $CallSource | Should -BeFalse                
            }
            
            It 'returns $False when LogFile.Name configuration is blank string' {
                $script:_logConfiguration.LogFile.Name = '  '
                Private_ShouldWriteToFile -CallerName $CallSource | Should -BeFalse                
            }
            
            It 'returns $True when LogFile.WriteFromHost configuration is $True and LogFile.Name is set' {
                $script:_logConfiguration.LogFile.Name = 'Test.log'
                Private_ShouldWriteToFile -CallerName $CallSource | Should -BeTrue                
            }
        }

        Context 'Called from script file' {
            BeforeEach {                
                $script:_logConfiguration.LogFile.WriteFromScript = $True
                $callerName = 'Somescript.ps1'
            }
            
            It 'returns $False when LogFile.WriteFromScript configuration is $False' {
                $script:_logConfiguration.LogFile.WriteFromScript = $False
                Private_ShouldWriteToFile -CallerName $callerName | Should -BeFalse                
            }
            
            It 'returns $False when LogFile.Name configuration is $Null' {
                $script:_logConfiguration.LogFile.Name = $Null
                Private_ShouldWriteToFile -CallerName $callerName | Should -BeFalse                
            }
            
            It 'returns $False when LogFile.Name configuration is empty string' {
                $script:_logConfiguration.LogFile.Name = ''
                Private_ShouldWriteToFile -CallerName $callerName | Should -BeFalse                
            }
            
            It 'returns $False when LogFile.Name configuration is blank string' {
                $script:_logConfiguration.LogFile.Name = '  '
                Private_ShouldWriteToFile -CallerName $callerName | Should -BeFalse                
            }
            
            It 'returns $True when LogFile.WriteFromScript configuration is $True and LogFile.Name is set' {
                $script:_logConfiguration.LogFile.Name = 'Test.log'
                Private_ShouldWriteToFile -CallerName $callerName | Should -BeTrue                
            }
        }

        Context 'Log file path validity' {     
            BeforeEach {                
                $script:_logConfiguration.LogFile.WriteFromScript = $True
                $callerName = 'Somescript.ps1'
            }       

            It 'returns $False when configuration LogFile.Name not a valid path' {
                $logFileName = 'CC:\Test\Test.log'
                # This scenario should never occur.  Pslogg should throw an exception when setting 
                # LogFileName to an invalid path via Set-LogConfiguration.
                $script:_logConfiguration.LogFile.Name = $logFileName
                $script:_logConfiguration.LogFile.FullPathReadOnly = $logFileName
                
                Private_ShouldWriteToFile -CallerName $callerName | Should -BeFalse    
            }      

            It 'returns $True when configuration LogFile.Name a valid path' {
                $logFileName = 'C:\Test\Test.log'
                $script:_logConfiguration.LogFile.Name = $logFileName
                $script:_logConfiguration.LogFile.FullPathReadOnly = $logFileName
                
                Private_ShouldWriteToFile -CallerName $callerName | Should -BeTrue    
            }
        }
    }
}
