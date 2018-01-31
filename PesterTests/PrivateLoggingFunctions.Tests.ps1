<#
.SYNOPSIS
Tests of the private logging functions in the Logging module.

.DESCRIPTION
Pester tests of the private functions in the Logging module that are called by the public 
logging functions.
#>

# PowerShell allows multiple modules of the same name to be imported from different locations.  
# This would confuse Pester.  So, to be sure there are not multiple Logging modules imported, 
# remove all Logging modules and re-import only one.
Get-Module Logging | Remove-Module -Force
Import-Module ..\Modules\Logging\Logging.psm1 -Force

InModuleScope Logging {

    Describe 'GetCallingFunctionName' {     

        Context 'Get-PSCallStack mocked to return $Null' {
            Mock Get-PSCallStack { return $Null }

            It 'returns "[UNKNOWN CALLER]"' {
                Private_GetCallingFunctionName | Should -Be "[UNKNOWN CALLER]"          
            } 
        }    

        Context 'Get-PSCallStack mocked to return empty collection' {
            Mock Get-PSCallStack { return @() }

            It 'returns "[UNKNOWN CALLER]"' {
                Private_GetCallingFunctionName | Should -Be "[UNKNOWN CALLER]"          
            } 
        }  

        Context 'Get-PSCallStack mocked to return single stack frame' {

            Mock Get-PSCallStack { 
                $callStack = @()
                $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Logging.psm1'; FunctionName='Private_GetCallingFunctionName' }
                $callStack += $stackFrame
                return $callStack
            }

            It 'returns "----"' {
                Private_GetCallingFunctionName | Should -Be "----"
            } 
        }  

        Context 'Get-PSCallStack mocked so second stack frame has no file name' {

            Mock Get-PSCallStack { 
                $callStack = @()
                $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Logging.psm1'; FunctionName='Private_GetCallingFunctionName' }
                $callStack += $stackFrame
                $stackFrame = New-Object PSObject -Property @{ ScriptName=$Null; FunctionName=$Null }
                $callStack += $stackFrame
                return $callStack
            }

            It 'returns "[CONSOLE]"' {
                Private_GetCallingFunctionName | Should -Be "[CONSOLE]"          
            } 
        }   

        Context 'Get-PSCallStack mocked so second stack frame has no script name' {

            Mock Get-PSCallStack { 
                $callStack = @()
                $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Logging.psm1'; FunctionName='Private_GetCallingFunctionName' }
                $callStack += $stackFrame
                $stackFrame = New-Object PSObject -Property @{ ScriptName=$Null; FunctionName='<ScriptBlock>' }
                $callStack += $stackFrame
                return $callStack
            }

            It 'returns "[CONSOLE]"' {
                Private_GetCallingFunctionName | Should -Be "[CONSOLE]"          
            } 
        }  

        Context 'Get-PSCallStack mocked so second stack frame function name is <ScriptBlock>' {

            Mock Get-PSCallStack { 
                $callStack = @()
                $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Logging.psm1'; FunctionName='Private_GetCallingFunctionName' }
                $callStack += $stackFrame
                $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Test.ps1'; FunctionName='<ScriptBlock>' }
                $callStack += $stackFrame
                return $callStack
            }

            It 'returns "Script <script name>"' {
                Private_GetCallingFunctionName | Should -Be "Script Test.ps1"
            } 
        } 

        Context 'Get-PSCallStack mocked so second stack frame has both script name and function name' {

            Mock Get-PSCallStack { 
                $callStack = @()
                $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Logging.psm1'; FunctionName='Private_GetCallingFunctionName' }
                $callStack += $stackFrame
                $stackFrame = New-Object PSObject -Property @{ ScriptName='C:\Test\Test.ps1'; FunctionName='TestFunction' }
                $callStack += $stackFrame
                return $callStack
            }

            It 'returns "<function name>"' {
                Private_GetCallingFunctionName | Should -Be "TestFunction"
            } 
        }
    }
}
