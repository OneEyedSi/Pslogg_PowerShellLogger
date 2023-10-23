function TestFunction ()
{
    Write-LogMessage 'Inside function TestFunction, using configuration MessageFormat, Line 3'
    Write-LogMessage 'Inside function TestFunction, using -MessageFormat, Line 4' -MessageFormat '{CallerName}, {CallerLineNumber}, {Message}'
}

function IntermediateFunction ()
{
    Write-LogMessage 'Inside function IntermediateFunction, using configuration MessageFormat, Line 9'
    Write-LogMessage 'Inside function IntermediateFunction, using -MessageFormat, Line 10' -MessageFormat '{CallerName}, {CallerLineNumber}, {Message}'
    TestFunction
}

Set-LogConfiguration -MessageFormat '{Timestamp:yyyy-MM-dd hh:mm:ss.fff} | {CallerName} | {CallerLineNumber} | {Category} | {MessageLevel} | {Message}'
Write-LogMessage 'In script root, using configuration MessageFormat, Line 15'
Write-LogMessage 'In script root, using -MessageFormat, Line 16' -MessageFormat '{CallerName}, {CallerLineNumber}, {Message}'

TestFunction
IntermediateFunction

# Also run another manual test, calling Write-LogMessage from the PowerShell console:
# {CallerName} should be set to [CONSOLE]
# {CallerLineNumber} should be set to [NONE]
<#
Set-LogConfiguration -MessageFormat '{Timestamp:yyyy-MM-dd hh:mm:ss.fff} | {CallerName} | {CallerLineNumber} | {Category} | {MessageLevel} | {Message}'
Write-LogMessage 'This is a test' -DoNotWriteToFile

Reset-LogConfiguration
Write-LogMessage 'This is a test' -DoNotWriteToFile

Write-LogMessage 'This is a test' -DoNotWriteToFile -MessageFormat '{Timestamp:yyyy-MM-dd hh:mm:ss.fff} | {CallerName} | {CallerLineNUmber} | {Category} | {MessageLevel} | {Message}'
#>
