<#
.SYNOPSIS
Function wrapper that tests the exceptions a function should generate.

.DESCRIPTION
Function wrapper that tests the exceptions a function should generate.  "Should -Throw" will 
only check the exception message, not the type of exception.  This function can test both 
exception message and type.

.NOTES
Copyright:		(c) 2018 Simon Elms
Requires:		PowerShell 5 (may work on earlier versions but untested)
				Pester 4 (may work on earlier versions but untested)
Version:		1.0.0 
Date:			22 Feb 2018

Usage is similar to that for Pester "Should -Throw": Wrap the function under test, along with 
any arguments, in curly braces and pipe it to Assert-ExceptionThrown.

.PARAMETER FunctionScriptBlock
A call to the function under test, including any arguments, wrapped in curly braces to form a 
scriptblock. 

.PARAMETER WithTypeName
The type name of the exception that is expected to be thrown by the function under test.  

The test passes if the type name specified via -WithTypeName matches the end of the full type 
name of the exception that is thrown.  This allows leading namespaces to be left out of the 
expected type name.  

So, for example, if the function under test throws a 
System.Management.Automation.ActionPreferenceStopException, 
the following will pass:
    -WithTypeName 'System.Management.Automation.ActionPreferenceStopException' 
    -WithTypeName 'Management.Automation.ActionPreferenceStopException' 
    -WithTypeName 'Automation.ActionPreferenceStopException'
    -WithTypeName 'ActionPreferenceStopException'

The test will fail if truncated namespaces or class names are specified.  For example, the 
following will fail:
    -WithTypeName 'mation.ActionPreferenceStopException' (truncated namespace 'Automation')
    -WithTypeName 'System.Management.Automation.ActionPref' (truncated class name 
                                                            'ActionPreferenceStopException')
    -WithTypeName 'StopException' (truncated class name 'ActionPreferenceStopException')

The comparison between the expected and actual exception type names is case insensitive.

.PARAMETER WithMessage
All or part of the exception message that is expected when the function under test is run. 
 
The test is effectively "Actual exception message must contain WithMessage".  The 
comparison between the expected and actual exception messages is case insensitive.

.PARAMETER Not
A switch parameter that inverts the test.  

The effect of the -Not parameter depends on whether -WithTypeName or 
-WithMessage are set.  If neither -WithTypeName nor 
-WithMessage are set then -Not means "Function should not throw any exception."  

If either -WithTypeName and/or -WithMessage are set then -Not means 
"Function should not throw an exception with the specified class name and/or message."  This 
means the test will pass if no exception is thrown, or if an exception is thrown which does not 
have the specified class name and/or message.

.EXAMPLE 
Test that a function taking two arguments throws an exception with a specified message

{ MyFunctionWithTwoArgs -Key 'header' -Value 10 } | 
    Assert-ExceptionThrown -WithMessage 'Value was of type int32, expected string'

The name of the function under test is MyFunctionWithTwoArgs.  The test will only pass if 
MyFunctionWithTwoArgs, with the specified arguments, throws an exception with a message 
that contains the specified text.

.EXAMPLE 
Test that a function taking no arguments throws an exception of a specified type

{ MyFunction } | 
    Assert-ExceptionThrown -WithTypeName System.ArgumentException

The test will only pass if MyFunction throws an System.ArgumentException.

.EXAMPLE 
Specify a short type name, without namespace, for the expected exception

{ MyFunction } | 
    Assert-ExceptionThrown -WithTypeName ArgumentException

The test will pass if MyFunction throws a System.ArgumentException.

.EXAMPLE 
Test that a function does not throw an exception

{ MyFunctionWithTwoArgs -Key 'header' -Value 'value' } | 
    Assert-ExceptionThrown -Not

The test will pass only if MyFunctionWithTwoArgs, with the specified arguments, does not throw 
any exception.

.EXAMPLE 
Test that a function does not throw an exception with a specified message

{ MyFunctionWithTwoArgs -Key 'header' -Value 10 } | 
    Assert-ExceptionThrown -Not -WithMessage 'Value was of type int32, expected string'

The test will fail if MyFunctionWithTwoArgs, with the specified arguments, throws an exception 
with a message that contains the specified text.  It will pass if MyFunctionWithTwoArgs does not 
throw an exception, or if it throws an exception with a different message.

.EXAMPLE 
Test that a function does not throw an exception of a specified type

{ MyFunction } | 
    Assert-ExceptionThrown -Not -WithTypeName ArgumentException

The test will fail if MyFunction throws a System.ArgumentException.  It will pass if MyFunction 
does not throw an exception, or if it throws an exception of a different type.

.LINK
https://github.com/AnotherSadGit/PesterAssertExceptionThrown
#>
function Assert-ExceptionThrown 
(
    [Parameter(
        Position=0, 
        ValueFromPipeline=$true)
    ]
    [scriptblock]$FunctionScriptBlock, 

    [string]$WithTypeName,
    [string]$WithMessage, 
    [switch]$UseFullTypeName, 

    [switch]$Not
)
{
    process
    {
        if ($FunctionScriptBlock -eq $Null)
        {
            throw [ArgumentNullException] `
                "Script block for function under test not found. Input to 'Assert-ExceptionThrown' must be enclosed in curly braces."
        }

        try
        {
            Invoke-Command -ScriptBlock $FunctionScriptBlock
        }
        catch
        {
            $errorMessages = Private_GetExceptionError -Exception $_.Exception `
                -WithTypeName $WithTypeName `
                -WithMessage $WithMessage `
                -Not:$Not
                
            if ($errorMessages.Count -gt 0)
            {
                throw [System.Exception] ($errorMessages -join [Environment]::NewLine)
            }            

            return
        }

        # Will only get here if no exception was thrown by the function under test...
        
        if ([string]::IsNullOrWhiteSpace($WithTypeName) -and 
            [string]::IsNullOrWhiteSpace($WithMessage))
        {
            # No expectations were specified for the exception type or the exception message.  
            # So in this case -Not means "Should not throw any exception."
            if ($Not)
            {
                return
            }
            
            # No exception expectations were specified and -Not was also not specified.  This 
            # means "Should throw an exception".
            throw [System.Exception] "Expected an exception to be thrown but none was."
        }

        # If any exception expectation was specified then -Not means "Should not throw an 
        # exception of the specified type and/or with the specified message, as appropriate."
        # So not throwing an exception is a pass. 
        if ($Not)
        {
            return
        }

        $errorMessages = @()

        if (-not [string]::IsNullOrWhiteSpace($WithTypeName))
        {
            $errorMessages += 
                "Expected $WithTypeName to be thrown but it wasn't."
        }

        if (-not [string]::IsNullOrWhiteSpace($WithMessage))
        {
            $errorMessages += 
                "Expected exception with message '$WithMessage' to be thrown but it wasn't."
        }

        if ($errorMessages.Count -gt 0)
        {
            throw [System.Exception] ($errorMessages -join [Environment]::NewLine)
        }
    }
}

<#
.SYNOPSIS
Gets an array of error messages where an exception does not meet expectations.

.DESCRIPTION
Gets an array of error messages where an exception does not meet expectations.  If all 
expectations are met the array will be empty.
#>
function Private_GetExceptionError
(
    [Exception]$Exception,

    [string]$WithTypeName,
    [string]$WithMessage, 

    [switch]$Not
)
{
    if ([string]::IsNullOrWhiteSpace($WithTypeName) -and 
        [string]::IsNullOrWhiteSpace($WithMessage))
    {
        # No expectations were specified for the exception type or the exception message.  
        # So in this case -Not means "Should not throw any exception."
        if ($Not)
        {
            # Note leading comma to turn this error message into an array with a single item.
            return ,"Expected no exception but $($Exception.GetType().FullName) was thrown with message '$($Exception.Message)'."
        }
            
        # No exception expectations were specified and -Not was also not specified.  This 
        # means "Should throw an exception, any exception".
        return @()
    }

    $errorMessages = @()
    $exceptionTypeMatched = $True
    $exceptionMessageMatched = $True

    if (-not [string]::IsNullOrWhiteSpace($WithTypeName))
    {
        $exceptionTypeMatched = $False
        $actualExceptionTypeName = $Exception.GetType().FullName
        
        if ($actualExceptionTypeName.EndsWith($WithTypeName.Trim(), 
                                                [StringComparison]::CurrentCultureIgnoreCase))
        {
            # Don't allow silliness like 
            #   Assert-ExceptionThrown -WithTypeName 'stem.ArgumentException'
            # or 
            #   Assert-ExceptionThrown -WithTypeName 'tion'
            # to pass.  Although namespaces can be left out of the expected type name, we don't 
            # want truncated namespaces or truncated class names.

            $actualTypeNameParts = $actualExceptionTypeName -split '.', 0, 'simplematch'
            $expectedTypeNameParts = $WithTypeName.Trim() -split '.', 0, 'simplematch'
            if ($actualTypeNameParts.Count -lt $expectedTypeNameParts.Count)
            {
                $exceptionTypeMatched = $False
            }
            else
            {
                # To handle leading parts of the namespace being missed from the expected type 
                # name, compare the parts of the type name in reverse.
                [array]::Reverse($actualTypeNameParts)
                [array]::Reverse($expectedTypeNameParts)
                
                $exceptionTypeMatched = $True
                for ($i = 0; $i -lt $expectedTypeNameParts.Count; $i++) 
                {
                    # -ine is case insensitive not equals.
                    if ($actualTypeNameParts[$i] -ine $expectedTypeNameParts[$i])
                    {
                        $exceptionTypeMatched = $False
                        break
                    }
                }
            }
        }
        if ($Not -and $exceptionTypeMatched)
        {
            $errorMessages += `
                "Expected an exception of a different type than $WithTypeName but exception thrown was of that type."
        }
        elseif (-not $Not -and -not $exceptionTypeMatched)
        {
            $errorMessages += `
                "Expected $WithTypeName but exception thrown was $actualExceptionTypeName."
        }

    }

    if (-not [string]::IsNullOrWhiteSpace($WithMessage))
    {
        # -ilike is case insensitive like.
        $exceptionMessageMatched = ($Exception.Message -ilike "*$WithMessage*")
        if ($Not -and $exceptionMessageMatched)
        {
            $errorMessages += `
                "Expected exception message different than '$WithMessage' but exception thrown had that message."
        }
        elseif (-not $Not -and -not $exceptionMessageMatched)
        {
            $errorMessages += `
                "Expected exception message '$WithMessage' but actual exception message was '$($Exception.Message)'."
        }
    }

    if ($Not)
    {
        if ($exceptionTypeMatched -and $exceptionMessageMatched)
        {
            return $errorMessages
        }
        else
        {
            return @()
        }
    }
    else
    {
        if ($exceptionTypeMatched -and $exceptionMessageMatched)
        {
            return @()
        }
        else
        {
            return $errorMessages
        }
    }
}