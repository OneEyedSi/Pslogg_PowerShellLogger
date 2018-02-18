<#
.SYNOPSIS
Function wrapper that tests the exceptions a function should generate.

.DESCRIPTION
Function wrapper that tests the exceptions a function should generate.  "Should -Throw" will 
only check the exception message, not the type of exception.  This function can test both 
exception message and type.

.NOTES
Usage is similar to that for Pester "Should -Throw": Wrap the function under test, along with 
any arguments, in curly braces and pipe it to Assert-ExceptionThrown.

.PARAMETER FunctionScriptBlock
The function under test, including any arguments, wrapped in curly braces to form a scriptblock. 

.PARAMETER ExpectedExceptionTypeName
The type name of the exception that is expected to be thrown by the function under test.  

By default this is the short name of the exception class, without a namespace.  If the 
-UseFullTypeName switch is set then the full name of the exception class, including namespace, 
must be specified.  The comparison between the expected and actual exception class names is 
case insensitive.

.PARAMETER ExpectedExceptionMessage
All or part of the exception message that is expected when the function under test is run. 
 
The test is effectively "Actual exception message must contain ExpectedExceptionMessage".  The 
comparison between the expected and actual exception messages is case insensitive.

.PARAMETER UseFullTypeName
A switch parameter that affects the behaviour of the ExpectedExceptionTypeName comparison.  

If UseFullTypeName is not set then the short name of the actual exception class, excluding any 
namespace, is compared against ExpectedExceptionTypeName.  If UseFullTypeName is set then the 
full name of the actual exception class, including namespace, is compared against 
ExpectedExceptionTypeName.

If ExpectedExceptionTypeName is not specified or is $Null or is an empty or blank string then 
UseFullTypeName is ignored.

.PARAMETER Not
A switch parameter that inverts the test when set.  

The effect of the Not parameter depends on whether ExpectedExceptionTypeName or 
ExpectedExceptionMessage are set.  If neither ExpectedExceptionTypeName nor 
ExpectedExceptionMessage are set then Not means "Function should not throw any exception."  

If either ExpectedExceptionTypeName and/or ExpectedExceptionMessage are set then Not means 
"Function should not throw an exception with the specified class name and/or message."  This 
means the test will pass if no exception is thrown, or if an exception is thrown which does not 
have the specified class name and/or message.

.EXAMPLE 
Test that a function taking two arguments throws an exception with a specified message

{ MyFunctionWithTwoArgs -Key MyKey -Value 10 } | 
    Assert-ExceptionThrown -ExpectedExceptionMessage 'Value was of type int32, expected string'

The name of the function under test is MyFunctionWithTwoArgs.  The test will only pass if 
MyFunctionWithTwoArgs, with the specified arguments, throws an exception with a message 
that contains the specified text.

.EXAMPLE 
Test that a function taking no arguments throws an exception of a specified type

{ MyFunction } | 
    Assert-ExceptionThrown -ExpectedExceptionTypeName ArgumentException

The test will only pass if MyFunction throws an ArgumentException.

.EXAMPLE 
Test that a function throws an exception with the type name specified in full

{ MyFunction } | 
    Assert-ExceptionThrown -ExpectedExceptionTypeName System.ArgumentException `
                            -UseFullTypeName

The test will only pass if MyFunction throws a System.ArgumentException.

.EXAMPLE 
Test that a function does not throw an exception

{ MyFunctionWithTwoArgs -Key MyKey -Value MyValue } | 
    Assert-ExceptionThrown -Not

The test will pass only if MyFunctionWithTwoArgs, with the specified arguments, does not throw 
any exception.

.EXAMPLE 
Test that a function does not throw an exception with a specified message

{ MyFunctionWithTwoArgs -Key MyKey -Value 10 } | 
    Assert-ExceptionThrown -Not -ExpectedExceptionMessage 'Value was of type int32, expected string'

The test will fail if MyFunctionWithTwoArgs, with the specified arguments, throws an exception 
with a message that contains the specified text.  It will pass if MyFunctionWithTwoArgs does not 
throw an exception, or if it throws an exception with a different message.

.EXAMPLE 
Test that a function does not throw an exception of a specified type

{ MyFunction } | 
    Assert-ExceptionThrown -Not -ExpectedExceptionTypeName ArgumentException

The test will fail if MyFunction throws an ArgumentException.  It will pass if MyFunction does 
not throw an exception, or if it throws an exception of a different type.
#>
function Assert-ExceptionThrown 
(
    [Parameter(
        Position=0, 
        ValueFromPipeline=$true)
    ]
    [scriptblock]$FunctionScriptBlock, 

    [string]$ExpectedExceptionTypeName,
    [string]$ExpectedExceptionMessage, 
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
            $errorMessages = Get-ExceptionError -Exception $_.Exception `
                -ExpectedExceptionTypeName $ExpectedExceptionTypeName `
                -ExpectedExceptionMessage $ExpectedExceptionMessage `
                -UseFullTypeName:$UseFullTypeName -Not:$Not
                
            if ($errorMessages.Count -gt 0)
            {
                throw [System.Exception] ($errorMessages -join [Environment]::NewLine)
            }            

            return
        }

        # Will only get here if no exception was thrown by the function under test...
        
        if ([string]::IsNullOrWhiteSpace($ExpectedExceptionTypeName) -and 
            [string]::IsNullOrWhiteSpace($ExpectedExceptionMessage))
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

        if (-not [string]::IsNullOrWhiteSpace($ExpectedExceptionTypeName))
        {
            $errorMessages += 
                "Expected $ExpectedExceptionTypeName to be thrown but it wasn't."
        }

        if (-not [string]::IsNullOrWhiteSpace($ExpectedExceptionMessage))
        {
            $errorMessages += 
                "Expected exception with message '$ExpectedExceptionMessage' to be thrown but it wasn't."
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
function Get-ExceptionError
(
    [Exception]$Exception,

    [string]$ExpectedExceptionTypeName,
    [string]$ExpectedExceptionMessage, 
    [switch]$UseFullTypeName, 

    [switch]$Not
)
{
    if ([string]::IsNullOrWhiteSpace($ExpectedExceptionTypeName) -and 
        [string]::IsNullOrWhiteSpace($ExpectedExceptionMessage))
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

    if (-not [string]::IsNullOrWhiteSpace($ExpectedExceptionTypeName))
    {
        $actualExceptionTypeName = $Exception.GetType().Name
        if ($UseFullTypeName)
        {
            $actualExceptionTypeName = $Exception.GetType().FullName
        }
        # -ieq is case insensitive equals.
        $exceptionTypeMatched = ($actualExceptionTypeName -ieq $ExpectedExceptionTypeName.Trim())
        if ($Not -and $exceptionTypeMatched)
        {
            $errorMessages += `
                "Expected an exception of a different type than $ExpectedExceptionTypeName but exception thrown was of that type."
        }
        elseif (-not $Not -and -not $exceptionTypeMatched)
        {
            $errorMessages += `
                "Expected $ExpectedExceptionTypeName but exception thrown was $actualExceptionTypeName."
        }

    }

    if (-not [string]::IsNullOrWhiteSpace($ExpectedExceptionMessage))
    {
        # -ilike is case insensitive like.
        $exceptionMessageMatched = ($Exception.Message -ilike "*$ExpectedExceptionMessage*")
        if ($Not -and $exceptionMessageMatched)
        {
            $errorMessages += `
                "Expected exception message different than '$ExpectedExceptionMessage' but exception thrown had that message."
        }
        elseif (-not $Not -and -not $exceptionMessageMatched)
        {
            $errorMessages += `
                "Expected exception message '$ExpectedExceptionMessage' but actual exception message was '$($Exception.Message)'."
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