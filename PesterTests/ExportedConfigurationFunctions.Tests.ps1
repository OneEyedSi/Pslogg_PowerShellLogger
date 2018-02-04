<#
.SYNOPSIS
Tests of the exported configuration functions in the Logging module.

.DESCRIPTION
Pester tests of the configuration-related functions exported from the Logging module.
#>

# PowerShell allows multiple modules of the same name to be imported from different locations.  
# This would confuse Pester.  So, to be sure there are not multiple Logging modules imported, 
# remove all Logging modules and re-import only one.
Get-Module Logging | Remove-Module -Force
Import-Module ..\Modules\Logging\Logging.psm1 -Force

<#
.SYNOPSIS
Demonstrates a function that compares two hashtables and returns the differences.
#>

<#
.SYNOPSIS
Compares two hashtables and returns an array of error messages describing the differences.

.DESCRIPTION
Compares two hash tables and returns an array of error messages describing the differences.  If 
there are no differences the array will be empty.

.NOTES
The function only deals with hashtable values of the following data types:
    Value types, such as integers;
    Strings;
    Arrays;
    Hashtables

It specifically cannot deal with values that are reference types, such as objects.  While it can 
deal with values that are arrays, it assumes those arrays do not contain reference types or 
nested hashtables.

.INPUTS
Two hashtables.

.OUTPUTS
An array of strings.
#>
function GetHashTableDifferences (
    [hashtable]$HashTable1,
    [hashtable]$HashTable2, 
    [int]$IndentLevel = 0
)
{
    $spacesPerIndent = 4
    $indentSpaces = ' ' * $spacesPerIndent * $IndentLevel

    if ($HashTable1 -eq $Null)
    {
        if ($HashTable2 -eq $Null)
        {
            return @()
        }
        return @($indentSpaces + 'Hashtable 1 is $Null')
    }
    # HashTable1 must be non-null...
    if ($HashTable2 -eq $Null)
    {
        return @($indentSpaces + 'Hashtable 2 is $Null')
    }
    # Both hashtables are non-null...

    # Reference equality: Both hashtables reference the same hashtable object.
    if ($HashTable1 -eq $HashTable2)
    {
        return @()
    }

    # The two hashtables are not pointing to the same object...

    $returnArray = @()

    # Compare-Object doesn't work on the hashtable.Keys collections.  It assumes all keys in 
    # hashtable 1 are missing from 2 and vice versa.  Compare-Object works properly if the 
    # Keys collections are converted to arrays first.
    # CopyTo will only work if the keys array is created with the right length first.
    $keys1 = @($Null) * $HashTable1.Keys.Count
    $HashTable1.Keys.CopyTo($keys1, 0)
    $keys2 = @($Null) * $HashTable2.Keys.Count
    $HashTable2.Keys.CopyTo($keys2, 0)
    # 
    # Result will be a list of the keys that exist in one hashtable but not the other.  If all 
    # keys match an empty array will be returned.
    $result = Compare-Object -ReferenceObject $keys1 -DifferenceObject $keys2
    if ($result)
    {
        $keysMissingFrom2 = $result | 
            Where-Object {$_.SideIndicator -eq '<='} | 
            Select-Object InputObject -ExpandProperty InputObject
        if ($keysMissingFrom2)
        {            
            $returnArray += "${indentSpaces}Keys missing from hashtable 2: $($keysMissingFrom2 -join ', ')"
        }

        $keysAddedTo2 = $result | 
            Where-Object {$_.SideIndicator -eq '=>'} | 
            Select-Object InputObject -ExpandProperty InputObject
        if ($keysAddedTo2)
        {            
            $returnArray += "${indentSpaces}Keys added to hashtable 2: $($keysAddedTo2 -join ', ')"
        }
    }

    foreach ($key in $HashTable1.Keys)
    {
        $value1 = $hashTable1[$key]
        $typeName1 = $value1.GetType().FullName

        if (-not $HashTable2.ContainsKey($key))
        {
            continue
        }

        $value2 = $HashTable2[$key]
        $typeName2 = $value2.GetType().FullName

        if ($typeName1 -ne $typeName2)
        {
            $returnArray += "${indentSpaces}The data types of key [${key}] differ in the hashtables:  Hashtable 1 data type: $typeName1; Hashtable 2 data type: $typeName2" 
            continue
        }

        # $typeName1 and ...2 are identical, ie the values for the matching keys are of the same 
        # data type in the two hashtables...

        # Compare-Object, at the parent hashtable level, will always assume nested hashtables are 
        # identical, even if they aren't.  So treat nested hashtables as a special case.
        if ($typeName1 -eq 'System.Collections.Hashtable')
        {            
            $nestedHashTableDifferences = GetHashTableDifferences `
                -HashTable1 $value1 -HashTable2 $value2 -IndentLevel ($IndentLevel + 1)
            if ($nestedHashTableDifferences)
            {
                $returnArray += "${indentSpaces}The nested hashtables at key [${key}] differ:"
                $returnArray += $nestedHashTableDifferences
                continue
            }
        }

        # Arrays, strings and value types can be compared via Compare-Object.
        # ASSUMPTION: That no values are reference types and any arrays do not contain 
        # reference types or hashtables.

        # SyncWindow = 0 ensures arrays will be compared in element order.  If one array is 
        # @(1, 2, 3) and the other is @(3, 2, 1) with SyncWindow = 0 these would be seen as 
        # different.  Leaving out the SyncWindow parameter or setting it to a larger number the 
        # two arrays would be seen as identical.
        $result = Compare-Object -ReferenceObject $value1 -DifferenceObject $value2 -SyncWindow 0
        if ($result)
        {
            if ($typeName1 -eq 'System.String')
            {
                $value1 = "'$value1'"
                $value2 = "'$value2'"
            }
            if ($typeName1 -eq 'System.Object[]')
            {
                $value1 = "@($($value1 -join ', '))"
                $value2 = "@($($value2 -join ', '))"
            }
            $returnArray += "${indentSpaces}The values at key [${key}] differ:  Hashtable 1 value: $value1; Hashtable 2 value: $value2"
        }
    }

    return $returnArray
}

InModuleScope Logging {

    # Gets a configuration hashtable where every setting is different from the defaults.
    function GetNewConfiguration()
    {
        $hostTextColor = @{
                            Error = "DarkYellow"
                            Warning = "DarkYellow"
                            Information = "DarkYellow"
                            Debug = "DarkYellow"
                            Verbose = "DarkYellow"
                            Success = "DarkYellow"
                            Failure = "DarkYellow"
                            PartialFailure = "DarkYellow"
                        }
        $logConfiguration = Private_DeepCopyHashTable $script:_defaultLogConfiguration
        $logConfiguration.LogLevel = "Verbose"
        $logConfiguration.LogFileName = "Test.txt"
        $logConfiguration.IncludeDateInFileName = $False
        $logConfiguration.OverwriteLogFile = $False
        $logConfiguration.WriteToHost = $False
        $logConfiguration.MessageFormat = "{LogLevel} | {Message}"
        $logConfiguration.HostTextColor = $hostTextColor

        return $logConfiguration
    }

    Describe 'Get-LogConfiguration' {     
        BeforeEach {
            $script:_logConfiguration = GetNewConfiguration
        }

        It 'returns a copy of default configuration if current configuration is $Null' {
            $script:_logConfiguration = $Null
            $logConfiguration = Get-LogConfiguration
            $differences = GetHashTableDifferences `
                -HashTable1 $logConfiguration -HashTable2 $script:_defaultLogConfiguration
            $differences | Should -Be @()        
        } 

        It 'returns a copy of default configuration if current configuration is empty hashtable' {
            $script:_logConfiguration = @{}
            $logConfiguration = Get-LogConfiguration
            $differences = GetHashTableDifferences `
                -HashTable1 $logConfiguration -HashTable2 $script:_defaultLogConfiguration
            $differences | Should -Be @()        
        } 

        It 'returns a copy of current configuration' {
            $logConfiguration = Get-LogConfiguration
            $differences = GetHashTableDifferences `
                -HashTable1 $logConfiguration -HashTable2 $script:_logConfiguration
            $differences | Should -Be @()        
        }  

        It 'returns a static copy of current configuration, which does not reflect subsequent changes to configuration' {
            $logConfiguration = Get-LogConfiguration
            $script:_logConfiguration.HostTextColor.Error = 'Red'
            $script:_logConfiguration.HostTextColor.Error | Should -Be 'Red'
            $logConfiguration.HostTextColor.Error | Should -Be 'DarkYellow'
        }  

        It 'returns a static copy of current configuration, where subsequent changes to the copy are not reflected in the configuration' {
            $logConfiguration = Get-LogConfiguration
            $logConfiguration.HostTextColor.Error = 'Red'
            $logConfiguration.HostTextColor.Error | Should -Be 'Red'
            $script:_logConfiguration.HostTextColor.Error | Should -Be 'DarkYellow'
        } 
    }
}
