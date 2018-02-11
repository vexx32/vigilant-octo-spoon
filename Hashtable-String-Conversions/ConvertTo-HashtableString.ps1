<#
.SYNOPSIS
    ConvertTo-HashtableString
    Converts a standalone or nested hashtable object into string form as you would see in raw PS code.
.DESCRIPTION
    ConvertTo-HashtableString is only capable of handling simple hashtable objects where the contents
    are either value types or can be easily converted into strings with the object's .ToString()
    method. It will not properly handle complex hashtables or cases where the value's .ToString()
    method does not return useable data.
.EXAMPLE
    PS C:\> ConvertTo-HashtableString -Hashtable @{Value1 = "Value1";Value2 = 2}
    
    Output:
    @{
        "Value1" = "Value1"
        "Value2" = 2
    }
.INPUTS
    ConvertTo-HashtableString will accept pipeline input of any amount of hashtable objects.
.OUTPUTS
    Outputs a string representation of the input object.
.NOTES
    Will handle nested hashtables without difficulty, but will not be able to handle other data types that do not
    cleanly convert to string values with .ToString()
#>
filter ConvertTo-HashtableString {
    [CmdletBinding()]
    [Alias(cths)]
    param (
        [Alias('Hashtable')]
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [hashtable]$InputObject
    )
    
    process {
        $StringBuilder = New-Object 'System.Text.StringBuilder'
        $StringBuilder.AppendLine("@{")

        $InputObject.Keys | ForEach-Object {
            switch ($InputObject.GetType()) {
                [Hashtable] {
                    $ValueString = ConvertTo-HashtableString -InputObject $InputObject[$_]
                }
                [string] {
                    $ValueString = "'$($InputObject[$_.ToString()])'"
                }
                default {
                    $ValueString = $InputObject[$_.ToString()]
                }
            }

            $StringBuilder.AppendFormat('`t"{0}" = {1}', $_, $ValueString)
        }

        $StringBuilder.AppendLine("}")

        Write-Output $StringBuilder.ToString()
    }
}