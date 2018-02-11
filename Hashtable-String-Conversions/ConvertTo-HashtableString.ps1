<#
.SYNOPSIS
    Short description
.DESCRIPTION
    Long description
.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>
filter ConvertTo-HashtableString {
    [CmdletBinding()]
    param (
        [Alias('Hashtable')]
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [hashtable]$InputObject
    )
    
    process {
        $StringBuilder = New-Object 'System.Text.StringBuilder'
        $StringBuilder.AppendLine("@{")

        $InputObject.Keys | ForEach-Object {
            if ($InputObject[$_] -is [Hashtable]) {
                $ValueString = ConvertTo-HashtableString -InputObject $InputObject[$_]
            }
            else {
                $ValueString = $InputObject[$_.ToString()]
            }

            $StringBuilder.AppendFormat('"{0}" = {1}', $_, $ValueString)
        }

        $StringBuilder.AppendLine("}")

        Write-Output $StringBuilder.ToString()
    }
}