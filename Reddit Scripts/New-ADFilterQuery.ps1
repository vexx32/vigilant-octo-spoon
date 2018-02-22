function New-ADFilterQuery {
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
    [CmdletBinding()]
    [OutputType([scriptblock])]
    param(
        [Parameter(Position = 0, Mandatory, ValueFromPipeline)]
        [Alias('Items', 'Values')]
        [string[]]
        $FilterValues,

        [Parameter(Position = 1, Mandatory)]
        [ValidatePattern('\w')]
        [string]
        $FilterProperty,

        [Parameter(Position = 1)]
        [Alias('Operator')]
        [ValidateSet(
            "-eq", "-ne",
            "-like", "-notlike", 
            "-le", "-ge", 
            "-lt", "-gt", 
            "-approx", "-recursivematch", 
            "-bor", "-band"
        )]
        [string]
        $FilterOperator = "-eq",

        [Parameter(Position = 2)]
        [ValidateSet("-and", "-or")]
        [string]
        $JoinOperator = "-and"
    )

    begin {
        $FilterString = New-Object 'System.Text.StringBuilder'
    }
    process {
        $FilterValues | ForEach-Object {
            if ($FilterString.Length -gt 0) {
                $FilterString.Append(" $JoinOperator ")
            }
            $FilterString.Append("$FilterProperty $FilterOperator '$_'")
        }
    }
    end {
        [scriptblock]::Create($FilterString.ToString())
    }

}