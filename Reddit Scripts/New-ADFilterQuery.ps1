function New-ADFilterQuery {
    <#
    .SYNOPSIS
        New-ADFilterQuery
        Takes an array of items and turns them into a structured AD query to be passed to the filter parameter.
    .DESCRIPTION
        Pipe in or supply an array of values to construct a filter query from. All values must target the same AD property.
    .EXAMPLE
        PS C:\> $Query = New-ADFilterQuery -Items 'WS-Example01' -FilterProperty 'SamAccountName'
        PS C:\> Get-ADComputer -Filter $Query
        
        Constructs a filter query to search for anything with a SamAccountName of 'WS-Example01'. Passing it to Get-ADComputer will
        retrieve computer objects with this name. The operation can be reversed or altered using the -FilterOperator and -JoinOperator
        parameters.
    .INPUTS
        Takes pipeline input as a collection of string values containing the target values.
    .OUTPUTS
        Outputs a script block that can be passed to the -Filter parameter of AD cmdlets.
    .NOTES
        
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
