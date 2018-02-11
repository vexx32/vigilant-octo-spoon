<#
.SYNOPSIS
    ConvertFrom-HashtableString
    Converts a string representation of a hashtable into a PowerShell hashtable object.
.DESCRIPTION
    Takes an input string or file path, and invokes it as an expression, casting it to a hashtable object.
    Will throw an error on any non-hashtable object sequences, but any arbitrary commands may be present in
    the hashtable value fields. Use with caution.
.EXAMPLE
    PS C:\> ConvertFrom-HashtableString -FilePath .\hashtable.txt
    
    Takes the file at the path specified (in this case, 'C:\hashtable.txt') and parses the text into a useable PowerShell object.
.INPUTS
    ConvertFrom-HashtableString will take pipeline input as a list of file paths or direct string input.
.OUTPUTS
    Outputs a single hashtable object.
.NOTES
    Multiple hashtables in a single file or string are not supported.
#>
function ConvertFrom-HashtableString {
    [CmdletBinding(DefaultParameterSetName = "DirectInput")]
    [Alias('cfhs')]
    [OutputType("Hashtable")]
    param (
        [Alias('Path', 'InFile')]
        [Parameter(Mandatory, ValueFromPipeline, Position = 0, ParameterSetName = "FileInput")]
        [System.IO.FileInfo]$FilePath,


        [Parameter(Mandatory, ValueFromPipeline, Position = 0, ParameterSetName = "DirectInput")]
        [string]$HashtableString
    )
    process {
        if ($PSCmdlet.ParameterSetName -eq "FileInput") {
            $HashtableString = Get-Content -Raw -Path $FilePath
        }

        [Hashtable] $OutputObject = Invoke-Expression -Command $HashtableString
        Write-Output $OutputObject
    }
}