function Invoke-Struggle {
    [CmdletBinding()]
    [Alias('struggle', 'grumble')]
    param(
        [Parameter(Position = 1, Mandatory)]
        [scriptblock]
        $ScriptBlock
    )
    $Error.StackTrace | ForEach-Object -Process $ScriptBlock
}
struggle { $_ }