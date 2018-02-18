<#
.SYNOPSIS
    Continually checks the state of the network by testing a given IP, rebooting on failure. 
.DESCRIPTION
    Preiodically tests the network interface by pinging the given IP address. Recommend using the 
    local router or other locally accessible network device to avoid unnecessary reboots on 
    uncontrollable wider network failures. Will first reset the network adapter and then re-test, 
    before rebooting the computer on consistent network failure to attempt to thoroughly reset the 
    network devices and services.

    The script will test the connection once every 5 minutes by default
.EXAMPLE
    PS C:\> .\Restart-OnNetworkFailure.ps1 -IPAddress 192.168.1.254
    
    Tests the local router at given address to see if it is reachable. If not, resets the network
    adapter and tries again. If it continues to fail, the computer is restarted.
.INPUTS
    Does not take any pipeline input. Parameters only.
.OUTPUTS
    All status messages are output to the host directly. No useable output.
.NOTES
    Be cautious and use -WhatIf if you need to test the script.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0, Mandatory)]
    [ValidatePattern("[1-255]([0-255]\.){2}([1-255])")]
    [string]
    $IPAddress,

    [Parameter(Position = 1)]
    [int]
    $SecondsBeforeRetest = 300
)
begin {
    $ActiveNic = Get-NetAdapter | 
        Where-Object { $_.Status -eq 'up' }

    $RebootIfFail = $false
}
process {
    while ($true) {
        if (-not (Test-connection $IPAddress -Quiet)) {
            if ($RebootIfFail -and $PSCmdlet.ShouldProcess("Local Machine","Reboot computer")) {
                Write-Host "Connection test failed after resetting adapter; rebooting..." -ForegroundColor Red
                Restart-Computer -Force
            }
            else {
                Write-Host "Connection test failed: $(Get-Date -Format "MM/dd hh:mm tt")"

                Write-Host "Resetting $($ActiveNic.InterfaceDescription)..." -Foregroundcolor Yellow 

                $active_nic | 
                    Disable-NetAdapter -Confirm:$false -PassThru |
                    Enable-NetAdapter -Confirm:$false

                Write-Host "Resetted $($ActiveNic.InterfaceDescription)" -Foregroundcolor Green
                $RebootIfFail = $true
                Start-Sleep -Seconds 30
                continue
            }
        }
        else {
            Write-Host "Connection test passed: $(Get-Date -Format "MM/dd hh:mm tt")"
            $RebootIfFail = $false

            Start-Sleep -Seconds $SecondsBeforeRetest
        }
    }
}