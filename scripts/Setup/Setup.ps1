<#
.SYNOPSIS
    Backwards compatibility wrapper for SetupWorkload.ps1

.NOTES
    This script is provided for backwards compatibility only.
    Please use SetupWorkload.ps1 directly for new implementations.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$WorkloadName,
    
    [Parameter()]
    [string]$FrontendAppId,
    
    [Parameter()]
    [switch]$Force
)

# Forward all arguments to SetupWorkload.ps1
& (Join-Path $PSScriptRoot "SetupWorkload.ps1") @PSBoundParameters