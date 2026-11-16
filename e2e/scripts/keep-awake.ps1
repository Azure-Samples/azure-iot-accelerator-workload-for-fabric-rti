$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class E2ePowerRequest {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint flags);
}
'@

# ES_CONTINUOUS | ES_SYSTEM_REQUIRED; no display or screen-lock override.
if ([E2ePowerRequest]::SetThreadExecutionState([uint32]2147483649) -eq 0) {
  throw 'Could not acquire the process-scoped system execution request.'
}
try {
  [Console]::Out.WriteLine('E2E_AWAKE_READY')
  [Console]::Out.Flush()
  $null = [Console]::ReadLine()
} finally {
  $null = [E2ePowerRequest]::SetThreadExecutionState([uint32]2147483648)
}
