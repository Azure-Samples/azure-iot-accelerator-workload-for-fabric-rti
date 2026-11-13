param(
  [Parameter(Mandatory=$true)][int]$TargetPid,
  [Parameter(Mandatory=$true)][long]$StartTicks
)
$ErrorActionPreference = 'Stop'
$processes = [System.Collections.Generic.Dictionary[int,System.Diagnostics.Process]]::new()
$parents = [System.Collections.Generic.Dictionary[int,int]]::new()

function Capture-Children([System.Diagnostics.Process]$Parent) {
  $start = $Parent.StartTime.ToUniversalTime().Ticks
  $end = if ($Parent.HasExited) { $Parent.ExitTime.ToUniversalTime().Ticks } else { [long]::MaxValue }
  foreach ($entry in @(Get-CimInstance Win32_Process -Filter "ParentProcessId = $($Parent.Id)")) {
    $id = [int]$entry.ProcessId
    if ($processes.ContainsKey($id)) { continue }
    $child = Get-Process -Id $id -ErrorAction SilentlyContinue
    if ($null -eq $child) { continue }
    try {
      $null = $child.SafeHandle
      $born = $child.StartTime.ToUniversalTime().Ticks
      if ($born -lt $start -or $born -gt $end) { $child.Dispose(); continue }
      if ($processes.Count -ge 128) { throw '[ownership] Owned process tree exceeded the bounded shutdown limit.' }
      $processes.Add($id, $child)
      $parents.Add($id, $Parent.Id)
      Capture-Children $child
    } catch {
      if (-not $processes.ContainsKey($id)) { $child.Dispose() }
      throw
    }
  }
}

function Stop-OwnedTree([System.Diagnostics.Process]$Process) {
  Capture-Children $Process
  if (-not $Process.HasExited) {
    Stop-Process -Id $Process.Id -Force -ErrorAction Stop
    if (-not $Process.WaitForExit(5000)) { throw '[runtime] Owned process termination was not confirmed.' }
  }
  Capture-Children $Process
  foreach ($id in @($parents.Keys | Where-Object { $parents[$_] -eq $Process.Id })) {
    Stop-OwnedTree $processes[$id]
  }
}

try {
  $root = Get-Process -Id $TargetPid -ErrorAction SilentlyContinue
  if ($null -eq $root) { exit 3 }
  $null = $root.SafeHandle
  if ($root.StartTime.ToUniversalTime().Ticks -ne $StartTicks) {
    $root.Dispose()
    throw '[ownership] Recorded process ID now belongs to another process generation.'
  }
  # Retaining native handles prevents PID reuse while the owned tree is stopped.
  $processes.Add($root.Id, $root)
  Stop-OwnedTree $root
  foreach ($process in $processes.Values) {
    if (-not $process.HasExited) { throw '[runtime] An owned descendant is still running.' }
  }
} finally {
  foreach ($process in $processes.Values) { $process.Dispose() }
}
