# Called by Setup.cmd and Run.cmd at the end. When the .cmd file was started
# by a double-click in Explorer, its window would close at once; then this
# waits for Enter. Started from a terminal it returns immediately.
# Process chain: this PowerShell -> cmd.exe -> explorer.exe (double-click).
try {
    $self = Get-CimInstance Win32_Process -Filter "ProcessId=$PID" -ErrorAction Stop
    $cmd = Get-CimInstance Win32_Process -Filter "ProcessId=$($self.ParentProcessId)" -ErrorAction Stop
    $parent = Get-Process -Id $cmd.ParentProcessId -ErrorAction Stop
    if ($parent.ProcessName -eq 'explorer') {
        Write-Host ''
        Read-Host 'Press Enter to close / Enter zum Schliessen' | Out-Null
    }
} catch { }
