Get-NetTCPConnection -State Established | ForEach-Object {
    $proc = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue
    [PSCustomObject]@{
        LocalAddress  = $_.LocalAddress
        LocalPort     = $_.LocalPort
        RemoteAddress = $_.RemoteAddress
        RemotePort    = $_.RemotePort
        State         = $_.State
        PID           = $_.OwningProcess
        ProcessName   = if ($proc) { $proc.ProcessName } else { "N/A" }
    }
} | Format-Table -AutoSize
