function StopWordProcess {
    [CmdletBinding(SupportsShouldProcess=$true)]
    # [CmdletBinding(ConfirmImpact="Medium")]
    param($pName="WINWORD")
    # foreach($p in Get-Process | where {$_.Name -match $pName}){
    #     if($PSCmdlet.ShouldProcess(
    #         "process $($p.Name) " +
    #         " (id: $($p.ProcessId))","Stop Process")){
    #         $p | Stop-Process
    #     }
    # }
     foreach($p in Get-WmiObject Win32_Process | 
            where {$_.Name -match $pName}){
        if($PSCmdlet.ShouldProcess(
            "process $($p.Name) " +
            " (id: $($p.ProcessId))",
            "Stop Process")){
            # $p | Stop-Process
            $p.Terminate() > $null
        }
    }
}

StopWordProcess -confirm