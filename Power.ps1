<#
  ╓──────────────────────────────────────────────────────────────────────────────────────
  ║   PowerShell.Module.Core            
  ║   
  ║   Power.ps1
  ╙──────────────────────────────────────────────────────────────────────────────────────
 #>

[CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Position=0,Mandatory=$false)]
        [int]$StatusBarTime = 5
    ) 

function Get-BatteryLevel {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory=$false)]
        [switch]$ShowStatusBar, 
        [Parameter(Mandatory=$false)]
        [int]$StatusBarTime = 3
    )  
    [int]$PercentBattery = 0
    $WmicExe = (get-command wmic).Source  
    [array]$PowerData=  &"$WmicExe" "PATH" "Win32_Battery" "Get" "EstimatedChargeRemaining"
    $PowerDataLen = $PowerData.Length
    $Data = [System.Collections.ArrayList]::new();
    ForEach($line in $PowerData){
        if($line.Length) { $Null=$Data.Add($line); } 
    }
    if( $Data.Count -eq 2 ){
        $PercentBattery = $Data[1]
    }


    $StatusBarTime = $StatusBarTime * 100
    if($ShowStatusBar){
        while($StatusBarTime){
            Write-Progress -Activity "BATTERY LEVEL INDICATOR --> $PercentBattery" -Status "$PercentBattery PERCENT" -PercentComplete $PercentBattery
            Sleep -Milliseconds 10
            $StatusBarTime--
        }
    }

    
    return $PercentBattery 
}



Get-BatteryLevel -ShowStatusBar -StatusBarTime $StatusBarTime