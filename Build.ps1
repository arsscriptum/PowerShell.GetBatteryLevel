
    [CmdletBinding(SupportsShouldProcess)]
    param()

function Invoke-BuildPowerApp{
    [CmdletBinding(SupportsShouldProcess)]
    param()

    try{
        Import-Module PowerShell.Module.Compiler -Force 
        $m = Get-Module "PowerShell.Module.Compiler" -ea Ignore
        if($Null -eq $m){throw "no compiler module" }
        rm Power.exe -EA Ignore -Force
        $Null = cl .\Power.ps1 -noConsole -noError -noOutput
        if(Test-Path ./Power.exe){
                Write-Host '[OK] ' -f DarkGreen -NoNewLine
                Write-Host "Power.exe compiled" -f Gray 
        }
        cp Power.exe $ENV:ToolsRoot -Force -EA Ignore
          Write-Host '[OK] ' -f DarkGreen -NoNewLine
         Write-Host "cp Power.exe $ENV:ToolsRoot" -f Gray 
    }catch{
        Write-Error "$_"
    }
}


Invoke-BuildPowerApp