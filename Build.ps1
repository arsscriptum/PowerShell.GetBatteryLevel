rm Power.exe -EA Ignore -Force
$Null = cl .\Power.ps1 -noConsole -noError -noOutput
if(Test-Path ./Power.exe){
        Write-Host '[OK] ' -f DarkGreen -NoNewLine
        Write-Host "Power.exe compiled" -f Gray 
}
cp Power.exe $ENV:ToolsRoot -Force -EA Ignore
  Write-Host '[OK] ' -f DarkGreen -NoNewLine
 Write-Host "cp Power.exe $ENV:ToolsRoot" -f Gray 