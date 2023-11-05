if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

Write-Host
$Confirm = $(Write-Host -NoNewLine) + $(Write-Host "Are you sure you want to uninstall all Office apps? [Y] Yes [N] No:" -ForegroundColor Green -NoNewLine; Read-Host)

if($Confirm -match "[yY]") {
    $null = New-Item -Path $env:temp\SaRA -ItemType Directory -Force
    Set-Location $env:temp\SaRA
    Write-Host
    Write-Host 'Downloading Microsoft Support and Recovery Assistant........' -ForegroundColor Green
    #$null = Invoke-WebRequest -Uri "https://aka.ms/SaRA_EnterpriseVersionFiles" -OutFile $env:temp\SaRA\SaRAcmd.zip
    (New-Object Net.WebClient).DownloadFile("https://aka.ms/SaRA_EnterpriseVersionFiles", "$env:temp\SaRA\SaRAcmd.zip")
    Expand-Archive .\SaRAcmd.zip SaRAcmd -Force | Out-Null
    Set-Location SaRAcmd

    Write-Host 'Removing Micorosft Office Application........' -ForegroundColor Yellow
    $closingApps = Get-Process -ProcessName lync, winword, excel, msaccess, mstore, infopath, setlang, msouc, ois, onenote, outlook, powerpnt, mspub, groove, visio, winproj, graph, teams -ErrorAction SilentlyContinue
    $closingApps | Stop-Process -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    .\SaRAcmd.exe -S OfficeScrubScenario -AcceptEula -Officeversion All
    Write-Host 'Done........' -ForegroundColor Green
    Write-Host
}
