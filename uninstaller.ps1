if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "irm msgang.com/uninstaller | iex"
    break
}

Write-Host

$method = $(Write-Host -NoNewLine) + $(Write-Host " Select the method for uninstallation (Enter either 1 or 2):`n
 1 = Office Deployment Tool (ODT) (Recommended)
 2 = Microsoft Support and Recovery Assistant (SaRA) `n
Your option: " -ForegroundColor Cyan -NoNewLine; Read-Host)

Write-Host

function Uninstall-ODT {
    Write-Host 'Downloading Office Deployment Tool...'
    $null = New-Item -Path $env:temp\c2r -ItemType Directory -Force
    Set-Location $env:temp\c2r
    $fileName = 'configuration.xml'
    $null = New-Item $fileName -ItemType File -Force
    Add-Content $fileName -Value '<Configuration>'
    Add-Content $fileName -Value '<Remove All="True"/>'
    Add-Content $fileName -Value '</Configuration>'
    $uri = 'https://github.com/bonben365/office-installer/raw/main/setup.exe'
    (New-Object Net.WebClient).DownloadFile($uri, "$env:temp\c2r\setup.exe")
    .\setup.exe /configure .\configuration.xml
    Write-Host 'Done. You can close this PowerShell window.'
}

function Uninstall-SaRA {
    $null = New-Item -Path $env:temp\SaRA -ItemType Directory -Force
    Set-Location $env:temp\SaRA
    
    Write-Host 'Downloading Microsoft Support and Recovery Assistant........'
    #$null = Invoke-WebRequest -Uri "https://aka.ms/SaRA_EnterpriseVersionFiles" -OutFile $env:temp\SaRA\SaRAcmd.zip
    (New-Object Net.WebClient).DownloadFile("https://aka.ms/SaRA_EnterpriseVersionFiles", "$env:temp\SaRA\SaRAcmd.zip")
    $null = Expand-Archive .\SaRAcmd.zip SaRAcmd -Force
    Set-Location SaRAcmd

    Write-Host 'Removing Micorosft Office Application........'
    $closingApps = Get-Process -ProcessName lync, winword, excel, msaccess, mstore, infopath, setlang, msouc, ois, onenote, outlook, powerpnt, mspub, groove, visio, winproj, graph, teams -ErrorAction SilentlyContinue
    $closingApps | Stop-Process -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    .\SaRAcmd.exe -S OfficeScrubScenario -AcceptEula -Officeversion All > result.txt

    #Grab the result then print the output
    $result = Get-Content .\result.txt
    if (($result | Select-String -SimpleMatch "something went wrong" | Measure-Object).Count -gt 0){
        Write-Host "Unintall failed. Re-run the script then select another method."
        Write-Host 
    } else {
        Write-Host 'Done. You can close this PowerShell window.'
        Write-Host 
    }
    #Cleanup
    Remove-Item result.txt -Force
}

if($method -match "[1]") {
    Write-Host "Uninstaling using Office Deployment Tool..." 
    Start-Sleep -Seconds 1
    Uninstall-ODT
}

if($method -match "[2]") {
    Write-Host "Uninstaling using Microsoft Support and Recovery Assistant (SaRA)..."
    Start-Sleep -Seconds 1
    Uninstall-SaRA
}

