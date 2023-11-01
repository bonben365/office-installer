if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}
Write-Host
Write-Host "Remove all installed Office apps using Office Deployment Tool ODT." -ForegroundColor yellow
$Confirm= Read-Host Are you sure you want to uninstall all Office apps? [Y] Yes [N] No
if($Confirm -match "[yY]") {
    New-Item -Path $env:temp\c2r -Type Directory -Force | Out-Null
    Set-Location $env:temp\c2r
    $fileName = 'configuration.xml'
    New-Item $fileName -ItemType File -Force | Out-Null
    Add-Content $fileName -Value '<Configuration>'
    Add-Content $fileName -Value '<Remove All="True"/>'
    Add-Content $fileName -Value '</Configuration>'
    $uri = 'https://github.com/bonben365/office-installer/raw/main/setup.exe'
    (New-Object Net.WebClient).DownloadFile($uri, "$env:temp\c2r\setup.exe")
    Write-host "Removing the Office apps"
    .\setup.exe /configure .\configuration.xml
} else {
    Write-Host ""
}
