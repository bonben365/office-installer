if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}



$uninstall = {
   New-Item -Path $env:temp\c2r -ItemType Directory -Force
   Set-Location $env:temp\c2r
   $fileName = 'configuration.xml'
   New-Item $fileName -ItemType File -Force
   Add-Content $fileName -Value '<Configuration>'
   Add-Content $fileName -Value '<Remove All="True"/>'
   Add-Content $fileName -Value '</Configuration>'
   $uri = 'https://github.com/bonben365/office-installer/raw/main/setup.exe'
   (New-Object Net.WebClient).DownloadFile($uri, "$env:temp\c2r\setup.exe")
   .\setup.exe /configure .\configuration.xml
}

