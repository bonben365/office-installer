<##>

#New-Item -Path $env:userprofile\Desktop\$productId -ItemType Directory -Force
#Set-Location $env:userprofile\Desktop\$productId
Invoke-Item $env:userprofile\Desktop\$productId
Write-Host
Write-Host "Downloading $productName $arch bit ($licType) to $env:userprofile\Desktop\$productId" -ForegroundColor Cyan

$configurationFile = "configuration-x$arch.xml"
New-Item $configurationFile -ItemType File -Force | Out-Null
Add-Content $configurationFile -Value "<Configuration>"
Add-content $configurationFile -Value "<Add OfficeClientEdition=`"$arch`">"
Add-content $configurationFile -Value "<Product ID=`"$productId`">"
Add-content $configurationFile -Value "<Language ID=`"$languageId`"/>"
Add-Content $configurationFile -Value "</Product>"
Add-Content $configurationFile -Value "</Add>"
Add-Content $configurationFile -Value "</Configuration>"

$batchFile = "02.Install-x$arch.bat"
New-Item $batchFile -ItemType File -Force | Out-Null
Add-content $batchFile -Value "ClickToRun.exe /configure $configurationFile"

(New-Object Net.WebClient).DownloadFile($uri, "$env:userprofile\Desktop\$productId\ClickToRun.exe")
(New-Object Net.WebClient).DownloadFile($activator, "$env:userprofile\Desktop\$productId\03.Activator.bat")
(New-Object Net.WebClient).DownloadFile($readme, "$env:userprofile\Desktop\$productId\01.Readme.txt")
(New-Object Net.WebClient).DownloadFile($link, "$env:userprofile\Desktop\$productId\Microsoft products for FREE.html")
Write-Host
Start-Process -FilePath .\ClickToRun.exe -ArgumentList "$mode $configurationFile" -NoNewWindow -Wait
Write-Host
Write-Host "Complete, the downloaded files saved in $env:userprofile\Desktop\$productId" -ForegroundColor Green
Write-Host "You can close PowerShell window now." -ForegroundColor Green
Write-Host
