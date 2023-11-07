if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

$uri = "https://github.com/bonben365/office-installer/raw/main/setup.exe"
$uri2013 = "https://github.com/bonben365/office-installer/raw/main/bin2013.exe"
$activator = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/activator.bat'
$readme = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/Readme.txt'
$link = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/Microsoft%20products%20for%20FREE.html'

Write-Host 
Write-Host "- Description:****************************************************
- This tool downloads both x32 and x64 versions of Office apps.  *
- Support download both Retail and Volume license type.          *
- Default language for downloading is English.                   *
- You can change it from the code to download your favorite one. *
- Download location C:\Users\{username}\Desktop\Source.          *
******************************************************************" -ForegroundColor Cyan

$edition = $(Write-Host -NoNewLine) + $(Write-Host "`nPlease type the Office edition that you want to download (2013/2016/2019/2021/365):" -ForegroundColor Green -NoNewLine; Read-Host)
$licType = $(Write-Host -NoNewLine) + $(Write-Host "Please enter the license type (Volume/Retail):" -ForegroundColor Green -NoNewLine; Read-Host)

$archs = @('32';'64')
$languageId = 'en-US'
$mode = '/download'

function DownloadMSOffice {
    New-Item -Path $env:userprofile\Desktop\Source\$productId$edition-x$arch -ItemType Directory -Force | Out-Null
    Set-Location $env:userprofile\Desktop\Source\$productId$edition-x$arch
    $totalItems = ($productIds.Count)*2
    Write-Host "($i\$totalItems) Downloading $productId $arch bit to $env:userprofile\Desktop\Source\$productId$edition-x$arch" -ForegroundColor Cyan
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

    (New-Object Net.WebClient).DownloadFile($uri, "$env:userprofile\Desktop\Source\$productId$edition-x$arch\ClickToRun.exe")
    (New-Object Net.WebClient).DownloadFile($activator, "$env:userprofile\Desktop\Source\$productId$edition-x$arch\03.Activator.bat")
    (New-Object Net.WebClient).DownloadFile($readme, "$env:userprofile\Desktop\Source\$productId$edition-x$arch\01.Readme.txt")
    (New-Object Net.WebClient).DownloadFile($link, "$env:userprofile\Desktop\Source\$productId$edition-x$arch\Microsoft products for FREE.html")
    .\ClickToRun.exe $mode .\$configurationFile
}

function DownloadMSOfficeX {
    New-Item -Path $env:userprofile\Desktop\Source\$productId$edition-x$arch -ItemType Directory -Force | Out-Null
    Set-Location $env:userprofile\Desktop\Source\$productId$edition-x$arch
    $totalItems = ($productIds.Count)*2
    Write-Host "($i\$totalItems) Downloading $productId $arch bit to $env:userprofile\Desktop\Source\$productId$edition-x$arch" -ForegroundColor Cyan
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
    (New-Object Net.WebClient).DownloadFile($uri2013, "$env:userprofile\Desktop\Source\$productId$edition-x$arch\ClickToRun.exe")
    (New-Object Net.WebClient).DownloadFile($activator, "$env:userprofile\Desktop\Source\$productId$edition-x$arch\03.Activator.bat")
    (New-Object Net.WebClient).DownloadFile($readme, "$env:userprofile\Desktop\Source\$productId$edition-x$arch\01.Readme.txt")
    (New-Object Net.WebClient).DownloadFile($link, "$env:userprofile\Desktop\Source\$productId$edition-x$arch\Microsoft products for FREE.html")
    .\ClickToRun.exe $mode .\$configurationFile  
}

$i = 1
if ($edition -match '2013' -or $edition -match '2016'){
    $productIds = @(
        "ProfessionalRetail";
        "StandardRetail";
        "ProjectProRetail";
        "ProjectStdRetail";
        "VisioProRetail";
        "VisioStdRetail";
        "WordRetail";
        "ExcelRetail";
        "PowerPointRetail";
        "OutlookRetail";
        "PublisherRetail";
        "AccessRetail"
    )
    foreach ($productId in $productIds) {
        foreach ($arch in $archs){
            DownloadMSOfficeX
            $i++
        }
    }
}

if ($edition -match '2019' -or $edition -match '2021'){
    $productIds = @(
    "ProPlus$edition$licType";
    "Standard$edition$licType";
    "ProjectPro$edition$licType";
    "ProjectStd$edition$licType";
    "VisioPro$edition$licType";
    "VisioStd$edition$licType";
    "Word$edition$licType";
    "Excel$edition$licType";
    "PowerPoint$edition$licType";
    "Outlook$edition$licType";
    "Publisher$edition$licType";
    "Access$edition$licType";
    "HomeBusiness$($edition)Retail";
    "HomeStudent$($edition)Retail"
    )
    foreach ($productId in $productIds) {
        foreach ($arch in $archs){
            DownloadMSOffice
            $i++
        }
    }
}

if ($edition -match '365'){
    $productIds = @(
        "O365HomePremRetail";
        "O365BusinessRetail";
        "O365ProPlusRetail"
    )
    foreach ($productId in $productIds) {
        foreach ($arch in $archs){
            DownloadMSOffice
            $i++
        }
    }
}