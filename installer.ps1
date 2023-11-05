if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
   Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
   break
}

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
[void] [Reflection.Assembly]::LoadWithPartialName("PresentationCore")

$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(1000,480)
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow 
$Form.Text = "Microsoft Office Installation Tool - www.msgang.com"
$Form.Font = New-Object System.Drawing.Font("Consolas",8,[System.Drawing.FontStyle]::Regular)
$Form.ShowInTaskbar = $True
$Form.KeyPreview = $True
$Form.AutoSize = $True
$Form.FormBorderStyle = "Fixed3D"
$Form.MaximizeBox = $False
$Form.MinimizeBox = $False
$Form.ControlBox = $True
$Form.Icon = $Icon

$uri = "https://github.com/bonben365/office-installer/raw/main/setup.exe"
$uri2013 = "https://github.com/bonben365/office-installer/raw/main/bin2013.exe"
$activator = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/activator.bat'
$readme = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/Readme.txt'
$link = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/Microsoft%20products%20for%20FREE.html'

$install = { 
  New-Item -Path $env:userprofile\Desktop\$productId -ItemType Directory -Force
  Set-Location $env:userprofile\Desktop\$productId
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
  .\ClickToRun.exe $mode .\$configurationFile
  Write-Host
  Write-Host "Complete, the downloaded files saved in $env:userprofile\Desktop\$productId" -ForegroundColor Green
  Write-Host "You can close PowerShell window now." -ForegroundColor Green
  Write-Host
}

$install2013 = { 
  New-Item -Path $env:userprofile\Desktop\$productId -ItemType Directory -Force
  Set-Location $env:userprofile\Desktop\$productId
  Invoke-Item $env:userprofile\Desktop\$productId
  Write-Host
  Write-Host "Downloading $productName $arch bit (Retail) to $env:userprofile\Desktop\$productId" -ForegroundColor Cyan
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
  (New-Object Net.WebClient).DownloadFile($uri2013, "$env:userprofile\Desktop\$productId\ClickToRun.exe")
  (New-Object Net.WebClient).DownloadFile($activator, "$env:userprofile\Desktop\$productId\03.Activator.bat")
  (New-Object Net.WebClient).DownloadFile($readme, "$env:userprofile\Desktop\$productId\01.Readme.txt")
  (New-Object Net.WebClient).DownloadFile($link, "$env:userprofile\Desktop\$productId\Microsoft products for FREE.html")
  .\ClickToRun.exe $mode .\$configurationFile

  Write-Host
  Write-Host "Complete, the downloaded files saved in $env:userprofile\Desktop\$productId" -ForegroundColor Green
  Write-Host "You can close PowerShell window now." -ForegroundColor Green
  Write-Host
}

$uninstall = {
   New-Item -Path $env:temp\c2r -ItemType Directory -Force
   Set-Location $env:temp\c2r
   $fileName = 'configuration.xml'
   New-Item $fileName -ItemType File -Force
   Add-Content $fileName -Value '<Configuration>'
   Add-Content $fileName -Value '<Remove All="True"/>'
   Add-Content $fileName -Value '</Configuration>'
   (New-Object Net.WebClient).DownloadFile($uri, "$env:temp\c2r\ClickToRun.exe")
   .\ClickToRun.exe /configure .\configuration.xml
}

$activate = {
   New-Item -Path $env:temp\activate -ItemType Directory -Force
   Set-Location $env:temp\activate
   $fileName = 'configuration.xml'
   (New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/bonben365/microsoft/main/oatc.ps1', "$env:temp\activate\activate.ps1")
   .\activate.ps1
   Write-Host "Done. You can close this PowerShell window." -ForegroundColor Green
}

############################################## Start functions

  function microsoftInstaller {
  try {
  if ($arch32.Checked -eq $true) {$arch="32"}
  if ($arch64.Checked -eq $true) {$arch="64"}

  if ($licenseTypeVolume.Checked -eq $true) {$licType="Volume"}
  if ($licenseTypeRetail.Checked -eq $true) {$licType="Retail"}

  if ($installModeSetup.Checked -eq $true) {$mode='/configure'}
  if ($installModeDownload.Checked -eq $true) {$mode='/download'}
  if ($installModeActivate.Checked -eq $true) {Invoke-Command $activate}

  if ($English.Checked -eq $true) {$languageId="en-US"}
  if ($Japanese.Checked -eq $true) {$languageId="ja-JP"}
  if ($Korean.Checked -eq $true) {$languageId="ko-KR"}
  if ($Chinese.Checked -eq $true) {$languageId="zh-TW"}
  if ($French.Checked -eq $true) {$languageId="fr-FR"}
  if ($Spanish.Checked -eq $true) {$languageId="es-ES"}
  if ($Vietnamese.Checked -eq $true) {$languageId="vi-VN"}


  if ($m365Home.Checked -eq $true) {$productId = "O365HomePremRetail"; $productName = 'Microsoft 365 Home'; Invoke-Command $install}
  if ($m365Business.Checked -eq $true) {$productId = "O365BusinessRetail"; $productName = 'Microsoft 365 Apps for Business'; Invoke-Command $install}
  if ($m365Enterprise.Checked -eq $true) {$productId = "O365ProPlusRetail"; $productName = 'Microsoft 365 Apps for Enterprise'; Invoke-Command $install}

  if ($2021Pro.Checked -eq $true) {$productId = "ProPlus2021$licType"; $productName = 'Office 2021 Professional LTSC 2021'; Invoke-Command $install}
  if ($2021Std.Checked -eq $true) {$productId = "Standard2021$licType"; $productName = 'Office 2021 Standard LTSC'; Invoke-Command $install}
  if ($2021ProjectPro.Checked -eq $true) {$productId = "ProjectPro2021$licType"; $productName = 'Project Pro 2021'; Invoke-Command $install}
  if ($2021ProjectStd.Checked -eq $true) {$productId = "ProjectStd2021$licType"; $productName = 'Project Standard 2021'; Invoke-Command $install}
  if ($2021VisioPro.Checked -eq $true) {$productId = "VisioPro2021$licType"; $productName = 'Visio Pro 2021'; Invoke-Command $install}
  if ($2021VisioStd.Checked -eq $true) {$productId = "VisioStd2021$licType"; $productName = 'Visio Standard 2021'; Invoke-Command $install}
  if ($2021Word.Checked -eq $true) {$productId = "Word2021$licType"; $productName = 'Microsoft Word LTSC 2021'; Invoke-Command $install}
  if ($2021Excel.Checked -eq $true) {$productId = "Excel2021$licType"; $productName = 'Microsoft Excel LTSC 2021'; Invoke-Command $install}
  if ($2021PowerPoint.Checked -eq $true) {$productId = "PowerPoint2021$licType"; $productName = 'Microsoft PowerPoint LTSC 2021'; Invoke-Command $install}
  if ($2021Outlook.Checked -eq $true) {$productId = "Outlook2021$licType"; $productName = 'Microsoft Outlook LTSC 2021'; Invoke-Command $install}
  if ($2021Publisher.Checked -eq $true) {$productId = "Publisher2021$licType"; $productName = 'Microsoft Publisher LTSC 2021'; Invoke-Command $install}
  if ($2021Access.Checked -eq $true) {$productId = "Access2021$licType"; $productName = 'Microsoft Access LTSC 2021'; Invoke-Command $install}
  if ($2021HomeBusiness.Checked -eq $true) {$productId = "HomeBusiness2021Retail"; $productName = 'Office HomeBusiness 2021'; Invoke-Command $install}
  if ($2021HomeStudent.Checked -eq $true) {$productId = "HomeStudent2021Retail"; $productName = 'Office HomeStudent LTSC 2021'; Invoke-Command $install}

  if ($2019Pro.Checked -eq $true) {$productId = "ProPlus2019$licType"; $productName = 'Office 2019 Professional Plus'; Invoke-Command $install}
  if ($2019Std.Checked -eq $true) {$productId = "Standard2019$licType"; $productName = 'Office 2019 Standard'; Invoke-Command $install}
  if ($2019ProjectPro.Checked -eq $true) {$productId = "ProjectPro2019$licType"; $productName = 'Project Pro 2019'; Invoke-Command $install}
  if ($2019ProjectStd.Checked -eq $true) {$productId = "ProjectStd2019$licType"; $productName = 'Project Standard 2019'; Invoke-Command $install}
  if ($2019VisioPro.Checked -eq $true) {$productId = "VisioPro2019$licType"; $productName = 'Visio Pro 2019'; Invoke-Command $install}
  if ($2019VisioStd.Checked -eq $true) {$productId = "VisioStd2019$licType"; $productName = 'Visio Standard 2019'; Invoke-Command $install}
  if ($2019Word.Checked -eq $true) {$productId = "Word2019$licType"; $productName = 'Microsoft Word 2019'; Invoke-Command $install}
  if ($2019Excel.Checked -eq $true) {$productId = "Excel2019$licType"; $productName = 'Microsoft Excel 2019'; Invoke-Command $install}
  if ($2019PowerPoint.Checked -eq $true) {$productId = "PowerPoint2019$licType"; $productName = 'Microsoft PowerPoint 201p'; Invoke-Command $install}
  if ($2019Outlook.Checked -eq $true) {$productId = "Outlook2019$licType"; $productName = 'Microsoft Outlook 2019'; Invoke-Command $install}
  if ($2019Publisher.Checked -eq $true) {$productId = "Publisher2019$licType"; $productName = 'Microsoft Publisher 2019'; Invoke-Command $install}
  if ($2019Access.Checked -eq $true) {$productId = "Access2019$licType"; $productName = 'Microsoft Access 2019'; Invoke-Command $install}
  if ($2019HomeBusiness.Checked -eq $true) {$productId = "HomeBusiness2019Retail"; $productName = 'Office HomeBusiness 2019'; Invoke-Command $install}
  if ($2019HomeStudent.Checked -eq $true) {$productId = "HomeStudent2019Retail"; $productName = 'Office HomeStudent 2019'; Invoke-Command $install}

  if ($2016Pro.Checked -eq $true) {$productId = "ProfessionalRetail"; $productName = 'Office 2016 Professional Plus'; Invoke-Command $install}
  if ($2016Std.Checked -eq $true) {$productId = "StandardRetail"; $productName = 'Office 2016 Standard'; Invoke-Command $install}
  if ($2016ProjectPro.Checked -eq $true) {$productId = "ProjectProRetail"; $productName = 'Microsoft Project Pro 2016'; Invoke-Command $install}
  if ($2016ProjectStd.Checked -eq $true) {$productId = "ProjectStdRetail"; $productName = 'Microsoft Project Standard 2016'; Invoke-Command $install}
  if ($2016VisioPro.Checked -eq $true) {$productId = "VisioProRetail"; $productName = 'Microsoft Visio Pro 2016'; Invoke-Command $install}
  if ($2016VisioStd.Checked -eq $true) {$productId = "VisioStdRetail"; $productName = 'Microsoft Visio Standard 2016'; Invoke-Command $install}
  if ($2016Word.Checked -eq $true) {$productId = "WordRetail"; $productName = 'Microsoft Word 2016'; Invoke-Command $install}
  if ($2016Excel.Checked -eq $true) {$productId = "ExcelRetail"; $productName = 'Microsoft Excel 2016'; Invoke-Command $install}
  if ($2016PowerPoint.Checked -eq $true) {$productId = "PowerPointRetail"; $productName = 'Microsoft PowerPoint 2016'; Invoke-Command $install}
  if ($2016Outlook.Checked -eq $true) {$productId = "OutlookRetail"; $productName = 'Microsoft Outlook 2016'; Invoke-Command $install}
  if ($2016Publisher.Checked -eq $true) {$productId = "PublisherRetail"; $productName = 'Microsoft Publisher 2016'; Invoke-Command $install}
  if ($2016Access.Checked -eq $true) {$productId = "AccessRetail"; $productName = 'Microsoft Access 2016'; Invoke-Command $install}
  if ($2016OneNote.Checked -eq $true) {$productId = "OneNoteRetail"; $productName = 'Microsoft Onenote 2016'; Invoke-Command $install}

  if ($2013Pro.Checked -eq $true) {$productId = "ProfessionalRetail"; $productName = 'Office 2013 Professional Plus'; Invoke-Command $install2013}
  if ($2013Std.Checked -eq $true) {$productId = "StandardRetail"; $productName = 'Office 2013 Standard'; Invoke-Command $install2013}
  if ($2013ProjectPro.Checked -eq $true) {$productId = "ProjectProRetail"; $productName = 'Microsoft Project Pro 2013'; Invoke-Command $install2013}
  if ($2013ProjectStd.Checked -eq $true) {$productId = "ProjectStdRetail"; $productName = 'Microsoft Project Standard 2013'; Invoke-Command $install2013}
  if ($2013VisioPro.Checked -eq $true) {$productId = "VisioProRetail"; $productName = 'Microsoft Visio Pro 2013'; Invoke-Command $install2013}
  if ($2013VisioStd.Checked -eq $true) {$productId = "VisioStdRetail"; $productName = 'Microsoft Visio Standard 2013'; Invoke-Command $install2013}
  if ($2013Word.Checked -eq $true) {$productId = "WordRetail"; $productName = 'Microsoft Word 2013'; Invoke-Command $install2013}
  if ($2013Excel.Checked -eq $true) {$productId = "ExcelRetail"; $productName = 'Microsoft Excel 2013'; Invoke-Command $install2013}
  if ($2013PowerPoint.Checked -eq $true) {$productId = "PowerPointRetail"; $productName = 'Microsoft PowerPoint 2013'; Invoke-Command $install2013}
  if ($2013Outlook.Checked -eq $true) {$productId = "OutlookRetail"; $productName = 'Microsoft Outlook 2013'; Invoke-Command $install2013}
  if ($2013Publisher.Checked -eq $true) {$productId = "PublisherRetail"; $productName = 'Microsoft Publisher 2013'; Invoke-Command $install2013}
  if ($2013Access.Checked -eq $true) {$productId = "AccessRetail"; $productName = 'Microsoft Access 2013'; Invoke-Command $install2013}
  }
  catch {}

}

function Uninstall-AllOffice {
  try {
    if ($uninstallcb.Checked -eq $true) {Invoke-Command $uninstall}
  }
  catch {
    
  }
  
}

############################################## end functions

############################################## Start group boxes

  $arch = New-Object System.Windows.Forms.GroupBox
  $arch.Location = New-Object System.Drawing.Size(10,10) 
  $arch.size = New-Object System.Drawing.Size(130,70) 
  $arch.text = "Arch:"
  $arch.Font = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Regular)
  $arch.ForeColor = [System.Drawing.Color]::DarkBlue
  $Form.Controls.Add($arch)

  $licenseType = New-Object System.Windows.Forms.GroupBox
  $licenseType.Location = New-Object System.Drawing.Size(10,90) 
  $licenseType.size = New-Object System.Drawing.Size(130,70) 
  $licenseType.text = "License Type:"
  $licenseType.Font = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Regular)
  $licenseType.ForeColor = [System.Drawing.Color]::DarkBlue
  $Form.Controls.Add($licenseType)

  $installMode = New-Object System.Windows.Forms.GroupBox
  $installMode.Location = New-Object System.Drawing.Size(10,170) 
  $installMode.size = New-Object System.Drawing.Size(130,110) 
  $installMode.text = "Mode:"
  $installMode.Font = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Regular)
  $installMode.ForeColor = [System.Drawing.Color]::DarkBlue
  $Form.Controls.Add($installMode) 

  $language = New-Object System.Windows.Forms.GroupBox
  $language.Location = New-Object System.Drawing.Size(155,110) 
  $language.size = New-Object System.Drawing.Size(130,170) 
  $language.text = "Language:"
  $language.ForeColor = [System.Drawing.Color]::DarkBlue
  $language.Font = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Regular)
  $Form.Controls.Add($language) 

  $groupBox365 = New-Object System.Windows.Forms.GroupBox
  $groupBox365.Location = New-Object System.Drawing.Size(155,10) 
  $groupBox365.size = New-Object System.Drawing.Size(130,90) 
  $groupBox365.text = "Microsoft 365:"
  $groupBox365.Font = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Regular)
  $groupBox365.ForeColor = [System.Drawing.Color]::DarkRed
  $Form.Controls.Add($groupBox365) 

  $groupBox2021 = New-Object System.Windows.Forms.GroupBox
  $groupBox2021.Location = New-Object System.Drawing.Size(300,10) 
  $groupBox2021.size = New-Object System.Drawing.Size(150,310) 
  $groupBox2021.text = "Office 2021 Apps:"
  $groupBox2021.Font = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Regular)
  $groupBox2021.ForeColor = [System.Drawing.Color]::DarkRed
  $Form.Controls.Add($groupBox2021)

  $groupBox2019 = New-Object System.Windows.Forms.GroupBox
  $groupBox2019.Location = New-Object System.Drawing.Size(465,10) 
  $groupBox2019.size = New-Object System.Drawing.Size(150,310) 
  $groupBox2019.text = "Office 2019 Apps:"
  $groupBox2019.Font = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Regular)
  $groupBox2019.ForeColor = [System.Drawing.Color]::DarkRed
  $Form.Controls.Add($groupBox2019)

  $groupBox2016 = New-Object System.Windows.Forms.GroupBox
  $groupBox2016.Location = New-Object System.Drawing.Size(630,10) 
  $groupBox2016.size = New-Object System.Drawing.Size(150,310) 
  $groupBox2016.text = "Office 2016 Apps:"
  $groupBox2016.Font = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Regular)
  $groupBox2016.ForeColor = [System.Drawing.Color]::DarkRed
  $Form.Controls.Add($groupBox2016)

  $groupBox2013 = New-Object System.Windows.Forms.GroupBox
  $groupBox2013.Location = New-Object System.Drawing.Size(795,10) 
  $groupBox2013.size = New-Object System.Drawing.Size(150,310)
  $groupBox2013.text = "Office 2013 Apps:"
  $groupBox2013.Font = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Regular)
  $groupBox2013.ForeColor = [System.Drawing.Color]::DarkRed
  $Form.Controls.Add($groupBox2013)

  $removeButton = New-Object System.Windows.Forms.Button 
  $removeButton.Cursor = [System.Windows.Forms.Cursors]::Hand
  $removeButton.Location = New-Object System.Drawing.Size(830,345) 
  $removeButton.Size = New-Object System.Drawing.Size(90,30) 
  $removeButton.Text = "Remove All"
  $removeButton.BackColor = [System.Drawing.Color]::Red
  $removeButton.ForeColor = [System.Drawing.Color]::White
  $removeButton.Font = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Bold)
  $removeButton.Add_Click({Uninstall-AllOffice})
  $Form.Controls.Add($removeButton)

  $groupBoxUninstall = New-Object System.Windows.Forms.GroupBox
  $groupBoxUninstall.Location = New-Object System.Drawing.Size(630,330) 
  $groupBoxUninstall.size = New-Object System.Drawing.Size(315,60) 
  $groupBoxUninstall.text = "Remove All Office Apps:"
  $groupBoxUninstall.Font = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Regular)
  $groupBoxUninstall.ForeColor = [System.Drawing.Color]::Red
  $Form.Controls.Add($groupBoxUninstall)

  ############################################## Start buttons

  $submitButton = New-Object System.Windows.Forms.Button 
  $submitButton.Cursor = [System.Windows.Forms.Cursors]::Hand
  $submitButton.Location = New-Object System.Drawing.Size(10,290) 
  $submitButton.Size = New-Object System.Drawing.Size(130,40) 
  $submitButton.Text = "Submit"
  $submitButton.BackColor = [System.Drawing.Color]::Green
  $submitButton.ForeColor = [System.Drawing.Color]::White
  $submitButton.Font = New-Object System.Drawing.Font("Consolas",9,[System.Drawing.FontStyle]::Bold)
  $submitButton.Add_Click({microsoftInstaller})
  $Form.Controls.Add($submitButton)

  $AboutLabel = New-Object System.Windows.Forms.Label
  $AboutLabel.Location = New-Object System.Drawing.Size(10,340)
  $AboutLabel.AutoSize = $True 
  $AboutLabel.Text = "(*) Default mode is Install. If you want to download only, select the Download mode."
  $Form.Controls.Add($AboutLabel)

  $AboutLabel = New-Object System.Windows.Forms.Label
  $AboutLabel.Location = New-Object System.Drawing.Size(10,360)
  $AboutLabel.AutoSize = $True 
  $AboutLabel.Text = "(*) By default, this script installs the 64-bit version in English."
  $Form.Controls.Add($AboutLabel)

  $AboutLabel2 = New-Object System.Windows.Forms.Label
  $AboutLabel2.Location = New-Object System.Drawing.Size(10,380)
  $AboutLabel2.AutoSize = $True 
  $AboutLabel2.Text = "(*) To install other apps, re-run the script."
  $Form.Controls.Add($AboutLabel2)

########################################
  $linklabel = New-Object System.Windows.Forms.LinkLabel
  $linklabel.Text = "(*) For more: https://msgang.com - Free Microsoft products for everyone."
  $linklabel.Location = New-Object System.Drawing.Size(10,400) 
  $linklabel.AutoSize = $True

  #Sample hyperlinks to add to the text of the link label control.
  $URLInfo = [pscustomobject]@{
    StartPos = 14;
    LinkLength = 18;
    Url = 'http://msgang.com'
  }
  #Add them.
  foreach ($URL in $URLinfo) {
    $null = $linklabel.Links.Add($URL.StartPos, $URL.LinkLength, $URL.URL)
  }
  #Register a handler for when the user clicks a link.
  $linklabel.add_LinkClicked({
    param($evtSender, $evtArgs)
    #Launch the default browser with the target URL.
    Start-Process $evtArgs.Link.LinkData
  })

  $form.Controls.Add($linklabel)

  $RemoveLable = New-Object System.Windows.Forms.Label
  $RemoveLable.Location = New-Object System.Drawing.Size(625,395)
  $RemoveLable.AutoSize = $True 
  $RemoveLable.Text = "(*) This option removes all installed Office apps."
  $Form.Controls.Add($RemoveLable)

############################################## end group boxes

############################################## Start Arch checkboxes

  $arch64 = New-Object System.Windows.Forms.RadioButton
  $arch64.Location = New-Object System.Drawing.Size(10,20)
  $arch64.Size = New-Object System.Drawing.Size(110,20)
  $arch64.Checked = $true
  $arch64.Text = "64 bit"
  $arch.Controls.Add($arch64)

  $arch32 = New-Object System.Windows.Forms.RadioButton
  $arch32.Location = New-Object System.Drawing.Size(10,40)
  $arch32.Size = New-Object System.Drawing.Size(110,20)
  $arch32.Checked = $false
  $arch32.Text = "32 bit"
  $arch.Controls.Add($arch32)

############################################## Start LicenseType checkboxes

  $licenseTypeVolume = New-Object System.Windows.Forms.RadioButton
  $licenseTypeVolume.Location = New-Object System.Drawing.Size(10,20)
  $licenseTypeVolume.Size = New-Object System.Drawing.Size(110,20)
  $licenseTypeVolume.Checked = $true
  $licenseTypeVolume.Text = "Volume"
  $licenseType.Controls.Add($licenseTypeVolume)

  $licenseTypeRetail = New-Object System.Windows.Forms.RadioButton
  $licenseTypeRetail.Location = New-Object System.Drawing.Size(10,40)
  $licenseTypeRetail.Size = New-Object System.Drawing.Size(110,20)
  $licenseTypeRetail.Checked = $false
  $licenseTypeRetail.Text = "Retail"
  $licenseType.Controls.Add($licenseTypeRetail)

  ############################################## Start installMode checkboxes

  $installModeSetup = New-Object System.Windows.Forms.RadioButton
  $installModeSetup.Location = New-Object System.Drawing.Size(10,20)
  $installModeSetup.Size = New-Object System.Drawing.Size(110,20)
  $installModeSetup.Checked = $true
  $installModeSetup.Text = "Install"
  $installMode.Controls.Add($installModeSetup)

  $installModeDownload = New-Object System.Windows.Forms.RadioButton
  $installModeDownload.Location = New-Object System.Drawing.Size(10,40)
  $installModeDownload.Size = New-Object System.Drawing.Size(110,20)
  $installModeDownload.Checked = $false
  $installModeDownload.Text = "Download"
  $installMode.Controls.Add($installModeDownload)

  $installModeActivate = New-Object System.Windows.Forms.RadioButton
  $installModeActivate.Location = New-Object System.Drawing.Size(10,60)
  $installModeActivate.Size = New-Object System.Drawing.Size(110,20)
  $installModeActivate.Checked = $false
  $installModeActivate.Text = "Activate"
  $installMode.Controls.Add($installModeActivate)

############################################## Start Arch checkboxes

  $English = New-Object System.Windows.Forms.RadioButton
  $English.Location = New-Object System.Drawing.Size(10,20)
  $English.Size = New-Object System.Drawing.Size(110,20)
  $English.Checked = $true
  $English.Text = "English"
  $language.Controls.Add($English)

  $Japanese = New-Object System.Windows.Forms.RadioButton
  $Japanese.Location = New-Object System.Drawing.Size(10,40)
  $Japanese.Size = New-Object System.Drawing.Size(110,20)
  $Japanese.Text = "Japanese"
  $language.Controls.Add($Japanese)

  $Korean = New-Object System.Windows.Forms.RadioButton
  $Korean.Location = New-Object System.Drawing.Size(10,60)
  $Korean.Size = New-Object System.Drawing.Size(110,20)
  $Korean.Text = "Korean"
  $language.Controls.Add($Korean)

  $Chinese = New-Object System.Windows.Forms.RadioButton
  $Chinese.Location = New-Object System.Drawing.Size(10,80)
  $Chinese.Size = New-Object System.Drawing.Size(110,20)
  $Chinese.Text = "Chinese"
  $language.Controls.Add($Chinese)

  $French = New-Object System.Windows.Forms.RadioButton
  $French.Location = New-Object System.Drawing.Size(10,100)
  $French.Size = New-Object System.Drawing.Size(110,20)
  $French.Text = "French"
  $language.Controls.Add($French)

  $Spanish = New-Object System.Windows.Forms.RadioButton
  $Spanish.Location = New-Object System.Drawing.Size(10,120)
  $Spanish.Size = New-Object System.Drawing.Size(110,20)
  $Spanish.Text = "Spanish"
  $language.Controls.Add($Spanish)

  $Vietnamese = New-Object System.Windows.Forms.RadioButton
  $Vietnamese.Location = New-Object System.Drawing.Size(10,140)
  $Vietnamese.Size = New-Object System.Drawing.Size(110,20)
  $Vietnamese.Text = "Vietnamese"
  $language.Controls.Add($Vietnamese)

############################################## End Arch checkboxes 

############################################## Start Microsoft 365 checkboxes
  $m365Home = New-Object System.Windows.Forms.RadioButton
  $m365Home.Location = New-Object System.Drawing.Size(10,20)
  $m365Home.Size = New-Object System.Drawing.Size(110,20)
  $m365Home.Checked = $false
  $m365Home.Text = "Home"
  $groupBox365.Controls.Add($m365Home)

  $m365Business = New-Object System.Windows.Forms.RadioButton
  $m365Business.Location = New-Object System.Drawing.Size(10,40)
  $m365Business.Size = New-Object System.Drawing.Size(110,20)
  $m365Business.Text = "Business"
  $groupBox365.Controls.Add($m365Business)

  $m365Enterprise = New-Object System.Windows.Forms.RadioButton
  $m365Enterprise.Location = New-Object System.Drawing.Size(10,60)
  $m365Enterprise.Size = New-Object System.Drawing.Size(110,20)
  $m365Enterprise.Text = "Enterprise"
  $groupBox365.Controls.Add($m365Enterprise)
############################################## End Microsoft 365 checkboxes

############################################## Start Office 2021 checkboxes
  $2021Pro = New-Object System.Windows.Forms.RadioButton
  $2021Pro.Location = New-Object System.Drawing.Size(10,20)
  $2021Pro.Size = New-Object System.Drawing.Size(110,20)
  $2021Pro.Checked = $false
  $2021Pro.Text = "Professional"
  $groupBox2021.Controls.Add($2021Pro)

  $2021Std = New-Object System.Windows.Forms.RadioButton
  $2021Std.Location = New-Object System.Drawing.Size(10,40)
  $2021Std.Size = New-Object System.Drawing.Size(110,20)
  $2021Std.Text = "Standard"
  $groupBox2021.Controls.Add($2021Std)

  $2021ProjectPro = New-Object System.Windows.Forms.RadioButton
  $2021ProjectPro.Location = New-Object System.Drawing.Size(10,60)
  $2021ProjectPro.Size = New-Object System.Drawing.Size(110,20)
  $2021ProjectPro.Text = "Project Pro"
  $groupBox2021.Controls.Add($2021ProjectPro)

  $2021ProjectStd = New-Object System.Windows.Forms.RadioButton
  $2021ProjectStd.Location = New-Object System.Drawing.Size(10,80)
  $2021ProjectStd.Size = New-Object System.Drawing.Size(110,20)
  $2021ProjectStd.AutoSize = $true
  $2021ProjectStd.Text = "Project Standard"
  $groupBox2021.Controls.Add($2021ProjectStd)

  $2021VisioPro = New-Object System.Windows.Forms.RadioButton
  $2021VisioPro.Location = New-Object System.Drawing.Size(10,100)
  $2021VisioPro.Size = New-Object System.Drawing.Size(110,20)
  $2021VisioPro.Text = "Visio Pro"
  $groupBox2021.Controls.Add($2021VisioPro)

  $2021VisioStd = New-Object System.Windows.Forms.RadioButton
  $2021VisioStd.Location = New-Object System.Drawing.Size(10,120)
  $2021VisioStd.Size = New-Object System.Drawing.Size(110,20)
  $2021VisioStd.Text = "Visio Standard"
  $groupBox2021.Controls.Add($2021VisioStd)

  $2021Word = New-Object System.Windows.Forms.RadioButton
  $2021Word.Location = New-Object System.Drawing.Size(10,140)
  $2021Word.Size = New-Object System.Drawing.Size(110,20)
  $2021Word.Text = "Word"
  $groupBox2021.Controls.Add($2021Word)

  $2021Excel = New-Object System.Windows.Forms.RadioButton
  $2021Excel.Location = New-Object System.Drawing.Size(10,160)
  $2021Excel.Size = New-Object System.Drawing.Size(110,20)
  $2021Excel.Text = "Excel"
  $groupBox2021.Controls.Add($2021Excel)

  $2021PowerPoint = New-Object System.Windows.Forms.RadioButton
  $2021PowerPoint.Location = New-Object System.Drawing.Size(10,180)
  $2021PowerPoint.Size = New-Object System.Drawing.Size(110,20)
  $2021PowerPoint.Text = "PowerPoint"
  $groupBox2021.Controls.Add($2021PowerPoint)

  $2021Outlook = New-Object System.Windows.Forms.RadioButton
  $2021Outlook.Location = New-Object System.Drawing.Size(10,200)
  $2021Outlook.Size = New-Object System.Drawing.Size(110,20)
  $2021Outlook.Text = "Outlook"
  $groupBox2021.Controls.Add($2021Outlook)

  $2021Publisher = New-Object System.Windows.Forms.RadioButton
  $2021Publisher.Location = New-Object System.Drawing.Size(10,220)
  $2021Publisher.Size = New-Object System.Drawing.Size(110,20)
  $2021Publisher.Text = "Publisher"
  $groupBox2021.Controls.Add($2021Publisher)

  $2021Access = New-Object System.Windows.Forms.RadioButton
  $2021Access.Location = New-Object System.Drawing.Size(10,240)
  $2021Access.Size = New-Object System.Drawing.Size(110,20)
  $2021Access.Text = "Access"
  $groupBox2021.Controls.Add($2021Access)

  $2021HomeBusiness = New-Object System.Windows.Forms.RadioButton
  $2021HomeBusiness.Location = New-Object System.Drawing.Size(10,260)
  $2021HomeBusiness.Size = New-Object System.Drawing.Size(110,20)
  $2021HomeBusiness.Text = "HomeBusiness"
  $groupBox2021.Controls.Add($2021HomeBusiness)

  $2021HomeStudent = New-Object System.Windows.Forms.RadioButton
  $2021HomeStudent.Location = New-Object System.Drawing.Size(10,280)
  $2021HomeStudent.Size = New-Object System.Drawing.Size(110,20)
  $2021HomeStudent.Text = "HomeStudent"
  $groupBox2021.Controls.Add($2021HomeStudent)
############################################## End Office 2021 checkboxes

############################################## Start Office 2019 checkboxes
  $2019Pro = New-Object System.Windows.Forms.RadioButton
  $2019Pro.Location = New-Object System.Drawing.Size(10,20)
  $2019Pro.Size = New-Object System.Drawing.Size(110,20)
  $2019Pro.Checked = $false
  $2019Pro.Text = "Professional"
  $groupBox2019.Controls.Add($2019Pro)

  $2019Std = New-Object System.Windows.Forms.RadioButton
  $2019Std.Location = New-Object System.Drawing.Size(10,40)
  $2019Std.Size = New-Object System.Drawing.Size(110,20)
  $2019Std.Text = "Standard"
  $groupBox2019.Controls.Add($2019Std)

  $2019ProjectPro = New-Object System.Windows.Forms.RadioButton
  $2019ProjectPro.Location = New-Object System.Drawing.Size(10,60)
  $2019ProjectPro.Size = New-Object System.Drawing.Size(110,20)
  $2019ProjectPro.Text = "Project Pro"
  $groupBox2019.Controls.Add($2019ProjectPro)

  $2019ProjectStd = New-Object System.Windows.Forms.RadioButton
  $2019ProjectStd.Location = New-Object System.Drawing.Size(10,80)
  $2019ProjectStd.Size = New-Object System.Drawing.Size(110,20)
  $2019ProjectStd.Text = "Project Standard"
  $2019ProjectStd.AutoSize = $true
  $groupBox2019.Controls.Add($2019ProjectStd)

  $2019VisioPro = New-Object System.Windows.Forms.RadioButton
  $2019VisioPro.Location = New-Object System.Drawing.Size(10,100)
  $2019VisioPro.Size = New-Object System.Drawing.Size(110,20)
  $2019VisioPro.Text = "Visio Pro"
  $groupBox2019.Controls.Add($2019VisioPro)

  $2019VisioStd = New-Object System.Windows.Forms.RadioButton
  $2019VisioStd.Location = New-Object System.Drawing.Size(10,120)
  $2019VisioStd.Size = New-Object System.Drawing.Size(110,20)
  $2019VisioStd.Text = "Visio Standard"
  $groupBox2019.Controls.Add($2019VisioStd)

  $2019Word = New-Object System.Windows.Forms.RadioButton
  $2019Word.Location = New-Object System.Drawing.Size(10,140)
  $2019Word.Size = New-Object System.Drawing.Size(110,20)
  $2019Word.Text = "Word"
  $groupBox2019.Controls.Add($2019Word)

  $2019Excel = New-Object System.Windows.Forms.RadioButton
  $2019Excel.Location = New-Object System.Drawing.Size(10,160)
  $2019Excel.Size = New-Object System.Drawing.Size(110,20)
  $2019Excel.Text = "Excel"
  $groupBox2019.Controls.Add($2019Excel)

  $2019PowerPoint = New-Object System.Windows.Forms.RadioButton
  $2019PowerPoint.Location = New-Object System.Drawing.Size(10,180)
  $2019PowerPoint.Size = New-Object System.Drawing.Size(110,20)
  $2019PowerPoint.Text = "PowerPoint"
  $groupBox2019.Controls.Add($2019PowerPoint)

  $2019Outlook = New-Object System.Windows.Forms.RadioButton
  $2019Outlook.Location = New-Object System.Drawing.Size(10,200)
  $2019Outlook.Size = New-Object System.Drawing.Size(110,20)
  $2019Outlook.Text = "Outlook"
  $groupBox2019.Controls.Add($2019Outlook)

  $2019Publisher = New-Object System.Windows.Forms.RadioButton
  $2019Publisher.Location = New-Object System.Drawing.Size(10,220)
  $2019Publisher.Size = New-Object System.Drawing.Size(110,20)
  $2019Publisher.Text = "Publisher"
  $groupBox2019.Controls.Add($2019Publisher)

  $2019Access = New-Object System.Windows.Forms.RadioButton
  $2019Access.Location = New-Object System.Drawing.Size(10,240)
  $2019Access.Size = New-Object System.Drawing.Size(110,20)
  $2019Access.Text = "Access"
  $groupBox2019.Controls.Add($2019Access)

  $2019HomeBusiness = New-Object System.Windows.Forms.RadioButton
  $2019HomeBusiness.Location = New-Object System.Drawing.Size(10,260)
  $2019HomeBusiness.Size = New-Object System.Drawing.Size(110,20)
  $2019HomeBusiness.Text = "HomeBusiness"
  $groupBox2019.Controls.Add($2019HomeBusiness)

  $2019HomeStudent = New-Object System.Windows.Forms.RadioButton
  $2019HomeStudent.Location = New-Object System.Drawing.Size(10,280)
  $2019HomeStudent.Size = New-Object System.Drawing.Size(110,20)
  $2019HomeStudent.Text = "HomeStudent"
  $groupBox2019.Controls.Add($2019HomeStudent)
############################################## End Office 2019 checkboxes


############################################## Start Office 2016 checkboxes
  $2016Pro = New-Object System.Windows.Forms.RadioButton
  $2016Pro.Location = New-Object System.Drawing.Size(10,20)
  $2016Pro.Size = New-Object System.Drawing.Size(110,20)
  $2016Pro.Checked = $false
  $2016Pro.Text = "Professional"
  $groupBox2016.Controls.Add($2016Pro)

  $2016Std = New-Object System.Windows.Forms.RadioButton
  $2016Std.Location = New-Object System.Drawing.Size(10,40)
  $2016Std.Size = New-Object System.Drawing.Size(110,20)
  $2016Std.Text = "Standard"
  $groupBox2016.Controls.Add($2016Std)

  $2016ProjectPro = New-Object System.Windows.Forms.RadioButton
  $2016ProjectPro.Location = New-Object System.Drawing.Size(10,60)
  $2016ProjectPro.Size = New-Object System.Drawing.Size(110,20)
  $2016ProjectPro.Text = "Project Pro"
  $groupBox2016.Controls.Add($2016ProjectPro)

  $2016ProjectStd = New-Object System.Windows.Forms.RadioButton
  $2016ProjectStd.Location = New-Object System.Drawing.Size(10,80)
  $2016ProjectStd.Size = New-Object System.Drawing.Size(110,20)
  $2016ProjectStd.Text = "Project Standard"
  $2016ProjectStd.AutoSize = $true
  $groupBox2016.Controls.Add($2016ProjectStd)

  $2016VisioPro = New-Object System.Windows.Forms.RadioButton
  $2016VisioPro.Location = New-Object System.Drawing.Size(10,100)
  $2016VisioPro.Size = New-Object System.Drawing.Size(110,20)
  $2016VisioPro.Text = "Visio Pro"
  $groupBox2016.Controls.Add($2016VisioPro)

  $2016VisioStd = New-Object System.Windows.Forms.RadioButton
  $2016VisioStd.Location = New-Object System.Drawing.Size(10,120)
  $2016VisioStd.Size = New-Object System.Drawing.Size(110,20)
  $2016VisioStd.Text = "Visio Standard"
  $groupBox2016.Controls.Add($2016VisioStd)

  $2016Word = New-Object System.Windows.Forms.RadioButton
  $2016Word.Location = New-Object System.Drawing.Size(10,140)
  $2016Word.Size = New-Object System.Drawing.Size(110,20)
  $2016Word.Text = "Word"
  $groupBox2016.Controls.Add($2016Word)

  $2016Excel = New-Object System.Windows.Forms.RadioButton
  $2016Excel.Location = New-Object System.Drawing.Size(10,160)
  $2016Excel.Size = New-Object System.Drawing.Size(110,20)
  $2016Excel.Text = "Excel"
  $groupBox2016.Controls.Add($2016Excel)

  $2016PowerPoint = New-Object System.Windows.Forms.RadioButton
  $2016PowerPoint.Location = New-Object System.Drawing.Size(10,180)
  $2016PowerPoint.Size = New-Object System.Drawing.Size(110,20)
  $2016PowerPoint.Text = "PowerPoint"
  $groupBox2016.Controls.Add($2016PowerPoint)

  $2016Outlook = New-Object System.Windows.Forms.RadioButton
  $2016Outlook.Location = New-Object System.Drawing.Size(10,200)
  $2016Outlook.Size = New-Object System.Drawing.Size(110,20)
  $2016Outlook.Text = "Outlook"
  $groupBox2016.Controls.Add($2016Outlook)

  $2016Publisher = New-Object System.Windows.Forms.RadioButton
  $2016Publisher.Location = New-Object System.Drawing.Size(10,220)
  $2016Publisher.Size = New-Object System.Drawing.Size(110,20)
  $2016Publisher.Text = "Publisher"
  $groupBox2016.Controls.Add($2016Publisher)

  $2016Access = New-Object System.Windows.Forms.RadioButton
  $2016Access.Location = New-Object System.Drawing.Size(10,240)
  $2016Access.Size = New-Object System.Drawing.Size(110,20)
  $2016Access.Text = "Access"
  $groupBox2016.Controls.Add($2016Access)

  $2016OneNote = New-Object System.Windows.Forms.RadioButton
  $2016OneNote.Location = New-Object System.Drawing.Size(10,260)
  $2016OneNote.Size = New-Object System.Drawing.Size(110,20)
  $2016OneNote.Text = "OneNote"
  $groupBox2016.Controls.Add($2016OneNote)

############################################## End Office 2016 checkboxes

############################################## Start Office 2013 checkboxes
  $2013Pro = New-Object System.Windows.Forms.RadioButton
  $2013Pro.Location = New-Object System.Drawing.Size(10,20)
  $2013Pro.Size = New-Object System.Drawing.Size(110,20)
  $2013Pro.Checked = $false
  $2013Pro.Text = "Professional"
  $groupBox2013.Controls.Add($2013Pro)

  $2013Std = New-Object System.Windows.Forms.RadioButton
  $2013Std.Location = New-Object System.Drawing.Size(10,40)
  $2013Std.Size = New-Object System.Drawing.Size(110,20)
  $2013Std.Text = "Standard"
  $groupBox2013.Controls.Add($2013Std)

  $2013ProjectPro = New-Object System.Windows.Forms.RadioButton
  $2013ProjectPro.Location = New-Object System.Drawing.Size(10,60)
  $2013ProjectPro.Size = New-Object System.Drawing.Size(110,20)
  $2013ProjectPro.Text = "Project Pro"
  $groupBox2013.Controls.Add($2013ProjectPro)

  $2013ProjectStd = New-Object System.Windows.Forms.RadioButton
  $2013ProjectStd.Location = New-Object System.Drawing.Size(10,80)
  $2013ProjectStd.Size = New-Object System.Drawing.Size(110,20)
  $2013ProjectStd.Text = "Project Standard"
  $2013ProjectStd.AutoSize = $true
  $groupBox2013.Controls.Add($2013ProjectStd)

  $2013VisioPro = New-Object System.Windows.Forms.RadioButton
  $2013VisioPro.Location = New-Object System.Drawing.Size(10,100)
  $2013VisioPro.Size = New-Object System.Drawing.Size(110,20)
  $2013VisioPro.Text = "Visio Pro"
  $groupBox2013.Controls.Add($2013VisioPro)

  $2013VisioStd = New-Object System.Windows.Forms.RadioButton
  $2013VisioStd.Location = New-Object System.Drawing.Size(10,120)
  $2013VisioStd.Size = New-Object System.Drawing.Size(110,20)
  $2013VisioStd.Text = "Visio Standard"
  $groupBox2013.Controls.Add($2013VisioStd)

  $2013Word = New-Object System.Windows.Forms.RadioButton
  $2013Word.Location = New-Object System.Drawing.Size(10,140)
  $2013Word.Size = New-Object System.Drawing.Size(110,20)
  $2013Word.Text = "Word"
  $groupBox2013.Controls.Add($2013Word)

  $2013Excel = New-Object System.Windows.Forms.RadioButton
  $2013Excel.Location = New-Object System.Drawing.Size(10,160)
  $2013Excel.Size = New-Object System.Drawing.Size(110,20)
  $2013Excel.Text = "Excel"
  $groupBox2013.Controls.Add($2013Excel)

  $2013PowerPoint = New-Object System.Windows.Forms.RadioButton
  $2013PowerPoint.Location = New-Object System.Drawing.Size(10,180)
  $2013PowerPoint.Size = New-Object System.Drawing.Size(110,20)
  $2013PowerPoint.Text = "PowerPoint"
  $groupBox2013.Controls.Add($2013PowerPoint)

  $2013Outlook = New-Object System.Windows.Forms.RadioButton
  $2013Outlook.Location = New-Object System.Drawing.Size(10,200)
  $2013Outlook.Size = New-Object System.Drawing.Size(110,20)
  $2013Outlook.Text = "Outlook"
  $groupBox2013.Controls.Add($2013Outlook)

  $2013Publisher = New-Object System.Windows.Forms.RadioButton
  $2013Publisher.Location = New-Object System.Drawing.Size(10,220)
  $2013Publisher.Size = New-Object System.Drawing.Size(110,20)
  $2013Publisher.Text = "Publisher"
  $groupBox2013.Controls.Add($2013Publisher)

  $2013Access = New-Object System.Windows.Forms.RadioButton
  $2013Access.Location = New-Object System.Drawing.Size(10,240)
  $2013Access.Size = New-Object System.Drawing.Size(110,20)
  $2013Access.Text = "Access"
  $groupBox2013.Controls.Add($2013Access)

############################################## Start uninstall checkbox

$uninstallcb = New-Object System.Windows.Forms.RadioButton
$uninstallcb.Location = New-Object System.Drawing.Size(10,25)
$uninstallcb.Size = New-Object System.Drawing.Size(200,20)
$uninstallcb.Text = "I Agree (Be careful)"
$groupBoxUninstall.Controls.Add($uninstallcb)

############################################## end buttons

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
