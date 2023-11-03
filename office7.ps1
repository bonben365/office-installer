[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
[void] [Reflection.Assembly]::LoadWithPartialName("PresentationCore")

$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(590,250)
$Form.StartPosition = "CenterScreen" 
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow 
$Form.Text = "Microsoft Office Installation Toool for Windows 7- www.bonguides.com" 
$Form.ShowInTaskbar = $True
$Form.KeyPreview = $True
$Form.AutoSize = $True
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $False
$Form.MinimizeBox = $False
$Form.ControlBox = $True
$Form.Icon = $Icon

$install = { 
   New-Item -Path $env:temp\c2r -ItemType Directory -Force
   Set-Location $env:temp\c2r
   $fileName = "configuration-x$arch.xml"
   New-Item $fileName -ItemType File -Force | Out-Null
   Add-Content $fileName -Value '<Configuration>'
   Add-content $fileName -Value "<Add OfficeClientEdition=`"$arch`">"
   Add-content $fileName -Value "<Product ID=`"$productId`">"
   Add-content $fileName -Value "<Language ID=`"$languageId`"/>"
   Add-Content $fileName -Value '</Product>'
   Add-Content $fileName -Value '</Add>'
   Add-Content $fileName -Value '</Configuration>'

   (New-Object System.Net.WebClient).DownloadFile('https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/Files/7/setup.exe',"$env:temp\c2r\setup.exe")
   .\setup.exe /configure .\$fileName
}

$install365 = { 
   New-Item -Path $env:temp\c2r -ItemType Directory -Force
   Set-Location $env:temp\c2r
   $fileName = "configuration-x$arch.xml"
   New-Item $fileName -ItemType File -Force | Out-Null
   Add-Content $fileName -Value '<Configuration>'
   Add-content $fileName -Value "<Add OfficeClientEdition=`"$arch`"> Channel='SemiAnnualPreview' Version='16.0.12527.20880'"
   Add-content $fileName -Value "<Product ID=`"$productId`">"
   Add-content $fileName -Value "<Language ID=`"$languageId`"/>"
   Add-Content $fileName -Value '</Product>'
   Add-Content $fileName -Value '</Add>'
   Add-Content $fileName -Value '</Configuration>'

   (New-Object System.Net.WebClient).DownloadFile('https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/Files/7/setup.exe',"$env:temp\c2r\setup.exe")
   .\setup.exe /configure .\$fileName
}

$install2016 = { 
   New-Item -Path $env:temp\c2r -ItemType Directory -Force
   Set-Location $env:temp\c2r
   $fileName = "configuration-x$arch.xml"
   New-Item $fileName -ItemType File -Force | Out-Null
   Add-Content $fileName -Value '<Configuration>'
   Add-content $fileName -Value "<Add OfficeClientEdition=`"$arch`">"
   Add-content $fileName -Value "<Product ID=`"$productId`">"
   Add-content $fileName -Value "<Language ID=`"$languageId`"/>"
   Add-Content $fileName -Value '</Product>'
   Add-Content $fileName -Value '</Add>'
   Add-Content $fileName -Value '</Configuration>'

   (New-Object System.Net.WebClient).DownloadFile('https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/Files/7/setup.exe',"$env:temp\c2r\setup.exe")
   .\setup.exe /configure .\$fileName
}

$install2013 = { 
   New-Item -Path $env:temp\c2r -ItemType Directory -Force
   Set-Location $env:temp\c2r
   $fileName = "configuration-x$arch.xml"
   New-Item $fileName -ItemType File -Force | Out-Null
   Add-Content $fileName -Value '<Configuration>'
   Add-content $fileName -Value "<Add OfficeClientEdition=`"$arch`">"
   Add-content $fileName -Value "<Product ID=`"$productId`">"
   Add-content $fileName -Value "<Language ID=`"$languageId`"/>"
   Add-Content $fileName -Value '</Product>'
   Add-Content $fileName -Value '</Add>'
   Add-Content $fileName -Value '</Configuration>'

   (New-Object System.Net.WebClient).DownloadFile('https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/Files/7/bin2013.exe',"$env:temp\c2r\bin2013.exe")
   .\bin2013.exe /configure .\$fileName
}

############################################## Start functions

   function microsoftInstaller {
   try {
   if ($arch32.Checked -eq $true) {$arch='32'}
   if ($arch64.Checked -eq $true) {$arch='64'}

   if ($English.Checked -eq $true) {$languageId='en-US'}
   if ($Japanese.Checked -eq $true) {$languageId='ja-JP'}
   if ($Korean.Checked -eq $true) {$languageId='ko-KR'}
   if ($Chinese.Checked -eq $true) {$languageId='zh-TW'}
   if ($French.Checked -eq $true) {$languageId='fr-FR'}
   if ($Spanish.Checked -eq $true) {$languageId='es-ES'}
   if ($Vietnamese.Checked -eq $true) {$languageId='vi-VN'}


   if ($m365Home.Checked -eq $true) {$productId = 'O365HomePremRetail'; Invoke-Command $install365}
   if ($m365Business.Checked -eq $true) {$productId = 'O365BusinessRetail'; Invoke-Command $install365}
   if ($m365Enterprise.Checked -eq $true) {$productId = 'O365ProPlusRetail'; Invoke-Command $install365}

   if ($2016Pro.Checked -eq $true) {$productId = 'ProfessionalRetail';Invoke-Command $install2016}
   if ($2016Std.Checked -eq $true) {$productId = 'StandardRetail';Invoke-Command $install2016}
   if ($2016ProjectPro.Checked -eq $true) {$productId = 'VisioProRetail';Invoke-Command $install2016}
   if ($2016VisioPro.Checked -eq $true) {$productId = 'ProjectProRetail';Invoke-Command $install2016}

   if ($2013Pro.Checked -eq $true) {$productId = 'ProfessionalRetail';Invoke-Command $install2013}
   if ($2013Std.Checked -eq $true) {$productId = 'StandardRetail';Invoke-Command $install2013}
   if ($2013ProjectPro.Checked -eq $true) {$productId = 'VisioProRetail';Invoke-Command $install2013}
   if ($2013VisioPro.Checked -eq $true) {$productId = 'ProjectProRetail';Invoke-Command $install2013}


   } #end try

   catch {$outputBox.text = "`nOperation could not be completed"}

} 
############################################## end functions

############################################## Start group boxes

   $arch = New-Object System.Windows.Forms.GroupBox
   $arch.Location = New-Object System.Drawing.Size(150,100) 
   $arch.size = New-Object System.Drawing.Size(130,80) 
   $arch.text = "Arch:"
   $Form.Controls.Add($arch) 

   $language = New-Object System.Windows.Forms.GroupBox
   $language.Location = New-Object System.Drawing.Size(10,10) 
   $language.size = New-Object System.Drawing.Size(130,170) 
   $language.text = "Language:"
   $Form.Controls.Add($language) 

   $groupBox365 = New-Object System.Windows.Forms.GroupBox
   $groupBox365.Location = New-Object System.Drawing.Size(150,10) 
   $groupBox365.size = New-Object System.Drawing.Size(130,90) 
   $groupBox365.text = "Microsoft 365:"
   $Form.Controls.Add($groupBox365) 

   $groupBox2016 = New-Object System.Windows.Forms.GroupBox
   $groupBox2016.Location = New-Object System.Drawing.Size(290,10) 
   $groupBox2016.size = New-Object System.Drawing.Size(130,110) 
   $groupBox2016.text = "Office 2016 Apps:"
   $Form.Controls.Add($groupBox2016)

   $groupBox2013 = New-Object System.Windows.Forms.GroupBox
   $groupBox2013.Location = New-Object System.Drawing.Size(430,10) 
   $groupBox2013.size = New-Object System.Drawing.Size(130,110) 
   $groupBox2013.text = "Office 2013 Apps:"
   $Form.Controls.Add($groupBox2013)

############################################## end group boxes

############################################## Start Arch checkboxes

   $arch64 = New-Object System.Windows.Forms.RadioButton
   $arch64.Location = New-Object System.Drawing.Size(10,20)
   $arch64.Size = New-Object System.Drawing.Size(100,20)
   $arch64.Checked = $true
   $arch64.Text = "64 bit"
   $arch.Controls.Add($arch64)

   $arch32 = New-Object System.Windows.Forms.RadioButton
   $arch32.Location = New-Object System.Drawing.Size(10,40)
   $arch32.Size = New-Object System.Drawing.Size(100,20)
   $arch32.Checked = $false
   $arch32.Text = "32 bit"
   $arch.Controls.Add($arch32)

############################################## End Arch checkboxes 

############################################## Start Arch checkboxes

   $English = New-Object System.Windows.Forms.RadioButton
   $English.Location = New-Object System.Drawing.Size(10,20)
   $English.Size = New-Object System.Drawing.Size(100,20)
   $English.Checked = $true
   $English.Text = "English"
   $language.Controls.Add($English)

   $Japanese = New-Object System.Windows.Forms.RadioButton
   $Japanese.Location = New-Object System.Drawing.Size(10,40)
   $Japanese.Size = New-Object System.Drawing.Size(100,20)
   $Japanese.Text = "Japanese"
   $language.Controls.Add($Japanese)

   $Korean = New-Object System.Windows.Forms.RadioButton
   $Korean.Location = New-Object System.Drawing.Size(10,60)
   $Korean.Size = New-Object System.Drawing.Size(100,20)
   $Korean.Text = "Korean"
   $language.Controls.Add($Korean)

   $Chinese = New-Object System.Windows.Forms.RadioButton
   $Chinese.Location = New-Object System.Drawing.Size(10,80)
   $Chinese.Size = New-Object System.Drawing.Size(100,20)
   $Chinese.Text = "Chinese"
   $language.Controls.Add($Chinese)

   $French = New-Object System.Windows.Forms.RadioButton
   $French.Location = New-Object System.Drawing.Size(10,100)
   $French.Size = New-Object System.Drawing.Size(100,20)
   $French.Text = "French"
   $language.Controls.Add($French)

   $Spanish = New-Object System.Windows.Forms.RadioButton
   $Spanish.Location = New-Object System.Drawing.Size(10,120)
   $Spanish.Size = New-Object System.Drawing.Size(100,20)
   $Spanish.Text = "Spanish"
   $language.Controls.Add($Spanish)

   $Vietnamese = New-Object System.Windows.Forms.RadioButton
   $Vietnamese.Location = New-Object System.Drawing.Size(10,140)
   $Vietnamese.Size = New-Object System.Drawing.Size(100,20)
   $Vietnamese.Text = "Vietnamese"
   $language.Controls.Add($Vietnamese)

############################################## End Arch checkboxes 

############################################## Start Microsoft 365 checkboxes
   $m365Home = New-Object System.Windows.Forms.RadioButton
   $m365Home.Location = New-Object System.Drawing.Size(10,20)
   $m365Home.Size = New-Object System.Drawing.Size(100,20)
   $m365Home.Checked = $false
   $m365Home.Text = "Home"
   $groupBox365.Controls.Add($m365Home)

   $m365Business = New-Object System.Windows.Forms.RadioButton
   $m365Business.Location = New-Object System.Drawing.Size(10,40)
   $m365Business.Size = New-Object System.Drawing.Size(100,20)
   $m365Business.Text = "Business"
   $groupBox365.Controls.Add($m365Business)

   $m365Enterprise = New-Object System.Windows.Forms.RadioButton
   $m365Enterprise.Location = New-Object System.Drawing.Size(10,60)
   $m365Enterprise.Size = New-Object System.Drawing.Size(100,20)
   $m365Enterprise.Text = "Enterprise"
   $groupBox365.Controls.Add($m365Enterprise)
############################################## End Microsoft 365 checkboxes

############################################## Start Office 2016 checkboxes
   $2016Pro = New-Object System.Windows.Forms.RadioButton
   $2016Pro.Location = New-Object System.Drawing.Size(10,20)
   $2016Pro.Size = New-Object System.Drawing.Size(100,20)
   $2016Pro.Checked = $false
   $2016Pro.Text = "Professional"
   $groupBox2016.Controls.Add($2016Pro)

   $2016Std = New-Object System.Windows.Forms.RadioButton
   $2016Std.Location = New-Object System.Drawing.Size(10,40)
   $2016Std.Size = New-Object System.Drawing.Size(100,20)
   $2016Std.Text = "Standard"
   $groupBox2016.Controls.Add($2016Std)

   $2016ProjectPro = New-Object System.Windows.Forms.RadioButton
   $2016ProjectPro.Location = New-Object System.Drawing.Size(10,60)
   $2016ProjectPro.Size = New-Object System.Drawing.Size(100,20)
   $2016ProjectPro.Text = "Project Pro"
   $groupBox2016.Controls.Add($2016ProjectPro)

   $2016VisioPro = New-Object System.Windows.Forms.RadioButton
   $2016VisioPro.Location = New-Object System.Drawing.Size(10,80)
   $2016VisioPro.Size = New-Object System.Drawing.Size(100,20)
   $2016VisioPro.Text = "Visio Pro"
   $groupBox2016.Controls.Add($2016VisioPro)

############################################## End Office 2016 checkboxes

############################################## Start Office 2013 checkboxes
   $2013Pro = New-Object System.Windows.Forms.RadioButton
   $2013Pro.Location = New-Object System.Drawing.Size(10,20)
   $2013Pro.Size = New-Object System.Drawing.Size(100,20)
   $2013Pro.Checked = $false
   $2013Pro.Text = "Professional"
   $groupBox2013.Controls.Add($2013Pro)

   $2013Std = New-Object System.Windows.Forms.RadioButton
   $2013Std.Location = New-Object System.Drawing.Size(10,40)
   $2013Std.Size = New-Object System.Drawing.Size(100,20)
   $2013Std.Text = "Standard"
   $groupBox2013.Controls.Add($2013Std)

   $2013ProjectPro = New-Object System.Windows.Forms.RadioButton
   $2013ProjectPro.Location = New-Object System.Drawing.Size(10,60)
   $2013ProjectPro.Size = New-Object System.Drawing.Size(100,20)
   $2013ProjectPro.Text = "Project Pro"
   $groupBox2013.Controls.Add($2013ProjectPro)

   $2013VisioPro = New-Object System.Windows.Forms.RadioButton
   $2013VisioPro.Location = New-Object System.Drawing.Size(10,80)
   $2013VisioPro.Size = New-Object System.Drawing.Size(100,20)
   $2013VisioPro.Text = "Visio Pro"
   $groupBox2013.Controls.Add($2013VisioPro)

############################################## End Office 2013 checkboxes


############################################## Start buttons

   $submitButton = New-Object System.Windows.Forms.Button 
   $submitButton.Cursor = [System.Windows.Forms.Cursors]::Hand
   $submitButton.BackColor = [System.Drawing.Color]::LightGreen
   $submitButton.Location = New-Object System.Drawing.Size(290,130) 
   $submitButton.Size = New-Object System.Drawing.Size(130,40) 
   $submitButton.Text = "Submit" 
   $submitButton.Add_Click({microsoftInstaller}) 
   $Form.Controls.Add($submitButton) 

############################################## end buttons

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
