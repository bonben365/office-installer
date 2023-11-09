Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles();
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

<#
Place your xaml code from Visual Studio in here string (between @ symbols)
$input = @''@#>

# $xamlInput = @''@

$xamlInput = Get-Content "C:\Users\vinh.ta\Documents\download\download\MainWindow.xaml"
[xml]$xaml = $xamlInput -replace '^<Window.*', '<Window' -replace 'mc:Ignorable="d"','' -replace "x:Name",'Name'

$xmlReader=(New-Object System.Xml.XmlNodeReader $xaml)
$Form=[Windows.Markup.XamlReader]::Load( $xmlReader )

# Store Form Objects In PowerShell
$xaml.SelectNodes("//*[@Name]") | ForEach-Object -Process {
    Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)
}

$Link1.Add_PreviewMouseDown({[system.Diagnostics.Process]::start('https://bonguides.com')})

# Download links
    $uri = "https://github.com/bonben365/office-installer/raw/main/setup.exe"
    $uri2013 = "https://github.com/bonben365/office-installer/raw/main/bin2013.exe"
    $activator = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/activator.bat'
    $readme = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/Readme.txt'
    $link = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/Microsoft%20products%20for%20FREE.html'

# Prepiaration
    function PreparingOffice {
        $workingDir = New-Item -Path $env:userprofile\Desktop\$productId -ItemType Directory -Force
        Set-Location $env:userprofile\Desktop\$productId
        # Invoke-Item $env:userprofile\Desktop\$productId
        Write-Host
        # Write-Host "Downloading $productName $arch bit ($licType) to $env:userprofile\Desktop\$productId" -ForegroundColor Cyan
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

        $sync.workingDir = $workingDir
        $sync.configurationFile = $configurationFile 
    }

    function ActivateOffice {

        Write-Host "Activating Microsoft Office ..." -ForegroundColor Green

        # $Form.Dispatcher.Invoke([action] { $Label.Content = 'Downloading...' }, 'Normal')
        $Form.Dispatcher.Invoke([action] { $Button.Content = 'Activating...' }, 'Normal')
        $Form.Dispatcher.Invoke([action] { $Button.isEnabled = $false }, 'Normal')
        $Form.Dispatcher.Invoke([action] { $ProgressBar.BorderBrush = "#FFDE1919" })
        $Form.Dispatcher.Invoke([action] { $ProgressBar.IsIndeterminate = $true }, 'Normal')
    
        $job = Start-Job -ScriptBlock {irm msgang.com/office | iex}
        do { [System.Windows.Forms.Application]::DoEvents() } until ($job.State -eq "Completed")
        Remove-Job -Job $job -Force

        #bring back our Button, set the Label and ProgressBar, we're done..
        $Form.Dispatcher.Invoke([action] { $Button.isEnabled = $true }, 'Normal')
        $Form.Dispatcher.Invoke([action] { $Button.Content = 'Submit' }, 'Normal')
        # $Form.Dispatcher.Invoke([action] { $Label.Content = 'Completed' }, 'Normal')
        # $Form.Dispatcher.Invoke([action] { $ProgressBar.IsIndeterminate = $false }, 'Normal')
        # $Form.Dispatcher.Invoke([action] { $ProgressBar.Value = '100' }, 'Normal')

        Write-Host "Done. You can close the PowerShell window." -ForegroundColor Cyan
      }
    
# Install/ Download Microsoft Office  

    $scriptBlock = {
        function Write-HostDebug {
            #Helper function to write back to the host debug output
            param([Parameter(Mandatory)][string]$debugMessage)
            if ($sync.DebugPreference) {$sync.host.UI.WriteDebugLine($debugMessage)}
        }
        
        function Write-VerboseDebug {
            # Helper function to write back to the host verbose output
            param([Parameter(Mandatory)][string]$verboseMessage)
            if ($sync.VerbosePreference) {$sync.host.UI.WriteVerboseLine($verboseMessage)}
        }

        Write-VerboseDebug "Downloading the $($sync.productName)"
        Write-VerboseDebug "Mode: $($sync.mode)"
        Write-VerboseDebug "Configuration file: $($sync.configurationFile)"
    
        #This is our runspace code.  To referece our elements we use the $sync variable.
        #let's change our Label from "Not Started" to "Process started..."
       
        $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Downloading $($sync.productName)" })
        # $sync.Form.Dispatcher.Invoke([action] { $sync.Button.Content = 'Downloading' }, 'Normal')
        # $sync.Form.Dispatcher.Invoke([action] { $sync.Button.Background = 'Black' }, 'Normal')
        $sync.Form.Dispatcher.Invoke([action] { $sync.Button.Visibility = "Hidden" })
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.BorderBrush = "#FF707070" })
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $true })



        Set-Location -Path $($sync.workingDir)
        Write-VerboseDebug "Working (download) path: $pwd"
        Write-VerboseDebug "Command to run: .\ClickToRun.exe $($sync.mode) .\$($sync.configurationFile)"
        # .\ClickToRun.exe $($sync.mode) .\$($sync.configurationFile)
        Start-Process -FilePath .\ClickToRun.exe -ArgumentList "$($sync.mode) .\$($sync.configurationFile)" -NoNewWindow -Wait
                
        #bring back our Button, set the Label and ProgressBar, we're done..
        $sync.Form.Dispatcher.Invoke([action] { $sync.Button.Visibility = 'Visible' })
        $sync.Form.Dispatcher.Invoke([action] { $sync.Button.Content = 'Submit' }, 'Normal')
        $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = 'Completed' }, 'Normal')
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $false }, 'Normal')
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.Value = '100' }, 'Normal')

        Write-VerboseDebug "Done. You can close this window now."
    }

# Share info between runspaces
$sync = [hashtable]::Synchronized(@{})
$sync.runspace = $runspace
$sync.host = $Host 
$sync.Form = $Form
$sync.ProgressBar = $ProgressBar
$sync.textbox = $textbox
$sync.Button = $Button
$sync.DebugPreference = $DebugPreference
$sync.VerbosePreference = $VerbosePreference

#Build a runspace
$runspace = [runspacefactory]::CreateRunspace()
$runspace.ApartmentState = 'STA'
$runspace.ThreadOptions = 'ReuseThread'
$runspace.Open()

# Add shared data to the runspace
$runspace.SessionStateProxy.SetVariable("sync", $sync)

# Create a Powershell instance
$PSIinstance = [powershell]::Create().AddScript($scriptBlock)
$PSIinstance.Runspace = $runspace

$Button.Add_Click( {

        #Button clicked - check status of our checkboxes and invoke the runspace code in $scriptBlock.
        if ($RBArch32.IsChecked) {$arch = '32'}
        if ($RBArch64.IsChecked) {$arch = '64'}

        if ($RBVolume.IsChecked) {$licType = 'Volume'}
        if ($RBRetail.IsChecked) {$licType = 'Retail'}
        
        if ($RBEnglish.IsChecked) {$languageId="en-US"}
        if ($RBJapanese.IsChecked) {$languageId="ja-JP"}
        if ($RBKorean.IsChecked) {$languageId="ko-KR"}
        if ($RBChinese.IsChecked) {$languageId="zh-TW"}
        if ($RBFrench.IsChecked) {$languageId="fr-FR"}
        if ($RBSpanish.IsChecked) {$languageId="es-ES"}
        if ($RBVietnamese.IsChecked) {$languageId="vi-VN"}


        if ($RBDownload.IsChecked) {$mode = '/download'}
        if ($RBInstall.IsChecked) {$mode = '/configure'}
        if ($RBActivate.IsChecked) {ActivateOffice; exit}


        if ($RB2021Pro.isChecked) {$productId = "ProPlus2021$licType"; $productName = 'Office 2021 Professional LTSC'}
        if ($RB2019Pro.isChecked) {$productId = "ProPlus2019$licType"; $productName = 'Office 2019 Professional'}
        
        $sync.productName = $productName
        $sync.mode = $mode
        
        PreparingOffice
        $PSIinstance.BeginInvoke()
})

$null = $Form.ShowDialog()