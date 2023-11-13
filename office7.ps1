if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
   Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
   Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "irm msgang.com/install | iex"
   break
}

# Load ddls to the current session.
Add-Type -AssemblyName PresentationFramework, System.Drawing, PresentationFramework, System.Windows.Forms, WindowsFormsIntegration, PresentationCore
[System.Windows.Forms.Application]::EnableVisualStyles()

# Place your xaml code from Visual Studio in here string (between @ symbols)
# $xamlinput = @'<xaml code here'@

$xamlInput = @'
<Window x:Class="office7.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:office7"
        mc:Ignorable="d"
        Title="Microsoft Installation Tool for Windows 7 - www.msgang.com" Height="515" Width="649" WindowStartupLocation="CenterScreen" Icon="https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/PowerShell/images/msgang_microsoft_icon.png">
    <Grid HorizontalAlignment="Left" VerticalAlignment="Top">
        <GroupBox x:Name="groupBoxMicrosoftOffice" Header="Select version to install:" BorderBrush="#FF164A69" Margin="125,10,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Height="458" Width="498" FontFamily="Consolas" FontSize="11">
            <Canvas HorizontalAlignment="Left" VerticalAlignment="Top">
                <Rectangle Height="81" Stroke="#FF164A69" Width="135" UseLayoutRounding="True" RadiusX="5" RadiusY="5" Canvas.Left="11" Canvas.Top="20" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton365Home" Content="Home" Canvas.Left="19" Canvas.Top="35" HorizontalAlignment="Left" VerticalAlignment="Top" VerticalContentAlignment="Center" Margin="0,5,0,0"/>
                <RadioButton x:Name="radioButton365Business" Content="Business" Canvas.Left="19" Canvas.Top="54" HorizontalAlignment="Left" VerticalAlignment="Center" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" Margin="0,5,0,0"/>
                <RadioButton x:Name="radioButton365Enterprise" Content="Enterprise" Canvas.Left="19" Canvas.Top="73" HorizontalAlignment="Left" VerticalAlignment="Top" VerticalContentAlignment="Center" Margin="0,5,0,0"/>
                <Label x:Name="Label365" Content="Microsoft 365" FontWeight="Bold" Canvas.Left="19" Background="#FFDA2323" HorizontalAlignment="Center" VerticalAlignment="Top" Canvas.Top="8" Foreground="White" Padding="8,4,8,4"/>
                <Rectangle Height="306" Stroke="#FF164A69" Width="150" UseLayoutRounding="True" RadiusX="5" RadiusY="5" Canvas.Left="159" Canvas.Top="20" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016Pro" Content="Professional" IsChecked="False" Padding="5,5,5,5" VerticalContentAlignment="Center" Canvas.Left="172" Canvas.Top="35" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016Std" Content="Standard" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="55" Canvas.Left="172" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016ProjectPro" Content="Project Pro" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="75" Canvas.Left="172" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016ProjectStd" Content="Project Standard" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="95" Canvas.Left="172" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016VisioPro" Content="Visio Pro" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="115" Canvas.Left="172" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016VisioStd" Content="Visio Standard" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="135" Canvas.Left="172" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016Word" Content="Word" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="155" Canvas.Left="172" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016Excel" Content="Excel" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="175" Canvas.Left="172" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016PowerPoint" Content="PowerPoint" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="195" Canvas.Left="172" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016Outlook" Content="Outlook" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="212" Canvas.Left="172" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016Access" Content="Access" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="235" Canvas.Left="172" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016Publisher" Content="Publisher" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="255" Canvas.Left="172" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016OneNote" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Content="OneNote" Canvas.Top="275" Canvas.Left="172" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <Label x:Name="Label2016" Content="Office 2016" FontWeight="Bold" Canvas.Left="169" Background="#FFA28210" Canvas.Top="8" HorizontalAlignment="Left" VerticalAlignment="Center" Padding="8,4,8,4" Foreground="White"/>
                <Rectangle Height="306" Stroke="#FF164A69" Width="150" UseLayoutRounding="True" RadiusX="5" RadiusY="5" Canvas.Left="324" Canvas.Top="19" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <Label x:Name="Label2013" Content="Office 2013" FontWeight="Bold" Canvas.Left="335" Background="#FF1B0F0F" Canvas.Top="7" HorizontalAlignment="Center" VerticalAlignment="Top" Foreground="White" Padding="8,4,8,4"/>
                <RadioButton x:Name="radioButton2013Pro" Content="Professional" IsChecked="False" Padding="5,5,5,5" VerticalContentAlignment="Center" Canvas.Left="335" Canvas.Top="34" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton2013Std" Content="Standard" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="54" Canvas.Left="335" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton2013ProjectPro" Content="Project Pro" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="74" Canvas.Left="335" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton2013ProjectStd" Content="Project Standard" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="94" Canvas.Left="335" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton2013VisioPro" Content="Visio Pro" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="114" Canvas.Left="335" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton2013VisioStd" Content="Visio Standard" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="134" Canvas.Left="335" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton2013Word" Content="Word" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="154" Canvas.Left="335" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton2013Excel" Content="Excel" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="174" Canvas.Left="335" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton2013PowerPoint" Content="PowerPoint" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="194" Canvas.Left="335" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton2013Outlook" Content="Outlook" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="214" Canvas.Left="335" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton2013Access3" Content="Access" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="234" Canvas.Left="335" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton2013Publisher" Content="Publisher" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="254" Canvas.Left="335" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <TextBox x:Name="textBox1" TextWrapping="Wrap" Text="(*) By default, this script installs the 32-bit version in English." Canvas.Top="338" FontSize="10.5" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalAlignment="Center" VerticalAlignment="Top" Canvas.Left="-3" Margin="0,2,0,2" Padding="0,0,0,2"/>
                <TextBox x:Name="textBox2" TextWrapping="Wrap" Text="(*) Default mode is Install. If you want to download only, select Download mode." Canvas.Top="360" FontSize="10.5" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalAlignment="Center" VerticalAlignment="Top" Canvas.Left="-3" Padding="0,0,0,2"/>
                <TextBox x:Name="textBox3" TextWrapping="Wrap" Text="(*) The downloaded files would be saved on the current user's desktop." Canvas.Top="377" FontSize="10.5" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalAlignment="Center" VerticalAlignment="Top" Canvas.Left="-3" Padding="0,0,0,2" Margin="0,2,0,2"/>
                <TextBox x:Name="textBox4" TextWrapping="Wrap" Text="(*) To activate the Office apps, select Activate mode then click on Activate button." Canvas.Top="395" FontSize="10.5" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalAlignment="Center" VerticalAlignment="Top" Canvas.Left="-3" Margin="0,2,0,2" Padding="0,0,0,2"/>
                <TextBox x:Name="textBox5" TextWrapping="Wrap" Text="(*) Created by: Leo Nguyen | Website: " Canvas.Top="414" FontSize="10.5" BorderBrush="{x:Null}" Background="{x:Null}" Canvas.Left="-3" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0,2,0,2" Padding="0,0,0,2"/>
                <Image x:Name="image" Height="81" Width="78" Canvas.Left="40" Canvas.Top="112" Source="https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/Temp/download.png" HorizontalAlignment="Center" VerticalAlignment="Top" Visibility="Hidden"/>
            </Canvas>
        </GroupBox>
        <GroupBox x:Name="groupBoxArch" Header="Arch:" Margin="10,10,0,0" BorderBrush="#FF0D4261" HorizontalAlignment="Left" VerticalAlignment="Top" FontFamily="Consolas" FontSize="11" Width="104">
            <StackPanel HorizontalAlignment="Left" VerticalAlignment="Top">
                <RadioButton x:Name="radioButtonArch32" Content="x32" Width="37" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="5,6,0,0" VerticalContentAlignment="Center" IsChecked="True"/>
                <RadioButton x:Name="radioButtonArch64" Content="x64" Width="37" VerticalContentAlignment="Center" IsChecked="False" Margin="5,6,0,5"/>
            </StackPanel>
        </GroupBox>
        <GroupBox x:Name="groupBoxMode" Header="Mode:" Margin="10,80,0,0" BorderBrush="#FF0D4261" HorizontalAlignment="Left" VerticalAlignment="Top" FontFamily="Consolas" FontSize="11" Width="104" Height="84" ToolTip="When selecting the Activate mode...">
            <StackPanel HorizontalAlignment="Left" VerticalAlignment="Top">
                <RadioButton x:Name="radioButtonInstall" Content="Install" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="5,6,0,0" VerticalContentAlignment="Center" IsChecked="True"/>
                <RadioButton x:Name="radioButtonDownload" Content="Download" VerticalContentAlignment="Center" Margin="5,6,0,0"/>
                <RadioButton x:Name="radioButtonActivate" Content="Activate" VerticalContentAlignment="Center" Margin="5,6,0,0"/>
            </StackPanel>
        </GroupBox>
        <GroupBox x:Name="groupBoxLanguage" Header="Language:" Margin="10,169,0,0" BorderBrush="#FF0D4261" HorizontalAlignment="Left" FontFamily="Consolas" FontSize="11" Width="104">
            <StackPanel HorizontalAlignment="Left" VerticalAlignment="Top">
                <RadioButton x:Name="radioButtonEnglish" Content="English" VerticalContentAlignment="Center" IsChecked="True" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="5,8,0,0"/>
                <RadioButton x:Name="radioButtonJapanese" Content="Japanese" VerticalContentAlignment="Center" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="5,6,0,0"/>
                <RadioButton x:Name="radioButtonKorean" Content="Korean" VerticalContentAlignment="Center" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="5,6,0,0"/>
                <RadioButton x:Name="radioButtonChinese" Content="Chinese" VerticalContentAlignment="Center" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="5,6,0,0"/>
                <RadioButton x:Name="radioButtonFrench" Content="French" VerticalContentAlignment="Center" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="5,6,0,0"/>
                <RadioButton x:Name="radioButtonSpanish" Content="Spanish" VerticalContentAlignment="Center" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="5,6,0,0"/>
                <RadioButton x:Name="radioButtonHindi" Content="Hindi" VerticalContentAlignment="Center" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="5,6,0,0"/>
                <RadioButton x:Name="radioButtonGerman" Content="German" VerticalContentAlignment="Center" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="5,6,0,0"/>
                <RadioButton x:Name="radioButtonGreek" Content="Greek" VerticalContentAlignment="Center" Margin="5,6,0,0"/>
                <RadioButton x:Name="radioButtonHungarian" Content="Hungarian" VerticalContentAlignment="Center" Margin="5,6,0,0"/>
                <RadioButton x:Name="radioButtonItalian" Content="Italian" VerticalContentAlignment="Center" Margin="5,6,0,0"/>
                <RadioButton x:Name="radioButtonPortuguese" Content="Portuguese" VerticalContentAlignment="Center" Margin="5,6,0,0"/>
                <RadioButton x:Name="radioButtonVietnamese" Content="Vietnamese" VerticalContentAlignment="Center" Margin="5,6,0,0"/>
            </StackPanel>
        </GroupBox>
        <Button x:Name="buttonSubmit" Content="Submit" HorizontalAlignment="Left" Margin="147,187,0,0" VerticalAlignment="Top" Width="118" Height="28" Background="#FF168E12" Foreground="White" FontFamily="Consolas" FontSize="13" FontWeight="Bold" UseLayoutRounding="True" BorderBrush="#FF168E12"/>
        <ProgressBar x:Name="progressbar" HorizontalAlignment="Left" Height="10" Margin="147,226,0,0" VerticalAlignment="Top" Width="118" IsEnabled="False" Background="{x:Null}" BorderBrush="{x:Null}"/>
        <TextBox x:Name="textbox" TextWrapping="Wrap" Width="120" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="147,248,0,0" FontFamily="Consolas" FontSize="11" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" Background="{x:Null}" BorderBrush="{x:Null}" AllowDrop="False" Focusable="False" IsHitTestVisible="False" IsTabStop="False" IsUndoEnabled="False"/>
        <Label x:Name="Link1" HorizontalAlignment="Left" Margin="342,436,0,0" VerticalAlignment="Top" Width="172" FontSize='10.5' FontFamily="Consolas" Padding="5,5,5,2">
            <Hyperlink NavigateUri="https://msgang.com">https://msgang.com</Hyperlink>
        </Label>

    </Grid>
</Window>
'@

[xml]$xaml = $xamlInput -replace '^<Window.*', '<Window' -replace 'mc:Ignorable="d"','' -replace "x:Name",'Name'
$xmlReader = (New-Object System.Xml.XmlNodeReader $xaml)
$Form = [Windows.Markup.XamlReader]::Load( $xmlReader )

# Store form objects (variables) in PowerShell
   $xaml.SelectNodes("//*[@Name]") | ForEach-Object -Process {
       Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)
   }

$Link1.Add_PreviewMouseDown({[system.Diagnostics.Process]::start('https://msgang.com')})

# Download links
   $uri2013 = "https://github.com/bonben365/office-installer/raw/main/bin2013.exe"
   $activator = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/activator.bat'
   $uninstall = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/uninstall.bat'
   $readme = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/Readme.txt'
   $link = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/Microsoft%20products%20for%20FREE.html'

# Prepiaration for download and install
   function PreparingOffice {
       if ($radioButtonDownload.IsChecked) {
           $workingDir = New-Item -Path $env:userprofile\Desktop\$productId -ItemType Directory -Force
           Set-Location $workingDir
           Invoke-Item $workingDir
       }

       if ($radioButtonInstall.IsChecked) {
           $workingDir = New-Item -Path $env:temp\ClickToRun\$productId -ItemType Directory -Force
           Set-Location $workingDir
       }

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

       (New-Object Net.WebClient).DownloadFile($uri, "$workingDir\ClickToRun.exe")
       (New-Object Net.WebClient).DownloadFile($activator, "$workingDir\03.Activator.bat")
       (New-Object Net.WebClient).DownloadFile($readme, "$workingDir\01.Readme.txt")
       (New-Object Net.WebClient).DownloadFile($link, "$workingDir\Microsoft products for FREE.html")

       $sync.configurationFile = $configurationFile
       $sync.workingDir = $workingDir
   }
   
# Creating script block for download and install
   $DownloadInstallOffice = {
       function Write-HostDebug {
           #Helper function to write back to the host debug output
           param([Parameter(Mandatory)]
           [string]
           $debugMessage)
           if ($sync.DebugPreference) {
               $sync.host.UI.WriteDebugLine($debugMessage)
           }
       }

       function Write-VerboseDebug {
           param([Parameter(Mandatory)]
           [string]
           $verboseMessage)
           if ($sync.VerbosePreference) {
               $sync.host.UI.WriteVerboseLine($verboseMessage)
           }
       }

       Write-VerboseDebug "Downloading the $($sync.productName)"
       Write-VerboseDebug "Mode: $($sync.mode)"
       Write-VerboseDebug "Configuration file: $($sync.configurationFile)"

       # To referece our elements we use the $sync variable from hashtable.
       $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = "Hidden" })
       $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "$($sync.UIstatus) $($sync.productName) $($sync.arch)-bit ($($sync.language))" })
       $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.BorderBrush = "#FF707070" })
       $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $true })
       $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Visible" })

       Set-Location -Path $($sync.workingDir)
       Write-VerboseDebug "Working (download) path: $pwd"
       Write-VerboseDebug "Command to run: .\ClickToRun.exe $($sync.mode) .\$($sync.configurationFile)"

       Start-Process -FilePath .\ClickToRun.exe -ArgumentList "$($sync.mode) .\$($sync.configurationFile)" -NoNewWindow -Wait
               
       # Bring back our Button, set the Label and ProgressBar, we're done..
       $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Hidden" })
       $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = 'Visible' })
       $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Content = 'Submit' })
       $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = 'Completed' })
       $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $false })
       $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.Value = '100' })

       Write-VerboseDebug "Done. You can close this window now."
   }

# Share info between runspaces
   $sync = [hashtable]::Synchronized(@{})
   $sync.runspace = $runspace
   $sync.host = $host
   $sync.Form = $Form
   $sync.ProgressBar = $ProgressBar
   $sync.textbox = $textbox
   $sync.image = $image
   $sync.buttonSubmit = $buttonSubmit
   $sync.DebugPreference = $DebugPreference
   $sync.VerbosePreference = $VerbosePreference

# Build a runspace
   $runspace = [runspacefactory]::CreateRunspace()
   $runspace.ApartmentState = 'STA'
   $runspace.ThreadOptions = 'ReuseThread'
   $runspace.Open()

# Add shared data to the runspace
   $runspace.SessionStateProxy.SetVariable("sync", $sync)

# Create a Powershell instance
   $PSIinstance = [powershell]::Create().AddScript($scriptBlock)
   $PSIinstance.Runspace = $runspace


# INSTALL/DOWNLOAD/ACTIVATE Microsoft Office with a runspace

   $ActivateOffice = {
       $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Activating Microsoft Office..." })
       $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = "Hidden" })
       $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.BorderBrush = "#FF707070" })
       $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $true })
       $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Visible" })

       Set-Location -Path $($sync.workingDir)
       (New-Object Net.WebClient).DownloadFile($($sync.activator), "$($sync.workingDir)\03.Activator.bat")
       Start-Process -FilePath .\03.Activator.bat -Wait

       $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Hidden" })
       $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = 'Visible' })
       $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Content = 'Submit' })
       $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = 'Completed' })
       $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $false })
       $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.Value = '100' })

       # Cleanup
       Set-Location ..
       Remove-Item ClickToRunA -Recurse -Force
   }

   $buttonSubmit.Add_Click( {

       if ($radioButtonActivate.IsChecked) {

           $workingDir = New-Item -Path $env:temp\ClickToRunA -ItemType Directory -Force
           Set-Location $workingDir
           $sync.workingDir = $workingDir
           $sync.activator = $activator

           $PSIinstance = [powershell]::Create().AddScript($ActivateOffice)
           $PSIinstance.Runspace = $runspace
           $PSIinstance.BeginInvoke()

       } else {
           $i = 0
           if ($radioButtonArch32.IsChecked) {$arch = '32'}
           if ($radioButtonArch64.IsChecked) {$arch = '64'}

           if ($radioButtonVolume.IsChecked) {$licType = 'Volume'}
           if ($radioButtonRetail.IsChecked) {$licType = 'Retail'}

           if ($radioButtonEnglish.IsChecked) {$languageId="en-US"; $language = 'English'}
           if ($radioButtonJapanese.IsChecked) {$languageId="ja-JP"; $language = 'Japanese'}
           if ($radioButtonKorean.IsChecked) {$languageId="ko-KR"; $language = 'Korean'}
           if ($radioButtonChinese.IsChecked) {$languageId="zh-TW"; $language = 'Chinese'}
           if ($radioButtonFrench.IsChecked) {$languageId="fr-FR"; $language = 'French'}
           if ($radioButtonSpanish.IsChecked) {$languageId="es-ES"; $language = 'Spanish'}
           if ($radioButtonHindi.IsChecked) {$languageId="hi-IN"; $language = 'Hindi'}
           if ($radioButtonGerman.IsChecked) {$languageId="de-DE"; $language = 'German'}
           if ($radioButtonGreek.IsChecked) {$languageId="el-GR"; $language = 'Greek'}
           if ($radioButtonHungarian.IsChecked) {$languageId="hu-HU"; $language = 'Hungarian'}
           if ($radioButtonItalian.IsChecked) {$languageId="it-IT"; $language = 'Italian'}
           if ($radioButtonPortuguese.IsChecked) {$languageId="pt-PT"; $language = 'Portuguese'}
           if ($radioButtonVietnamese.IsChecked) {$languageId="vi-VN"; $language = 'Vietnamese'}

           if ($radioButtonDownload.IsChecked) {$mode = '/download'; $UIstatus = 'Downlading'}
           if ($radioButtonInstall.IsChecked) {$mode = '/configure'; $UIstatus = 'Installing'}

           if ($radioButton365Home.IsChecked -eq $true) {$productId = "O365HomePremRetail"; $productName = 'Microsoft 365 Home'; $i++}
           if ($radioButton365Business.IsChecked -eq $true) {$productId = "O365BusinessRetail"; $productName = 'Microsoft 365 Apps for Business'; $i++}
           if ($radioButton365Enterprise.IsChecked -eq $true) {$productId = "O365ProPlusRetail"; $productName = 'Microsoft 365 Apps for Enterprise'; $i++}

       # For Office 2016
           if ($radioButton2016Pro.IsChecked -eq $true) {$productId = "ProfessionalRetail"; $productName = 'Office 2016 Professional Plus'; $i++}
           if ($radioButton2016Std.IsChecked -eq $true) {$productId = "StandardRetail"; $productName = 'Office 2016 Standard'; $i++}
           if ($radioButton2016ProjectPro.IsChecked -eq $true) {$productId = "ProjectProRetail"; $productName = 'Microsoft Project Pro 2016'; $i++}
           if ($radioButton2016ProjectStd.IsChecked -eq $true) {$productId = "ProjectStdRetail"; $productName = 'Microsoft Project Standard 2016'; $i++}
           if ($radioButton2016VisioPro.IsChecked -eq $true) {$productId = "VisioProRetail"; $productName = 'Microsoft Visio Pro 2016'; $i++}
           if ($radioButton2016VisioStd.IsChecked -eq $true) {$productId = "VisioStdRetail"; $productName = 'Microsoft Visio Standard 2016'; $i++}
           if ($radioButton2016Word.IsChecked -eq $true) {$productId = "WordRetail"; $productName = 'Microsoft Word 2016'; $i++}
           if ($radioButton2016Excel.IsChecked -eq $true) {$productId = "ExcelRetail"; $productName = 'Microsoft Excel 2016'; $i++}
           if ($radioButton2016PowerPoint.IsChecked -eq $true) {$productId = "PowerPointRetail"; $productName = 'Microsoft PowerPoint 2016'; $i++}
           if ($radioButton2016Outlook.IsChecked -eq $true) {$productId = "OutlookRetail"; $productName = 'Microsoft Outlook 2016'; $i++}
           if ($radioButton2016Publisher.IsChecked -eq $true) {$productId = "PublisherRetail"; $productName = 'Microsoft Publisher 2016'; $i++}
           if ($radioButton2016Access.IsChecked -eq $true) {$productId = "AccessRetail"; $productName = 'Microsoft Access 2016'; $i++}
           if ($radioButton2016OneNote.IsChecked -eq $true) {$productId = "OneNoteRetail"; $productName = 'Microsoft Onenote 2016'; $i++}

       # For Office 2013
           if ($radioButton2013Pro.IsChecked -eq $true) {$productId = "ProfessionalRetail"; $uri = $uri2013; $productName = 'Office 2013 Professional Plus'; $i++}
           if ($radioButton2013Std.IsChecked -eq $true) {$productId = "StandardRetail"; $uri = $uri2013; $productName = 'Office 2013 Standard'; $i++}
           if ($radioButton2013ProjectPro.IsChecked -eq $true) {$productId = "ProjectProRetail"; $uri = $uri2013; $productName = 'Microsoft Project Pro 2013'; $i++}
           if ($radioButton2013ProjectStd.IsChecked -eq $true) {$productId = "ProjectStdRetail"; $uri = $uri2013; $productName = 'Microsoft Project Standard 2013'; $i++}
           if ($radioButton2013VisioPro.IsChecked -eq $true) {$productId = "VisioProRetail"; $uri = $uri2013; $productName = 'Microsoft Visio Pro 2013'; $i++}
           if ($radioButton2013VisioStd.IsChecked -eq $true) {$productId = "VisioStdRetail"; $uri = $uri2013; $productName = 'Microsoft Visio Standard 2013'; $i++}
           if ($radioButton2013Word.IsChecked -eq $true) {$productId = "WordRetail"; $uri = $uri2013; $productName = 'Microsoft Word 2013'; $i++}
           if ($radioButton2013Excel.IsChecked -eq $true) {$productId = "ExcelRetail"; $uri = $uri2013; $productName = 'Microsoft Excel 2013'; $i++}
           if ($radioButton2013PowerPoint.IsChecked -eq $true) {$productId = "PowerPointRetail"; $uri = $uri2013; $productName = 'Microsoft PowerPoint 2013'; $i++}
           if ($radioButton2013Outlook.IsChecked -eq $true) {$productId = "OutlookRetail"; $uri = $uri2013; $productName = 'Microsoft Outlook 2013'; $i++}
           if ($radioButton2013Publisher.IsChecked -eq $true) {$productId = "PublisherRetail"; $uri = $uri2013; $productName = 'Microsoft Publisher 2013'; $i++}
           if ($radioButton2013Access.IsChecked -eq $true) {$productId = "AccessRetail"; $uri = $uri2013; $productName = 'Microsoft Access 2013'; $i++}
       # Update the shared hashtable
           $sync.arch = $arch
           $sync.mode = $mode
           $sync.language = $language
           $sync.UIstatus = $UIstatus
           $sync.productName = $productName

           if ($i -eq '1') {
               PreparingOffice
               $PSIinstance = [powershell]::Create().AddScript($DownloadInstallOffice)
               $PSIinstance.Runspace = $runspace
               $PSIinstance.BeginInvoke()
           } else {
               $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Foreground = "Red" })
               $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.FontWeight = "Bold" })
               $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Please select an Office version." })
           } 
       }
   })

$null = $Form.ShowDialog()