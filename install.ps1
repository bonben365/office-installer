$MainMenu = {
   Write-Host " ******************************************************"
   Write-Host " * Microsoft Office Installation Script               *" 
   Write-Host " * Date:    26/02/2023                                *" 
   Write-Host " * Author:  https://github.com/bonben365              *" 
   Write-Host " * Website: https://bonben365.com/                    *" 
   Write-Host " ******************************************************" 
   Write-Host 
   Write-Host " 1. Office 2019" 
   Write-Host " 2. Microsoft Office 365"
   Write-Host " 3. Quit or Press Ctrl + C"
   Write-Host 
   Write-Host " Select an option and press Enter: "  -nonewline
   }
   cls
   
   Do { 
   cls
   Invoke-Command $MainMenu
   $Select = Read-Host
   if ($Select -eq 1) {$version = '2019'}
   Switch ($Select)
      {
         #1. Office 2019
          1 {
               $SubMenu = {
               Write-Host " *************************************************"
               Write-Host " * Select a Microsoft Office Product             *" 
               Write-Host " *************************************************" 
               Write-Host 
               Write-Host " 1. ProPlus2019Volume" 
               Write-Host " 2. Office 365 Personal"
               Write-Host " 3. Microsoft 365 Apps for Business" 
               Write-Host " 4. Microsoft 365 Apps for Enterprise" 
               Write-Host " 5. Go Back"
               Write-Host 
               Write-Host " Select an option and press Enter: "  -nonewline
               } 
               cls
                        
            Do { 
            cls
            Invoke-Command $SubMenu
            $SubSelect = Read-Host
            
            if ($SubSelect -eq 1) {$productId = "ProPlus$($version)Volume"}
            if ($SubSelect -eq 2) {$productId = 'O365HomePremRetail'}
            if ($SubSelect -eq 3) {$productId = 'O365BusinessRetail'}
            if ($SubSelect -eq 4) {$productId = 'O365ProPlusRetail'}
   
            $ofin = {
               $null = New-Item -Path $env:temp\c2r -ItemType Directory -Force
               Set-Location $env:temp\c2r
               $fileName = 'configuration.xml'
               New-Item $fileName -ItemType File -Force | Out-Null
               Add-Content $fileName -Value '<Configuration>'
               Add-content $fileName -Value '<Add OfficeClientEdition="64"'
               Add-content $fileName -Value "<Product ID=`"$productId`">"
               Add-content $fileName -Value '<Language ID="en-us" />'
               Add-Content $fileName -Value '</Product>'
               Add-Content $fileName -Value '</Add>'
               Add-Content $fileName -Value '</Configuration>'
               Write-Host
               Write-Host ============================================================
               Write-Host "Installing $productId...."
               Write-Host ============================================================
               Write-Host

               $uri = 'https://github.com/bonben365/office365-installer/raw/main/setup.exe'
               Invoke-WebRequest -Uri $uri -OutFile 'setup.exe' -ErrorAction:SilentlyContinue | Out-Null
               .\setup.exe /configure .\configuration.xml
               
               # Cleanup
               Set-Location "$env:temp"
               Remove-Item $env:temp\c2r -Recurse -Force
           }
       
           Switch ($SubSelect)
           {
                1 { Invoke-Command $ofin }
                2 { Invoke-Command $ofin }
                3 { Invoke-Command $ofin }
                4 { Invoke-Command $ofin }
        
            }
               }
            While ($SubSelect -ne 5)
            cls
            }  
      
         #2. Office 365 / Microsoft 365
          2 {
               $SubMenu = {
               Write-Host " *************************************************"
               Write-Host " * Select a Microsoft Office 365 Product         *" 
               Write-Host " *************************************************" 
               Write-Host 
               Write-Host " 1. Office 365 Home" 
               Write-Host " 2. Office 365 Personal"
               Write-Host " 3. Microsoft 365 Apps for Business" 
               Write-Host " 4. Microsoft 365 Apps for Enterprise" 
               Write-Host " 5. Go Back"
               Write-Host 
               Write-Host " Select an option and press Enter: "  -nonewline
               } 
               cls
                        
            Do { 
            cls
            Invoke-Command $SubMenu
            $SubSelect = Read-Host
   
            if ($SubSelect -eq 1) {$productId = 'O365HomePremRetail'}
            if ($SubSelect -eq 2) {$productId = 'O365HomePremRetail'}
            if ($SubSelect -eq 3) {$productId = 'O365BusinessRetail'}
            if ($SubSelect -eq 4) {$productId = 'O365ProPlusRetail'}
   
            $o365 = {
               $null = New-Item -Path $env:temp\c2r -ItemType Directory -Force
               Set-Location $env:temp\c2r
               $fileName = 'configuration.xml'
               New-Item $fileName -ItemType File -Force | Out-Null
               Add-Content $fileName '<Configuration>'
               Add-content $fileName '<Add OfficeClientEdition="64" Channel="Current">'
               Add-content $fileName "<Product ID=`"$productId`">"
               Add-content $fileName '<Language ID="en-us" />'
               Add-Content $fileName -Value '</Product>'
               Add-Content $fileName -Value '</Add>'
               Add-Content $fileName -Value '</Configuration>'
               Write-Host
               Write-Host ============================================================
               Write-Host "Installing $productId...."
               Write-Host ============================================================
               Write-Host

               $uri = 'https://github.com/bonben365/office365-installer/raw/main/setup.exe'
               Invoke-WebRequest -Uri $uri -OutFile 'setup.exe' -ErrorAction:SilentlyContinue | Out-Null
               .\setup.exe /configure .\configuration.xml
               
               # Cleanup
               Set-Location "$env:temp"
               Remove-Item $env:temp\c2r -Recurse -Force
           }
       
           Switch ($SubSelect)
           {
                1 { Invoke-Command $o365 }
                2 { Invoke-Command $o365 }
                3 { Invoke-Command $o365 }
                4 { Invoke-Command $o365 }
        
            }
               }
            While ($SubSelect -ne 5)
            cls
            } 
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      }
   }
   While ($Select -ne 3)
