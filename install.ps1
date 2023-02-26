$MainMenu = {
   Write-Host " ******************************************************"
   Write-Host " * Microsoft Office Installation Script               *" 
   Write-Host " * Date:    26/02/2023                                *" 
   Write-Host " * Author:  https://github.com/bonben365              *" 
   Write-Host " * Website: https://bonben365.com/                    *" 
   Write-Host " ******************************************************" 
   Write-Host 
   Write-Host " 1. Office 365 / Microsoft 365" 
   Write-Host " 2. Uninstall All Previous Versions of Microsoft Office"
   Write-Host " 3. Quit or Press Ctrl + C"
   Write-Host 
   Write-Host " Select an option and press Enter: "  -nonewline
   }
   cls
   
   Do { 
   cls
   Invoke-Command $MainMenu
   $Select = Read-Host
   Switch ($Select)
      {
         #1. Office 365 / Microsoft 365
          1 {
               $365SubMenu = {
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
            Invoke-Command $365SubMenu
            $365SubSelect = Read-Host
   
            if ($365SubSelect -eq 1) {$productId = 'O365HomePremRetail'}
            if ($365SubSelect -eq 2) {$productId = 'O365HomePremRetail'}
            if ($365SubSelect -eq 3) {$productId = 'O365BusinessRetail'}
            if ($365SubSelect -eq 4) {$productId = 'O365ProPlusRetail'}
   
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
       
           Switch ($365SubSelect)
           {
                1 { Invoke-Command $o365 }
                2 { Invoke-Command $o365 }
                3 { Invoke-Command $o365 }
                4 { Invoke-Command $o365 }
        
            }
               }
            While ($365SubSelect -ne 5)
            cls
            }  
      }
   }
   While ($Select -ne 3)
