<#
        .SYNOPSIS
        This script imports the Chrome Bookmarks files and File Explorer Quick Links files into the appropriate locations
        
        .DESCRIPTION
        This script will utilise the outputs from CutoverExportFiles as the inputs for this script (CutoverImportFiles)

        - Notably this will import the Bookmarks, Bookmarks.Bak, Favicons and Favicons-journal files from Downloads/temp/Boomarks into the "$LocalAppDataPath\Google\Chrome\User Data\Default file
        - This will also import the File Explorer Quick Access files from Downloads/temp/FileExplorerLinks into $AppDataPath\Microsoft\Windows\Recent\AutomaticDestinations

        .PARAMETER Name
        The name of the file is CutoverImportFiles

        .PARAMETER Extension
        The extension is .ps1. No direct output from the file.  

        .INPUTS
        Inputs from this file are the outputs from the CutoverExportFiles Script. 
        
        Inputs Include:
        - Input Bookmarks, Bookmarks.Bak, Favicons and Favicons-journal files.
        - Input ChromeBookmarks.html file will be created as well as a backup
        - Input Note also the File Explorer Quick Access files will also be inputs from FileExplorerLinks
       
        .OUTPUTS
        Outputs include:
        - Outputs include a restored Chrome Bookmarks Bar
        - Outputs include a restored File Explorer Quick Access tool Bar

        .Notes
        Completed as part of my graduate journey in Infrastructure Operations to assist with the Surface Laptop 4 cutover project
#>

########################################################################################################
# Unzip Contents
########################################################################################################

function UnzipContents {
        
$UserDownloadFile = "$env:userprofile\Downloads"

Expand-Archive -Path "$UserDownloadFile\temp.Zip" -DestinationPath $UserDownloadFile\temp

}

########################################################################################################
# Open Chrome - First Time
########################################################################################################

$ChromePath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"

start-process -FilePath "$ChromePath\Google Chrome.lnk"


########################################################################################################
# Close Chrome
########################################################################################################

function CloseChrome {

Start-Sleep -Seconds 4

Stop-Process -Name chrome

}

########################################################################################################
# Import Chrome Bookmarks from C:\Users\user\Downloads
########################################################################################################

function ImportBookmarks {

Start-Sleep -Seconds 5

# The relative path for the Downloads file of user

$ChromeSave = "$env:userprofile\Downloads\temp\Bookmarks"

# The relative path for Local App Data ==> C:\Users\user\AppData\Local

$AppDataFolderBookmarks = "$env:LOCALAPPDATA"
$ChromeFiles = "$AppDataFolderBookmarks\Google\Chrome\User Data\Default" 

# Copy Chrome Files (Bookmarks, Bookmarks.Bak, Favicons & Favicons-Journals) into $AppDataFolderBookmarks\Google\Chrome\User Data\Default" 

Copy-Item -Path "$ChromeSave\*" -Destination $ChromeFiles -Force

}

########################################################################################################
# Import Email Templates
########################################################################################################

function ImportEmailTemplates {

Start-Sleep -Seconds 5

# The relative path for the Downloads file of user

$EmailTemplateSave = "$env:userprofile\Downloads\temp\QuickPartsAutotext"

# The relative path for Local App Data ==> C:\Users\user\AppData\Local

$QuickPartsAutotext =  "$env:appdata\Microsoft\Templates"

Copy-Item -Path "$EmailTemplateSave\*" -Destination $QuickPartsAutotext -Force
    
}

########################################################################################################
# Import Email Signatures
########################################################################################################

function ImportEmailSignatures {

Start-Sleep -Seconds 8

# The relative path for the Downloads file of user

$EmailSignatureSave = "$env:userprofile\Downloads\temp\CopyEmailSignatures"

# The relative path for Local App Data ==> C:\Users\user\AppData\Local

$EmailSignatures =  "$env:appdata\Microsoft"

New-Item -Path "$EmailSignatures" -Name "Signatures" -ItemType "directory"


Copy-Item -Path "$EmailSignatureSave\*" -Destination "$EmailSignatures\Signatures" -Force -Recurse

}


########################################################################################################
# Open Chrome - Second Time
########################################################################################################

function OpenChromeSecond {

Start-Sleep -Seconds 8

$ChromePath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"

start-process -FilePath "$ChromePath\Google Chrome.lnk"


}

########################################################################################################
# Import File Explorer Quick Access from C:\Users\user\Downloads
########################################################################################################

# The relative path for the logged in user's Downloads file  ==> C:\Users\user\Downloads ==> Changes per user

$UserDownloadFile = "$env:userprofile\Downloads"

# The relative path for App Data ==> C:\Users\user\AppData\Roaming ==> Changes per user  

$AppDataFolder = "$env:appdata"

$QuickLinks = "$UserDownloadFile\temp\FileExplorerLinks\*"


function ImportQuickLinks {

Start-Sleep -Seconds 5

# Copy File Explorer quick links into $AppDataFolder\Microsoft\Windows\Recent\AutomaticDestination

Copy-Item -Path $QuickLinks -Destination "$AppDataFolder\Microsoft\Windows\Recent\AutomaticDestinations" -Force

}



########################################################################################################
# Run GPUpdate
########################################################################################################

#Create BAT file to execute GPUpdate

function GPUpdate {

Start-Sleep -Seconds 5

$GroupPolicy = "$env:userprofile\Downloads\temp\GroupPolicy.bat"
Start-Process -FilePath $GroupPolicy

}

########################################################################################################
# Call functions 
########################################################################################################

# Call the functions defined above 

UnzipContents
CloseChrome
ImportEmailTemplates
ImportEmailSignatures
ImportBookmarks
OpenChromeSecond
ImportQuickLinks
GPUpdate
