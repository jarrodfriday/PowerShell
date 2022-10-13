
$user =  $env:UserName

########################################################################################################
# Unzip Contents
########################################################################################################

function UnzipContents {
        
$UserDownloadFile = "$env:userprofile\Downloads"

Expand-Archive -Path "$env:userprofile\Downloads\testing.Zip" -DestinationPath $UserDownloadFile\testing

}

########################################################################################################
# Unzip Contents
########################################################################################################
    
function ImportBookmarks {

Start-Sleep -Seconds 5

# The relative path for the Downloads file of user

$Bookmarks = "$env:userprofile\Downloads\testing\$user\Bookmarks"

# The relative path for Local App Data ==> C:\Users\user\AppData\Local

$ChromeFiles = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default" 

# Copy Chrome Files (Bookmarks, Bookmarks.Bak, Favicons & Favicons-Journals) into $AppDataFolderBookmarks\Google\Chrome\User Data\Default" 

Copy-Item -Path "$Bookmarks\*" -Destination $ChromeFiles -Force

}

########################################################################################################
# Unzip Contents
########################################################################################################

function ImportQuickLinks {

$QuickLinks = "$env:userprofile\Downloads\$env:UserName\testing\QuickLinks"


Start-Sleep -Seconds 5
# Copy File Explorer quick links into $AppDataFolder\Microsoft\Windows\Recent\AutomaticDestination

Copy-Item -Path $QuickLinks\* -Destination "$env:appdata\Microsoft\Windows\Recent\AutomaticDestinations" -Force

}


UnzipContents
ImportBookmarks
ImportQuickLinks


