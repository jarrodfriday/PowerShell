$UserDownloadFile = "$env:userprofile\Downloads"

$users = Get-ChildItem "C:\users"


function CopyAllUserBookmarks {

foreach ($user in $Users) {

New-item -Path $UserDownloadFile -Name $user -ItemType "directory"

New-item -Path $UserDownloadFile\$user -Name "Bookmarks" -ItemType "directory"

<# 1 #> Copy-Item -Path "C:\Users\$user\AppData\Local\Google\Chrome\User Data\Default\Bookmarks" -Destination "$UserDownloadFile\$user\Bookmarks"

<# 2 #> Copy-Item -Path "C:\Users\$user\AppData\Local\Google\Chrome\User Data\Default\Bookmarks.bak" -Destination "$UserDownloadFile\$user\Bookmarks"

<# 3 #> Copy-Item -Path "C:\Users\$user\AppData\Local\Google\Chrome\User Data\Default\Favicons" -Destination "$UserDownloadFile\$user\Bookmarks"

<# 4 #> Copy-Item -Path "C:\Users\$user\AppData\Local\Google\Chrome\User Data\Default\Favicons-journal" -Destination "$UserDownloadFile\$user\Bookmarks"

}

}

function CopyAllUserQuickAccess {

foreach ($user in $Users) {

New-item -Path $UserDownloadFile\$user -Name "QuickLinks" -ItemType "directory"

<# 1 #> Copy-Item -Path "C:\Users\$user\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations\*" -Destination "$UserDownloadFile\$user\QuickLinks"
    
}

}

function ZipAllContents { 

Start-Sleep -Seconds 1.6

foreach ($user in $Users) {

$compress = @{
  
  Path = "$UserDownloadFile\$user"

  CompressionLevel = "Fastest"
  
  DestinationPath = "$UserDownloadFile\testing.Zip"
}
Compress-Archive @compress -update
}

}

CopyAllUserBookmarks
CopyAllUserQuickAccess
ZipAllContents 
