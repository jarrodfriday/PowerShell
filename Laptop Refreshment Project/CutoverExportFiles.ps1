 <#
        .SYNOPSIS
        Copy to a directory in the C:\ drive called temp the files required to cutover Chrome Bookmarks and Filexplorer quick access paths
        
        .DESCRIPTION
        Copy to temp the Chrome Bookmark files (Bookmarks, Bookmarks.Bak, Favicons and Favicons-journal) + ChromeBookmarks.HTML file for import as a backup
        This is located in "$AppDataFolder\Microsoft\Windows\Recent\AutomaticDestinations\*"

        This will also Copy to a directory the users File Explorer Quick access located in "$AppDataFolder\Microsoft\Windows\Recent\AutomaticDestinations\*"

        .PARAMETER Name
        The name of the file is CutoverExportFiles

        .PARAMETER Extension
        The extension is .ps1. No direct output from the file.  

        .INPUTS
        Inputs Include: 

        - Input $ChromeFiles = "$LocalAppDataPath\Google\Chrome\User Data\Profile 2" will be on Surface Laptop 4's "$LocalAppDataPath\Google\Chrome\User Data\Default"
        - Input $ExplorerLinks =  "$AppDataPath\Microsoft\Windows\Recent\AutomaticDestinations\*"
        - Input $QuickPartsAutotext =  "$AppdataPath\Microsoft\Templates\*"

        .OUTPUTS
        Outputs Include:

        - Output Bookmarks, Bookmarks.Bak, Favicons and Favicons-journal to Bookmarks direcotry in C:/temp
        - Output  File Explorer quick links file to FileExplorerLinks directory in C:/temp
        - Output  File Explorer quick links file to QuickPartsAutotext directory in C:/temp

        .Notes
        Completed as part of my graduate journey in Infrastructure Operations to assist with the Surface Laptop 4 cutover project
#>

########################################################################################################
# Global Variable
########################################################################################################

$LocalAppDataPath = "$env:localappdata" # ---> \Users\JFR03\AppData\Local

$AppDataPath = "$env:appdata" # ---> C:\Users\JFR03\AppData\Roaming

    
########################################################################################################
# Find and Delete previous copies of this script 
########################################################################################################

function RemovovePreviousFileInstances {

    <# 1 #> remove-item -path "C:\Users\JFR03\Downloads\temp.zip"

    <# 2 #> remove-item -path "C:\Users\JFR03\Downloads\GroupPolicy.bat"

    <# 3 #> remove-item -path "C:\Users\JFR03\Downloads\ChromeBookmarks.html"

    <# 4 #> remove-item -path "C:\Users\JFR03\Downloads\CopyEmailSignatures" -Recurse

    <# 5 #> remove-item -path "C:\Users\JFR03\Downloads\FileExplorerLinks" -Recurse

    <# 6 #> remove-item -path "C:\Users\JFR03\Downloads\QuickPartsAutotext" -Recurse

    <# 7 #> remove-item -path "C:\Users\JFR03\Downloads\Bookmarks" -Recurse
}

########################################################################################################
# Copy Chrome Bookmarks - c:\Users\JFR03\AppData\Local\Chrome\User Data\Default\Bookmarks
########################################################################################################

function CopyChromeBookmarkFiles {

    # Create a temp directory in C: drive called Bookmarks
    New-Item -Path "$env:userprofile\Downloads" -Name "Bookmarks" -ItemType "directory"

    # Relative path for local app data + location of Bookmark files

    $ChromeFiles = "$LocalAppDataPath\Google\Chrome\User Data\Default" 

    # Copy Bookmarks, Bookmarks.Bak, Favicons & Favicons-journal (Note Favicons is for Favourite imaages in Browser)

    <# 1 #> Copy-Item -Path "$ChromeFiles\Bookmarks" -Destination "$env:userprofile\Downloads\Bookmarks" 
   
    <# 2 #> Copy-Item -Path "$ChromeFiles\Bookmarks.bak" -Destination "$env:userprofile\Downloads\Bookmarks" 
   
    <# 3 #> Copy-Item -Path "$ChromeFiles\Favicons" -Destination "$env:userprofile\Downloads\Bookmarks" 
   
    <# 4 #> Copy-Item -Path "$ChromeFiles\Favicons-journal" -Destination "$env:userprofile\Downloads\Bookmarks"
    

    }

########################################################################################################
# Copy File Explorer Quick Access - %AppData%\Microsoft\windows\recent\automaticdestinations
########################################################################################################

function CopyQuickAccess {

    # The relative path for App Data ==> C:\Users\user\AppData\Roaming ==> Changes per user 

    $CreateFileExplorerLinksFolder = New-Item -Path "$env:userprofile\Downloads" -Name "FileExplorerLinks" -ItemType "directory"

    # The location for the File Explorer quicklinks 

    $ExplorerLinks =  "$env:appdata\Microsoft\Windows\Recent\AutomaticDestinations\*"

    # Copy Quicklinks to created directory 

    Copy-Item -Path $ExplorerLinks -Destination $env:userprofile\Downloads\FileExplorerLinks
    
    # Call defined variables from above 

    $CreateFileExplorerLinksFolder 
    $ExplorerLinks


}

########################################################################################################
# Copy QuickPartsAutotext - %AppData%\Roaming\Microsoft\Templates
########################################################################################################

function CopyQuickPartsAutotext {

    $CreateQuickPartsAutotextFolder = New-Item -Path "$env:userprofile\Downloads" -Name "QuickPartsAutotext" -ItemType "directory"
    $QuickPartsAutotext =  "$env:appdata\Microsoft\Templates\*"
    Copy-Item -Path $QuickPartsAutotext -Destination $env:userprofile\Downloads\QuickPartsAutotext
    
    $CreateQuickPartsAutotextFolder
    $QuickPartsAutotext

}

########################################################################################################
# Copy Email Signatures - %appdata%\Roaming\Microsoft\Signatures
########################################################################################################

function CopyEmailSignatures {

    $CreateEmailSignaturesFolder = New-Item -Path "$env:userprofile\Downloads" -Name "CopyEmailSignatures" -ItemType "directory"
    $EmailSignatures =  "$env:appdata\Microsoft\Signatures\*"
    Copy-Item -Path $EmailSignatures -Destination $env:userprofile\Downloads\CopyEmailSignatures -Recurse
    
    $CreateEmailSignaturesFolder
    $EmailSignatures

}

########################################################################################################
# Copy Chrome Bookmarks - HTML
########################################################################################################

function CopyChromeBookmarksHTML {

    #credits: Mostly to tobibeer and Snak3d0c @ https://stackoverflow.com/questions/47345612/export-chrome-bookmarks-to-csv-file-using-powershell
    #Path to chrome bookmarks
  
$pathToJsonFile = "$LocalAppDataPath\Google\Chrome\User Data\Default\Bookmarks" 

$htmlOut = "$env:userprofile\Downloads\ChromeBookmarks.html"
$htmlHeader = @'
<!DOCTYPE NETSCAPE-Bookmark-file-1>
<!--This is an automatically generated file.
    It will be read and overwritten.
    Do Not Edit! -->
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<Title>Bookmarks</Title>
<H1>Bookmarks</H1>
<DL><p>
'@

    $htmlHeader | Out-File -FilePath $htmlOut -Force -Encoding utf8 #line59

    # A nested function to enumerate bookmark folders
    Function Get-BookmarkFolder {
    [cmdletbinding()]
    Param(
    [Parameter(Position=0,ValueFromPipeline=$True)]
    $Node
    )

    Process 
    {

    foreach ($child in $node.children) 
    {
    $da = [math]::Round([double]$child.date_added / 1000000) #date_added - from microseconds (Google Chrome {dates}) to seconds 'standard' epoch.
    $dm = [math]::Round([double]$child.date_modified / 1000000) #date_modified - from microseconds (Google Chrome {dates}) to seconds 'standard' epoch.
    if ($child.type -eq 'Folder') 
    {
        "    <DT><H3 FOLDED ADD_DATE=`"$($da)`">$($child.name)</H3>" | Out-File -FilePath $htmlOut -Append -Force -Encoding utf8
        "       <DL><p>" | Out-File -FilePath $htmlOut -Append -Force -Encoding utf8
        Get-BookmarkFolder $child
        "       </DL><p>" | Out-File -FilePath $htmlOut -Append -Force -Encoding utf8
    }
    else 
    {
            "       <DT><a href=`"$($child.url)`" ADD_DATE=`"$($da)`">$($child.name)</a>" | Out-File -FilePath $htmlOut -Append -Encoding utf8
    } #else url
    } #foreach
    } #process
    } #end function

    $data = Get-content $pathToJsonFile -Encoding UTF8 | out-string | ConvertFrom-Json
    $sections = $data.roots.PSObject.Properties | select -ExpandProperty name
    ForEach ($entry in $sections) {
        $data.roots.$entry | Get-BookmarkFolder
    }
    '</DL>' | Out-File -FilePath $htmlOut -Append -Force -Encoding utf8

}

########################################################################################################
# Create GPUpdate Bat File
########################################################################################################

function CreateBATFile {

Start-Sleep -Seconds 1.5

$CreateScript = echo "echo n | gpupdate /force" | out-file -encoding ascii -FilePath "$env:userprofile\Downloads\GroupPolicy.bat"
$CreateScript

}


########################################################################################################
# ZipContents
########################################################################################################

function ZipContents {


Start-Sleep -Seconds 1.6

$UserDownloadFile = "$env:userprofile\Downloads"

$compress = @{
  
  Path = "$UserDownloadFile\Bookmarks", "$UserDownloadFile\QuickPartsAutotext", "$UserDownloadFile\CopyEmailSignatures", "$UserDownloadFile\FileExplorerLinks","$UserDownloadFile\GroupPolicy.bat","$UserDownloadFile\ChromeBookmarks.html"

  CompressionLevel = "Fastest"
  
  DestinationPath = "$UserDownloadFile\temp.Zip"
}
Compress-Archive @compress

}

########################################################################################################
# Call functions 
########################################################################################################

# Call the functions defined above 

RemovovePreviousFileInstances
CopyChromeBookmarkFiles
CopyChromeBookmarksHTML
CopyEmailSignatures
CopyQuickAccess
CopyQuickPartsAutotext
CreateBATFile
ZipContents


<#
    
$UserDownloadFile = "$env:userprofile\Downloads"


$compress = @{
  Path = "$UserDownloadFile\Bookmarks", "$UserDownloadFile\FileExplorerLinks", "$UserDownloadFile\QuickPartsAutotext", "$UserDownloadFile\GroupPolicy.bat","$UserDownloadFile\ChromeBookmarks.html"

  CompressionLevel = "Fastest"
  DestinationPath = "$UserDownloadFile\temp.Zip"
}
Compress-Archive @compress

#>

<#

$UserDownloadFile = "$env:userprofile\Downloads"

Expand-Archive -Path "$UserDownloadFile\temp.Zip" -DestinationPath $UserDownloadFile\temp

#>
