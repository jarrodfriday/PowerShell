 <#
        .SYNOPSIS
        Query all Graphics Drivers & Processors on an Asset and returns key details. 
        .DESCRIPTION
        
        This script queries a system for key details.
        
            - For Graphics Cards = The name of the video controller, the current driver version it is running and when it was installed.
            - For Processors = The name of the Processor, it's the manufacturer and a short description

        
        For example, the query is set to identify Intel GPU's + Drivers and Processors (targetted drivers & Processors) ($i.name -like "*Intel*")

        This query can be customised to target specific brands / types of Video Controller and Processors such as AMD, Intel or Nvidia 

        .PARAMETER Name
        The name of the file is GraphicsCardFinder.ps1. No direct output from the file.  

        .PARAMETER Extension
        The extension is .ps1. No direct output from the file.  

        .INPUTS
        $VideoController variable stores the Get-WmiObject win32_VideoController for easier access and cleaner code 

        Similarly the $Processor variable stores the win32_Processor for easier access and cleaner code 

        $AssetID variable identifies the asset the script is being run on and returns the -last 1 result
        As the system name will be the same on all returned drivers, returning the last result is fine 

        .OUTPUTS
        This script identifies many key attributes as detailed below:
        
            - The number of identified Graphics Cards on an asset
            - The system name the script is running on

        For all found Graphics Cards the following will be returned

            - The name of the video controller, the current driver version and when it was installed 

        For all found processors the following will be returned

        - The name of the Processor, the Manufacturer and a short description 

        .Notes
        Completed as part of my graduate journey in Security Operations 
#>


########################################################################################################
# Static Variables 
########################################################################################################

$VideoController = Get-WmiObject win32_VideoController
$AssetID = $VideoController | select -ExpandProperty systemname -last 1
$Processor = Get-WMIObject win32_Processor 

#Assign VideoController and Processor to array. This will return length no matter how many are on the system
$AssignGpuToArray = @($VideoController)
$AssignCPUToArray = @($Processor)

########################################################################################################
# Check count of Graphics Cards  
########################################################################################################

function GraphicsCardCount {

    Write-host `n"Querying system and identifying Graphics Cards(s)"
    Write-host `n"There is $($AssignGpuToArray.Length) video controller(s) on this system | AssetID: $AssetID"

}

########################################################################################################
# Display number of Graphics Cards found along with AssetID / Number 
########################################################################################################

function GraphicsDriverCheck {

# Display results

if ($VideoController) {

# Loop through all Graphics Cards from the WMI query and display results

    foreach ($i in $VideoController) {

        # Query is able to be tailored. For example -like "*Intel*"

          if ($i.name -like "*Intel*") {

            # Write Name of Graphics Card, Current Driver Version and Date of Installation for targetted Graphics Card (i.e Intel)
                 
            Write-host `n"Found Intel controller! >> $($i.name)", ">> Running Version: $($i.driverversion)", ">> Date Installed:" $i.ConvertToDateTime($i.driverdate)
                  

        }

        else {

            # Write Name of Graphics Card, Current Driver Version and Date of Installation for all non targetted Graphics Cards

            Write-host `n"Name: $($i.name)", ">> Running Version: $($i.driverversion)", ">> Date Installed:" $i.ConvertToDateTime($i.driverdate)

        }


    }

}

else {

    Write-host `n"Something went wrong - unable to query via WMI"

}

}

########################################################################################################
# Check count of Graphics Cards  
#######################################################################################################

function ProcessorCount {
    
    Write-host `n"Querying system and identifying processor(s)"
    Write-host `n"There is $($AssignCPUToArray.Length) processor(s) on this system | AssetID: $AssetID"
    
}

########################################################################################################
# CPU Finder 
########################################################################################################

function CpuFinder {

    if ($Processor) {

        # Loop through all processors from the WMI query and display results
        
            foreach ($i in $Processor) {
        
                # Query is able to be tailored. For example -like "*Intel*"
        
                  if ($i.name -like "*Intel*") {
        
                    # Write Name of Processor, Manufacturer and Description for targetted Graphics Card (i.e Intel)
                         
                    Write-host `n"Found Intel Processor! >> $($i.name)", ">> Manufacturer: $($i.manufacturer)", ">> Description:" $($i.caption)
                          
        
                }
        
                else {
        
                    # Write Name of Processor, Manufacturer and Description of all non targetted Graphics Cards

                    Write-host `n"Found Intel Processor! >> $($i.name)", ">> Manufacturer: $($i.manufacturer)", ">> Description:" $($i.caption)
                          
                }
        
        
            }
        
        }
        
        else {
        
            Write-host `n"Something went wrong - unable to query via WMI"
        
        }

}


########################################################################################################
# Called Functions are shown below 
########################################################################################################

GraphicsCardCount
GraphicsDriverCheck
ProcessorCount
CpuFinder




