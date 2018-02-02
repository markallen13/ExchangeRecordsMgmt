<#
.SYNOPSIS
Get-ERMFolderTag takes in a *live* mailbox folder and retrieves the current retention tag 
applied to it.  
.DESCRIPTION
Get-ERMFolderTag  takes an instance of Microsoft.Exchange.WebServices.Data.Folder and 
retrieves the current retention tag name as a string.  Errors are printed to standard 
output.   

This module has multiple requirements that are documented in the README.md file, including 
(1) Exchange Web Services us be installed and available, (2) the user must be able to open 
a remote session with the Exchange server, and (3) the user must be granted the 
ApplicationImpersonation role in Exchange.  Read the README.md for more information.  

Much of the module code is based on several scripts created by David Barrett, Microsoft.
.PARAMETER Folder
The instance of Microsoft.Exchange.WebServices.Data.Folder to use for retrival.  Folders
used for containing contacts, calendar items, and other non-email items are not accepted.   
.PARAMETER ArchiveTag
If set, script gets the folder's archive policy tag instead of the retention policy tag.  
.OUTPUTS
A string representing the name of the retention policy that was located, or 'None' if no
tag was found.  If an error occurs an exception is thrown.  
.EXAMPLE
Get-ERMFolderTag -Folder $MyFolder 
#>

Function Get-ERMFolderTag {
    [CmdletBinding()] 
    param (
        [Parameter(Mandatory=$true,Position=0)]
        [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
        [switch]$ArchiveTag
    )

    $ErrorActionPreference = "Stop"

    # Get the tag. 

    $CurrentTag = GetFolderRetentionTag $Folder $ArchiveTag

    # Close the Exchange PSSession if it exists.  

    Write-Verbose "Finished."
    return $CurrentTag
}

