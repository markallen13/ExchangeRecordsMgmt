
<#
.SYNOPSIS
Get-ERMChildFolder retrieves the folders in the specified mailbox that are children of
the folder name given by the user. 
.DESCRIPTION
Get-ERMChildFolder returns an array of Exchange mailbox folder objects that represent the
child items in a given mailbox folder.  If no folder name is provided, the root of the 
mailbox (or the archive) is searched.

This function can return folders that are not used for holding e-mail, such as Contacts and 
Calendar.  

This module has multiple requirements that are documented in the README.md file, including 
(1) Exchange Web Services us be installed and available, (2) the user must be able to open 
a remote session with the Exchange server, and (3) the user must be granted the 
ApplicationImpersonation role in Exchange.  Read the README.md for more information.  

Much of the module code is based on several scripts created by David Barrett, Microsoft.
.PARAMETER EmailAddress
The email address of the mailbox for processing, also used to discover the Exchange Server.      
.PARAMETER Folder
The name of the folder to start the search.  Subfolders can be accessed using a backslash 
from the mailbox root, e.g. Inbox\Folder1\Folder2.   
.PARAMETER TraceFile
If this parameter is set, then EWS Tracing is enabled and is written to the provided file
name.  This shows the communication between the exchange server and this client.  
.PARAMETER Archive
If set, the script works with the Archive Folder root directory instead of the mailbox
root directory.   
.PARAMETER Recurse
If set, the function will recurse through all child directories, not just the first level.
.OUTPUTS
An array of Microsoft.Exchange.WebServices.Data.Folder objects found under indicated folder, 
or an exception if nothing found or an error occured.  
.EXAMPLE
Get-ERMChildFolder -EmailAddress 'bob@bob.com' -Folder 'Inbox\TestMe' 
#>
Function Get-ERMChildFolder {
    [CmdletBinding()] 
    param (
        [Parameter(Mandatory=$true)][string]$EmailAddress,
        [string]$Folder,
        [string]$TraceFile,
        [switch]$Archive,
        [switch]$Recurse
    )

    $ErrorActionPreference = "Stop"

    # Set up the EWS Enviroment

    $Service = SetupEWSImpersonationService $EmailAddress $TraceFile

    # Get the Root Folder to start the search.  

    if ($Archive) {
        $RootFolderId = 
            [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot
    }
    else {
        $RootFolderId = 
            [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot
    }
    $TopFolderObj = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $RootFolderId)

    if ($Folder) {

        Write-Verbose "Checking for existence of $Folder"
        
        # Do a basic sanity check on the path given to the folder.  

        $PathElements = $Folder -split '\\'
        foreach ($Element in $PathElements) {
            if ($Element -eq $null -or $Element.Length -eq 0) {
                throw "Empty folder name found in path."
            }
        }

        # Iterate through the path to get the folder one level higher than the new folder.

        for ($i = 0; $i -lt $PathElements.Count; $i++) {
            $TopFolderObj = GetFolder  $TopFolderObj $PathElements[$i]
            if ($TopFolderObj -eq $null) {
                throw "Cannot locate folder $($PathElements[$i]) in the given hierarchy." 
            }
        }
    }

    $AllResults = @()

    # Set up the search, and iterate through the results to remove any items that 
    # are not applicable for retention.    

    $View = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
    if ($Recurse) {
        $View.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    }

    $PropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
    $View.PropertySet = $Propset

    do {
        $FolderResults = $TopFolderObj.FindFolders($View)
        foreach ($ChildFolder in $FolderResults.Folders) {
            $AllResults += $ChildFolder
        }
        $View.Offset += $FolderResults.Folders.Count
    } while  ($FolderResults.MoreAvailable)

    Write-Verbose "Finished Successfully!"
    return $AllResults
}






