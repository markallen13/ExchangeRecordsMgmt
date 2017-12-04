<#
.SYNOPSIS
Get-ERMFolder retrieves a mail folder from the Exchange server.    
.DESCRIPTION
Get-ERMFolderchecks for the existence of a given mail folder in an outlook mailbox, and 
returns that mail folder to the caller as an Exchange object.  The name of the folder must
be relative to the top of the mailbox; the function does not search for subfolders.

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
The name of the folder to be retrieved.  Subfolders can be accessed using a backslash from
the mailbox root, e.g. Inbox\Folder1\Folder2.   
.PARAMETER TraceFile
If this parameter is set, then EWS Tracing is enabled and is written to the provided file
name.  This shows the communication between the exchange server and this client.  
.PARAMETER Archive
If set, the script works with the Archive Folder root directory instead of the mailbox
root directory.   
.OUTPUTS
An instance of Microsoft.Exchange.WebServices.Data.Folder that was found, or an exception if 
the folder doesn't exist or an error occured.  
.EXAMPLE
Get-ERMFolder -EmailAddress 'bob@bob.com' -Folder 'Inbox\TestMe' 
#>
Function Get-ERMFolder {
    [CmdletBinding()] 
    param (
        [Parameter(Mandatory=$true)][string]$EmailAddress,
        [Parameter(Mandatory=$true)][string]$Folder,
        [string]$TraceFile,
        [switch]$Archive
    )

    $ErrorActionPreference = "Stop"

    # Set up the EWS Enviroment for Impersonation.

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

    Write-Verbose "Checking for existence of $Folder"

    # Do a basic sanity check on the path given to the folder.  

    $PathElements = $Folder -split '\\'
    foreach ($Element in $PathElements) {
        if ($Element -eq $null -or $Element.Length -eq 0) {
            throw "Empty folder name found in path."
        }
    }

    # Iterate through the path to get the folder requested.

    for ($i = 0; $i -lt $PathElements.Count; $i++) {
        $TopFolderObj = GetFolder $TopFolderObj $PathElements[$i]
        if ($TopFolderObj -eq $null) {
            throw "Cannot locate folder $($PathElements[$i]) in the given hierarchy."
        }
    }

    Write-Verbose "Finished Successfully!"
    $TopFolderObj
}






