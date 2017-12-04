<#
.SYNOPSIS
New-ERMMailFolder creates a mail folder in an Exchange Inbox.     
.DESCRIPTION
New-ERMMailFolder checks for the existence of a given mail folder in an outlook mailbox, 
and creates it if it doesn't exist.  No additional parameters are applied to the folder.
The Exchange object representing the object is returned upon success.    

This module has multiple requirements that are documented in the README.md file, including 
(1) Exchange Web Services us be installed and available, (2) the user must be able to open 
a remote session with the Exchange server, and (3) the user must be granted the 
ApplicationImpersonation role in Exchange.  Read the README.md for more information.  

Much of the module code is based on several scripts created by David Barrett, Microsoft.
.PARAMETER EmailAddress
The email address of the mailbox for processing, also used to discover the Exchange Server.       
.PARAMETER Folder
The name of the folder to be created.  Subfolders can be accessed using a backslash from
the mailbox root, e.g. Inbox\Folder1\Folder2.   
.PARAMETER TraceFile
If this parameter is set, then EWS Tracing is enabled and is written to the provided file
name.  This shows the communication between the exchange server and this client.  
.PARAMETER Archive
If set, the script works with the Archive Folder root directory instead of the mailbox
root directory.   
.OUTPUTS
An instance of Microsoft.Exchange.WebServices.Data.Folder that was created, or an exception 
if an error occured.  
.EXAMPLE
New-ERMFolder -EmailAddress 'bob@bob.com' -Folder 'Inbox\TestMe' 
#>
Function New-ERMFolder {
    [CmdletBinding()] 
    param (
        [Parameter(Mandatory=$true)][string]$EmailAddress,
        [Parameter(Mandatory=$true)][string]$Folder,
        [string]$TraceFile,
        [switch]$Archive
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

    Write-Verbose "Checking for existence of $Folder"

    # Do a basic sanity check on the path given to the folder.  

    $PathElements = $Folder -split '\\'
    foreach ($Element in $PathElements) {
        if ($Element -eq $null -or $Element.Length -eq 0) {
            throw "Empty folder name found in path."
        }
    }

    # Iterate through the path to get the folder one level higher than the new folder.

    for ($i = 0; $i -lt $PathElements.Count - 1; $i++) {
        $TopFolderObj = GetFolder $TopFolderObj $PathElements[$i]
        if ($TopFolderObj -eq $null) {
            throw "Cannot locate folder $($PathElements[$i]) in the given hierarchy." 
        }
    }

    # Attempt to create the folder.  

    $NewFolderName = $PathElements[$PathElements.Count - 1]
    $FolderObj = GetFolder $TopFolderObj $NewFolderName

    if ($FolderObj -eq $null) {
        $FolderObj = New-Object Microsoft.Exchange.WebServices.Data.Folder($Service)
        $FolderObj.DisplayName = $NewFolderName 
        $FolderObj.FolderClass = "IPF.Note"
        $FolderObj.Save($TopFolderObj.Id)
        Write-Verbose "Created new folder $NewFolderName"
    }
    else {
        throw "Mail Folder '$Folder' already exists in mailbox $Identity." 
    }

    # Run a check to make sure the folder was created.  

    $FolderObj = GetFolder $TopFolderObj $NewFolderName
    if ($FolderObj -ne $null) {
        Write-Verbose "Successfully verified new folder $Folder"
    }
    else {
        throw "Mail Folder '$Folder' already exists in mailbox $Identity." 
    }

    Write-Verbose "Finished Successfully!"
    $FolderObj
}
