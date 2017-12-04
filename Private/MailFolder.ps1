<#
File contains shared functions for working with Outlook mailbox folders.  These are all  
helper functions.

This module has multiple requirements that are documented in the README.md file, including 
(1) Exchange Web Services us be installed and available, (2) the user must be able to open 
a remote session with the Exchange server, and (3) the user must be granted the 
ApplicationImpersonation role in Exchange.  Read the README.md for more information.  

Much of the module code is based on several scripts created by David Barrett, Microsoft.
#>

#-----------------------------------------------------------------------------------------

<#
.SYNOPSIS
Checks for the existance of the given folder just below the given "top level" folder.  
.PARAMETER TopFolderObj
The top level folder to search under.   
.PARAMETER FolderName
The name of the folder to check for. 
.OUTPUTS
Returns the folder object for further processing as a Exchange Folder object.  
#>
Function GetFolder() {
    [CmdletBinding()] 
    param (
        [Microsoft.Exchange.WebServices.Data.Folder]$TopFolderObj, 
        [string]$FolderName
    )

    # Add MAPI properies to be included when items are found in the search.

    $View = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
    $PropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
    $View.PropertySet = $Propset

    $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
        [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $FolderName)
    $FolderResults = $TopFolderObj.FindFolders($SearchFilter, $View)

    if ($FolderResults.TotalCount -eq 1) {
        return $FolderResults.Folders[0]
    }
    return $null
}