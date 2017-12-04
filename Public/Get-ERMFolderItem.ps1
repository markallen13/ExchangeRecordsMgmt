<#
.SYNOPSIS
Get-ERMFolderItem  takes in a *live* mailbox folder and retrieves the individual items 
(typically emails) in the mail folder.   
.DESCRIPTION
Get-ERMFolderItem is used to retrieve the individual items in a given folder.  It allows
you to submit a query string to limit the results that are returned.  

This function can accept folders that are not used for holding e-mail, such as Contacts and 
Calendar.  

This module has multiple requirements that are documented in the README.md file, including 
(1) Exchange Web Services us be installed and available, (2) the user must be able to open 
a remote session with the Exchange server, and (3) the user must be granted the 
ApplicationImpersonation role in Exchange.  Read the README.md for more information.  

Much of the module code is based on several scripts created by David Barrett, Microsoft.
.PARAMETER Folder
The instance of Microsoft.Exchange.WebServices.Data.Folder to use for retrival.
.PARAMETER QueryString
 Allows the user to retrieve emails based on Advanced Query Syntax as defined in 
 https://msdn.microsoft.com/en-us/library/office/ee693615(v=exchg.150).aspx.  
.OUTPUTS
An array of Microsoft.Exchange.WebServices.Data.Item objects (Typical an EmailMessage but
may be other types of objects), or an exception there was an error.  
.EXAMPLE
Get-ERMFolderItem $DeletedItems -QueryString "Subject:Ducks" -Verbose
#>
Function Get-ERMFolderItem {
    [CmdletBinding()] 
    param (
        [Parameter(Mandatory=$true,Position=0)]
        [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
        [string]$QueryString
    )
    
    $ErrorActionPreference = "Stop"

    # Run the query, allowing for large result sizes.  

    $AllResults = @()
    $View = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)

    do {    
        if ($QueryString) {
            $ItemResults = $Folder.FindItems($QueryString, $View)
        }
        else {
            $ItemResults = $Folder.FindItems($View)
        }
        $View.Offset += $ItemResults.Items.Count
        $AllResults += $ItemResults

    } while ($ItemResults.MoreAvailable)

    Write-Verbose "$($AllResults.Count) items found."
    Write-Verbose "Finished Successfully!"
    return $AllResults
}
