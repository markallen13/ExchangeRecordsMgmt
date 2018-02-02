<#
.SYNOPSIS
Clear-ERMFolderTag takes in a *live* mailbox folder and clears the retention policy tag on 
that folder.  
.DESCRIPTION
Clear-ERMFolderTag takes a instance of Microsoft.Exchange.WebServices.Data.Folder and 
removes any retention tag settings on it.  This causes the retention tag to revert back to 
the parent folder. 

If you wish to clear an Archive tag, use the -Archive flag, as these are not automatically
removed.  

To apply changes from this script quickly, run the following command after this one to
start the Managed Folder Assistant process immediately.  

    Start-ManagedFolderAssistant <MailboxName>

This module has multiple requirements that are documented in the README.md file, including 
(1) Exchange Web Services us be installed and available, (2) the user must be able to open 
a remote session with the Exchange server, and (3) the user must be granted the 
ApplicationImpersonation role in Exchange.  Read the README.md for more information.  

Much of the module code is based on several scripts created by David Barrett, Microsoft.
.PARAMETER Folder
The instance of Microsoft.Exchange.WebServices.Data.Folder to be modified.  Folders
used for containing contacts, calendar items, and other non-email items are not accepted.  
.PARAMETER ArchiveTag
If set, script clears the folder's archive policy tag instead of the retention policy tag. 
.EXAMPLE
Clear-ERMFolderTag -Folder $MyFolder -Verbose
#>

Function Clear-ERMFolderTag {
    [CmdletBinding()] 
    param (
        [Parameter(Mandatory=$true,Position=0)]
        [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
        [switch]$ArchiveTag
    )

    $ErrorActionPreference = "Stop"

    # Define which MAPI properties we want to clear and the new retention flags. 

    if ($ArchiveTag) {
        $PeriodDef = $ArchivePeriodDef
        $PolicyTagDef = $ArchivePolicyTagDef
        $Flags = (GetFolderRetentionFlags $Folder) -band (-bnot $ExplicitArchiveFlags)
    }
    else {
        $PeriodDef = $RetentionPeriodDef
        $PolicyTagDef = $RetPolicyTagDef
        $Flags = (GetFolderRetentionFlags $Folder) -band (-bnot $ExplicitRetentionFlags)
        Write-Verbose "Note:  To clear an Archive Tag, use the -Archive Switch"
    }

    # Remove the policy name and period properties. 

    $PropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
    $PropSet.Add($PeriodDef) 
    $PropSet.Add($PolicyTagDef) 
    $Folder.Load($PropSet)

    [void]$Folder.RemoveExtendedProperty($PeriodDef) 
    [void]$Folder.RemoveExtendedProperty($PolicyTagDef) 
    $Folder.Update()

    # Update the retnetion flags.

    $Folder.SetExtendedProperty($RetentionFlagsDef, $Flags) 
    $Folder.Update()
    Write-Verbose "Requested tag removed for folder $($Folder.DisplayName)" 

    # Check to make sure the tag was applied.

    $TagName = GetFolderRetentionTag -Folder $Folder -ArchiveTag $ArchiveTag
    if ($TagName -ne "None") {
        throw "Tag could not be cleared on folder $($Folder.DisplayName)!" 
    }

    Write-Verbose "Retention tag verification successful."
    Write-Verbose "Finished!"
}
