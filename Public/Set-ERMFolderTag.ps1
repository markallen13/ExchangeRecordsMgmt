<#
.SYNOPSIS
Set-ERMFolderTag takes in a *live* mailbox folder and sets a retention policy tag on that 
folder.  
.DESCRIPTION
Set-ERMFolderTag takes a string that is the name of a folder retention tag that exists in 
the exchange server and applies that tag to the given instance of 
Microsoft.Exchange.WebServices.Data.Folder.  Script fails if the folder object doesn't 
exist or the tag is not unique or does not exist. 

To apply changes from this script quickly, run the following command after this one to
start the Managed Folder Assistant process immediately.  

    Start-ManagedFolderAssistant <MailboxName>

This module has multiple requirements that are documented in the README.md file, including 
(1) Exchange Web Services us be installed and available, (2) the user must be able to open 
a remote session with the Exchange server, and (3) the user must be granted the 
ApplicationImpersonation role in Exchange.  Read the README.md for more information.  

Much of the module code is based on several scripts created by David Barrett, Microsoft.
.PARAMETER Folder
The instance of Microsoft.Exchange.WebServices.Data.Folder that will be stamped with the 
tag.  Folders used for containing contacts, calendar items, and other non-email items are 
not accepted.  
.PARAMETER TagName
The name of the tag to be resoved in the Exchange server and applied to the folder.  The
format of this is the same as the 'Identity' flag from the Get-RetentionPolicyTag cmdlet
from Exchange Management Shell.  Therefore it can take wildcards, although the resolved
name must be unique.         
.EXAMPLE
Set-ERMFolderTag -Folder $MyFolder -TagName '5 Years Delete' -Verbose
#>
Function Set-ERMFolderTag {
    [CmdletBinding()] 
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
        [Parameter(Mandatory = $true, Position = 1)]
        [string]$TagName
    )

    $ErrorActionPreference = "Stop"

    if (($Folder.FolderClass -ne 'IPF.Note') -or ($Folder.GetType().Name -ne 'Folder')) {
        throw "Folder $($Folder.DisplayName) is not used for E-Mail Messages."
    }

    # Search for the tag to be applied and verify it is a unique tag.   

    $Tag = Get-RetentionPolicyTag -Identity $TagName
    if ($Tag -eq $null) {
        throw "Indicated retention tag cold not be found.  Please clarify."
    }
    if ($Tag.GetType().FullName -eq 'System.Object[]') {
        throw "Indicated retention tag matches more than one value.  Please clarify."
    }

    Write-Verbose "Applying Retention Tag '$($Tag.Name)'"

    # Determine MAPI Property settings based on Archive vs Retention stamp.  

    [int]$TagPeriod = $Tag.AgeLimitForRetention.Split('.')[0]
    $TagGuid = $Tag.Guid.ToByteArray()

    $IsArchiveTag = ($Tag.RetentionAction -eq 'MoveToArchive')
    if ($IsArchiveTag) {
        $Flags = (GetFolderRetentionFlags $Folder) -bor $ExplicitArchiveFlags
        $Folder.SetExtendedProperty($RetentionFlagsDef, $Flags) 
        $Folder.SetExtendedProperty($ArchivePeriodDef, $TagPeriod)
        $Folder.SetExtendedProperty($ArchivePolicyTagDef, $TagGuid)
    }
    else {
        $Flags = (GetFolderRetentionFlags $Folder) -bor $ExplicitRetentionFlags
        $Folder.SetExtendedProperty($RetentionFlagsDef, $Flags) 
        $Folder.SetExtendedProperty($RetentionPeriodDef, $TagPeriod)
        $Folder.SetExtendedProperty($RetPolicyTagDef, $TagGuid)
    } 

    # Time to apply the retention policy tag. 

    $FlagsStr = [string]::Format("0x{0:x}", $Flags)
    Write-Verbose ("Stamping Folder $($Folder.DisplayName) with flags $($FlagsStr)" +  
        ", period $($Tag.AgeLimitForRetention.TotalDays), and tag name '$($Tag.Name)'")
    $Folder.Update()

    # Check to make sure the tag was applied.

    $CurrentTag = GetFolderRetentionTag $Folder $IsArchiveTag
    if ($CurrentTag -ne $Tag.Name) {
        throw "Tag $($Tag.Name) could not be applied to folder $($Folder.DisplayName)!" 
    }
    else {
        Write-Verbose "Retention tag verification successful."
    }

    Write-Verbose "Finished!"
}






