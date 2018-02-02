<#
File contains shared functions for working with Outlook mailbox retention tags.  These are 
all helper functions.

This module has multiple requirements that are documented in the README.md file, including 
(1) Exchange Web Services us be installed and available, (2) the user must be able to open 
a remote session with the Exchange server, and (3) the user must be granted the 
ApplicationImpersonation role in Exchange.  Read the README.md for more information.  

Much of the module code is based on several scripts created by David Barrett, Microsoft.
#>

#-----------------------------------------------------------------------------------------

# MAPI property constants used for accessing Exchange Web Services API.
# This information is defined in in MS-OXCMSG 2.2.1.58, which can be found at:  
# https://msdn.microsoft.com/en-us/library/ee158272(v=exchg.80).aspx

$RetentionFlagsDef = 
    New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(
    0x301D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$RetentionPeriodDef = 
    New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(
    0x301A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$RetPolicyTagDef = 
    New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(
    0x3019, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$ArchivePeriodDef = 
    New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(
    0x301E, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$ArchivePolicyTagDef = 
    New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(
    0x3018, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)

Set-Variable ExplicitArchiveFlags -option Constant -value 0x90
Set-Variable ExplicitRetentionFlags -option Constant -value 0x89

<#
.SYNOPSIS
Use to obtain the tag name from the provided folder.  
.PARAMETER Folder
An instance of a Folder object to be checked. 
.PARAMETER ArchiveTag
Retrieves the folder's archive policy tag instead of it's retention policy. 
.OUTPUTS
A string representing the name of the retention policy that was located, or 'None' if 
no tag was found.  
#>
Function GetFolderRetentionTag() {
    [CmdletBinding()] 
    param (
        [Microsoft.Exchange.WebServices.Data.Folder]$Folder, 
        [bool]$ArchiveTag
    )

    # If there are no retention flags, then no policy is applied regardless.   

    if ((GetFolderRetentionFlags -Folder $Folder) -eq 0) {
        return 'None'
    }

    if ($ArchiveTag) {
        $PolicyTag = $ArchivePolicyTagDef
    }
    else {
        $PolicyTag = $RetPolicyTagDef
    }

    # Load the MAPI property for the folder. 

    $PropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
    $Propset.Add($PolicyTag) 
    $Folder.Load($Propset)

    # Attempt to retrieve the property.

    $PropValue = $null
    $Result = $Folder.TryGetProperty($PolicyTag, [ref]$PropValue)
    if (!$result) {
        Write-Verbose "A tag was not found on Folder $($Folder.DisplayName)."
        return "None"
    }
    $Guid = New-Object Guid @(,$PropValue)
    $TagObj = Get-RetentionPolicyTag | ? { $_.Guid -eq $Guid }

    Write-Verbose "Tag named '$($TagObj.Name)' was found on Folder $($Folder.DisplayName)." 
    return $TagObj.Name
}

<#
.SYNOPSIS
Use to obtain the current retention flags for the folder.  The meaning of the flags 
can be found at https://msdn.microsoft.com/en-us/library/ee202166(v=exchg.80).aspx.
.PARAMETER Folder
An instance of a Folder object to be checked.   
.OUTPUTS
A 32 bit integer with the flags that are set, as defined in the above article.
#>
Function GetFolderRetentionFlags() {
    [CmdletBinding()] 
    param (
        [Microsoft.Exchange.WebServices.Data.Folder]$Folder
    )
    # Load the MAPI property for the folder. 

    $PropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
    $Propset.Add($RetentionFlagsDef) 
    $Folder.Load($Propset)

    # Attempt to retrieve the property.

    $PropValue = $null
    $Result = $Folder.TryGetProperty($RetentionFlagsDef, [ref]$PropValue)
    if (!$result) {
        Write-Verbose "Retention Flags not set on Folder $($Folder.DisplayName)."
        return 0
    }

    $Flags = [string]::Format("0x{0:x}", $PropValue)
    Write-Verbose "Flags $($Flags) found on Folder $($Folder.DisplayName)"
    return $PropValue
}

<#
.SYNOPSIS
Use to obtain the tag name from the provided EmailMessage.  
.PARAMETER EmailMessage
An instance of a EmailMessage object to be checked. 
.PARAMETER ArchiveTag
Retrieves the Email's archive policy tag instead of it's retention policy.
.OUTPUTS
A string representing the name of the retention policy that was located, or 'None' if 
no tag was found.  
#>
Function GetEmailRetentionTag() {
    [CmdletBinding()] 
    param (
        [Microsoft.Exchange.WebServices.Data.EmailMessage]$EmailMessage, 
        [bool]$ArchiveTag
    )

    # If there are no retention flags, then no policy is applied regardless.   

    if ((GetEmailRetentionFlags -EmailMessage $EmailMessage) -eq 0) {
        return 'None'
    }

    if ($ArchiveTag) {
        $PolicyTag = $ArchivePolicyTagDef
    }
    else {
        $PolicyTag = $RetPolicyTagDef
    }

    # Load the MAPI property for the folder. 

    $PropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
    $Propset.Add($PolicyTag) 
    $EmailMessage.Load($Propset)

    # Attempt to retrieve the property.

    $PropValue = $null
    $Result = $EmailMessage.TryGetProperty($PolicyTag, [ref]$PropValue)
    if (!$result) {
        Write-Verbose "A tag was not found on Email with subject $($EmailMessage.Subject)."
        return "None"
    }
    $Guid = New-Object Guid @(,$PropValue)
    $TagObj = Get-RetentionPolicyTag | ? { $_.Guid -eq $Guid }

    Write-Verbose "Tag named '$($TagObj.Name)' was found on Email with subject $($EmailMessage.Subject)." 
    return $TagObj.Name
}

<#
.SYNOPSIS
Use to obtain the current retention flags for the Email.  The meaning of the flags 
can be found at https://msdn.microsoft.com/en-us/library/ee202166(v=exchg.80).aspx.
.PARAMETER EmailMessage
An instance of an EmailMessage object to be checked.   
.OUTPUTS
A 32 bit integer with the flags that are set, as defined in the above article.
#>
Function GetEmailRetentionFlags() {
    [CmdletBinding()] 
    param (
        [Microsoft.Exchange.WebServices.Data.EmailMessage]$EmailMessage
    )

    # Load the MAPI property for the folder. 

    $PropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
    $Propset.Add($RetentionFlagsDef) 
    $EmailMessage.Load($Propset)

    # Attempt to retrieve the property.

    $PropValue = $null
    $Result = $EmailMessage.TryGetProperty($RetentionFlagsDef, [ref]$PropValue)
    if (!$result) {
        Write-Verbose "Retention Flags not set on Email with subject $($EmailMessage.Subject)."
        return 0
    }

    $Flags = [string]::Format("0x{0:x}", $PropValue)
    Write-Verbose "Flags $($Flags) found on Email with subject $($EmailMessage.Subject)."
    return $PropValue
}
