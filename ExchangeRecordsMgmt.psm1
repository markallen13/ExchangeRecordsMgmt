<#
.SYNOPSIS
The ExcangeRecordsManagement module is used to manipulate Exhange user folder structures
(typically viewed in Outlook) to create folder based records retention systems across an
organization.

.DESCRIPTION
This module encapsulates functionality in Exchange Web Services and the Exchange 
Management Shell to perform folder and item manipulation tasks on arbitrary user mailboxes 
to create a basic retention system, using Retention Policy Tags already defined in the 
Exchange server.  The primary use is to create folder structures and apply or remove 
Retention Policy Tags and Archive Policy Tags in folder and items in the Outlook mailbox 
folder hierarchy.  

This module was originally developed and tested to work on Exchange Server 2010 SP2.  

This module has multiple requirements that are documented in the README.md file, including 
(1) Exchange Web Services us be installed and available, (2) the user must be able to open 
a remote session with the Exchange server, and (3) the user must be granted the 
ApplicationImpersonation role in Exchange.  Read the README.md for more information.  

The module structure is based on work by by Warren F. (RamblingCookieMonster) found at:  
http://ramblingcookiemonster.github.io/Building-A-PowerShell-Module/.

Much of the module code is based on several scripts created by David Barrett, Microsoft.
#>

#-----------------------------------------------------------------------------------------

# If there are any problems loading the required enviroment, stop processing.  

$ErrorActionPreference = "Stop"

# Import constants used throughout this module.  These can be set using Settings.xml

[xml]$ConfigFile = Get-Content "$($PSScriptRoot)\Settings.xml"

$EWSPath = ($ConfigFile.Settings.EWSPath).Trim()
$EWSExchangeVersion = ($ConfigFile.Settings.EWSExchangeVersion).Trim()
$ExchPowerShellURI = ($ConfigFile.Settings.ExchPowerShellURI).Trim()
$ExchPowerShellAuth = ($ConfigFile.Settings.ExchPowerShellAuth).Trim()

# Search for the EWS libraries within Program Files (x64 and x86)

$ProgramDir = $Env:ProgramFiles
$ProgramDirx86 = [environment]::GetEnvironmentVariable("ProgramFiles(x86)")

if (Get-Item -Path ($programDir + $EWSPath) -ErrorAction SilentlyContinue) {
    $EWSLibrary = $programDir + $EWSPath
}
elseif (Get-Item -Path ($ProgramDirx86 + $EWSPath) -ErrorAction SilentlyContinue) {
    $EWSLibrary = $ProgramDirx86 + $EWSPath
}
else {
    throw "Failed to locate EWS Managed API, cannot continue."    
}
      
# Add the EWS library before we try to read the remainder of the module. 

try {
    Write-Host ([string]::Format("Using managed API found at: {0}", $EWSLibrary))
    Add-Type -Path $EWSLibrary
} 
catch {
    throw "Cannot Import EWS Library at $($EWSLibrary): `n" + $_
}

# Import Exchange Shell Cmdlets if Necessary.

try {
    $CmdletsNeeded = (Get-Command -Name 'Get-RetentionPolicyTag' -ErrorAction 'SilentlyContinue') -eq $null
    if ($CmdletsNeeded) {
        Write-Host "Importing Required Cmdlets from $($ExchPowerShellURI)"
        $ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri `
            $ExchPowerShellURI -Authentication $ExchPowerShellAuth
        Import-PSSession $ExchSession -CommandName 'Get-RetentionPolicyTag'
    }
}
catch {
    throw "Cannot Connect to Exchange at $($ExchPowerShellURI): `n" + $_
}

# Get public and private function definition files.

$Public = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue)
$Private = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue)

# Dot source the files

foreach ($import in @($Public + $Private)) {
    try {
        . $import.fullname
    }
    catch {
        Write-Error -Message "Failed to import function $($import.fullname): $_"
    }
}

Export-ModuleMember -Function $Public.Basename


