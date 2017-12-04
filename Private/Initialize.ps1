<#
File contains functions that are used for initializing the Exchange Web Services enviroment 
and connecting to the exchange server.  These are all helper functions.

This module has multiple requirements that are documented in the README.md file, including 
(1) Exchange Web Services us be installed and available, (2) the user must be able to open 
a remote session with the Exchange server, and (3) the user must be granted the 
ApplicationImpersonation role in Exchange.  Read the README.md for more information.  

Much of the module code is based on several scripts created by David Barrett, Microsoft.
#>

#-----------------------------------------------------------------------------------------

<#
.SYNOPSIS
This is a derived class of ITraceListener used to write debugging trace information 
to a specified file.  This is used to show the communication between the exchange server 
and this client.  It's an optional feature.  
#>
Class TraceListener : Microsoft.Exchange.WebServices.Data.ITraceListener {
    hidden [string]$TraceFile = $null

    InitializeTraceFile([string]$TraceFile)
    {
        if (Get-Item -Path ($TraceFile) -ErrorAction SilentlyContinue) {
            Remove-Item $TraceFile
        }
        $this.TraceFile = $TraceFile
    }

    Trace([string]$traceType, [string]$traceMessage)
    {
        if ($this.TraceFile) {
            $Output = "****Trace Message Type:  $($traceType)****`n`n"
            $Output += $traceMessage
            $Output += "`n"

            $Output | Out-File -FilePath $this.TraceFile -Append 
        }
    }
}

<#
.SYNOPSIS
Used to do the initial setup of the EWS enviroment for processing. 
.PARAMETER EmailAddress
The email address of the mailbox for processing, also used to discover the Exchange Server.  
.PARAMETER TraceFile
If this is set, EWS Tracing is enabled and written to the indicated file path.
.OUTPUTS
Returns a set up instance of Microsoft.Exchange.WebServices.Data.ExchangeService, or an 
exception if there was an error.
#>
Function SetupEWSImpersonationService() {
    [CmdletBinding()] 
    param (
        [string]$EmailAddress,
        [string]$TraceFile
    )

    # Set up the Exchange Service for use.

    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(
        [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$EWSExchangeVersion)

    # Autodiscover the Web Service URL using the provided email.  

    Write-Host "Processing mailbox $EmailAddress" 
    try {
        $Service.AutodiscoverUrl($EmailAddress)
    }
    catch {
        throw "Mailbox $($EmailAddress) cannot be found for autodiscover."
    }
    if ([string]::IsNullOrEmpty($Service.Url)) {
        throw "Mailbox $($EmailAddress) cannot be found for autodiscover."
    }
    Write-Verbose ([string]::Format("EWS Url found: {0}", $Service.Url))

    # Start Impersonating the Mailbox.

    Write-Verbose "Impersonating $EmailAddress"
    $Service.ImpersonatedUserId = 
        New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
        [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
    
    # Set Up EWS Tracing.

    if (![String]::IsNullOrEmpty($TraceFile))
    {
        Write-Verbose "Starting trace and creating trace file at $($TraceFile)"

        $Service.TraceListener = New-Object TraceListener
        $Service.TraceListener.InitializeTraceFile($TraceFile)
        $Service.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
        $Service.TraceEnabled = $True
    }

    return $Service
}










