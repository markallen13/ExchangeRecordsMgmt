# Exchange-Records-Mgmt

PowerShell module that works with Exchange Web Services to create folders, set retention periods, and other tasks as needed related to setting up records management policies for an organization.

### Overview

This module encapsulates functionality in Exchange Web Services and the Exchange Management Shell to perform folder and item manipulation tasks on arbitrary user mailboxes to create a basic retention system, using Retention Policy Tags already defined in the Exchange server.  The primary use is to create folder structures and apply or remove Retention Policy Tags and Archive Policy Tags in folder and items in the Outlook mailbox folder hierarchy.  

This module was originally developed and tested to work on Exchange Server 2010 SP3.  It has not been tested or verified on later versions of Exchange or Exchange Online.

### Module Requirements 

The user will need to update the configuration parameters defined in Settings.xml before using the module.

The module requires the Exchange Web Services library to be installed and available.  You can download this package from:  

https://www.microsoft.com/en-us/download/details.aspx?id=42951

The module requires that either the user have permissions to open a remote session with the Exchange Server, either by using the module in the Exchange Management Shell, or by automatically connecting the Exchange Server directly through a PSSession.  The module checks for this when it is initialized, and only opens a PSSession with Exchange server if needed.  You will need to update the module configuration with the domain name of the Exchange Server and the authentication method to be used.  The following authentication methods can be used:

https://docs.microsoft.com/en-us/dotnet/api/system.management.automation.runspaces.authenticationmechanism?view=powershellsdk-1.1.0

The module requires that the local network logon account also be granted the ApplicationImpersonation role in Exchange.  I did this by setting up a RoleAssigneeType of 'User', which means my user is directly assigned to this role as opposed to a group.  Here's the PowerShell I used (this does require Exchange Management Console):

`New-ManagementRoleAssignment -Role 'ApplicationImpersonation' -User <Your Domain User Name>`

### How to Use the Module

The user must already have defined a set or retention policies in his or her Exchange environment using the management consol or the management shell.  For more information , see: 

https://technet.microsoft.com/en-us/library/dd297955(v=exchg.141).aspx

The user first must obtain an instance of an Microsoft.Exchange.WebServices.Data.Folder object that he or she will be working with, using one of these cmdlets:

    Get-ERMFolder
    Get-ERMChildFolder
    New-ERMFolder

The user can then manipulate retention tags by passing in the folder object.  The user can also obtain access to individual e-mail items using Get-ERMFolderItem.  There are 3 sets of cmdlets for folders and e-email items used to obtain, clear, and set the retention policy, which are named accordingly. 

An important note is that the EWS objects returned by this module (specifically the Microsoft.Exchange.WebServices.Data.Folder and Microsoft.Exchange.WebServices.Data.EmailMessage objects) remain connected to the Exchange Server once they are returned to the user.  It follows that the user can manipulate these items directly using the object's methods.  For example, you can change the name of the objects or delete them using the methods defined in the API.  See the following for more information:

https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.folder_methods(v=exchg.80).aspx
https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.emailmessage_methods(v=exchg.80).aspx

Changes to the Retention Policy Tags will likely not apply immediately.  To apply changes from this script quickly, run the following command after this one to
start the Managed Folder Assistant process.  

    Start-ManagedFolderAssistant <MailboxName>

When retrieving folder objects, you have the option of creating a Tracefile.  This will show all communication between the server and your client.  Note that the tracefile will continue to be populated anywhere you use the folder object (in other words, not just when you specify the Tracefile parameter). 

### Why is Exchange Web Services API required?

All the information I was able to obtain says the Exchange Management Console doesn't support what this script does, at least on Exchange 2010 SP3. To wit:

https://social.technet.microsoft.com/Forums/office/en-US/69d4d386-03c3-4c82-a8db-cb7819f63016/encounter-error-with-command-newmailboxfolder-the-specified-mailbox-doesnt-exist?forum=exchangesvrgeneral

https://blogs.technet.microsoft.com/exchange/2013/05/20/using-exchange-web-services-to-apply-a-personal-tag-to-a-custom-folder/

### Why does the module use functions from the Exchange Management Shell?

Because it was developed for Exchange Server 2010 SP3.  More to the point, the ability to obtain a list of the current Retention Policy Tags in EWS is only supported on Exchange Online and Exchange Server 2013 and above, so I had to find a different way.  I'm interested to update the module to do both, but I don't have that environment available to me.    

### A Few Credits

The module structure is based on work by by Warren F. (RamblingCookieMonster) found at:  

http://ramblingcookiemonster.github.io/Building-A-PowerShell-Module/.

Much of the module code is based on several scripts created by David Barrett, Microsoft.
