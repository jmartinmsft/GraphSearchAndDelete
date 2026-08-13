# Graph-SearchAndDelete

Search and delete content from a user's mailbox using the Graph API.

## Description
This script can be used to search and delete content from a mailbox. The search criteria can include sender's email address, subject, created/received time, or message body. A report is generated with a list of items that will be/are deleted from the mailbox. The Delete parameter must be included for the script to delete the items.

## Requirements
1. The script requires an application registration in Entra ID that has the Microsoft Graph Mail.ReadWrite and MailboxSettings.Read permission.

## Note
Message body searches are limited to 275 results per folder. Multiple runs are needed to delete more than 275 items from a folder.

## Usage
Search the Inbox for items from a sender and only generate a CSV file with the results:
```powershell
$secret = ConvertTo-SecureString -String "xxxxxxxxxxxxxxxxxxxxxxxxxx" -AsPlainText -Force
.\Graph-SearchAndDelete.ps1 -Mailbox jim@contoso.com -OutputPath C:\Temp\ -SenderAddress kelly@contoso.com -IncludeFolderList Inbox -OAuthClientId 2e542266-a1b2-4567-8901-abcdccd61976 -OAuthTenantId 9101fc97-a2e6-2255-a2d5-83e051e52057 -OAuthClientSecret $secret
```
Search the entire mailbox for items containing a subject and message body and delete those items in batches of 10 items:
```powershell
.\Graph-SearchAndDelete.ps1 -Mailbox jim@contoso.com -OutputPath C:\Temp\ -Subject Microsoft -MessageBody Exchange -DeleteContent -OAuthClientId 2e542266-a1b2-4567-8901-abcdccd61976 -OAuthTenantId 9101fc97-a2e6-2255-a2d5-83e051e52057 -OAuthClientSecret $secret -BatchSize 10
```
Search the recoverable items for items within a date range and delete those items:
```powershell
.\Graph-SearchAndDelete.ps1 -Mailbox jim@contoso.com -OutputPath C:\Temp\ -CreatedAfter 2024-01-01 -CreatedBefore 2024-01-31 -SearchDumpster -DeleteContent -OAuthClientId 2e542266-a1b2-4567-8901-abcdccd61976 -OAuthTenantId 9101fc97-a2e6-2255-a2d5-83e051e52057 -OAuthClientSecret $secret
```
Search the archive mailbx (including aux archive mailboxes) for items containing the subject Graph from sender shared@contoso.com that were sent before a date in the GraphArchive folder
```powershell
.\Graph-SearchAndDelete.ps1 -OAuthClientId 7fc9c210-fa39-4c7a-83b8-9c6970b3c16a -OAuthTenantId 9101fc97-a2e6-2255-a2d5-83e051e52057 -OAuthCertificate 7765BEC834A02110DF8686D13436ABC8BE265917 -CertificateStore CurrentUser -PermissionType Application -Mailbox jim@contoso.com -Archive -IncludeFolderList GraphArchive -CreatedBefore (Get-Date -Date '7/23/2026') -Sender shared@contoso.com -Subject Graph
```

## Parameters

**Mailbox** - The Mailbox parameter specifies the mailbox to be accessed

**Archive** - The Archive parameter is a switch to search the archive mailbox (otherwise, the main mailbox is searched).

**ProcessSubfolders** - The ProcessSubfolders parameter is a switch to enable searching the subfolders of any specified folder

**IncludeFolderList** - The IncludeFolderList parameter specifies the folder(s) to be searched (if not present, then the Inbox folder will be searched).  Any exclusions override this list.

**ExcludeFolderList** - The ExcludeFolderList parameter specifies the folder(s) to be excluded (these folders will not be searched).

**SearchDumpster** - The SearchDumpster parameter is a switch to search the recoverable items.

**CreatedBefore** - The CreatedBefore parameter specifies only messages created before this date will be searched.

**CreatedAfter** - The CreatedAfter parameter specifies only messages created after this date will be searched.

**Subject** - The Subject paramter specifies the subject string used by the search.

**Sender** - The Sender paramter specifies the sender email address used by the search.

**MessageBody** - The MessageBody parameter specifies the body string used by the search.

**DeleteContent** - The DeleteContent parameter is a switch to delete the items found in the search results (moved to Deleted Items).

**HardDelete** The HardDelete parameter is a switch to hard-delete the items found in the search results

**AzureEnvironment** - The AzureEnvironment parameter specified the Azure environment for the tenant.

**PermissionType** - The PermissionType parameter specifies whether the app registrations uses delegated or application permissions.

**OAuthClientId** - The OAuthClientId parameter is the Azure Application Id that this script uses to obtain the OAuth token.  Must be registered in Azure AD.

**OAuthTenantId** - The OAuthTenantId paramter is the tenant Id where the application is registered (Must be in the same tenant as mailbox being accessed).

**OAuthRedirectUri** - The OAuthRedirectUri parameter is the redirect Uri of the Azure registered application.

**OAuthClientSecret** - The OAuthClientSecret parameter is the the secret for the registered application.

**OAuthCertificate** - The OAuthCertificate parameter is the certificate for the registerd application. Certificate auth requires MSAL libraries to be available..

**CertificateStore** - The CertificateStore parameter specifies the certificate store where the certificate is loaded.

**Scope** - The Scope parameter specifies the API permissions

**OutputPath** - The OutputPath parameter specifies the path for the EWS usage report.

**LogFile** - The LogFile parameter specifies the full path for the script log file. If not specified, a log file is created in the OutputPath.

**BatchSize** - The BatchSize parameter specifies how many items to delete within a batch request.