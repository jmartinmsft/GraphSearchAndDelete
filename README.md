# Graph-SearchAndDelete

Search and delete content from a user's mailbox using the Graph API.

## Description
This script can be used to search and delete content from a mailbox. The search criteria can include sender's email address, subject, created/received time, or message body. A report is generated with a list of items that will be/are deleted from the mailbox. The Delete parameter must be included for the script to delete the items.

## Requirements
1. The script requires an application registration in Entra ID that has the Microsoft Graph Mail.ReadWrite application permission.

## Note
Message body searches are limited to 275 results per folder. Multiple runs are needed to delete more than 275 items from a folder.

## Usage
Search the Inbox for items from a sender and only generate a CSV file with the results:
```powershell
$secret = ConvertTo-SecureString -String "xxxxxxxxxxxxxxxxxxxxxxxxxx" -AsPlainText -Force
.\Graph-SearchAndDelete.ps1 -Mailbox jim@contoso.com -OutputPath C:\Temp\ -SenderAddress kelly@contoso.com -IncludeFolderList Inbox -OAuthClientId 2e542266-a1b2-4567-8901-abcdccd61976 -OAuthTenantId 9101fc97-a2e6-2255-a2d5-83e051e52057 -OAuthClientSecret $secret
```
Search the entire mailbox for items containing a subject and message body and delete those items:
```powershell
.\Graph-SearchAndDelete.ps1 -Mailbox jim@contoso.com -OutputPath C:\Temp\ -Subject Microsoft -MessageBody Exchange -DeleteContent -OAuthClientId 2e542266-a1b2-4567-8901-abcdccd61976 -OAuthTenantId 9101fc97-a2e6-2255-a2d5-83e051e52057 -OAuthClientSecret $secret
```
Search the recoverable items for items within a date range and delete those items:
```powershell
.\Graph-SearchAndDelete.ps1 -Mailbox jim@contoso.com -OutputPath C:\Temp\ -CreatedAfter 2024-01-01 -CreatedBefore 2024-01-31 -SearchDumpster -DeleteContent -OAuthClientId 2e542266-a1b2-4567-8901-abcdccd61976 -OAuthTenantId 9101fc97-a2e6-2255-a2d5-83e051e52057 -OAuthClientSecret $secret
```

## Parameters

**Mailbox** - The Mailbox parameter specifies the mailbox to be accessed

**ProcessSubfolders** - The ProcessSubfolders parameter is a switch to enable searching the subfolders of any specified folder

**IncludeFolderList** - The IncludeFolderList parameter specifies the folder(s) to be searched (if not present, then the Inbox folder will be searched).  Any exclusions override this list.

**ExcludeFolderList** - The ExcludeFolderList parameter specifies the folder(s) to be excluded (these folders will not be searched).

**SearchDumpster** - The SearchDumpster parameter is a switch to search the recoverable items.

**CreatedBefore** - The CreatedBefore parameter specifies only messages created before this date will be searched.

**CreatedAfter** - The CreatedAfter parameter specifies only messages created after this date will be searched.

**Subject** - The Subject paramter specifies the subject string used by the search.

**Sender** - The Sender paramter specifies the sender email address used by the search.

**MessageBody** - The MessageBody parameter specifies the body string used by the search.

**MessageId** - The MessageId parameter specified the MessageId used by the search.

**DeleteContent** - The DeleteContent parameter is a switch to delete the items found in the search results (moved to Deleted Items).

**AzureEnvironment** - The AzureEnvironment parameter specified the Azure environment for the tenant.

**PermissionType** - The PermissionType parameter specifies whether the app registrations uses delegated or application permissions.

**OAuthClientId** - The OAuthClientId parameter is the Azure Application Id that this script uses to obtain the OAuth token.  Must be registered in Azure AD.

**OAuthTenantId** - The OAuthTenantId paramter is the tenant Id where the application is registered (Must be in the same tenant as mailbox being accessed).

**OAuthRedirectUri** - The OAuthRedirectUri parameter is the redirect Uri of the Azure registered application.

**OAuthClientSecret** - The OAuthClientSecret parameter is the the secret for the registered application.

**OAuthCertificate** - The OAuthCertificate parameter is the certificate for the registerd application. Certificate auth requires MSAL libraries to be available..

**CertificateStore** - The CertificateStore parameter specifies the certificate store where the certificate is loaded.

**OutputPath** - The OutputPath parameter specifies the path for the EWS usage report.

**ThrottlingDelay** - The ThrottlingDelay parameter specifies the throttling delay (time paused between sending EWS requests) - note that this will be increased automatically if throttling is detected"