Graph-SearchAndDelete

This PowerShell script can be used to search and delete content from a mailbox using the Graph API. The search can be performed using the sender email address, subject, created/received time, or message body. The search can only be executed against the primary mailbox. The Graph API currently does not support accessing an archive mailbox.

Requirements
An application registration must be created in Azure AD for the tenant and this application must have the Mail.ReadWrite API permission for Microsoft Graph. The script also required the MSAL.PS PowerShell module.

NOTE:
Message body searches are limited to 275 results per folder. Multiple runs are needed to delete more than 275 items from a folder.

How To Run
$secret = ConvertTo-SecureString -String "RQD8Q~DAMByQMCeLMgg_QGfIMS3y" -AsPlainText -Force
This syntax will search the Inbox for items from a sender with the email address kelly@contoso.com and generate a CSV file with the results.

.\Graph-SearchAndDelete.ps1 -Mailbox jim@contoso.com -OutputPath C:\Temp\ -SenderAddress kelly@contoso.com -IncludeFolderList Inbox -OAuthClientId 2e542266-a1b2-4567-8901-abcdccd61976 -OAuthTenantId 9101fc97-a2e6-2255-a2d5-83e051e52057 -OAuthClientSecret $secret

This syntax will search the entire mailbox from items where the subject contains the word Microsoft and the message body contains the word Exchange, generate a CSV file with the results, and delete the items.

.\Graph-SearchAndDelete.ps1 -Mailbox jim@contoso.com -OutputPath C:\Temp\ -Subject Microsoft -MessageBody Exchange -DeleteContent -OAuthClientId 2e542266-a1b2-4567-8901-abcdccd61976 -OAuthTenantId 9101fc97-a2e6-2255-a2d5-83e051e52057 -OAuthClientSecret $secret

This syntax will search the Recoverable Items for items created between 01 Jan 2024 and 31 Jan 2024, generate a CSV file with the results, and delete the items.

.\Graph-SearchAndDelete.ps1 -Mailbox jim@contoso.com -OutputPath C:\Temp\ -CreatedAfter 2024-01-01 -CreatedBefore 2024-01-31 -SearchDumpster -DeleteContent -OAuthClientId 2e542266-a1b2-4567-8901-abcdccd61976 -OAuthTenantId 9101fc97-a2e6-2255-a2d5-83e051e52057 -OAuthClientSecret $secret

Parameters

Mailbox
The Mailbox parameter specifies the mailbox to be accessed

ProcessSubfolders
The ProcessSubfolders parameter is a switch to enable searching the subfolders of any specified folder

IncludeFolderList
The IncludeFolderList parameter specifies the folder(s) to be searched (if not present, then the Inbox folder will be searched).  Any exclusions override this list.

ExcludeFolderList
The ExcludeFolderList parameter specifies the folder(s) to be excluded (these folders will not be searched).

SearchDumpster
The SearchDumpster parameter is a switch to search the recoverable items.

CreatedBefore
The CreatedBefore parameter specifies only messages created before this date will be searched.

CreatedAfter
The CreatedAfter parameter specifies only messages created after this date will be searched.

Subject
The Subject paramter specifies the subject string used by the search.

Sender
The Sender paramter specifies the sender email address used by the search.

MessageBody
The MessageBody parameter specifies the body string used by the search.

MessageId
The MessageId parameter specified the MessageId used by the search.

DeleteContent
The DeleteContent parameter is a switch to delete the items found in the search results (moved to Deleted Items).

AzureEnvironment
The AzureEnvironment parameter specified the Azure environment for the tenant.

PermissionType
The PermissionType parameter specifies whether the app registrations uses delegated or application permissions.

OAuthClientId
The OAuthClientId parameter is the Azure Application Id that this script uses to obtain the OAuth token.  Must be registered in Azure AD.

OAuthTenantId
The OAuthTenantId paramter is the tenant Id where the application is registered (Must be in the same tenant as mailbox being accessed).

OAuthRedirectUri
The OAuthRedirectUri parameter is the redirect Uri of the Azure registered application.

OAuthClientSecret
The OAuthClientSecret parameter is the the secret for the registered application.

OAuthCertificate
The OAuthCertificate parameter is the certificate for the registerd application. Certificate auth requires MSAL libraries to be available..

CertificateStore
The CertificateStore parameter specifies the certificate store where the certificate is loaded.

OutputPath
The OutputPath parameter specifies the path for the EWS usage report.

ThrottlingDelay
The ThrottlingDelay parameter specifies the throttling delay (time paused between sending EWS requests) - note that this will be increased automatically if throttling is detected"