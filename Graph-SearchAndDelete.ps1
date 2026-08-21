<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 20260821.1529

[CmdletBinding()]
param (
    [Parameter(Position=0,Mandatory=$false,HelpMessage="The Mailbox parameter specifies the mailbox to be accessed.")]
    [ValidateNotNullOrEmpty()]
    # Reject characters that could alter the Graph request path (path traversal / query injection).
    [ValidatePattern('^[^\\/?#&\s]+$')]
    [string]$Mailbox,

    [Parameter(Mandatory=$False, HelpMessage="The Archive parameter is a switch to search the archive mailbox (otherwise, the main mailbox is searched).")]
    [switch]$Archive,

    [Parameter(Mandatory=$False, HelpMessage="The ProcessSubfolders parameter is a switch to enable searching the subfolders of any specified folder.")]
    [switch]$ProcessSubfolders,

    [Parameter(Mandatory=$False, HelpMessage="The IncludeFolderList parameter specifies the folder(s) to be searched (if not present, then the entire mailbox will be searched).  Any exclusions override this list.")]
    [object]$IncludeFolderList,

    [Parameter(Mandatory=$False, HelpMessage="The ExcludeFolderList parameter specifies the folder(s) to be excluded (these folders will not be searched).")]
    [object]$ExcludeFolderList,

    [Parameter(Mandatory=$false,HelpMessage="The SearchDumpster parameter is a switch to search the recoverable items.")] 
    [switch]$SearchDumpster,
    
    [ValidateSet("Global", "USGovernmentL4", "USGovernmentL5", "ChinaCloud")]
    [Parameter(Mandatory = $false)]
    [string]$AzureEnvironment = "Global",

    [Parameter(Mandatory=$false, HelpMessage="The PermissionType parameter specifies whether the app registrations uses delegated or application permissions")] [ValidateSet('Application','Delegated')]
    [string]$PermissionType="Application",
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthClientId parameter is the Azure Application Id that this script uses to obtain the OAuth token.  Must be registered in Azure AD.")] 
    [string]$OAuthClientId,
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthTenantId parameter is the tenant Id where the application is registered (Must be in the same tenant as mailbox being accessed).")] 
    [string]$OAuthTenantId,
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthRedirectUri parameter is the redirect Uri of the Azure registered application.")] 
    [string]$OAuthRedirectUri = "http://localhost:8004",
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthSecretKey parameter is the the secret for the registered application.")] 
    [SecureString]$OAuthClientSecret,
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthCertificate parameter is the certificate thumbprint for the registered application. Certificate auth requires MSAL libraries to be available.")] 
    # A thumbprint is hex only. Rejecting anything else stops the value being used to traverse the Cert: provider.
    [ValidatePattern('^[0-9a-fA-F]+$')]
    [string]$OAuthCertificate,
  
    [Parameter(Mandatory=$False,HelpMessage="The CertificateStore parameter specifies the certificate store where the certificate is loaded.")] [ValidateSet("CurrentUser", "LocalMachine")]
     [string] $CertificateStore = "CurrentUser",

    [Parameter(Mandatory=$false)]
    [object]$Scope= @("Mail.ReadWrite"),

    [Parameter(Mandatory=$false, HelpMessage="The ReceivedBefore parameter specifies only messages received before this date will be searched.")] 
    [DateTime]$ReceivedBefore,
    
    [Parameter(Mandatory=$false, HelpMessage="The ReceivedAfter parameter specifies only messages received after this date will be searched.")] 
    [DateTime]$ReceivedAfter,
    
    [Parameter(Mandatory=$False,HelpMessage="The Subject parameter specifies the subject string used by the search.")] 
    [string]$Subject=$null,
    
    [Parameter(Mandatory=$False,HelpMessage="The Sender parameter specifies the sender email address used by the search.")]
    [ValidatePattern("^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$|^$")]
    [string]$Sender=$null,

    [Parameter(Mandatory=$False,HelpMessage="The MessageBody parameter specifies the body string used by the search.")] 
    [string]$MessageBody=$null,

    [Parameter(Mandatory=$False,HelpMessage="The AttachmentName parameter specifies the name of the attachment used to filter search results.")]
    [string]$AttachmentName,

    [Parameter(Mandatory=$False,HelpMessage="The DeleteContent parameter is a switch to delete the content found by the search. If not specified, the script will only report the number of items that would be deleted.")]
    [switch]$DeleteContent,
    
    [Parameter(Mandatory=$False,HelpMessage="The HardDelete parameter is a switch to permanently delete the content found by the search. If not specified, the script will soft delete the items (move to Deleted Items folder).")]
    [switch]$HardDelete,

    [Parameter(Mandatory = $false, HelpMessage="The BatchSize parameter specifies the number of items to process in each batch.")]
    [ValidateRange(1, 20)]
    [int]$BatchSize = 20,

    [Parameter(Mandatory = $false, HelpMessage="The ResultSize parameter specifies the number of items to returned in each query.")]
    [ValidateRange(1, 500)]
    [int]$ResultSize = 500,

    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [Parameter(Mandatory = $true, HelpMessage="The OutputPath parameter specifies the path for the EWS usage report.")]
    [string]$OutputPath,

    [Parameter(Mandatory = $false, HelpMessage="The Confirm switch specifies whether to prompt for confirmation before performing delete actions.")]
    [boolean]$ConfirmDelete=$true
)

#region Logging
function Initialize-Log {
    param([Parameter(Mandatory)][string]$Path)

    $encoding = [System.Text.UTF8Encoding]::new($false)
    $Script:LogWriter = [System.IO.StreamWriter]::new($Path, $true, $encoding, 65536)
    $Script:LogWriter.AutoFlush = $false
    $Script:LogLinesSinceFlush = 0
}
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )

    # Defence in depth: strip credentials from every entry regardless of what the caller passed.
    # The cheap IndexOf checks avoid running two regexes over every one of the (many) DEBUG entries.
    if ($Message.IndexOf('=') -ge 0) {
        $Message = $Message -replace '(?i)(([?&]code|id_token|access_token|refresh_token|client_secret|client_assertion|assertion)=)[^&\s"'']+', '$1<redacted>'
    }
    if ($Message.IndexOf('Bearer', [StringComparison]::OrdinalIgnoreCase) -ge 0) {
        $Message = $Message -replace '(?i)(Bearer\s+)[A-Za-z0-9\-\._~\+\/]+=*', '$1<redacted>'
    }

    $timestamp = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "[$timestamp] [$Level] $Message"

    if ($null -ne $Script:LogWriter) {
        $Script:LogWriter.WriteLine($entry)
        $Script:LogLinesSinceFlush++

        # Periodically flush buffered entries; flush errors immediately.
        if ($Level -eq "ERROR" -or $Script:LogLinesSinceFlush -ge 50) {
            $Script:LogWriter.Flush()
            $Script:LogLinesSinceFlush = 0
        }
    }

    switch ($Level) {
        "ERROR" { Write-Host $entry -ForegroundColor Red }
        "WARN"  { Write-Host $entry -ForegroundColor Yellow }
        "DEBUG" { Write-Verbose $entry }
        default { Write-Host $entry -ForegroundColor Cyan }
    }
}

function Close-Log {
    if ($null -ne $Script:LogWriter) {
        $Script:LogWriter.Flush()
        $Script:LogWriter.Dispose()
        $Script:LogWriter = $null
    }
}

# Initialize log file
$Script:RunId = "{0}_{1}_{2}" -f (Get-Date -Format "yyyyMMdd_HHmmss_fff"), $PID, ([guid]::NewGuid().ToString("N").Substring(0, 8))
$Script:LogFile = Join-Path $OutputPath "GraphSearch_$($Script:RunId).log"

Initialize-Log -Path $Script:LogFile

# Some hosts (Windows PowerShell 5.1) still negotiate TLS 1.0/1.1 by default, which AAD and Graph reject.
try {
    if (([System.Net.ServicePointManager]::SecurityProtocol -band [System.Net.SecurityProtocolType]::Tls12) -ne [System.Net.SecurityProtocolType]::Tls12) {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    }
} catch {
    Write-Log "Unable to enforce TLS 1.2: $($_.Exception.Message)" -Level DEBUG
}

Write-Log "Script started. Mailbox: $Mailbox | Archive: $Archive | SearchDumpster: $SearchDumpster | PermissionType: $PermissionType"
Write-Log "Action: DeleteContent=$DeleteContent | HardDelete=$HardDelete | ConfirmDelete=$ConfirmDelete | BatchSize=$BatchSize | ResultSize=$ResultSize"
Write-Log ("Criteria: Subject='{0}' | Sender='{1}' | AttachmentName='{2}' | MessageBody='{3}' | ReceivedAfter='{4}' | ReceivedBefore='{5}'" -f $Subject, $Sender, $AttachmentName, $MessageBody, $ReceivedAfter, $ReceivedBefore)
Write-Log ("Scope: ProcessSubfolders={0} | IncludeFolderList='{1}' | ExcludeFolderList='{2}'" -f $ProcessSubfolders, ($IncludeFolderList -join '; '), ($ExcludeFolderList -join '; '))
#endregion

function Get-CloudServiceEndpoint {
    [CmdletBinding()]
    param(
        [string]$EndpointName
    )

    <#
        This shared function is used to get the endpoints for the Azure and Microsoft 365 services.
        It returns a PSCustomObject with the following properties:
            GraphApiEndpoint: The endpoint for the Microsoft Graph API
            ExchangeOnlineEndpoint: The endpoint for Exchange Online
            AutoDiscoverSecureName: The endpoint for Autodiscover
            AzureADEndpoint: The endpoint for Azure Active Directory
            EnvironmentName: The name of the Azure environment
    #>

    begin {
        Write-Log "Calling $($MyInvocation.MyCommand)" -Level DEBUG
    }
    process {
        # https://learn.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints
        switch ($EndpointName) {
            "Global" {
                $environmentName = "AzureCloud"
                $graphApiEndpoint = "https://graph.microsoft.com"
                $exchangeOnlineEndpoint = "https://outlook.office.com"
                $autodiscoverSecureName = "https://autodiscover-s.outlook.com"
                $azureADEndpoint = "https://login.microsoftonline.com"
                $managementApiEndpoint = "https://manage.office.com"
                break
            }
            "USGovernmentL4" {
                $environmentName = "AzureUSGovernment"
                $graphApiEndpoint = "https://graph.microsoft.us"
                $exchangeOnlineEndpoint = "https://outlook.office365.us"
                $autodiscoverSecureName = "https://autodiscover-s.office365.us"
                $azureADEndpoint = "https://login.microsoftonline.us"
                $managementApiEndpoint = "https://manage.office365.us"
                break
            }
            "USGovernmentL5" {
                $environmentName = "AzureUSGovernment"
                $graphApiEndpoint = "https://dod-graph.microsoft.us"
                $exchangeOnlineEndpoint = "https://outlook-dod.office365.us"
                $autodiscoverSecureName = "https://autodiscover-s-dod.office365.us"
                $azureADEndpoint = "https://login.microsoftonline.us"
                $managementApiEndpoint = "https://manage.protection.apps.mil"
                break
            }
            "ChinaCloud" {
                $environmentName = "AzureChinaCloud"
                $graphApiEndpoint = "https://microsoftgraph.chinacloudapi.cn"
                $exchangeOnlineEndpoint = "https://partner.outlook.cn"
                $autodiscoverSecureName = "https://autodiscover-s.partner.outlook.cn"
                $azureADEndpoint = "https://login.partner.microsoftonline.cn"
                $managementApiEndpoint = "https://manage.office.cn"
                break
            }
        }
    }
    end {
        return [PSCustomObject]@{
            EnvironmentName        = $environmentName
            GraphApiEndpoint       = $graphApiEndpoint
            ExchangeOnlineEndpoint = $exchangeOnlineEndpoint
            AutoDiscoverSecureName = $autodiscoverSecureName
            AzureADEndpoint        = $azureADEndpoint
            ManagementApiEndpoint  = $managementApiEndpoint
        }
    }
}

function Get-NewJsonWebToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$CertificateThumbprint,
        [ValidateSet("CurrentUser", "LocalMachine")][Parameter(Mandatory = $false)][string]$CertificateStore = "CurrentUser",
        [Parameter(Mandatory = $false)][string]$Issuer,
        [Parameter(Mandatory = $false)][string]$Audience,
        [Parameter(Mandatory = $false)][string]$Subject,
        [Parameter(Mandatory = $false)][int]$TokenLifetimeInSeconds = 3600,
        [ValidateSet("RS256", "RS384", "RS512")][Parameter(Mandatory = $false)][string]$SigningAlgorithm = "RS256"
    )

    <#
        Shared function to create a signed Json Web Token (JWT) by using a certificate.
        It is also possible to use a secret key to sign the token, but that is not supported in this function.
        The function returns the token as a string if successful, otherwise it returns $null.
        https://www.rfc-editor.org/rfc/rfc7519
        https://learn.microsoft.com/azure/active-directory/develop/active-directory-certificate-credentials
        https://learn.microsoft.com/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
    #>

    begin {
        Write-Log "Calling $($MyInvocation.MyCommand)" -Level DEBUG
    }
    process {
        try {
            $certificate = Get-ChildItem -LiteralPath "Cert:\$CertificateStore\My\$CertificateThumbprint"
            if ($certificate.HasPrivateKey) {
                $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
                # Base64url-encoded SHA-1 thumbprint of the X.509 certificate's DER encoding
                $x5t = [System.Convert]::ToBase64String($certificate.GetCertHash())
                #$x5t = ((($x5t).Replace("\+", "-")).Replace("/", "_")).Replace("=", "")
                $x5t = ((($x5t).Replace("+", "-")).Replace("/", "_")).Replace("=", "")                
                Write-Log "x5t is: $x5t" -Level DEBUG
            } else {
                Write-Log "We don't have a private key for certificate: $CertificateThumbprint and so cannot sign the token" -Level DEBUG
                return
            }
        } catch {
            Write-Log "Unable to import the certificate - Exception: $($Error[0].Exception.Message)" -Level ERROR
            return
        }

        $header = [ordered]@{
            alg = $SigningAlgorithm
            typ = "JWT"
            x5t = $x5t
        }

        # "iat" (issued at) and "exp" (expiration time) must be UTC and in UNIX time format
        $payload = @{
            iat = [Math]::Round((Get-Date).ToUniversalTime().Subtract((Get-Date -Date "01/01/1970")).TotalSeconds)
            exp = [Math]::Round((Get-Date).ToUniversalTime().Subtract((Get-Date -Date "01/01/1970")).TotalSeconds) + $TokenLifetimeInSeconds
        }

        # Issuer, Audience and Subject are optional as per RFC 7519
        if (-not([System.String]::IsNullOrEmpty($Issuer))) {
            Write-Log "Issuer: $Issuer will be added to payload" -Level DEBUG
            $payload.Add("iss", $Issuer)
        }

        if (-not([System.String]::IsNullOrEmpty($Audience))) {
            Write-Log "Audience: $Audience will be added to payload" -Level DEBUG
            $payload.Add("aud", $Audience)
        }

        if (-not([System.String]::IsNullOrEmpty($Subject))) {
            Write-Log "Subject: $Subject will be added to payload" -Level DEBUG
            $payload.Add("sub", $Subject)
        }

        $headerJson = $header | ConvertTo-Json -Compress
        $payloadJson = $payload | ConvertTo-Json -Compress

        $headerBase64 = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($headerJson)).Split("=")[0].Replace("+", "-").Replace("/", "_")
        $payloadBase64 = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($payloadJson)).Split("=")[0].Replace("+", "-").Replace("/", "_")

        $signatureInput = [System.Text.Encoding]::ASCII.GetBytes("$headerBase64.$payloadBase64")

        Write-Log "Header (Base64) is: $headerBase64" -Level DEBUG
        Write-Log "Payload (Base64) is: $payloadBase64" -Level DEBUG
        Write-Log "Signature input is: $signatureInput" -Level DEBUG

        $signingAlgorithmToUse = switch ($SigningAlgorithm) {
            ("RS384") { [Security.Cryptography.HashAlgorithmName]::SHA384 }
            ("RS512") { [Security.Cryptography.HashAlgorithmName]::SHA512 }
            default { [Security.Cryptography.HashAlgorithmName]::SHA256 }
        }
        Write-Log "Signing the Json Web Token using: $SigningAlgorithm" -Level DEBUG

        try {
            $signature = $privateKey.SignData($signatureInput, $signingAlgorithmToUse, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
        }
        finally {
            # Release the private key handle rather than waiting for the finalizer.
            if ($null -ne $privateKey) { $privateKey.Dispose() }
        }
        $signature = [Convert]::ToBase64String($signature).Split("=")[0].Replace("+", "-").Replace("/", "_")
    }
    end {
        if ((-not([System.String]::IsNullOrEmpty($headerBase64))) -and
            (-not([System.String]::IsNullOrEmpty($payloadBase64))) -and
            (-not([System.String]::IsNullOrEmpty($signature)))) {
            Write-Log "Returning Json Web Token" -Level DEBUG
            return ("$headerBase64.$payloadBase64.$signature")
        } else {
            Write-Log "Unable to create Json Web Token" -Level ERROR
            return
        }
    }
}

function Get-ApplicationAccessToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$TenantID,
        [Parameter(Mandatory = $true)][string]$ClientID,
        [Parameter(Mandatory = $true, ParameterSetName = 'Secret')][SecureString]$ClientSecret,
        [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')][string]$ClientAssertion,
        [Parameter(Mandatory = $true)][string]$Endpoint,
        [Parameter(Mandatory = $false)][string]$TokenService = "oauth2/v2.0/token",
        [Parameter(Mandatory = $false)][switch]$CertificateBasedAuthentication,
        [Parameter(Mandatory = $true)][string]$Scope
    )

    <#
        Shared function to create an OAuth token by using a JWT or secret.
        If you want to use a certificate, set the CertificateBasedAuthentication switch and pass a JWT token as the Secret parameter.
        You can use the Get-NewJsonWebToken function to create a JWT token.
        If you want to use a secret, pass the secret as the Secret parameter.
        This function returns a PSCustomObject with the OAuth token, status and the time the token was created.
        If the request fails, the PSCustomObject will contain the exception message.
    #>

    begin {
        Write-Log "Calling $($MyInvocation.MyCommand)" -Level DEBUG
        $oAuthTokenCallSuccess = $false
        $exceptionMessage = $null

        Write-Log "TenantID: $TenantID - ClientID: $ClientID - Endpoint: $Endpoint - TokenService: $TokenService - Scope: $Scope" -Level DEBUG
        $body = @{
            scope      = $Scope
            client_id  = $ClientID
            grant_type = "client_credentials"
        }

        if ($CertificateBasedAuthentication) {
            Write-Log "Function was called with CertificateBasedAuthentication switch" -Level DEBUG
            $body.Add("client_assertion_type", "urn:ietf:params:oauth:client-assertion-type:jwt-bearer")
            $body.Add("client_assertion", $ClientAssertion)
        } else {
            Write-Log "Authentication is based on a secret" -Level DEBUG
            $bstr = [IntPtr]::Zero
            $plainSecret = $null
            try {
                $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
                $plainSecret = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
                $body.client_secret = $plainSecret
            }
            finally{
                if ($bstr -ne [IntPtr]::Zero) {
                    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
                }
                # Drop the last managed reference to the plaintext secret.
                $plainSecret = $null
            }
        }

        $invokeRestMethodParams = @{
            ContentType = "application/x-www-form-urlencoded"
            Method      = "POST"
            Body        = $body # Create string by joining bodyList with '&'
            Uri         = "$Endpoint/$TenantID/$TokenService"
        }
    }
    process {
        try {
            Write-Log "Now calling the Invoke-RestMethod cmdlet to create an OAuth token" -Level DEBUG
            $oAuthToken = Invoke-RestMethod @invokeRestMethodParams
            Write-Log "Invoke-RestMethod call was successful" -Level DEBUG
            $oAuthTokenCallSuccess = $true
        } catch {
            Write-Log "We fail to create an OAuth token - Exception: $($_.Exception.Message)" -Level ERROR
            $exceptionMessage = $_.Exception.Message
        }
        
        finally{
            $body.Remove("client_secret")
            $body.Remove("client_assertion")
        }
    }
    end {
        return [PSCustomObject]@{
            OAuthToken           = $oAuthToken
            Successful           = $oAuthTokenCallSuccess
            ExceptionMessage     = $exceptionMessage
            LastTokenRefreshTime = (Get-Date)
        }
    }
}

function New-ClientAssertion {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][string]$Thumbprint,
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$AzureADEndpoint
    )

    <#
        Shared function that builds the client assertion used for certificate-based application
        (app-only) authentication against Entra ID.

        It is a thin wrapper over Get-NewJsonWebToken that fills in the claims Entra ID expects:
            Issuer / Subject: the application (client) ID
            Audience:         the tenant token endpoint, "$AzureADEndpoint/$TenantId/oauth2/v2.0/token"
    #>

    $jwt = Get-NewJsonWebToken -CertificateThumbprint $Thumbprint `
                               -CertificateStore      $CertificateStore `
                               -Issuer                $ClientId `
                               -Subject               $ClientId `
                               -Audience              "$AzureADEndpoint/$TenantId/oauth2/v2.0/token"

    if ([string]::IsNullOrEmpty($jwt)) {
        throw "Unable to generate a client assertion from certificate $Thumbprint."
    }
    return $jwt
}
function Update-AccessTokenIfNeeded {
    param(
        [switch]$Force
    )

    <#
        Keeps the OAuth access token in $Script:Token valid for the life of the script.

        Invoke-GraphApiRequest calls this before every Graph request, so a long-running search or
        delete does not fail part way through when the original token expires. The call is a no-op
        while the current token is still considered fresh - it is only renewed once
        $Script:tokenLastRefreshTime is more than 55 minutes old, which leaves headroom before the
        usual 60 minute lifetime. Pass -Force to renew immediately regardless of age; this is what
        the 401 retry path in Invoke-GraphApiRequest uses when Graph rejects a token early, for
        example after a revocation.

        How the token is renewed depends on $PermissionType:
            Application: requests a brand new token via the client credentials flow, using either a
                         certificate-backed client assertion (New-ClientAssertion) or the stored
                         client secret.
            Delegated:   redeems $Script:RefreshToken against the token endpoint and stores the
                         rotated refresh token that comes back with the response.

        On success $Script:Token and $Script:tokenLastRefreshTime are updated in place (plus
        $Script:RefreshToken in the delegated case) and nothing is returned. Any failure to obtain a
        token throws, deliberately stopping the caller rather than letting it continue unauthenticated.
    #>

    if($null -ne $Script:tokenLastRefreshTime) {
        $refreshAt = $Script:tokenLastRefreshTime.AddMinutes(55)
    }

    if (-not $Force -and $null -ne $Script:tokenLastRefreshTime -and (Get-Date) -lt $refreshAt) {
        return
    }

    Write-Log "Refreshing OAuth access token." -Level DEBUG

    if ($PermissionType -eq 'Application') {
        $tokenParams = @{
            TenantID = $Script:applicationInfo.TenantID
            ClientID = $Script:applicationInfo.ClientID
            Endpoint = $cloudService.AzureADEndpoint
            Scope    = $Script:GraphScope
        }

        if (-not [string]::IsNullOrEmpty($Script:applicationInfo.CertificateThumbprint)) {
            $tokenParams.CertificateBasedAuthentication = $true
            $tokenParams.ClientAssertion = New-ClientAssertion -Thumbprint $Script:applicationInfo.CertificateThumbprint -ClientId $Script:applicationInfo.ClientID -TenantId $Script:applicationInfo.TenantID -AzureADEndpoint $cloudService.AzureADEndpoint
        }
        else {
            $tokenParams.ClientSecret = $Script:applicationInfo.ClientSecret
        }

        $result = Get-ApplicationAccessToken @tokenParams

        if (-not $result.Successful) {
            throw "Unable to refresh access token: $($result.ExceptionMessage)"
        }

        $Script:Token = $result.OAuthToken.access_token
        $Script:tokenLastRefreshTime = $result.LastTokenRefreshTime
    }
    else {
        $params = @{
            Uri         = "$($cloudService.AzureADEndpoint)/organizations/oauth2/v2.0/token"
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Body        = @{
                client_id     = $Script:applicationInfo.ClientID
                scope         = $Script:GraphScope
                grant_type    = 'refresh_token'
                refresh_token = $Script:RefreshToken
            }
            UseBasicParsing = $true
        }

        $response = Invoke-WebRequestWithProxyDetection `
            -ParametersObject $params

        if ($response.StatusCode -ne 200) {
            throw "Unable to refresh delegated access token."
        }

        $tokens = $response.Content | ConvertFrom-Json
        $Script:Token = $tokens.access_token
        $Script:RefreshToken = $tokens.refresh_token
        $Script:tokenLastRefreshTime = Get-Date
    }

    Write-Log "OAuth access token refreshed." -Level INFO
}

function Get-DelegatedAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$AzureADEndpoint = "https://login.microsoftonline.com",
        [Parameter(Mandatory = $false)][string]$GraphApiUrl = "https://graph.microsoft.com",
        [Parameter(Mandatory = $false)][string]$Scope = "$($GraphApiUrl)//Mail.Read email openid profile offline_access",
        [Parameter(Mandatory = $false)][string]$ClientID,
        [Parameter(Mandatory = $false)][string]$RedirectUri
    )

    <#
        This function is used to get an access token for the Azure Graph API by using the OAuth 2.0 authorization code flow
        with PKCE (Proof Key for Code Exchange). The OAuth 2.0 authorization code grant type, or auth code flow,
        enables a client application to obtain authorized access to protected resources like web APIs.
        The auth code flow requires a user-agent that supports redirection from the authorization server
        (the Microsoft identity platform) back to your application.

        More information about the auth code flow with PKCE can be found here:
        https://learn.microsoft.com/azure/active-directory/develop/v2-oauth2-auth-code-flow#protocol-details
    #>

    begin {
        Write-Log "Calling $($MyInvocation.MyCommand)" -Level DEBUG
       
        $responseType = "code" # Provides the code as a query string parameter on our redirect URI
        $prompt = "select_account" # We want to show the select account dialog
        $codeChallengeMethod = "S256" # The code challenge method is S256 (SHA256)
        $codeChallengeVerifier = Get-NewS256CodeChallengeVerifier
        $state = ([guid]::NewGuid()).Guid
        $connectionSuccessful = $false
    }
    process {
        $codeChallenge = $codeChallengeVerifier.CodeChallenge
        $codeVerifier = $codeChallengeVerifier.Verifier

        # Request an authorization code from the Microsoft Azure Active Directory endpoint.
        # Values are percent-encoded so the URL survives Start-Process argument parsing (scope contains spaces).
        $authCodeRequestUrl = "$AzureADEndpoint/organizations/oauth2/v2.0/authorize?client_id=$([Uri]::EscapeDataString($ClientID))" +
        "&response_type=$responseType&redirect_uri=$([Uri]::EscapeDataString($RedirectUri))&scope=$([Uri]::EscapeDataString($Scope))&state=$state&prompt=$prompt" +
        "&code_challenge_method=$codeChallengeMethod&code_challenge=$codeChallenge"

        # Listen on the port declared by the redirect URI rather than assuming the default.
        $listenerPort = 8004
        $parsedRedirect = $null
        if ([Uri]::TryCreate($RedirectUri, [UriKind]::Absolute, [ref]$parsedRedirect) -and $parsedRedirect.Port -gt 0) {
            $listenerPort = $parsedRedirect.Port
        }

        # Start-Process will happily launch anything; make sure we only ever hand it a web URL.
        if ($authCodeRequestUrl -notmatch '^https://') {
            Write-Log "Refusing to launch a non-HTTPS authorization URL." -Level ERROR
            return
        }

        Start-Process -FilePath $authCodeRequestUrl
        $authCodeResponse = Start-LocalListener -Port $listenerPort

        if ($null -ne $authCodeResponse) {
            $returnedCode  = Get-UrlQueryParameter -Url $authCodeResponse -Name 'code'
            $returnedState = Get-UrlQueryParameter -Url $authCodeResponse -Name 'state'

            if ([string]::IsNullOrEmpty($returnedCode)) {
                Write-Log "The redirect did not contain an authorization code." -Level ERROR
                return
            }

            # The value taken from RawUrl is still percent-encoded; decode it so the form post does not double-encode it.
            $returnedCode = [Uri]::UnescapeDataString($returnedCode)

            # Verify the CSRF state so an injected authorization code cannot be redeemed.
            if (-not [string]::Equals([Uri]::UnescapeDataString([string]$returnedState), $state, [StringComparison]::Ordinal)) {
                Write-Log "OAuth state mismatch detected. The authorization response did not originate from this request and will be discarded." -Level ERROR
                return
            }

            # Redeem the returned code for an access token
            $redeemAuthCodeParams = @{
                Uri             = "$AzureADEndpoint/organizations/oauth2/v2.0/token"
                Method          = "POST"
                ContentType     = "application/x-www-form-urlencoded"
                Body            = @{
                    client_id     = $ClientID
                    scope         = $Scope
                    code          = $returnedCode
                    redirect_uri  = $RedirectUri
                    grant_type    = "authorization_code"
                    code_verifier = $codeVerifier
                }
                UseBasicParsing = $true
            }
            $redeemAuthCodeResponse = Invoke-WebRequestWithProxyDetection -ParametersObject $redeemAuthCodeParams

            if ($redeemAuthCodeResponse.StatusCode -eq 200) {
                $tokens = $redeemAuthCodeResponse.Content | ConvertFrom-Json
                $connectionSuccessful = $true
            } else {
                Write-Log "Unable to redeem the authorization code for an access token." -Level ERROR
            }
        } else {
            Write-Log "Unable to acquire an authorization code from the Microsoft Azure Active Directory endpoint." -Level ERROR
        }
    }
    end {
        if ($connectionSuccessful) {
            return [PSCustomObject]@{
                AccessToken = $tokens.access_token
                RefreshToken = $tokens.refresh_token
                LastTokenRefreshTime = (Get-Date)
                Successful           = $true
            }
        }
        exit
    }
}

function Get-NewS256CodeChallengeVerifier {
    param()

    <#
        This function can be used to generate a new SHA256 code challenge and verifier following the PKCE specification.
        The Proof Key for Code Exchange (PKCE) extension describes a technique for public clients to mitigate the threat
        of having the authorization code intercepted. The technique involves the client first creating a secret,
        and then using that secret again when exchanging the authorization code for an access token.

        The function returns a PSCustomObject with the following properties:
        Verifier: The verifier that was generated
        CodeChallenge: The code challenge that was generated

        It returns $null if the code challenge and verifier generation fails.

        More information about the auth code flow with PKCE can be found here:
        https://www.rfc-editor.org/rfc/rfc7636
    #>

    Write-Log "Calling $($MyInvocation.MyCommand)" -Level DEBUG

    $bytes = [System.Byte[]]::new(64)
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $rng.GetBytes($bytes)
    }
    finally {
        $rng.Dispose()
    }

    $verifier = ([Convert]::ToBase64String($bytes).TrimEnd("=")).Replace("+", "-").Replace("/", "_")

    # Hash the verifier directly. Get-FileHash required a MemoryStream/StreamWriter (never disposed) and
    # returned a hex string that then had to be parsed back into bytes.
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $challengeBytes = $sha256.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($verifier))
    }
    finally {
        $sha256.Dispose()
    }

    $base64UrlEncoded = ([Convert]::ToBase64String($challengeBytes).TrimEnd("=")).Replace("+", "-").Replace("/", "_")

    if ((-not([System.String]::IsNullOrEmpty($verifier))) -and
        (-not([System.String]::IsNullOrEmpty(($base64UrlEncoded))))) {
        Write-Log "Verifier and CodeChallenge generated successfully" -Level DEBUG
        return [PSCustomObject]@{
            Verifier      = $verifier
            CodeChallenge = $base64UrlEncoded
        }
    }

    Write-Log "Verifier and CodeChallenge generation failed" -Level ERROR
    return $null
}

function Start-LocalListener {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Only non-destructive operations are performed in this function.')]
    param(
        [Parameter(Mandatory = $false)][int]$Port = 8004,
        [Parameter(Mandatory = $false)][int]$TimeoutSeconds = 60,
        [Parameter(Mandatory = $false)][string]$UrlContains = "code=",
        [Parameter(Mandatory = $false)][string]$ExpectedHttpMethod = "GET",
        [Parameter(Mandatory = $false)][string]$ResponseOutput = "Authentication complete. You can return to the application. Feel free to close this browser tab."
    )

    <#
        This function is used to start a local listener on the specified port (default: 8004).
        It will wait for the specified amount of seconds (default: 60) for a request to be made.
        The function will return the URL of the request that was made.
    #>

    begin {
        Write-Log "Calling $($MyInvocation.MyCommand)" -Level DEBUG
        $url = $null
        $signalled = $false
        $stopwatch = New-Object System.Diagnostics.Stopwatch
        $listener = New-Object Net.HttpListener
    }
    process {
        $listener.Prefixes.add("http://localhost:$($Port)/")
        try {
            Write-Log "Starting listener..." -Level DEBUG
            Write-Log "Listening on port: $($Port)" -Level DEBUG
            Write-Log "Waiting $($TimeoutSeconds) seconds for request to be made to url that contains: $($UrlContains)" -Level DEBUG
            $stopwatch.Start()
            $listener.Start()

            while ($listener.IsListening) {
                $task = $listener.GetContextAsync()

                # WaitOne already blocks for the poll interval; an extra Start-Sleep only doubled the response latency.
                while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
                    if ($task.AsyncWaitHandle.WaitOne(100)) {
                        $signalled = $true
                        break
                    }
                }

                if ($signalled) {
                    $context = $task.GetAwaiter().GetResult()
                    $request = $context.Request
                    $response = $context.Response
                    $url = $request.RawUrl
                    $content = [byte[]]@()

                    if (($url.Contains($UrlContains)) -and
                        ($request.HttpMethod -eq $ExpectedHttpMethod)) {
                        Write-Log "Request made to listener and url that was called is as expected. HTTP Method: $($request.HttpMethod)" -Level DEBUG
                        $content = [System.Text.Encoding]::UTF8.GetBytes($ResponseOutput)
                        $response.StatusCode = 200 # OK
                        $response.OutputStream.Write($content, 0, $content.Length)
                        $response.Close()
                        break
                    } else {
                        #Write-Log "Request made to listener but the url that was called is not as expected. URL: $($url)" -Level DEBUG
                        Write-Log "Request made to listener but the url that was called is not as expected. HTTP method: $($request.HttpMethod) | URL: $(Get-RedactedUrl $url)" -Level DEBUG
                        $response.StatusCode = 404 # Not Found
                        $response.OutputStream.Write($content, 0, $content.Length)
                        $response.Close()
                        break
                    }
                } else {
                    Write-Log "Timeout of $($TimeoutSeconds) seconds reached..." -Level DEBUG
                    break
                }
            }
        } finally {
            Write-Log "Stopping listener..." -Level DEBUG
            Start-Sleep -Seconds 2
            $stopwatch.Stop()
            $listener.Stop()
        }
    }
    end {
        return $url
    }
}

function Get-UrlQueryParameter {
    <#
        Returns the raw (still percent-encoded) value of a named query string parameter.
        Matching by name avoids relying on parameter ordering in the redirect response.
        Returns $null when the parameter is not present.
    #>
    param(
        [AllowNull()][string]$Url,
        [Parameter(Mandatory)][string]$Name
    )

    if ([string]::IsNullOrEmpty($Url)) { return $null }

    $queryIndex = $Url.IndexOf('?')
    if ($queryIndex -lt 0) { return $null }

    foreach ($pair in $Url.Substring($queryIndex + 1).Split('&')) {
        if ([string]::IsNullOrWhiteSpace($pair)) { continue }

        $separator = $pair.IndexOf('=')
        if ($separator -lt 0) { continue }

        if ([string]::Equals($pair.Substring(0, $separator), $Name, [StringComparison]::OrdinalIgnoreCase)) {
            return $pair.Substring($separator + 1)
        }
    }

    return $null
}

function Get-RedactedUrl {
    <#
        Returns a URL safe for logging: known credential-bearing query parameters have their
        values replaced with a length placeholder. Everything else is preserved for diagnostics.
    #>
    param([AllowNull()][string]$Url)

    if ([string]::IsNullOrEmpty($Url)) { return '<empty>' }

    $sensitive = @(
        'code', 'id_token', 'access_token', 'refresh_token',
        'client_secret', 'client_assertion', 'assertion', 'session_state'
    )

    $queryIndex = $Url.IndexOf('?')
    if ($queryIndex -lt 0) { return $Url }

    $path  = $Url.Substring(0, $queryIndex)
    $query = $Url.Substring($queryIndex + 1)

    $redacted = foreach ($pair in $query.Split('&')) {
        if ([string]::IsNullOrWhiteSpace($pair)) { continue }

        $separator = $pair.IndexOf('=')
        if ($separator -lt 0) { $pair; continue }

        $name  = $pair.Substring(0, $separator)
        $value = $pair.Substring($separator + 1)

        if ($sensitive -contains $name.ToLowerInvariant()) {
            "$name=<redacted:$($value.Length) chars>"
        }
        else {
            "$name=$value"
        }
    }

    return "$path`?$($redacted -join '&')"
}

function Invoke-GraphApiRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Query,

        [ValidateSet("v1.0", "beta")]
        [Parameter(Mandatory = $false)]
        [string]$Endpoint = "v1.0",

        [Parameter(Mandatory = $false)]
        [string]$Method = "GET",

        [Parameter(Mandatory = $false)]
        [string]$ContentType = "application/json",

        [Parameter(Mandatory = $false)]
        [string]$Body,

        [Parameter(Mandatory = $false)]
        [int]$ExpectedStatusCode = 200,

        [Parameter(Mandatory = $false)]
        [string]$GraphApiUrl,

        [Parameter(Mandatory = $false)]
        [hashtable]$Headers = @{}
    )

    <#
        This shared function is used to make requests to the Microsoft Graph API.
        It returns a PSCustomObject with the following properties:
            Content: The content of the response (converted from JSON to a PSCustomObject)
            Response: The full response object
            StatusCode: The status code of the response
            Successful: A boolean indicating whether the request was successful
    #>

    begin {
        Write-Log "Calling $($MyInvocation.MyCommand)" -Level DEBUG
        $successful = $false
        $content = $null
    }
    process {
        #Check for expired token before making Graph request
        Update-AccessTokenIfNeeded

        $requestUri = if ([Uri]::IsWellFormedUriString($Query, [UriKind]::Absolute)) {
           $Query
        }
        else {
            "$GraphApiUrl/$Endpoint/$($Query.TrimStart('/'))"
        }

        $requestHeaders = @{
            Authorization = "Bearer $Script:Token"
        }
        foreach ($name in $Headers.Keys) {
           $requestHeaders[$name] = $Headers[$name]
        }

        $graphApiRequestParams = @{
            Uri             = $requestUri
            Headers         = $requestHeaders
            Method          = $Method
            ContentType     = $ContentType
            UseBasicParsing = $true
            ErrorAction     = "Stop"
        }

        if (-not([System.String]::IsNullOrEmpty($Body))) {
            Write-Log "Body: $Body" -Level DEBUG
            $graphApiRequestParams.Add("Body", $Body)
        }
        if($PSVersionTable.PSVersion.Major -ge 7){
            #Need to prevent redirection as it will fail without sending Auth header again. This is a known issue with PS7 and Invoke-WebRequest.
            Write-Log "PSVersion is 7 or higher, setting MaximumRedirection to 0" -Level DEBUG
            $graphApiRequestParams.Add("MaximumRedirection", 0)
        }

        Write-Log "Graph API uri called: $($graphApiRequestParams.Uri)" -Level DEBUG
        $Script:graphApiResponse = Invoke-WebRequestWithProxyDetection -ParametersObject $graphApiRequestParams
        if($Script:graphApiResponse.StatusCode -eq 401){
            Write-Log "Graph returned 401; forcing token refresh" -Level WARN
            Update-AccessTokenIfNeeded -Force
            $graphApiRequestParams.Headers.Authorization = "Bearer $Script:Token"
            $Script:graphApiResponse = Invoke-WebRequestWithProxyDetection -ParametersObject $graphApiRequestParams
        }
        if (($null -eq $graphApiResponse) -or
            ([System.String]::IsNullOrEmpty($graphApiResponse.StatusCode))) {
            Write-Log "Graph API request failed - no response" -Level DEBUG
        } elseif ($graphApiResponse.StatusCode -ne $ExpectedStatusCode) {
            Write-Log "Graph API status code: $($graphApiResponse.StatusCode) does not match expected status code: $ExpectedStatusCode" -Level DEBUG
        } else {
            Write-Log "Graph API request successful" -Level DEBUG
            $successful = $true
            $content = $graphApiResponse.Content | ConvertFrom-Json
        }
    }
    end {
        return [PSCustomObject]@{
            Content    = $content
            Response   = $graphApiResponse
            StatusCode = $graphApiResponse.StatusCode
            Successful = $successful
            Headers   = $graphApiResponse.Headers
            ErrorMessage = $graphApiResponse.ErrorMessage
        }
    }
}

function Invoke-WebRequestWithProxyDetection {
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Default")][string]$Uri,
        [Parameter(Mandatory = $false, ParameterSetName = "Default")][switch]$UseBasicParsing,
        [Parameter(Mandatory = $true, ParameterSetName = "ParametersObject")][hashtable]$ParametersObject,
        [Parameter(Mandatory = $false, ParameterSetName = "Default")][string]$OutFile
    )

    <#
        Single wrapper around Invoke-WebRequest used for every outbound HTTP call in the script.
        It exists so proxy handling, throttling retries and error shaping live in one place instead
        of being repeated at each call site.

        It can be called two ways:
            Default:          pass -Uri (optionally -OutFile / -UseBasicParsing) for simple requests.
            ParametersObject: pass a hashtable that is splatted straight onto Invoke-WebRequest.
                              This is what Invoke-GraphApiRequest uses so it can control headers,
                              method, body and MaximumRedirection.

        Behavior it adds on top of Invoke-WebRequest:
            Proxy       - if Confirm-ProxyServer finds an explicit proxy in front of the target,
                          ProxyUseDefaultCredentials is turned on so the request can authenticate to
                          it. A value already supplied by the caller is left alone.
            Throttling  - HTTP 429 is retried up to 4 times, waiting for the interval Graph asks for
                          in Retry-After (read from typed headers on PS7 and string headers on PS5.1)
                          or falling back to exponential backoff when the header is absent.
            Errors      - all other failures stop immediately and are converted from a terminating
                          exception into a returned object, so callers can branch on a status code
                          rather than wrapping every call in try/catch:
                            308  ErrorCode "PermanentRedirect" with the Location header in
                                 ErrorMessage. This is how an auxiliary archive mailbox is
                                 discovered, and is why redirects must not be followed.
                            5xx  ErrorCode from the status description.
                            all others: the Graph JSON error body is parsed for code and message,
                                 falling back to the raw body or exception text, truncated for logging.

        Note the return type differs by outcome: a successful request returns the native response
        object from Invoke-WebRequest, while a failure returns a PSCustomObject with ErrorCode,
        ErrorMessage, StatusCode and Successful. Callers should test StatusCode rather than assume a
        web response was returned.
    #>

    Write-Log "Calling $($MyInvocation.MyCommand)" -Level DEBUG
    if ([System.String]::IsNullOrEmpty($Uri)) {
        $Uri = $ParametersObject.Uri
    }

    if ($null -eq $ParametersObject) {
        $params = @{
            Uri     = $Uri
            OutFile = $OutFile
        }

        if ($UseBasicParsing) {
            $params.UseBasicParsing = $true
        }
    } else {
        $params = $ParametersObject
    }

    # Must run after $params exists, otherwise the proxy credentials were silently dropped.
    if (-not $params.ContainsKey('ProxyUseDefaultCredentials') -and (Confirm-ProxyServer -TargetUri $Uri)) {
        $params.ProxyUseDefaultCredentials = $true
    }
    #Allow for maximum retries of 4 for throttling. This is a known issue with Graph API where it will return 429 for a period of time and then allow the request to go through.
    $maxAttempts = 4
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            return Invoke-WebRequest @params
        } 
        catch {
            $response = $_
            $httpResponse = $response.Exception.Response
            $responseBody = $_.ErrorDetails.Message
            $statusCode = if ($null -ne $httpResponse) {
                [int]$httpResponse.StatusCode
            }
            else {
                $null
            }
            #Exit the retry loop when the error is not a 429 (Too Many Requests)
            if($statusCode -ne 429){
                if($statusCode -eq 308){
                    Write-Log $response.Exception.Message -Level DEBUG
                    $responseType = $httpResponse.GetType().FullName
                    #Get the aux archive location from the response headers. The location header is different for PS7 and PS5 so we need to check the type of the response object.
                    $location = if ($responseType -eq 'System.Net.Http.HttpResponseMessage') {
                        [string]$httpResponse.Headers.Location
                    }
                    else {
                        [string]$httpResponse.Headers['Location']
                    }
                    Write-Log $location -Level DEBUG
                    return [PSCustomObject]@{
                            ErrorCode    = "PermanentRedirect"
                            ErrorMessage  = $location
                            StatusCode = 308
                            Successful = $false
                    }
                }
                if($statusCode -ge 500 -and $statusCode -lt 600){
                    Write-Log "Graph API request failed with status code: $statusCode - $($httpResponse.StatusDescription)" -Level ERROR
                    return [PSCustomObject]@{
                        ErrorCode    = $httpResponse.StatusDescription
                        ErrorMessage = $response.Exception.Message
                        StatusCode = [int]$httpResponse.StatusCode
                        Successful = $false
                    }
                }
                else{
                    #Error encountered, read the response body to get the error message and return it to the caller.
                    if ([string]::IsNullOrWhiteSpace($responseBody) -and $null -ne $httpResponse -and $httpResponse.PSObject.Methods.Name -contains 'GetResponseStream') {
                        $reader = [System.IO.StreamReader]::New($httpResponse.GetResponseStream())
                        try {
                            $responseBody = $reader.ReadToEnd()
                        }
                        finally {
                            $reader.Dispose()
                        }
                    }
                    $errorCode = $null
                    $errorMessage = $null

                    if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
                        try {
                            $responseContent = $responseBody | ConvertFrom-Json -ErrorAction Stop
                            $errorCode = $responseContent.error.code
                            $errorMessage = $responseContent.error.message
                        }
                        catch {
                            Write-Log "Graph API returned a non-JSON error body: $($_.Exception.Message)" -Level DEBUG
                        }
                    }
                    if ([string]::IsNullOrWhiteSpace($errorCode)) { $errorCode = "NonJsonErrorResponse" }
                    if ([string]::IsNullOrWhiteSpace($errorMessage)) {
                        $errorMessage = if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
                            $responseBody.Trim()
                        }
                        else {
                            $response.Exception.Message
                        }
                    }
                    # HTML error pages can be very large; keep the log readable.
                    if ($errorMessage.Length -gt 1000) {
                        $errorMessage = $errorMessage.Substring(0, 1000) + "...[truncated]"
                    }

                    Write-Log "Graph API request failed with status code: $statusCode" -Level DEBUG
                    Write-Log "Error message: $errorMessage" -Level DEBUG

                    return [PSCustomObject]@{
                        ErrorCode    = $errorCode
                        ErrorMessage = $errorMessage
                        StatusCode   = if ($null -ne $statusCode) { $statusCode } else { 0 }
                        Successful   = $false
                    }
                }
            }
            #If we reach the maximum number of attempts, return an error indicating that the Graph API is throttled.
            if($attempt -eq $maxAttempts){
                Write-Log "Graph remained throttled after $maxAttempts attempts." -Level ERROR
                return [PSCustomObject]@{
                    ErrorCode    = "TooManyRequests"
                    ErrorMessage  = $response.Exception.Message
                    StatusCode = 429
                    Successful = $false
                }
            }
            $retryAfterSeconds = 0
            if ($null -ne $httpResponse) {
                # PS7 surfaces HttpResponseMessage (typed headers); PS5.1 surfaces a WebResponse (string indexer).
                # The string indexer silently returns $null on PS7, which previously forced every retry onto backoff.
                if ($httpResponse.GetType().FullName -eq 'System.Net.Http.HttpResponseMessage') {
                    $retryAfterHeader = $httpResponse.Headers.RetryAfter
                    if ($null -ne $retryAfterHeader -and $null -ne $retryAfterHeader.Delta) {
                        $retryAfterSeconds = [int]$retryAfterHeader.Delta.TotalSeconds
                    }
                }
                else {
                    [void][int]::TryParse([string]$httpResponse.Headers['Retry-After'], [ref]$retryAfterSeconds)
                }
            }

            if ($retryAfterSeconds -le 0) {
                $retryAfterSeconds = [int][Math]::Pow(2, $attempt)
            }

            Write-Log "Graph returned 429. Attempt $attempt of $maxAttempts; retrying after $retryAfterSeconds seconds." -Level WARN
            Start-Sleep -Seconds ($retryAfterSeconds + 1)
        }
    }
}

function Confirm-ProxyServer {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)][string]$TargetUri
    )

    # The system proxy lookup is comparatively expensive and the result does not change
    # mid-run, so cache it per host instead of resolving it on every Graph request.
    if ($null -eq $Script:ProxyDetectionCache) {
        $Script:ProxyDetectionCache = @{}
    }

    $cacheKey = $TargetUri
    $parsedTarget = $null
    if ([Uri]::TryCreate($TargetUri, [UriKind]::Absolute, [ref]$parsedTarget)) {
        $cacheKey = $parsedTarget.GetLeftPart([System.UriPartial]::Authority)
    }

    if ($Script:ProxyDetectionCache.ContainsKey($cacheKey)) {
        return $Script:ProxyDetectionCache[$cacheKey]
    }

    Write-Log "Calling $($MyInvocation.MyCommand)" -Level DEBUG
    try {
        $proxyObject = ([System.Net.WebRequest]::GetSystemWebProxy()).GetProxy($TargetUri)
        if ($TargetUri -ne $proxyObject.OriginalString) {
            Write-Log "Proxy server configuration detected" -Level DEBUG
            Write-Log $proxyObject.OriginalString -Level DEBUG
            $Script:ProxyDetectionCache[$cacheKey] = $true
            return $true
        } else {
            Write-Log "No proxy server configuration detected" -Level DEBUG
            $Script:ProxyDetectionCache[$cacheKey] = $false
            return $false
        }
    } catch {
        Write-Log "Unable to check for proxy server configuration" -Level DEBUG
        $Script:ProxyDetectionCache[$cacheKey] = $false
        return $false
    }
}

function WriteErrorInformationBase {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0],
        [ValidateSet("Write-Host", "Write-Verbose")][string]$Cmdlet
    )

    if ($null -ne $CurrentError.OriginInfo) {
        & $Cmdlet "Error Origin Info: $($CurrentError.OriginInfo.ToString())"
    }

    & $Cmdlet "$($CurrentError.CategoryInfo.Activity) : $($CurrentError.ToString())"

    if ($null -ne $CurrentError.Exception -and
        $null -ne $CurrentError.Exception.StackTrace) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception.StackTrace)"
    } elseif ($null -ne $CurrentError.Exception) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception)"
    }

    if ($null -ne $CurrentError.InvocationInfo.PositionMessage) {
        & $Cmdlet "Position Message: $($CurrentError.InvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
        & $Cmdlet "Remote Position Message: $($CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.ScriptStackTrace) {
        & $Cmdlet "Script Stack: $($CurrentError.ScriptStackTrace)"
    }
}

function Write-HostErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Host"
}

function Get-OAuthToken {
    param(
        [array]$AppScope
    )

    <#
        Performs the initial sign-in for the script and seeds the token state that everything else
        depends on. It runs once during start-up; Update-AccessTokenIfNeeded takes over from there
        and handles renewal for the rest of the run.

        The flow depends on $PermissionType:
            Application: appends ".default" to the scope, as the client credentials flow requires,
                         and authenticates with either the client secret or a certificate-backed
                         client assertion. Exactly one must be supplied. The credential is recorded
                         in $Script:applicationInfo so the token can later be renewed unattended.
            Delegated:   ensures email, openid and offline_access are present in $AppScope -
                         offline_access is what causes Entra ID to issue the refresh token needed to
                         stay signed in - then runs the interactive authorization code flow with
                         PKCE via Get-DelegatedAccessToken.

        On success it populates $Script:Token and $Script:tokenLastRefreshTime, plus
        $Script:RefreshToken in the delegated case, and returns nothing.

        On failure it logs the underlying error and terminates the script with exit rather than
        throwing. Nothing downstream can run without a token, so there is no useful recovery.
    #>

    if($PermissionType -eq "Application") {
        $Script:GraphScope = "$($Script:GraphScope).default"
        $createOAuthTokenParams = @{
            TenantID = $OAuthTenantId
            ClientID = $OAuthClientId
            Scope    = $Script:GraphScope
            Endpoint = $azureADEndpoint
        }

        if ([System.String]::IsNullOrEmpty($OAuthCertificate)) {
            if ($null -eq $OAuthClientSecret) {
                Write-Log "Application permissions require either -OAuthClientSecret or -OAuthCertificate." -Level ERROR
                exit 1
            }
            # Keep it as a SecureString end to end.
            $Script:applicationInfo.ClientSecret  = $OAuthClientSecret
            $createOAuthTokenParams.ClientSecret  = $OAuthClientSecret
        }
        else {
            $createOAuthTokenParams.CertificateBasedAuthentication = $true
            $Script:applicationInfo.CertificateThumbprint = $OAuthCertificate
            $createOAuthTokenParams.ClientAssertion = New-ClientAssertion -Thumbprint $OAuthCertificate -ClientId $OAuthClientId -TenantId $OAuthTenantId -AzureADEndpoint $azureADEndpoint
        }

        #Create OAUTH token
        $oAuthReturnObject = Get-ApplicationAccessToken @createOAuthTokenParams
        if ($oAuthReturnObject.Successful -eq $false) {
            Write-Host ""
            Write-Log "Unable to fetch an OAuth token. Please review the error message below and re-run the script:" -Level ERROR
            Write-Log $oAuthReturnObject.ExceptionMessage -Level ERROR
            exit 1
        }
        $Script:Token = $oAuthReturnObject.OAuthToken.access_token
        $Script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime
    }
    elseif ($PermissionType -eq "Delegated") {
        if(-not(($AppScope.Contains("email")))) {
            $AppScope += "email"
        }
        if(-not(($AppScope.Contains("openid")))) {
            $AppScope += "openid"
        }
        if(-not(($AppScope.Contains("offline_access")))) {
            $AppScope += "offline_access"
        }
        $Script:GraphScope = "$($Script:GraphScope)$($AppScope)"
        $oAuthReturnObject = Get-DelegatedAccessToken -AzureADEndpoint $cloudService.AzureADEndpoint -GraphApiUrl $cloudService.GraphApiEndpoint -Scope $Script:GraphScope -ClientID $OAuthClientId -RedirectUri $OAuthRedirectUri
        if ($oAuthReturnObject.Successful -eq $false) {
            #Write-Host ""
            Write-Log "Unable to fetch an OAuth token for accessing EWS. Please review the error message below and re-run the script:" -Level ERROR
            Write-Log $oAuthReturnObject.ExceptionMessage -Level ERROR
            exit
        }    
        $Script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime
        $Script:Token = $oAuthReturnObject.AccessToken
        $Script:RefreshToken = $oAuthReturnObject.RefreshToken
    }    
}

function Clear-SensitiveState {
    [CmdletBinding()]
    param()

    <#
        Tears down the credential material the script has been holding in script scope.

        It is called from the finally block at the end of the script, so it runs whether the script
        completed normally, failed, or was interrupted. That matters most in an interactive session,
        where the PowerShell host stays alive afterwards and any variables left behind remain
        readable by anything else running in that session.

        Cleared here:
            $Script:Token and $Script:RefreshToken - the bearer and refresh tokens.
            $Script:applicationInfo - the ClientSecret and CertificateThumbprint entries are removed
            so a later call cannot silently mint a new token from leftover credentials.

        The garbage collection at the end is a best-effort prompt to drop unreferenced copies
        sooner. It is not a guarantee: .NET strings are immutable, so any plaintext that was already
        materialised may persist in memory until it happens to be overwritten. This is why secrets
        are handled as SecureString for as long as possible elsewhere in the script.
    #>

    $Script:Token        = $null
    $Script:RefreshToken = $null

    if ($null -ne $Script:applicationInfo) {
        $Script:applicationInfo.Remove('ClientSecret')
        $Script:applicationInfo.Remove('CertificateThumbprint')
    }

    [System.GC]::Collect()
}
function ConvertTo-SearchResult{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Result,

        [Parameter(Mandatory)]
        [string]$MailboxName,

        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    <#
        Turns a single raw message returned by Graph into the flat record the rest of the script
        works with, and applies the last of the search criteria while doing so.

        The filtering matters as much as the conversion. AttachmentName and MessageBody cannot be
        expressed in the Graph $filter used by CreateSearchQuery, so they are evaluated here against
        the message that came back:
            AttachmentName - case-insensitive exact match on an attachment name. Relies on the
                             attachments having been expanded on the request.
            MessageBody    - case-insensitive substring match against the message body content.
        Both are read from script scope, and either one failing to match returns $null so the
        caller can discard the message. A message is only a real hit if this function returns an
        object.

        MailboxName and FolderPath are supplied by the caller because a Graph message does not carry
        the mailbox it came from or the display path of the folder it was found in.

        The returned object is deliberately flat and ordered: it is written straight to the search
        results CSV and its id property is the message id later submitted in the $batch delete
        request, so field names are part of the contract with ConvertTo-DeleteResult and the CSV
        output.
    #>

    #Check if attachment name is specified
    $matchedAttachment = $null

    if (-not [string]::IsNullOrWhiteSpace($AttachmentName)) {
        foreach ($attachment in @($Result.attachments)) {
            if ([string]::Equals(
                [string]$attachment.name,
                $AttachmentName,
                [StringComparison]::OrdinalIgnoreCase
            )) {
                $matchedAttachment = [string]$attachment.name
                break
            }
        }

        if ($null -eq $matchedAttachment) {
            return $null
        }
    }

    #Check if message body is specified
    if (-not [string]::IsNullOrWhiteSpace($MessageBody)) {
        $bodyContent = [string]$Result.body.content

        if ([string]::IsNullOrEmpty($bodyContent) -or $bodyContent.IndexOf($MessageBody, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
            return $null
        }
    }
    $senderAddress = $null
    if ($null -ne $Result.from -and $null -ne $Result.from.emailAddress) {
        $senderAddress = [string]$Result.from.emailAddress.address
    }

    [PSCustomObject][ordered]@{
        mailbox          = $MailboxName
        id               = [string]$Result.id
        folder           = $FolderPath
        internetMessageId = [string]$Result.internetMessageId
        subject          = [string]$Result.subject
        receivedDateTime = [string]$Result.receivedDateTime
        from             = $senderAddress
        attachment       = $matchedAttachment
    }
}

function ConvertTo-DeleteResult{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Response,
        [Parameter(Mandatory)]
        [object]$currentBatch,[Parameter(Mandatory)]
        [bool]$IsFinalAttempt
    )

    <#
        Interprets one sub-response from a Graph $batch delete and produces the row that is written
        to the delete results CSV. It is called once per item in the batch response.

        A $batch response only reports an id and a status; it does not echo back any detail about
        the message that was deleted. The id is the index the request was submitted under, so it is
        used to look the original item up in $currentBatch - the lookup built from the
        ConvertTo-SearchResult records for this batch - and the two are merged into a single record
        carrying both what was targeted and what happened to it.

        The status drives the outcome:
            204     the delete succeeded.
            429     throttled. If attempts remain the item is marked Retry and the caller resubmits
                    it; $IsFinalAttempt is what turns a retryable throttle into a recorded failure.
            other   treated as a failure and logged.

        This is not a pure conversion. It also maintains the running totals the end of run summary
        reports - $Script:itemsDeleted, $Script:TotalItemsDeleted, $Script:itemsFailedDelete and
        $Script:TotalDeleteFailures - and records the folder name in $Script:DeleteFailureFolders
        when an item could not be deleted.

        Note that the folder name used for failure reporting comes from $MailboxFolder, which is
        read from the calling scope rather than passed in as a parameter. The function therefore
        only reports folders correctly when called from within the folder processing loop.
    #>

    #Create object with item information
    $item = [PSCustomObject]@{
        Mailbox=$currentBatch[[int]$Response.id].mailbox
        Id=$currentBatch[[int]$Response.id].id
        Folder=$currentBatch[[int]$Response.id].folder
        Subject=$currentBatch[[int]$Response.id].subject
        InternetMessageId=$currentBatch[[int]$Response.id].internetMessageId
        ReceivedDateTime=$currentBatch[[int]$Response.id].receivedDateTime
        From=$currentBatch[[int]$Response.id].from
        Attachment=$currentBatch[[int]$Response.id].attachment
        StatusCode=[int]$Response.Status
    }
    #Add the delete status to the object based on the response from Graph API
    switch($Response.Status){
        204 {
            $Script:itemsDeleted++
            $Script:TotalItemsDeleted++
            $item | Add-Member -MemberType NoteProperty -Name "DeleteStatus" -Value "Succeeded"
        }
        429 {
            if ($IsFinalAttempt) {
                Write-Log "Deletion remained throttled after the final attempt." -Level WARN
                $Script:itemsFailedDelete++
                $Script:TotalDeleteFailures++
                [void]$Script:DeleteFailureFolders.Add([string]$MailboxFolder.displayName)
                $item | Add-Member DeleteStatus "Failed"
            }
            else {
                Write-Log "Too many requests. Retrying deletion later." -Level WARN
                $item | Add-Member -MemberType NoteProperty -Name "DeleteStatus" -Value "Retry"
            }
        }
        default {
            Write-Log "Failed to delete item. Status code: $($Response.Status)" -Level WARN
            $Script:itemsFailedDelete++
            $Script:TotalDeleteFailures++
            [void]$Script:DeleteFailureFolders.Add([string]$MailboxFolder.displayName)
            $item | Add-Member -MemberType NoteProperty -Name "DeleteStatus" -Value "Failed"
        }
    }
    

    return $item
}

function Protect-CsvValue {
    param([AllowNull()][object]$Value)

    <#
        Sanitises a single field before it is written to one of the CSV reports.

        Values such as subject, sender and attachment name come from received mail, so they are
        attacker controlled and must not be trusted just because they are only being written to a
        file. Two problems are handled:

        Formula injection - Excel and other spreadsheet applications treat a cell beginning with
            =, +, - or @ as a formula, so simply opening the report could evaluate content that
            arrived in an email subject. Any such value is prefixed with a single quote, which
            forces the cell to be read as text. Leading whitespace is included in the check because
            spreadsheet parsers skip it when deciding whether a cell is a formula.

        Record splitting - embedded carriage returns and line feeds are replaced with spaces so one
            message always occupies one physical line. This keeps the report readable and stops a
            crafted subject from appearing to add extra rows.

        $null is passed straight through so blank fields stay blank rather than becoming an empty
        string. Callers do not use this directly; ConvertTo-SafeCsvRecord applies it to every string
        property of a record.

        Because the value can be altered, the CSV is intended for reporting and review. It is not a
        byte for byte copy of the original message properties.
    #>

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value

    # Keep each record on one physical CSV line.
    $text = $text -replace "[`r`n]", " "

    # Force potentially executable spreadsheet values to text.
    if ($text -match '^[\t ]*[=+\-@]') {
        return "'$text"
    }

    return $text
}

function ConvertTo-SafeCsvRecord {
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object]$InputObject
    )

    <#
        Pipeline wrapper that makes a whole record safe to export, so call sites can pipe straight
        into Export-Csv without having to remember which individual fields need sanitising:

            $pageResults | ConvertTo-SafeCsvRecord | Export-Csv ...

        Each object is rebuilt property by property into a new PSCustomObject. String properties are
        passed through Protect-CsvValue; everything else - status codes, numbers, booleans - is
        copied unchanged, since only text can carry a leading formula character or an embedded line
        break. The original object is never modified.

        The rebuild uses an ordered dictionary so the property order of the source object is
        preserved. That matters because Export-Csv takes its column order and header from the first
        object it receives, and these reports are written incrementally with -Append across several
        pages of results.

        Applying this to every record rather than to selected fields means new properties added to
        the search or delete result objects are protected automatically.
    #>

    process {
        $record = [ordered]@{}

        foreach ($property in $InputObject.PSObject.Properties) {
            if ($property.Value -is [string]) {
                $record[$property.Name] = Protect-CsvValue $property.Value
            }
            else {
                $record[$property.Name] = $property.Value
            }
        }

        [PSCustomObject]$record
    }
}
function SearchMailbox {
    param(
        [string]$uriQuery
    )

    <#
        Main worker of the script. Walks every folder collected in $Script:searchFolders, searches
        it for matching messages, writes the hits to the search results CSV, and - when
        -DeleteContent was supplied - deletes them before moving on to the next folder.

        $uriQuery is the base mailFolders path for the target mailbox. The folder id and /messages
        are appended to it for each folder in turn.

        Per folder the sequence is:

        1. Archive location. When -Archive is used the folder is first probed on the beta endpoint
           via admin/exchange/mailboxes/{mailbox}/folders/{id}. A 308 response means the folder
           actually lives in an auxiliary archive mailbox, and the Location header carries both the
           auxiliary mailbox GUID and the id of the folder inside it. Those values are extracted,
           validated, and used to retarget the request. This is the reason redirects must not be
           followed automatically - the Location header is the payload, not a detour.

        2. Query. The $filter is produced once by CreateSearchQuery and reused for every folder,
           since the criteria do not change. Requests ask for immutable ids
           (Prefer: IdType="ImmutableId") so an id obtained during the search is still valid when
           the delete is issued.

        3. Paging. Results are followed through @odata.nextLink to the end. Each message is passed
           through ConvertTo-SearchResult, which also applies the AttachmentName and MessageBody
           criteria that cannot be expressed in the Graph filter. Each page is appended to the CSV
           as it is retrieved rather than being held until the end.

        4. Delete. Deletes are performed per folder, immediately after that folder is searched,
           because the target mailbox can differ from folder to folder once auxiliary archives are
           involved. Items are removed in $batch requests of -BatchSize, using permanentDelete when
           -HardDelete is set and a normal DELETE otherwise. Throttled items are retried with
           backoff, and every outcome is recorded through ConvertTo-DeleteResult.
           The user is prompted before deleting when the folder search reported failures, because
           the result set is known to be incomplete, and again per folder when -ConfirmDelete is
           used - answering All stops further prompting.

        Nothing is returned. Progress is reported through the log and the two CSV files, and running
        state is kept in script scope: $Script:folderSearchResults for the current folder,
        $Script:TotalSearchResults, $Script:IncompleteSearchFolders, $Script:DeleteFailureFolders
        and the delete counters used by the end of run summary.
    #>

    Write-Log "Performing search against the mailbox..." -Level INFO
    $Script:IncompleteSearchFolders = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    [int64]$Script:TotalSearchResults = 0
    #Perform the search against each folder
    foreach($MailboxFolder in $Script:searchFolders) {
        #Array to hold the search results for the current folder to be used for deletion if the delete switch is set
        $Script:folderSearchResults = [System.Collections.Generic.List[object]]::new()
        Write-Log "Searching folder: $($MailboxFolder.displayName)" -Level INFO
        Write-Log "Processing folder: $($MailboxFolder.id)" -level DEBUG
        [int]$searchFailures=0
        if($Archive){
            #Check to see if the folder is in the main archive or an aux archive
            $Uri = "admin/exchange/mailboxes/$($Script:userMailbox)/folders/$($MailboxFolder.id)"
            $FolderCheckParams = @{
                GraphApiUrl     = $cloudService.graphApiEndpoint
                Endpoint        = "beta"
                Query           = $Uri
            }
            Write-Log "Checking the archive location for folder: $($MailboxFolder.displayName)" -Level DEBUG
            $Script:FolderCheck = Invoke-GraphApiRequest @FolderCheckParams
            #Check the response to see if the folder is in an aux archive mailbox
            if($Script:FolderCheck.StatusCode -eq 308){
                Write-Log $Script:FolderCheck.Response.ErrorMessage -Level DEBUG
                #Modify the URL using the aux archive guid and the folder id for the folder within the aux archive mailbox
                $redirect = [string]$Script:FolderCheck.Response.ErrorMessage
                if ($redirect -notmatch 'MBX:([a-fA-F0-9-]{36})@') {
                    throw "Unable to extract auxiliary archive mailbox ID from redirect."
                }
                $mailboxGuid = $Matches[1]

                if ($redirect -notmatch "folders\('([^']+)'\)") {
                    throw "Unable to extract auxiliary archive folder ID from redirect."
                }
                $folderValue = $Matches[1]

                # The redirect is server supplied, but it is concatenated straight into the request path,
                # so reject anything that could add a segment or a query string.
                if ($folderValue -match '[/\\?#&\s]') {
                    throw "Auxiliary archive folder ID from redirect contains unexpected characters."
                }

                $parsedGuid = [guid]::Empty
                if (-not [guid]::TryParse($mailboxGuid, [ref]$parsedGuid)) {
                    throw "Invalid auxiliary archive mailbox ID."
                }
                $mailboxTarget = "MBX:$mailboxGuid@$OAuthTenantId"
                Write-Log "Checking auxiliary archive mailbox $($mailboxGuid) for items in $($MailboxFolder.displayName)" -Level INFO
                $auxUriQuery = "/users/$($mailboxTarget)/mailFolders/$($folderValue)"
                $Uri = "$($auxUriQuery)/messages?"
                $mailboxName = "$mailboxGuid@$OAuthTenantId"
            }
            else {
                if ($Script:userMailbox -notmatch '^MBX:[0-9a-fA-F-]{36}@') { throw "Unexpected mailbox identifier format." }
                $mailboxName = $Script:userMailbox.Substring(4)
                $Uri = "$($uriQuery)/$($MailboxFolder.id)/messages?"
            }
        }
        else {
            if ($Script:userMailbox -notmatch '^MBX:[0-9a-fA-F-]{36}@') { throw "Unexpected mailbox identifier format." }
            $mailboxName = $script:userMailbox.Substring(4)
            $Uri = "$($uriQuery)/$($MailboxFolder.id)/messages?"
        }
        
        #Use the same search query for all folders, so only build it once if it hasn't been built yet
        if([string]::IsNullOrEmpty($UriFilter)) {
            #Build the search query based on the parameters provided to the script
            $UriFilter = CreateSearchQuery
            Write-Log "Search query: $UriFilter" -Level INFO
        }
        # Search the mailbox for items
        $SearchParams = @{
            GraphApiUrl     = $cloudService.graphApiEndpoint
            Query           =  "$($Uri)$UriFilter"
            Headers     = @{ Prefer = 'IdType="ImmutableId"' }
        }
        
        $SearchItems = Invoke-GraphApiRequest @SearchParams
        #Check for errors in the search request and log them if found, then continue to the next folder
        if($SearchItems.Successful -eq $false){
            Write-Log "Search failed for folder $($MailboxFolder.displayName)." -Level WARN
            Write-Log "Error: $($SearchItems.ErrorMessage)" -Level WARN
            [void]$Script:IncompleteSearchFolders.Add([string]$MailboxFolder.displayName)
            $searchFailures++
            continue
        }
        else{
            $pageResults = [System.Collections.Generic.List[object]]::new()
            foreach($Result in $SearchItems.Content.Value){
                $item = ConvertTo-SearchResult -Result $Result -MailboxName $mailboxName -FolderPath $MailboxFolder.displayName
                if($null -ne $item){
                    [void]$Script:folderSearchResults.Add($item)
                    [void]$pageResults.Add($item)
                    $Script:TotalSearchResults++
                }
            }
            if ($pageResults.Count -gt 0) {
                $pageResults | ConvertTo-SafeCsvRecord | Export-Csv -Path $Script:searchResultsCsvPath -NoTypeInformation -Append -Encoding UTF8
            }
            while($null -ne $SearchItems.Content.'@odata.nextLink'){
                $pageResults = [System.Collections.Generic.List[object]]::new()
                $Query = [string]$SearchItems.Content.'@odata.nextLink'
                $SearchItems = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -Query $Query -Headers @{ Prefer = 'IdType="ImmutableId"' }
                if($SearchItems.Successful -eq $false){
                    Write-Log "Search failed for folder $($MailboxFolder.displayName)." -Level WARN
                    Write-Log "Error: $($SearchItems.ErrorMessage)" -Level WARN
                    $searchFailures++
                    [void]$Script:IncompleteSearchFolders.Add([string]$MailboxFolder.displayName)
                    continue
                }
                else{
                    foreach($Result in $SearchItems.Content.Value){
                        $item = ConvertTo-SearchResult -Result $Result -MailboxName $mailboxName -FolderPath $MailboxFolder.displayName
                        if($null -ne $item){
                            [void]$Script:folderSearchResults.Add($item)
                            [void]$pageResults.Add($item)
                            $Script:TotalSearchResults++
                        }
                    }
                    if ($pageResults.Count -gt 0) {
                        $pageResults | ConvertTo-SafeCsvRecord | Export-Csv -Path $Script:searchResultsCsvPath -NoTypeInformation -Append -Encoding UTF8
                    }
                }
            }
        }
        Write-Log ([string]::Format("Found {0} items in the {1} folder.", $Script:folderSearchResults.Count, $MailboxFolder.displayName)) -Level INFO
        if($searchFailures -gt 0){
            Write-Log ([string]::Format("Search for {0} folder in mailbox {1} had {2} failures.", $MailboxFolder.displayName, $mailboxName, $searchFailures)) -Level WARN
        }
        
        $confirmation = $null
        #Delete items now to ensure correct mailbox using batches
        if($DeleteContent -and $Script:folderSearchResults.count -gt 0){
            if($searchFailures -gt 0){
                Write-Log "Warning: There were $searchFailures errors during the search process." -Level WARN
                Write-Log "These search results are incomplete and may not include all items that match the search criteria." -Level WARN
                Write-Log "Do you still want to delete the items found?" -Level WARN
                $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
                $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
                $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                $confirmation = $host.ui.PromptForChoice("Search errors detected", "Errors occurred during search. Do you still want to delete the items found?", $options, 1)
                if($confirmation -eq 1){
                    Write-Log "User chose not to delete items due to search errors." -Level INFO
                    continue
                }
            }
            #Confirm with the user that they want to continue with the delete since all folders will be searched
            elseif($ConfirmDelete){ 
                $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
                $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
                $all = New-Object System.Management.Automation.Host.ChoiceDescription "&All"
                $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $all)
                $confirmation = $host.ui.PromptForChoice("Confirmation to delete all items found in folder", "Do you want to continue?", $options, 1)
                if($confirmation -eq 2){
                    $ConfirmDelete = $false
                }
            }
            if(($confirmation -eq 0 -or $confirmation -eq 2) -or $ConfirmDelete -eq $false){
                #User has confirmed to continue with the delete, so proceed with the delete operation and don't prompt again
                Write-Log "Deleting $($Script:folderSearchResults.Count) items from $($MailboxFolder.displayName)..." -Level WARN
                [int]$Script:itemsDeleted = 0
                [int]$Script:itemsFailedDelete = 0
                [int]$itemsProcessed = 0
                #prevent batch size being reduced from previous folder search results
                $currentBatchSize = $BatchSize
                # Make sure the results are not less than the batch size
                if($Script:folderSearchResults.count -lt $currentBatchSize){
                    $currentBatchSize = $Script:folderSearchResults.Count
                }
                $Query = "`$batch"
                # Loop thru the results creating batches to delete
                while($itemsProcessed -lt $Script:folderSearchResults.Count){
                    # Make sure the batch size is not greater than the items left to process
                    if(($Script:folderSearchResults.Count - $itemsProcessed) -lt $currentBatchSize){
                        $currentBatchSize = $Script:folderSearchResults.Count - $itemsProcessed
                    }
                    #Create an array of requests to send to the batch endpoint
                    $requests = [System.Collections.Generic.List[object]]::new()
                    #Create a hash table to track delete status for each item in the batch.
                    $itemIdLookup = @{}

                    for($x=0; $x -lt $currentBatchSize; $x++){
                        if($HardDelete){
                            $Method = "POST"
                            $Url = "/users/MBX:$($mailboxName)/messages/$($Script:folderSearchResults[$itemsProcessed].id)/permanentDelete"
                        }
                        else {
                            $Method = "DELETE"
                            $Url = "/users/MBX:$($mailboxName)/messages/$($Script:folderSearchResults[$itemsProcessed].id)"
                        }

                        $request = @{
                            Id          = $x+1
                            Method      = $Method
                            Url         = $Url
                            Headers = @{
                               Prefer = 'IdType="ImmutableId"'
                            }
                        }
                        $itemIdLookup[($x+1)] = $Script:folderSearchResults[$itemsProcessed]
                        [void]$requests.Add($request)
                        $itemsProcessed++
                    }
                    
                    #Retry logic for batch delete requests. If any of the requests in the batch fail with a 429 (Too Many Requests) status code, we will retry those requests up to 4 times with exponential backoff.
                    $pendingRequests = @($requests)
                    $maxAttempts = 4

                    for ($attempt = 1; $attempt -le $maxAttempts -and $pendingRequests.Count; $attempt++) {
                        $batchRequest = @{
                            Requests = $pendingRequests
                            #Requests = $requests
                        } | ConvertTo-Json -Depth 6

                        Write-Log "Sending batch delete request ($($pendingRequests.Count) items, total deleted so far: $itemsDeleted)" -Level DEBUG
                        $batchDeleteResponse = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -Query $Query -Method POST -Body $batchRequest
                        #Setup for list of requests to retry if any of the responses return a 429 status code
                        $retryRequests = [System.Collections.Generic.List[object]]::new()
                        $retryAfterSeconds = 0
                        #Index the outstanding requests by id so retry correlation is O(1) rather than a scan per response.
                        $pendingRequestsById = @{}
                        foreach ($pendingRequest in $pendingRequests) {
                            $pendingRequestsById[[string]$pendingRequest.Id] = $pendingRequest
                        }
  
                        #Check the responses from the batch for any failures
                        if($batchDeleteResponse.Successful -eq $false){
                            Write-Log "Batch request to delete items failed." -Level WARN
                            Write-Log "Error: $($batchDeleteResponse.ErrorMessage)" -Level WARN
                            $Script:itemsFailedDelete += $pendingRequests.Count
                            $Script:TotalDeleteFailures += $pendingRequests.Count
                            [void]$Script:DeleteFailureFolders.Add([string]$MailboxFolder.displayName)
                            #Entire batch failed, so log all items in the batch as failed to delete
                            $deleteFailed = foreach ($request in $pendingRequests) {
                                $item = $itemIdLookup[[int]$request.Id]

                                if ($null -eq $item) {
                                    throw "Unable to correlate failed batch request ID '$($request.Id)' with its source message."
                                }

                                [PSCustomObject]@{
                                    Mailbox          = $item.mailbox
                                    Id               = $item.id
                                    Folder           = $item.folder
                                    InternetMessageId = $item.internetMessageId
                                    Subject          = $item.subject
                                    ReceivedDateTime = $item.receivedDateTime
                                    From             = $item.from
                                    Attachment       = $item.attachment
                                    StatusCode       = $batchDeleteResponse.StatusCode
                                    DeleteStatus     = "Failed"
                                }
                            }
                            $deleteFailed | ConvertTo-SafeCsvRecord | Export-Csv -Path $Script:deleteResultsCsvPath -NoTypeInformation -Append -Encoding UTF8
                            break
                        }
                        else{
                            $deleteResults = [System.Collections.Generic.List[object]]::new()
                            #Check the response for each delete request
                            foreach($response in $batchDeleteResponse.Content.Responses){
                                $result = ConvertTo-DeleteResult -Response $response -currentBatch $itemIdLookup -IsFinalAttempt ($attempt -eq $maxAttempts)
                                if($result.StatusCode -eq 429 -and $attempt -lt $maxAttempts){
                                    #If the request failed with a 429 status code, add it to the list of requests to retry and determine the delay before retrying based on the Retry-After header in the response.
                                    $requestToRetry = $pendingRequestsById[[string]$response.Id]
                                    if ($null -ne $requestToRetry) {
                                        $retryRequests.Add($requestToRetry)
                                    }
                                    $delay = 0
                                    if ([int]::TryParse([string]$response.headers.'Retry-After',[ref]$delay)) {
                                        $retryAfterSeconds = [Math]::Max($retryAfterSeconds,$delay)
                                    }
                                }
                                else{
                                    #Request either successful or failed with a status code other than 429, so add it to the delete results to be logged.
                                    $deleteResults.Add($result)
                                }
                            }
                            $deleteResults | ConvertTo-SafeCsvRecord | Export-Csv -Path $Script:deleteResultsCsvPath -NoTypeInformation -Append -Encoding UTF8
                            Write-Log (
                                "Batch delete completed: {0} responses, {1} succeeded, {2} failed." -f
                                    $deleteResults.Count,
                                    @($deleteResults.Where({ $_.DeleteStatus -eq "Succeeded" })).Count,
                                    @($deleteResults.Where({ $_.DeleteStatus -eq "Failed" })).Count
                            ) -Level DEBUG
                        }
                        #Update the request list for the batch to only include the requests that need to be retried due to a 429 status code. If there are no requests to retry, the loop will exit.
                        $pendingRequests = @($retryRequests)

                        #Determine retry delay based on the Retry-After header in the response. If no Retry-After header is present, use exponential backoff based on the attempt number.
                        if ($pendingRequests.Count -and $attempt -lt $maxAttempts) {
                            if ($retryAfterSeconds -le 0) {
                                $retryAfterSeconds = [Math]::Pow(2, $attempt)
                            }
                            Start-Sleep -Seconds ($retryAfterSeconds + 1)
                        }
                    }
                }
                Write-Log "Delete complete for folder $($MailboxFolder.displayName). $itemsProcessed processed; $Script:itemsDeleted succeeded; $Script:itemsFailedDelete failed; " -Level INFO
            }
        }
    }   
}    
function CreateSearchQuery {

    <#
        Builds the OData query string that is appended to each folder's /messages endpoint. It is
        called once per run by SearchMailbox and the result is reused for every folder, since the
        criteria do not change between them.

        The criteria are read from script scope rather than passed in: $Subject, $ReceivedBefore,
        $ReceivedAfter, $Sender, $AttachmentName, $MessageBody and $ResultSize. Only the ones that
        were supplied contribute a clause, and clauses are combined with "and" by rewriting the
        existing "filter=" prefix, so they can be assembled in any order. If no criteria are given,
        no filter is emitted and every item in the folder matches.

        Values that come from user input are escaped twice on the way in: single quotes are doubled,
        which is how a quote is represented inside an OData string literal, and the result is then
        URL encoded. Without the first step a value containing a quote could close the literal early
        and append predicates of its own. Dates are converted to UTC and written in ISO 8601 form,
        because receivedDateTime is evaluated in UTC.

        Two criteria cannot be fully evaluated by Graph and are only partially represented here:
            AttachmentName - contributes hasAttachments eq true and expands the attachment
                             collection so the names are present in the response. The name itself is
                             matched client side by ConvertTo-SearchResult.
            MessageBody    - not filterable at all. It only causes the body property to be added to
                             $select so the text is available for the client side match.

        The query always ends with a $select, to avoid pulling back properties the script does not
        use, and $top set to $ResultSize. Note that $top is the page size rather than a limit on the
        number of results: SearchMailbox continues to follow @odata.nextLink until the folder is
        exhausted.
    #>

    #Use filter if the message body is not specified, otherwise use search
            #Check if the subject is specified and build the filter query accordingly
            if(-not([string]::IsNullOrEmpty($Subject))) {
                $subjectValue = $Subject.Replace("'", "''")
                $subjectValue = [Uri]::EscapeDataString($subjectValue)
                $UriFilter = "`$filter=contains(subject,`'$subjectValue`')"
            }
            #Check if the received before and after dates are specified and build the filter query accordingly
            if(-not([string]::IsNullOrEmpty($ReceivedBefore))) {
                $TempStartDate = [datetime]$ReceivedBefore
                $TempStartDate = $TempStartDate.ToUniversalTime()
                $SearchStartDate = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $TempStartDate
                if($UriFilter -like '*filter*'){
                    $UriFilter = $UriFilter.Replace('filter=', "filter=receivedDateTime lt $($SearchStartDate) and ")
                }
                else {
                    $UriFilter = "`$filter=receivedDateTime lt $($SearchStartDate)"
                }
            }
            #Check if the received after date is specified and build the filter query accordingly
            if(-not([string]::IsNullOrEmpty($ReceivedAfter))){
                $TempEndDate = [datetime]$ReceivedAfter
                $TempEndDate = $TempEndDate.ToUniversalTime()
                $SearchEndDate = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $TempEndDate
                if($UriFilter -like '*filter*'){
                    $UriFilter = $UriFilter.Replace('filter=', "filter=receivedDateTime gt $($SearchEndDate) and ")
                }
                else {
                    $UriFilter = "`$filter=receivedDateTime gt $($SearchEndDate)"
                }
            }
            #Check if the sender is specified and build the filter query accordingly
            if(-not([string]::IsNullOrEmpty($Sender))){
                $senderValue = $Sender.Replace("'", "''")
                $senderValue = [Uri]::EscapeDataString($senderValue)
                if($UriFilter -like '*filter*'){
                    $UriFilter = $UriFilter.Replace('filter=', "filter=(from/emailAddress/address) eq `'$senderValue`' and ")
                }
                else {
                    $UriFilter = "`$filter=(from/emailAddress/address) eq `'$senderValue`'"
                }
            }
            if(-not([string]::IsNullOrEmpty($AttachmentName))){
                if($UriFilter -like '*filter*'){
                    $UriFilter = $UriFilter.Replace('filter=', "filter=hasAttachments eq true and ")
                    $UriFilter = "$($UriFilter)&`$expand=attachments(`$select=id,name,contentType,size,isInline)"
                }
                else {
                    $UriFilter = "`$filter=hasAttachments eq true&`$expand=attachments(`$select=id,name,contentType,size,isInline)"
                }
            }
            if([string]::IsNullOrEmpty($MessageBody)){
                $UriFilter = "$($UriFilter)&`$select=id,subject,receivedDateTime,from,internetMessageId,hasAttachments&`$top=$($ResultSize)"
            }
            else{
                $UriFilter = "$($UriFilter)&`$select=id,subject,receivedDateTime,from,internetMessageId,hasAttachments,body&`$top=$($ResultSize)"
            }
        return $UriFilter    
}
    
function GetFolderList{

    <#
        Enumerates every mail folder in the target mailbox and stores the raw results in
        $Script:folderList. This is the first step of folder selection: once the list is complete it
        calls BuildFolderListTree, which resolves each folder's parent chain into a full display
        path and produces $Script:folderListTree - the collection the include and exclude parameters
        are matched against.

        The mailFolders/delta endpoint is used rather than mailFolders because delta returns the
        entire hierarchy, nested subfolders included, as one paginated sequence. Enumerating through
        mailFolders instead would only return the children of whichever folder was asked for,
        requiring a request per level of the tree.

        The mailbox is identified by $Script:userMailbox, so whether the primary or the archive is
        enumerated has already been decided by the time this runs. Results are paged through
        @odata.nextLink until no further link is returned.

        Failure at any point is fatal: the error is logged and the script exits, because folder
        selection, searching and deleting all depend on having a complete folder list. Continuing
        with a partial list would risk reporting a folder as clean simply because it was never
        enumerated.
    #>

    Write-Log "Getting a list of folders in the mailbox..." -Level INFO
    #Create an arraylist to hold the folder results
    $Script:folderList = [System.Collections.Generic.List[object]]::new()
    [string]$Query = "users/$($Script:userMailbox)/mailFolders/delta"
    $params = @{
        GraphApiUrl         = $cloudService.graphApiEndpoint
        Query               = $Query
    }
    $Script:FolderResults = Invoke-GraphApiRequest @params
    #Check for errors in the folder enumeration request and log them if found, then exit the script
    if($Script:FolderResults.Successful -eq $false){
        Write-Log "Unable to get a list of folders in the mailbox. Please review the error message below and re-run the script:" -Level ERROR
        Write-Log $Script:FolderResults.ErrorMessage -Level ERROR
        exit 1
    }
    foreach($Result in $Script:FolderResults.Content.Value){
        [void]$Script:folderList.Add($Result)
    }   
    while($null -ne $Script:FolderResults.Content.'@odata.nextLink'){
        $Query = [string]$Script:FolderResults.Content.'@odata.nextLink'
        $Script:FolderResults = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -Query $Query
        if($Script:FolderResults.Successful -eq $false){
            Write-Log "Unable to get a list of folders in the mailbox. Please review the error message below and re-run the script:" -Level ERROR
            Write-Log $Script:FolderResults.ErrorMessage -Level ERROR
            exit 1
        }
        foreach($Result in $Script:FolderResults.Content.Value){
            [void]$Script:folderList.Add($Result)
        }
    }
    
    Write-Log "Folder enumeration complete. $($Script:folderList.Count) folders found." -Level INFO
    BuildFolderListTree
}

function BuildFolderListTree{

    <#
        Turns the flat folder list produced by GetFolderList into something the include and exclude
        parameters can be matched against.

        Graph returns each folder with only its own displayName and a parentFolderId, so nothing in
        the raw list says where a folder sits in the hierarchy - and display names are not unique, as
        several parents can each contain an "Archive" or "2024" subfolder. This function walks each
        folder's parent chain and rewrites displayName in place to the full backslash delimited path,
        for example "\Inbox\Projects\2024". Those paths are what -IncludeFolderList and
        -ExcludeFolderList are compared against, and what appears in the log and CSV output.

        Paths are resolved recursively with two safeguards. Results are memoised in $pathCache, so a
        deep tree resolves each ancestor once rather than once per descendant. A set of the folders
        currently being visited is carried down the recursion, and a folder that is encountered while
        already being resolved means the parent chain loops back on itself, which would otherwise
        recurse until the call stack was exhausted; that case throws instead. A folder whose parent
        is missing from the list is treated as a root, which is what happens for the mailbox root
        itself.

        The result is stored in $Script:folderListTree, sorted by path so that parents appear
        immediately before their children in output and prefix matching behaves predictably.

        Note that this mutates the objects in $Script:folderList rather than copying them, so after
        this runs displayName is a full path everywhere, not a leaf name.
    #>

    Write-Log "Generating folder hierarchy structure for folders" -Level DEBUG
    $foldersById = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::Ordinal)
    $namesById = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    $pathCache = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)

    foreach ($folder in $Script:folderList) {
        $foldersById[[string]$folder.id] = $folder
        $namesById[[string]$folder.id] = $folder.displayName
    }

    function Resolve-FolderPath {
        param(
            [object]$Folder,
            [System.Collections.Generic.HashSet[string]]$Visiting
        )

        $folderId = [string]$Folder.id

        if ($pathCache.ContainsKey($folderId)) {
            return $pathCache[$folderId]
        }

        if (-not $Visiting.Add($folderId)) {
            throw "Circular folder relationship detected at folder '$folderId'."
        }

        $name = $namesById[$folderId]
        $parentId = [string]$Folder.parentFolderId

        if ([string]::IsNullOrEmpty($parentId) -or -not $foldersById.ContainsKey($parentId)) {
            $path = "\$name"
        }
        else {
            $parentPath = Resolve-FolderPath -Folder $foldersById[$parentId] -Visiting $Visiting
            $path = "$parentPath\$name"
        }

        [void]$Visiting.Remove($folderId)
        $pathCache[$folderId] = $path

        return $path
    }

    foreach ($folder in $Script:folderList) {
        $visiting = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
        $folder.displayName = Resolve-FolderPath -Folder $folder -Visiting $visiting
    }

    $Script:folderListTree = @(
        $Script:folderList | Sort-Object displayName
    )
}

function GetRecoverableItemsFolderList{

    <#
        The -SearchDumpster counterpart to GetFolderList. It enumerates the Recoverable Items
        subtree - Deletions, Purges, Versions, DiscoveryHolds and so on - and leaves the results in
        $Script:folderList before calling BuildFolderListTree, so everything downstream (include and
        exclude matching, searching, deleting) works exactly as it does for a normal folder list.

        The delta endpoint used by GetFolderList is not an option here. Recoverable Items folders are
        hidden, so they are not returned by mailFolders/delta. Instead the tree is walked explicitly
        from the well known RecoverableItemsRoot folder using childFolders with
        includeHiddenFolders=true, which is what makes the hidden folders visible at all.

        Because childFolders only returns one level, the walk is breadth first: a queue holds the
        folders whose children still need to be requested, and each child found is both recorded and
        enqueued so its own children are collected on a later pass. A set of already visited ids
        guards against a folder being processed twice, which would otherwise duplicate entries or
        loop indefinitely. Each level is paged through @odata.nextLink before moving on.

        Note that RecoverableItemsRoot is seeded as a placeholder object purely to start the walk, so
        it is used as a parent but never added to the folder list itself - only its descendants are
        searchable.

        A failure to enumerate any level throws rather than exiting, so the script's outer try and
        finally still run and sensitive state is cleared and the log is closed.
    #>

    Write-Log "Getting a list of folders in the recoverable items..." -Level INFO
    #Create an arraylist to hold the folder results
    $Script:folderList = [System.Collections.Generic.List[object]]::new()
    $queue = [System.Collections.Generic.Queue[object]]::new()
    $visited = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)

    $queue.Enqueue([PSCustomObject]@{
        id = "RecoverableItemsRoot"
    })

    while ($queue.Count -gt 0) {
        $parent = $queue.Dequeue()

        if (-not $visited.Add([string]$parent.id)) {
            continue
        }

        $nextLink = "users/$($Script:userMailbox)/mailFolders/$($parent.id)/childFolders?includeHiddenFolders=true"

        while ($nextLink) {
            $page = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -Query $nextLink

            if (-not $page.Successful) {
                throw "Failed to enumerate children of Recoverable Items folder '$($parent.id)': $($page.ErrorMessage)"
            }

            foreach ($folder in @($page.Content.Value)) {
                [void]$Script:folderList.Add($folder)
                $queue.Enqueue($folder)
            }

            $nextLink = $page.Content.'@odata.nextLink'
        }
    }

    BuildFolderListTree
    Write-Log "Recoverable items enumeration complete. $($Script:folderList.Count) folders found." -Level INFO
}

function Add-SearchFolder {
    param([Parameter(Mandatory)][object]$Folder)

    <#
        The single gate through which every folder must pass to end up in $Script:searchFolders. It
        exists so the include list, the subfolder expansion and the aux archive expansion below can
        all add folders freely without any of them having to worry about whether another has already
        added the same one.

        Deduplication is keyed on the Graph folder id rather than the display name, because names are
        not unique across a mailbox and the tree paths built by BuildFolderListTree are only unique
        by convention. The $searchFolderIds HashSet does double duty: its Add returns false when the
        id is already present, so the membership test and the insert are a single operation. The
        comparison is ordinal because Graph folder ids are opaque, case sensitive tokens.

        Without this, a folder reachable through more than one rule - for example one named
        explicitly in -IncludeFolderList and also picked up by -ProcessSubfolders, or an aux archive
        subfolder matched by both the include pattern and the aux archive pattern - would be searched
        twice and its messages counted, reported and deleted twice over.

        A folder with no id is fatal rather than skipped. Silently dropping it would leave the caller
        believing a folder they asked for had been searched and found clean, which is the one failure
        mode this script must never produce.

        Note that this closes over $searchFolderIds and $Script:searchFolders from the enclosing
        script scope rather than taking them as parameters - it is defined inline in the main body
        and is not intended for use anywhere else.
    #>

    $folderId = [string]$Folder.id
    if ([string]::IsNullOrWhiteSpace($folderId)) {
        throw "A selected folder has no valid Graph folder ID."
    }

    if ($searchFolderIds.Add($folderId)) {
        $Script:searchFolders.Add($Folder)
    }
}

#Safety check to ensure the search is not against the entire mailbox and ConfirmDelete is set to false which would result in all items being deleted from the mailbox
if((([string]::IsNullOrEmpty($IncludeFolderList)) -and ([string]::IsNullOrEmpty($ExcludeFolderList)) -and $ConfirmDelete -eq $false) -and $DeleteContent){
    Write-Log "Both IncludeFolderList and ExcludeFolderList are not specified and ConfirmDelete is set to false. This could result in all items being deleted from the mailbox. Please review the parameters and try again." -Level ERROR
    Close-Log
    exit 0
}

#Peformance check to ensure MessageBody filter is no the only filter specified since this will result in a full mailbox search and could take a long time to complete
if(([string]::IsNullOrWhiteSpace($Subject) -and [string]::IsNullOrEmpty($ReceivedBefore) -and [string]::IsNullOrEmpty($ReceivedAfter) -and [string]::IsNullOrEmpty($Sender) -and [string]::IsNullOrWhiteSpace($AttachmentName)) -and (-not [string]::IsNullOrWhiteSpace($MessageBody))){
    Write-Log "MessageBody is the only filter specified. This will result in a full mailbox search and could take a long time to complete. Please review the parameters and try again." -Level WARN
    Close-Log
    exit 0
}

#Safety check for an unfiltered search and delete
if([string]::IsNullOrWhiteSpace($Subject) -and [string]::IsNullOrEmpty($ReceivedBefore) -and [string]::IsNullOrEmpty($ReceivedAfter) -and [string]::IsNullOrEmpty($Sender) -and [string]::IsNullOrWhiteSpace($AttachmentName) -and $DeleteContent){
    Write-Log "No search filters specified and DeleteContent is set to true. This could result in all items being deleted from the mailbox. Please review the parameters and try again." -Level WARN
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    $confirmation = $host.ui.PromptForChoice("Confirmation", "Do you want to continue with this search?", $options, 1)
    #Confirm with the user that they want to continue with the search since all content in folders will be deleted
    if($confirmation -ne 0){
        Write-Log "Search cancelled by user." -Level WARN
        Close-Log
        exit 0
    }
}

#Ensure the IncludeFolderList and ExcludeFolderList values all start with a backslash to match the folder display names
if(-not([string]::IsNullOrEmpty($IncludeFolderList))){
    # Convert to array if only one entry provided as a string
    if($IncludeFolderList -isnot [System.Collections.ArrayList] -and $IncludeFolderList -isnot [array]){
        $IncludeFolderList = @($IncludeFolderList)
    }
    for ($i = 0; $i -lt $IncludeFolderList.Count; $i++) {
        $normalized = ([string]$IncludeFolderList[$i]).Trim().Trim('\')

        if ([string]::IsNullOrWhiteSpace($normalized)) {
            Write-Log "IncludeFolderList contains an empty folder path." -Level ERROR
            Close-Log
            throw "IncludeFolderList contains an empty folder path."
        }
        $IncludeFolderList[$i] = "\$normalized"
    }
}
if(-not([string]::IsNullOrEmpty($ExcludeFolderList))){
    # Convert to array if only one entry provided as a string
    if($ExcludeFolderList -isnot [System.Collections.ArrayList] -and $ExcludeFolderList -isnot [array]){
        $ExcludeFolderList = @($ExcludeFolderList)
    }
    for ($i = 0; $i -lt $ExcludeFolderList.Count; $i++) {
        $normalized = ([string]$ExcludeFolderList[$i]).Trim().Trim('\')

        if ([string]::IsNullOrWhiteSpace($normalized)) {
            Write-Log "ExcludeFolderList contains an empty folder path." -Level ERROR
            Close-Log
            throw "ExcludeFolderList contains an empty folder path."
        }
        $ExcludeFolderList[$i] = "\$normalized"
    }
}

try{
#Get parameters and pass to obtain an OAuth token
$cloudService = Get-CloudServiceEndpoint $AzureEnvironment
$azureADEndpoint = $cloudService.AzureADEndpoint
$Script:applicationInfo = @{
    "TenantID" = $OAuthTenantId
    "ClientID" = $OAuthClientId
}
$Script:GraphScope = "$($cloudService.GraphApiEndpoint)/"
Get-OAuthToken -AppScope $Scope

#Get the mailbox settings to retrieve the primary and archive mailbox guids for the specified mailbox
[string]$Query = "users/$Mailbox/settings/exchange"
$params = @{
    GraphApiUrl         = $cloudService.graphApiEndpoint
    Query               = $Query
    Endpoint            = 'beta'
}
$Script:MailboxSettings = Invoke-GraphApiRequest @params
#Check for errors in the mailbox settings request and log them if found, then exit the script
if($MailboxSettings.Successful -eq $false){
    Write-Log "Unable to retrieve mailbox settings for $Mailbox. Please check the mailbox name and try again." -Level ERROR
    Write-Log "Error details: $($MailboxSettings.ErrorMessage)" -Level ERROR
    exit 1
}
#Capture the primary and archive mailbox guids for use in the search
$Script:archiveMailbox = $MailboxSettings.Content.inPlaceArchiveMailboxId
$Script:primaryMailbox = $MailboxSettings.content.primaryMailboxId
#Determine which mailbox to use based on the Archive switch parameter
if($Archive){
    $Script:userMailbox = $script:archiveMailbox
    Write-Log "Using archive mailbox: $($Script:userMailbox)" -Level INFO
}
else{
    $Script:userMailbox = $script:primaryMailbox
    Write-Log "Using primary mailbox: $($Script:userMailbox)" -Level INFO
}

#Get a list of folders in the mailbox
if(-not($SearchDumpster)){
    GetFolderList
}
else{
    GetRecoverableItemsFolderList
}

#Determine the folder to search based on include/exclude lists
Write-Log "Determining folders to search..." -Level INFO
#Create an arraylist to hold the folders to be searched
$Script:searchFolders = [System.Collections.Generic.List[object]]::new()
$searchFolderIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)

#Check is specific folders are specified in the include list, if not, search all folders under the root
if([string]::IsNullOrEmpty($IncludeFolderList)){
    #If no include list is specified, search all folders under the desired root
    #$Script:searchFolders = $Script:folderListTree
    foreach ($folder in $Script:folderListTree) {
        Add-SearchFolder -Folder $folder
    }
    if(-not($ExcludeFolderList)){
        Write-Log "No include list specified. Searching all folders ($($Script:searchFolders.Count) folders)..." -Level WARN
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $confirmation = $host.ui.PromptForChoice("Confirmation", "Do you want to continue?", $options, 1)
        #Confirm with the user that they want to continue with the search since all folders will be searched
        if($confirmation -ne 0){
            Write-Log "Search cancelled by user." -Level WARN
            exit 3
        }
    }
}
else {
    Write-Log "Building folder search list from include list..." -Level INFO
    #Add folders that match the include list
    foreach ($folderPath in $IncludeFolderList) {
        if ($ProcessSubfolders) {
            $matchedFolders = @(
                $Script:folderListTree | Where-Object {
                    $_.displayName -match "^$([regex]::Escape($folderPath))($|\\)"
                }
            )
        }
        else {
            $matchedFolders = @(
                $Script:folderListTree | Where-Object {
                    [string]::Equals([string]$_.displayName,[string]$folderPath,[StringComparison]::OrdinalIgnoreCase)
                }
            )
        }

        if ($matchedFolders.Count -eq 0) {
            throw "Included folder '$folderPath' was not found."
        }

        foreach ($match in $matchedFolders) {
            Add-SearchFolder -Folder $match
        }

        if ($Archive -and -not $ProcessSubfolders) {
            $auxSubfolders = @(
                $Script:folderListTree | Where-Object {
                    $_.displayName -match "^$([regex]::Escape($folderPath))\\"
                } | Where-Object {
                        $parts = $_.displayName -split '\\'
                        $rootFolder = $parts[-2]
                        $lastPart = $parts[-1]
                        $lastPart -match "^$([regex]::Escape($rootFolder))_\d{4}\s+\(Created on"
                    }
            )

            foreach ($auxSubfolder in $auxSubfolders) {
                Add-SearchFolder -Folder $auxSubfolder
            }
        }
    }
}

#Remove folders that match the exclude list and includes all subfolders
if($ExcludeFolderList){
    Write-Log "Removing excluded folders from the list..." -Level INFO
    #Find all folders that match the exclude list and add them to the remove list
    foreach ($exclude in $ExcludeFolderList) {
        $excludeExists = @(
            $Script:folderListTree | Where-Object {
                    [string]::Equals([string]$_.displayName,[string]$exclude,[StringComparison]::OrdinalIgnoreCase)
            }
        ).Count -gt 0

        if (-not $excludeExists) {
            throw "Excluded folder '$exclude' was not found. Deletion cannot safely continue."
        }
    }

    $newSearchFolders = [System.Collections.Generic.List[object]]::new()

    foreach ($folder in $Script:searchFolders) {
        $shouldExclude = $false

        foreach ($exclude in $ExcludeFolderList) {
            if ([string]::Equals([string]$folder.displayName,[string]$exclude,[StringComparison]::OrdinalIgnoreCase) -or
                $folder.displayName.StartsWith("$exclude\",[StringComparison]::OrdinalIgnoreCase)){
                    $shouldExclude = $true
                    break
            }
        }

        if (-not $shouldExclude) {
            $newSearchFolders.Add($folder)
        }
    }

    $Script:searchFolders = $newSearchFolders
}

if ($Script:searchFolders.Count -eq 0) {
    throw "No folders remain in the search list."
}


Write-Log "Final list of folders to be searched ($($Script:searchFolders.Count) folders):" -Level INFO
$Script:searchFolders | ForEach-Object { Write-Log "  $($_.displayName)" -Level DEBUG }
$Script:searchFolders | Format-Table displayName

#Create csv file to hold the search results
$Script:searchResultsCsvPath = Join-Path $OutputPath "SearchResults_$($Script:RunId).csv"
Write-Log "Search results will be saved to: $($Script:searchResultsCsvPath)" -Level INFO
if($DeleteContent){
    $Script:deleteResultsCsvPath = Join-Path $OutputPath "DeleteResults_$($Script:RunId).csv"
    Write-Log "Delete results will be saved to: $($Script:deleteResultsCsvPath)" -Level INFO
}

[int64]$Script:TotalItemsDeleted = 0
[int64]$Script:TotalDeleteFailures = 0
$Script:DeleteFailureFolders = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

#Initiate the search against the mailbox using the specified query and parameters
SearchMailbox -uriQuery "/users/$Script:userMailbox/mailFolders"
if ($Script:IncompleteSearchFolders.Count -gt 0) {
    Write-Log (
        "Search completed incompletely. {0} folder(s) could not be fully searched: {1}" -f
        $Script:IncompleteSearchFolders.Count,
        ($Script:IncompleteSearchFolders -join "; ")
    ) -Level ERROR
}
#Export the search results to a CSV file
if($Script:TotalSearchResults -gt 0){
    Write-Log "Search complete. $Script:TotalSearchResults item(s) found in total." -Level INFO
    Write-Log "Search results exported to: $($Script:searchResultsCsvPath)" -Level INFO
}
if ($DeleteContent -and ($Script:TotalItemsDeleted -gt 0 -or $Script:TotalDeleteFailures -gt 0)) {
    if ($Script:TotalDeleteFailures -gt 0) {
        Write-Log (
            "Deletion completed with failures. {0} succeeded; {1} failed. Affected folders: {2}" -f
            $Script:TotalItemsDeleted,
            $Script:TotalDeleteFailures,
            ($Script:DeleteFailureFolders -join "; ")
        ) -Level WARN
    }
    else {
        Write-Log ("Deletion completed successfully. {0} item(s) deleted; 0 failures." -f $Script:TotalItemsDeleted) -Level INFO
    }
    Write-Log "Delete results exported to: $($Script:deleteResultsCsvPath)" -Level INFO
}
Write-Log "Script completed. Log file: $($Script:LogFile)" -Level INFO
}
finally{
    Clear-SensitiveState
    Close-Log
}