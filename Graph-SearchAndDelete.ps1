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

# Version 20260819.1912

param (
    [Parameter(Position=0,Mandatory=$false,HelpMessage="The Mailbox parameter specifies the mailbox to be accessed.")]
    [ValidateNotNullOrEmpty()] 
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
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthCertificate parameter is the certificate for the registered application. Certificate auth requires MSAL libraries to be available.")] 
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

    [ValidateScript({ Test-Path $_ })]
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

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
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

if (-not([string]::IsNullOrEmpty($OutputPath))) {
    $Script:LogFile = Join-Path $OutputPath "GraphSearch_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
} else {
    $Script:LogFile = Join-Path $PSScriptRoot "GraphSearch_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
}
Initialize-Log -Path $Script:LogFile

Write-Log "Script started. Mailbox: $Mailbox | Archive: $Archive | SearchDumpster: $SearchDumpster | PermissionType: $PermissionType"
Write-Log "Output path: $OutputPath | Log file: $Script:LogFile"
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
            $certificate = Get-ChildItem Cert:\$CertificateStore\My\$CertificateThumbprint
            if ($certificate.HasPrivateKey) {
                $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
                # Base64url-encoded SHA-1 thumbprint of the X.509 certificate's DER encoding
                $x5t = [System.Convert]::ToBase64String($certificate.GetCertHash())
                $x5t = ((($x5t).Replace("\+", "-")).Replace("/", "_")).Replace("=", "")
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

        $signature = $privateKey.SignData($signatureInput, $signingAlgorithmToUse, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
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
        [Parameter(Mandatory = $true)][string]$Secret,
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
            $body.Add("client_assertion", $Secret)
        } else {
            Write-Log "Authentication is based on a secret" -Level DEBUG
            $bstr = [IntPtr]::Zero
            $plainSecret = $null
            try {
                $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($OAuthClientSecret)
                $plainSecret = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
                $body.client_secret = $plainSecret
            }
            finally{
                $plainSecret = $null
                if ($bstr -ne [IntPtr]::Zero) {
                    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
                }
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
        }
        #>
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

function Update-AccessTokenIfNeeded {
    param(
        [switch]$Force
    )
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
            CertificateBasedAuthentication = -not [string]::IsNullOrEmpty(
                $Script:applicationInfo.CertificateThumbprint
            )
        }

        if ($tokenParams.CertificateBasedAuthentication) {
            $jwt = Get-NewJsonWebToken `
                -CertificateThumbprint $Script:applicationInfo.CertificateThumbprint `
                -CertificateStore $CertificateStore `
                -Issuer $Script:applicationInfo.ClientID `
                -Subject $Script:applicationInfo.ClientID `
                -Audience "$($cloudService.AzureADEndpoint)/$($Script:applicationInfo.TenantID)/oauth2/v2.0/token"

            if (-not $jwt) {
                throw "Unable to generate a certificate assertion."
            }

            $tokenParams.Secret = $jwt
        }
        else {
            $tokenParams.Secret = $Script:applicationInfo.AppSecret
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

        # Request an authorization code from the Microsoft Azure Active Directory endpoint
        $authCodeRequestUrl = "$AzureADEndpoint/organizations/oauth2/v2.0/authorize?client_id=$clientId" +
        "&response_type=$responseType&redirect_uri=$redirectUri&scope=$scope&state=$state&prompt=$prompt" +
        "&code_challenge_method=$codeChallengeMethod&code_challenge=$codeChallenge"

        Start-Process -FilePath $authCodeRequestUrl
        $authCodeResponse = Start-LocalListener

        if ($null -ne $authCodeResponse) {
            # Redeem the returned code for an access token
            $redeemAuthCodeParams = @{
                Uri             = "$AzureADEndpoint/organizations/oauth2/v2.0/token"
                Method          = "POST"
                ContentType     = "application/x-www-form-urlencoded"
                Body            = @{
                    client_id     = $ClientID
                    scope         = $Scope
                    code          = ($($authCodeResponse.Split("=")[1]).Split("&")[0])
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
                #TenantId    = (Convert-JsonWebTokenToObject $tokens.id_token).Payload.tid
                LastTokenRefreshTime = (Get-Date)
                Successful           = $true
            }
        }
        exit
    }
}

function Convert-JsonWebTokenToObject {
    param(
        [Parameter(Mandatory = $true)][ValidatePattern("^([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_\-\+\/=]*)")][string]$Token
    )

    <#
        This function can be used to split a JSON web token (JWT) into its header, payload, and signature.
        The JWT is expected to be in the format of <header>.<payload>.<signature>.
        The function returns a PSCustomObject with the following properties:
            Header    - The header of the JWT
            Payload   - The payload of the JWT
            Signature - The signature of the JWT

            It returns $null if the JWT is not in the expected format or conversion fails.
    #>

    begin {
        Write-Log "Calling $($MyInvocation.MyCommand)" -Level DEBUG
        function ConvertJwtFromBase64StringWithoutPadding {
            param(
                [Parameter(Mandatory = $true)]
                [string]$Jwt
            )
            $Jwt = ($Jwt.Replace("-", "+")).Replace("_", "/")
            switch ($Jwt.Length % 4) {
                0 { return [System.Convert]::FromBase64String($Jwt) }
                2 { return [System.Convert]::FromBase64String($Jwt + "==") }
                3 { return [System.Convert]::FromBase64String($Jwt + "=") }
                default { throw "The JWT is not a valid Base64 string." }
            }
        }
    }
    process {
        $tokenParts = $Token.Split(".")
        $tokenHeader = $tokenParts[0]
        $tokenPayload = $tokenParts[1]
        $tokenSignature = $tokenParts[2]

        Write-Log "Now processing token header..." -Level DEBUG
        $tokenHeaderDecoded = [System.Text.Encoding]::UTF8.GetString((ConvertJwtFromBase64StringWithoutPadding $tokenHeader))

        Write-Log "Now processing token payload..." -Level DEBUG
        $tokenPayloadDecoded = [System.Text.Encoding]::UTF8.GetString((ConvertJwtFromBase64StringWithoutPadding $tokenPayload))

        Write-Log "Now processing token signature..." -Level DEBUG
        $tokenSignatureDecoded = [System.Text.Encoding]::UTF8.GetString((ConvertJwtFromBase64StringWithoutPadding $tokenSignature))
    }
    end {
        if (($null -ne $tokenHeaderDecoded) -and
            ($null -ne $tokenPayloadDecoded) -and
            ($null -ne $tokenSignatureDecoded)) {
            Write-Log "Conversion of the token was successful" -Level DEBUG
            return [PSCustomObject]@{
                Header    = ($tokenHeaderDecoded | ConvertFrom-Json)
                Payload   = ($tokenPayloadDecoded | ConvertFrom-Json)
                Signature = $tokenSignatureDecoded
            }
        }

        Write-Log "Conversion of the token failed" -Level DEBUG
        return $null
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
    ([System.Security.Cryptography.RandomNumberGenerator]::Create()).GetBytes($bytes)
    $b64String = [Convert]::ToBase64String($bytes)
    $verifier = (($b64String.TrimEnd("=")).Replace("+", "-")).Replace("/", "_")

    $newMemoryStream = [System.IO.MemoryStream]::new()
    $newStreamWriter = [System.IO.StreamWriter]::new($newMemoryStream)
    $newStreamWriter.write($verifier)
    $newStreamWriter.Flush()
    $newMemoryStream.Position = 0
    $hash = Get-FileHash -InputStream $newMemoryStream | Select-Object Hash
    $hex = $hash.Hash

    $bytesArray = [byte[]]::new($hex.Length / 2)

    for ($i = 0; $i -lt $hex.Length; $i+=2) {
        $bytesArray[$i/2] = [Convert]::ToByte($hex.Substring($i, 2), 16)
    }

    $base64Encoded = [Convert]::ToBase64String($bytesArray)
    $base64UrlEncoded = (($base64Encoded.TrimEnd("=")).Replace("+", "-")).Replace("/", "_")

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

                while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
                    if ($task.AsyncWaitHandle.WaitOne(100)) {
                        $signalled = $true
                        break
                    }
                    Start-Sleep -Milliseconds 100
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
                        Write-Log "Request made to listener but the url that was called is not as expected. URL: $($url)" -Level DEBUG
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
            #Uri             = "$GraphApiUrl/$Endpoint/$($Query.TrimStart("/"))"
            Uri             = $requestUri
            #Header          = @{ Authorization = "Bearer $Script:Token" }
            Header          = $requestHeaders
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
            $graphApiRequestParams.Header.Authorization = "Bearer $Script:Token"
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

    Write-Log "Calling $($MyInvocation.MyCommand)" -Level DEBUG
    if ([System.String]::IsNullOrEmpty($Uri)) {
        $Uri = $ParametersObject.Uri
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Confirm-ProxyServer -TargetUri $Uri) {
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell")
        $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
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
                    $responseContent = $responseBody | ConvertFrom-Json
                    Write-Log "Graph API request failed with status code: $statusCode" -Level DEBUG
                    Write-Log "Error message: $($responseContent.error.message)" -Level DEBUG
                    return [PSCustomObject]@{
                        ErrorCode    = $responseContent.error.code
                        ErrorMessage   = $responseContent.error.message
                        StatusCode = [int]$httpResponse.StatusCode
                        Successful = $false
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
            $retryAfter = [string]$httpResponse.Headers['Retry-After']

            if (-not [int]::TryParse($retryAfter, [ref]$retryAfterSeconds)) {
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

    Write-Log "Calling $($MyInvocation.MyCommand)" -Level DEBUG
    try {
        $proxyObject = ([System.Net.WebRequest]::GetSystemWebProxy()).GetProxy($TargetUri)
        if ($TargetUri -ne $proxyObject.OriginalString) {
            Write-Log "Proxy server configuration detected" -Level DEBUG
            Write-Log $proxyObject.OriginalString -Level DEBUG
            return $true
        } else {
            Write-Log "No proxy server configuration detected" -Level DEBUG
            return $false
        }
    } catch {
        Write-Log "Unable to check for proxy server configuration" -Level DEBUG
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
    if($PermissionType -eq "Application") {
        $Script:GraphScope = "$($Script:GraphScope).default"
        if ([System.String]::IsNullOrEmpty($OAuthCertificate)) {
            $Script:applicationInfo.Add("AppSecret", $OAuthClientSecret)
        }
        else {
            $jwtParams = @{
                CertificateThumbprint = $OAuthCertificate
                CertificateStore      = $CertificateStore
                Issuer                = $OAuthClientId
                Audience              = "$azureADEndpoint/$OAuthTenantId/oauth2/v2.0/token"
                Subject               = $OAuthClientId
            }
            $jwt = Get-NewJsonWebToken @jwtParams
    
            if ($null -eq $jwt) {
                Write-Log "Unable to generate Json Web Token by using certificate: $CertificateThumbprint" -Level ERROR
                exit 1
            }
    
            $Script:applicationInfo.Add("AppSecret", $jwt)
            $Script:applicationInfo.Add("CertificateThumbprint", $OAuthCertificate)
        }
    
        $createOAuthTokenParams = @{
            TenantID                       = $OAuthTenantId
            ClientID                       = $OAuthClientId
            Secret                         = $Script:applicationInfo.AppSecret
            Scope                          = $Script:GraphScope
            Endpoint                       = $azureADEndpoint
            CertificateBasedAuthentication = (-not([System.String]::IsNullOrEmpty($OAuthCertificate)))
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
        [object]$currentBatch
    )
    #Create object with item information
    $item = [PSCustomObject]@{
        Mailbox=$currentBatch[[int]$Response.id].mailbox
        Id=$currentBatch[[int]$Response.id].id
        Folder=$currentBatch[[int]$Response.id].folder
        Subject=$currentBatch[[int]$Response.id].subject
        ReceivedDateTime=$currentBatch[[int]$Response.id].receivedDateTime
        From=$currentBatch[[int]$Response.id].from
        Attachment=$currentBatch[[int]$Response.id].attachment
        StatusCode=[int]$Response.Status
    }
    #Add the delete status to the object based on the response from Graph API
    <#
    if($Response.Status -ne 204){
        Write-Log $Response.Body.Error.Message -Level WARN
        $Script:itemsFailedDelete++
        $item | Add-Member -MemberType NoteProperty -Name "DeleteStatus" -Value "Failed"
    }
    else{
        $Script:itemsDeleted++
        $item | Add-Member -MemberType NoteProperty -Name "DeleteStatus" -Value "Succeeded"
    }
    #>
    switch($Response.Status){
        204 {
            $Script:itemsDeleted++
            $item | Add-Member -MemberType NoteProperty -Name "DeleteStatus" -Value "Succeeded"
        }
        429 {
            Write-Log "Too many requests. Retrying deletion later." -Level WARN
            $Script:itemsRetry++
            $item | Add-Member -MemberType NoteProperty -Name "DeleteStatus" -Value "Retry"
        }
        default {
            Write-Log "Failed to delete item. Status code: $($Response.Status)" -Level WARN
            $Script:itemsFailedDelete++
            $item | Add-Member -MemberType NoteProperty -Name "DeleteStatus" -Value "Failed"
        }
    }
    

    return $item
}
function SearchMailbox {
    param(
        [string]$uriQuery
    )
    Write-Log "Performing search against the mailbox..." -Level INFO
    #Array to hold the search results for all folders
    [int64]$Script:TotalSearchResults = 0
    #Perform the search against each folder
    foreach($MailboxFolder in $Script:searchFolders) {
        #Array to hold the search results for the current folder to be used for deletion if the delete switch is set
        $Script:folderSearchResults = New-Object System.Collections.ArrayList
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
                $mailboxGuid = $FolderCheck.Response.ErrorMessage -match 'MBX:([a-f0-9\-]+)@' | ForEach-Object { $matches[1] }
                $folderValue = $FolderCheck.Response.ErrorMessage -match "folders\('([^']+)'\)" | ForEach-Object { $matches[1] }
                Write-Log "Checking auxiliary archive mailbox $($mailboxGuid) for items in $($MailboxFolder.displayName)" -Level INFO
                $auxUriQuery = "/users/MBX:$($mailboxGuid)@$($OAuthTenantId)/mailFolders/$($folderValue)"
                $Uri = "$($auxUriQuery)/messages?"
                $mailboxName = $mailboxGuid
            }
            else {
                $mailboxName = $Script:userMailbox.Substring(4)
                $Uri = "$($uriQuery)/$($MailboxFolder.id)/messages?"
            }
        }
        else {
            $mailboxName = $script:userMailbox.Substring(4)
            $Uri = "$($uriQuery)/$($MailboxFolder.id)/messages?"
        }
        
        #Use the same search query for all folders, so only build it once if it hasn't been built yet
        if([string]::IsNullOrEmpty($UriFilter)) {
            #Build the search query based on the parameters provided to the script
            $UriFilter = CreateSearchQuery
        }
        # Search the mailbox for items
        $SearchParams = @{
            GraphApiUrl     = $cloudService.graphApiEndpoint
            Query           =  "$($Uri)$UriFilter"
            #Headers     = @{ Prefer = 'IdType="ImmutableId"' }
        }
        
        $SearchItems = Invoke-GraphApiRequest @SearchParams
        #Check for errors in the search request and log them if found, then continue to the next folder
        if($SearchItems.Successful -eq $false){
            Write-Log "Search failed for folder $($MailboxFolder.displayName)." -Level WARN
            Write-Log "Error: $($SearchItems.ErrorMessage)" -Level WARN
            $searchFailures++
            continue
        }
        else{
            $pageResults = [System.Collections.Generic.List[object]]::new()
            foreach($Result in $SearchItems.Content.Value){
                $item = ConvertTo-SearchResult -Result $Result -MailboxName $mailboxName -FolderPath $MailboxFolder.displayName
                if($null -ne $item){
                    $Script:folderSearchResults.Add($item) | Out-Null
                    $pageResults.Add($item)
                    $Script:TotalSearchResults++
                }
            }
            if ($pageResults.Count -gt 0) {
                $pageResults | Export-Csv -Path $Script:searchResultsCsvPath -NoTypeInformation -Append -Encoding UTF8
            }
            while($null -ne $SearchItems.Content.'@odata.nextLink'){
                $pageResults = [System.Collections.Generic.List[object]]::new()
                $Query = [string]$SearchItems.Content.'@odata.nextLink'
                $SearchItems = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -Query $Query #-Headers @{ Prefer = 'IdType="ImmutableId"' }
                if($SearchItems.Successful -eq $false){
                    Write-Log "Search failed for folder $($MailboxFolder.displayName)." -Level WARN
                    Write-Log "Error: $($SearchItems.ErrorMessage)" -Level WARN
                    $searchFailures++
                    continue
                }
                else{
                    foreach($Result in $SearchItems.Content.Value){
                        $item = ConvertTo-SearchResult -Result $Result -MailboxName $mailboxName -FolderPath $MailboxFolder.displayName
                        if($null -ne $item){
                            $Script:folderSearchResults.Add($item) | Out-Null
                            $pageResults.Add($item)
                            $Script:TotalSearchResults++
                        }
                    }
                    if ($pageResults.Count -gt 0) {
                        $pageResults | Export-Csv -Path $Script:searchResultsCsvPath -NoTypeInformation -Append -Encoding UTF8
                    }
                }
            }
        }
        Write-Log ([string]::Format("Found {0} items in the {1} folder.", $Script:folderSearchResults.Count, $MailboxFolder.displayName)) -Level INFO
        if($searchFailures -gt 0){
            Write-Log ([string]::Format("Search for {0} folder in mailbox {1} had {2} failures.", $MailboxFolder.displayName, $mailboxName, $searchFailures)) -Level WARN
        }
        
        #Delete items now to ensure correct mailbox using batches
        if($DeleteContent -and $Script:folderSearchResults.count -gt 0){
            if($searchFailures -gt 0){
                Write-Log "Warning: There were $searchFailures errors during the search process." -Level WARN
                $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
                $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
                $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                $confirmation = $host.ui.PromptForChoice("Search errors detected", "Errors occurred during search. Do you still want to delete the items found?", $options, 1)
                if($confirmation -eq 1){
                    Write-Log "User chose not to delete items due to search errors." -Level INFO
                    continue
                }
                $ConfirmDelete = $false
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
            if($confirmation -eq 0 -or $ConfirmDelete -eq $false){
                #User has confirmed to continue with the delete, so proceed with the delete operation and don't prompt again
                Write-Log "Deleting $($Script:folderSearchResults.Count) items from $($MailboxFolder.displayName)..." -Level WARN
                [int]$Script:itemsDeleted = 0
                [int]$Script:itemsFailedDelete = 0
                [int]$Script:itemsRetry = 0
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
                    $requests = New-Object System.Collections.ArrayList
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
                        $requests.Add($request) | Out-Null
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
        
                        #Check the responses from the batch for any failures
                        if($batchDeleteResponse.Successful -eq $false){
                            Write-Log "Batch request to delete items failed." -Level WARN
                            Write-Log "Error: $($batchDeleteResponse.ErrorMessage)" -Level WARN
                            $Script:itemsFailedDelete = $Script:itemsFailedDelete + $currentBatchSize
                            #Entire batch failed, so log all items in the batch as failed to delete
                            $deleteFailed = foreach($entry in $itemIdLookup.GetEnumerator()) { 
                            [PSCustomObject]@{ 
                                Mailbox=$entry.value.mailbox
                                Id=$entry.value.id
                                Folder=$entry.value.folder
                                Subject=$entry.value.subject
                                ReceivedDateTime=$entry.value.receivedDateTime
                                From=$entry.value.from
                                Attachment=$entry.value.attachment
                                DeleteStatus='Failed'
                            }
                            }
                            $deleteFailed | Export-Csv -Path $Script:deleteResultsCsvPath -NoTypeInformation -Append -Encoding UTF8
                        }
                        else{
                            #Setup for list of requests to retry if any of the responses return a 429 status code
                            $retryRequests = [System.Collections.Generic.List[object]]::new()
                            $retryAfterSeconds = 0

                            $deleteResults = [System.Collections.Generic.List[object]]::new()
                            #Check the response for each delete request
                            foreach($response in $batchDeleteResponse.Content.Responses){
                                $result = ConvertTo-DeleteResult -Response $response -currentBatch $itemIdLookup
                                if($result.StatusCode -eq 429 -and $attempt -lt $maxAttempts){
                                    #If the request failed with a 429 status code, add it to the list of requests to retry and determine the delay before retrying based on the Retry-After header in the response.
                                    $requestToRetry = $pendingRequests |  Where-Object { [string]$_.Id -eq [string]$response.Id } | Select-Object -First 1
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
                            $deleteResults | Export-Csv -Path $Script:deleteResultsCsvPath -NoTypeInformation -Append -Encoding UTF8
                            Write-Log (
                                "Batch delete completed: {0} responses, {1} succeeded, {2} failed, {3} retried." -f
                                    $deleteResults.Count,
                                    @($deleteResults.Where({ $_.DeleteStatus -eq "Succeeded" })).Count,
                                    @($deleteResults.Where({ $_.DeleteStatus -eq "Failed" })).Count,
                                    @($deleteResults.Where({ $_.DeleteStatus -eq "Retry" })).Count
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
                Write-Log "Delete complete for folder $($MailboxFolder.displayName). $itemsProcessed processed; $Script:itemsDeleted succeeded; ($Script:itemsFailedDelete + $Script:itemsRetried) failed; " -Level INFO
            }
        }
    }   
}    
function CreateSearchQuery {
    #Use filter if the message body is not specified, otherwise use search
            #Check if the subject is specified and build the filter query accordingly
            if(-not([string]::IsNullOrEmpty($Subject))) {
                $Subject = $Subject.Replace("'", "''")
                $Subject = [Uri]::EscapeDataString($Subject)
                $UriFilter = "`$filter=contains(subject,`'$Subject`')"
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
                $Sender = $Sender.Replace("'", "''")
                $Sender = [Uri]::EscapeDataString($Sender)
                if($UriFilter -like '*filter*'){
                    $UriFilter = $UriFilter.Replace('filter=', "filter=(from/emailAddress/address) eq `'$Sender`' and ")
                }
                else {
                    $UriFilter = "`$filter=(from/emailAddress/address) eq `'$Sender`'"
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
    Write-Log "Getting a list of folders in the mailbox..." -Level INFO
    #Create an arraylist to hold the folder results
    $Script:folderList = New-Object System.Collections.ArrayList
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
        $Script:folderList.Add($Result) | Out-Null
    }   
    while($null -ne $Script:FolderResults.Content.'@odata.nextLink'){
        $Query = [string]$Script:FolderResults.Content.'@odata.nextLink'
        $Script:FolderResults = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -Query $Query
        foreach($Result in $Script:FolderResults.Content.Value){
            $Script:folderList.Add($Result) | Out-Null
        }
    }
    
    Write-Log "Folder enumeration complete. $($Script:folderList.Count) folders found." -Level INFO
    BuildFolderListTree
}

function BuildFolderListTree{
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
    Write-Log "Getting a list of folders in the recoverable items..." -Level INFO
    #Create an arraylist to hold the folder results
    $Script:folderList = New-Object System.Collections.ArrayList
    [string]$Query = "users/$($Script:userMailbox)/mailFolders/RecoverableItemsRoot/childfolders/?includeHiddenFolders=true"

    $params = @{
        GraphApiUrl         = $cloudService.graphApiEndpoint
        Query               = $Query
    }
    $Script:FolderResults = Invoke-GraphApiRequest @params
    #Check for errors in the folder enumeration request and log them if found, then exit the script
    if($Script:FolderResults.Successful -eq $false){
        Write-Log "Unable to get a list of folders in the recoverable items. Please review the error message below and re-run the script:" -Level ERROR
        Write-Log $Script:FolderResults.ErrorMessage -Level ERROR
        exit 1
    }
    foreach($Result in $Script:FolderResults.Content.Value){
        $Script:folderList.Add($Result) | Out-Null
    }    
    while($null -ne $Script:FolderResults.Content.'@odata.nextLink'){
        $Query = [string]$Script:FolderResults.Content.'@odata.nextLink'
        $Script:FolderResults = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -Query $Query
        foreach($Result in $Script:FolderResults.Content.Value){
            $Script:folderList.Add($Result) | Out-Null
        }
    }

    #Get subfolders for each folder in the recoverable items
    $subfolderList = New-Object System.Collections.ArrayList
    foreach($folder in $Script:folderList){
        $Query = "users/$($Script:userMailbox)/mailFolders/$($folder.id)/childfolders/?includeHiddenFolders=true"
        $params = @{
            GraphApiUrl         = $cloudService.graphApiEndpoint
            Query               = $Query
        }
        $FolderResults = Invoke-GraphApiRequest @params

        foreach($Result in $FolderResults.Content.Value){
            $subfolderList.Add($Result) | Out-Null
        }    
        while($null -ne $FolderResults.Content.'@odata.nextLink'){
            $Query = [string]$FolderResults.Content.'@odata.nextLink'
            $FolderResults = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -Query $Query
            foreach($Result in $FolderResults.Content.Value){
                $subfolderList.Add($Result) | Out-Null
            }
        }
    }
    foreach($subfolder in $subfolderList){
        $Script:folderList.Add($subfolder) | Out-Null
    }
    BuildFolderListTree
    Write-Log "Recoverable items enumeration complete. $($Script:folderList.Count) folders found." -Level INFO
}

#Safety check to ensure the search is not against the entire mailbox and ConfirmDelete is set to false which would result in all items being deleted from the mailbox
if((([string]::IsNullOrEmpty($IncludeFolderList)) -and ([string]::IsNullOrEmpty($ExcludeFolderList)) -and $ConfirmDelete -eq $false) -and $DeleteContent){
    Write-Log "Both IncludeFolderList and ExcludeFolderList are not specified and ConfirmDelete is set to false. This could result in all items being deleted from the mailbox. Please review the parameters and try again." -Level ERROR
    exit 0
}

#Peformance check to ensure MessageBody filter is no the only filter specified since this will result in a full mailbox search and could take a long time to complete
if(([string]::IsNullOrWhiteSpace($Subject) -and [string]::IsNullOrEmpty($ReceivedBefore) -and [string]::IsNullOrEmpty($ReceivedAfter) -and [string]::IsNullOrEmpty($Sender) -and [string]::IsNullOrWhiteSpace($AttachmentName)) -and (-not [string]::IsNullOrWhiteSpace($MessageBody))){
    Write-Log "MessageBody is the only filter specified. This will result in a full mailbox search and could take a long time to complete. Please review the parameters and try again." -Level WARN
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
        exit 0
    }
}

#Ensure the IncludeFolderList and ExcludeFolderList values all start with a backslash to match the folder display names
if(-not([string]::IsNullOrEmpty($IncludeFolderList))){
    # Convert to array if only one entry provided as a string
    if($IncludeFolderList -isnot [System.Collections.ArrayList] -and $IncludeFolderList -isnot [array]){
        $IncludeFolderList = @($IncludeFolderList)
    }
    foreach($folder in $IncludeFolderList){
        if(-not($folder.StartsWith("\"))){
            $IncludeFolderList[$IncludeFolderList.IndexOf($folder)] = "\" + $folder
        }
        if($folder.EndsWith("\")){
            $IncludeFolderList[$IncludeFolderList.IndexOf($folder)] = $folder.TrimEnd("\")
        }
    }
}
if(-not([string]::IsNullOrEmpty($ExcludeFolderList))){
    # Convert to array if only one entry provided as a string
    if($ExcludeFolderList -isnot [System.Collections.ArrayList] -and $ExcludeFolderList -isnot [array]){
        $ExcludeFolderList = @($ExcludeFolderList)
    }
    foreach($folder in $ExcludeFolderList){
        if(-not($folder.StartsWith("\"))){
            $ExcludeFolderList[$ExcludeFolderList.IndexOf($folder)] = "\" + $folder
        }
        if($folder.EndsWith("\")){
            $ExcludeFolderList[$ExcludeFolderList.IndexOf($folder)] = $folder.TrimEnd("\")
        }
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
$Script:searchFolders = New-Object System.Collections.ArrayList
#Check is specific folders are specified in the include list, if not, search all folders under the root
if([string]::IsNullOrEmpty($IncludeFolderList)){
    #If no include list is specified, search all folders under the desired root
    $Script:searchFolders = $Script:folderListTree
    if(-not($ExcludeFolderList)){
        Write-Log "No include list specified. Searching all folders ($($Script:searchFolders.Count) folders)..." -Level WARN
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $confirmation = $host.ui.PromptForChoice("Confirmation", "Do you want to continue?", $options, 1)
        #Confirm with the user that they want to continue with the search since all folders will be searched
        if($confirmation -ne 0){
            Write-Log "Search cancelled by user." -Level WARN
            exit
        }
    }
}
else {
    Write-Log "Building folder search list from include list..." -Level INFO
    #Add folders that match the include list
    if($ProcessSubfolders){
        foreach($folder in $IncludeFolderList){
            #Add all subfolders of the specified folder to the search list
            $includeFolders = ($Script:folderListTree | Where-Object { $_.displayName -match "^" + [regex]::Escape($folder) + "($|\\)"})
            foreach($iFolder in $includeFolders){
                $Script:searchFolders.Add($iFolder) | Out-Null
            }
        }
    }
    else {
        #Add only the specified folders to the search list
        foreach($folder in $IncludeFolderList){
            [void]$Script:searchFolders.Add(($Script:folderListTree | Where-Object { $_.displayName -eq $folder }))
            if($Archive){
                $subfolders = ($Script:folderListTree | Where-Object { $_.displayName -match "^" + [regex]::Escape($folder) + "($|\\)"})
                #Include any subfolders that were created by the aux archive mailbox
                $auxSubfolders = $subfolders | Where-Object {
                    $parts = $_ -split '\\'
                    $rootFolder = $parts[-2]
                    $lastPart = $parts[-1]
                    # Check if last part matches pattern: $rootFolder\$rootFolder_YYYY (Created on
                    $lastPart -match "^$([regex]::Escape($rootFolder))_\d{4}\s+\(Created on"
                }
                foreach($subfolder in $auxSubfolders){
                    $Script:searchFolders.Add($subfolder) | Out-Null
                }

            }
        }
    }
}

#Remove folders that match the exclude list and includes all subfolders
$removeFolderList = New-Object System.Collections.ArrayList
if($ExcludeFolderList){
    Write-Log "Removing excluded folders from the list..." -Level INFO
    #Find all folders that match the exclude list and add them to the remove list
    foreach ($exclude in $ExcludeFolderList) {
        [void]$removeFolderList.Add(($Script:searchFolders | Where-Object { $_.displayName -eq $exclude}))
    }
    #Remove the excluded folders and all subfolders from the search list
    if($removeFolderList.Count -gt 0){
        $newSearchFolders = New-Object System.Collections.ArrayList
        foreach($folder in $Script:searchFolders){
            $shouldExclude = $false
            foreach($exclude in $ExcludeFolderList){
                if($folder.displayName.StartsWith("$exclude\",[StringComparison]::OrdinalIgnoreCase) -or $folder.displayName -eq $exclude){
                    $shouldExclude = $true
                    break
                }
            }
            if(-not $shouldExclude){
                $newSearchFolders.Add($folder) | Out-Null
            }
        }
        $Script:searchFolders = $newSearchFolders
        $newSearchFolders = $null
    }
}

Write-Log "Final list of folders to be searched ($($Script:searchFolders.Count) folders):" -Level INFO
$Script:searchFolders | ForEach-Object { Write-Log "  $($_.displayName)" -Level DEBUG }
$Script:searchFolders | Format-Table displayName

#Create csv file to hold the search results
$Script:searchResultsCsvPath = "$($OutputPath)\SearchResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
Write-Log "Search results will be saved to: $($Script:searchResultsCsvPath)" -Level INFO
if($DeleteContent){
    $Script:deleteResultsCsvPath = "$($OutputPath)\DeleteResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    Write-Log "Delete results will be saved to: $($Script:deleteResultsCsvPath)" -Level INFO
}
#Initiate the search against the mailbox using the specified query and parameters
SearchMailbox -uriQuery "/users/$Script:userMailbox/mailFolders"

Write-Log "Search complete. $Script:TotalSearchResults item(s) found in total." -Level INFO
#Export the search results to a CSV file
Write-Log "Results exported to: $($Script:searchResultsCsvPath)" -Level INFO
Write-Log "Script completed. Log file: $($Script:LogFile)" -Level INFO
}
finally{
    Close-Log
}