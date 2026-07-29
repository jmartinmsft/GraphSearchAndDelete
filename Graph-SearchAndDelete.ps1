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

# Version 20260729.0838

param (
    [Parameter(Position=0,Mandatory=$false,HelpMessage="The Mailbox parameter specifies the mailbox to be accessed.")]
    [ValidateNotNullOrEmpty()] 
    [string]$Mailbox,

    [Parameter(Mandatory=$False, HelpMessage="The Archive parameter is a switch to search the archive mailbox (otherwise, the main mailbox is searched).")]
    [alias("SearchArchive")] [switch]$Archive,

    [Parameter(Mandatory=$False, HelpMessage="The ProcessSubfolders parameter is a switch to enable searching the subfolders of any specified folder.")]
    [switch]$ProcessSubfolders,

    [Parameter(Mandatory=$False, HelpMessage="The IncludeFolderList parameter specifies the folder(s) to be searched (if not present, then the Inbox folder will be searched).  Any exclusions override this list.")]
    $IncludeFolderList,

    [Parameter(Mandatory=$False, HelpMessage="The ExcludeFolderList parameter specifies the folder(s) to be excluded (these folders will not be searched).")]
    $ExcludeFolderList,

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

    [Parameter(Mandatory=$false, HelpMessage="The CreatedBefore parameter specifies only messages created before this date will be searched.")] 
    [DateTime]$CreatedBefore,
    
    [Parameter(Mandatory=$false, HelpMessage="The CreatedAfter parameter specifies only messages created after this date will be searched.")] 
    [DateTime]$CreatedAfter,
    
    [Parameter(Mandatory=$False,HelpMessage="The Subject parameter specifies the subject string used by the search.")] 
    [string]$Subject=$null,
    
    [Parameter(Mandatory=$False,HelpMessage="The Sender parameter specifies the sender email address used by the search.")] 
    [string]$Sender=$null,

    [Parameter(Mandatory=$False,HelpMessage="The MessageBody parameter specifies the body string used by the search.")] 
    [string]$MessageBody=$null,

    [Parameter(Mandatory=$False,HelpMessage="The DeleteContent parameter is a switch to delete the content found by the search. If not specified, the script will only report the number of items that would be deleted.")]
    [switch]$DeleteContent,
    
    [Parameter(Mandatory=$False,HelpMessage="The HardDelete parameter is a switch to permanently delete the content found by the search. If not specified, the script will soft delete the items (move to Deleted Items folder).")]
    [switch]$HardDelete,

    [Parameter(Mandatory = $false, HelpMessage="The BatchSize parameter specifies the number of items to process in each batch.")]
    [ValidateRange(1, 20)]
    [int]$BatchSize = 20,

    [ValidateScript({ Test-Path $_ })]
    [Parameter(Mandatory = $false, HelpMessage="The OutputPath parameter specifies the path for the EWS usage report.")]
    [string] $OutputPath,

    [Parameter(Mandatory = $false, HelpMessage="The LogFile parameter specifies the full path for the script log file. If not specified, a log file is created in the OutputPath.")]
    [string] $LogFile
)

#region Logging
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DEBUG")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"
    if ($Script:LogFile) {
        Add-Content -Path $Script:LogFile -Value $entry
    }
    switch ($Level) {
        "ERROR" { Write-Host $entry -ForegroundColor Red }
        "WARN"  { Write-Host $entry -ForegroundColor Yellow }
        "DEBUG" { Write-Verbose $entry }
        default { Write-Host $entry -ForegroundColor Cyan }
    }
}

# Initialize log file
if ([string]::IsNullOrEmpty($LogFile)) {
    if (-not([string]::IsNullOrEmpty($OutputPath))) {
        $Script:LogFile = Join-Path $OutputPath "GraphSearch_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    } else {
        $Script:LogFile = Join-Path $PSScriptRoot "GraphSearch_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    }
} else {
    $Script:LogFile = $LogFile
}

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
            $body.Add("client_secret", $Secret)
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

function CheckTokenExpiry {
    param(
            $ApplicationInfo,
            [ref]$EWSService,
            [ref]$Token,
            [string]$Environment,
            $EWSOnlineURL,
            $AuthScope,
            $AzureADEndpoint
        )

    # if token is going to expire in next 5 min then refresh it
    if ($null -eq $script:tokenLastRefreshTime -or $script:tokenLastRefreshTime.AddMinutes(55) -lt (Get-Date)) {
        Write-Verbose "Requesting new OAuth token as the current token expires at $($script:tokenLastRefreshTime)."
        if($PermissionType -eq "Application") {
        $createOAuthTokenParams = @{
            TenantID                       = $ApplicationInfo.TenantID
            ClientID                       = $ApplicationInfo.ClientID
            Endpoint                       = $AzureADEndpoint
            CertificateBasedAuthentication = (-not([System.String]::IsNullOrEmpty($ApplicationInfo.CertificateThumbprint)))
            #Scope                          = $AuthScope
            Scope                           = $Script:GraphScope
        }

        # Check if we use an app secret or certificate by using regex to match Json Web Token (JWT)
        if ($ApplicationInfo.AppSecret -match "^([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_\-\+\/=]*)") {
            $jwtParams = @{
                CertificateThumbprint = $ApplicationInfo.CertificateThumbprint
                CertificateStore      = $CertificateStore
                Issuer                = $ApplicationInfo.ClientID
                Audience              = "$AzureADEndpoint/$($ApplicationInfo.TenantID)/oauth2/v2.0/token"
                Subject               = $ApplicationInfo.ClientID
            }
            $jwt = Get-NewJsonWebToken @jwtParams

            if ($null -eq $jwt) {
                Write-Host "Unable to sign a new Json Web Token by using certificate: $($ApplicationInfo.CertificateThumbprint)" -ForegroundColor Red
                exit
            }

            $createOAuthTokenParams.Add("Secret", $jwt)
        } else {
            $createOAuthTokenParams.Add("Secret", $ApplicationInfo.AppSecret)
        }

        $oAuthReturnObject = Get-ApplicationAccessToken @createOAuthTokenParams
        if ($oAuthReturnObject.Successful -eq $false) {
            Write-Host ""
            Write-Host "Unable to refresh EWS OAuth token. Please review the error message below and re-run the script:" -ForegroundColor Red
            Write-Host $oAuthReturnObject.ExceptionMessage -ForegroundColor Red
            exit
        }
        Write-Log "Obtained a new OAuth token (Application)" -Level INFO
        $Script:Token = $oAuthReturnObject.OAuthToken.access_token
        $script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime
        #return $oAuthReturnObject.OAuthToken.access_token
        }
        else {
            #$connectionSuccessful = $false
    
            # Request an authorization code from the Microsoft Azure Active Directory endpoint
            $redeemAuthCodeParams = @{
                Uri             = "$AzureADEndpoint/organizations/oauth2/v2.0/token"
                Method          = "POST"
                ContentType     = "application/x-www-form-urlencoded"
                Body            = @{
                    client_id     = $ApplicationInfo.ClientID
                    scope         = $AuthScope
                    grant_type    = "refresh_token"
                    refresh_token =  $Script:RefreshToken
                }
                UseBasicParsing = $true
            }
            $redeemAuthCodeResponse = Invoke-WebRequestWithProxyDetection -ParametersObject $redeemAuthCodeParams

            if ($redeemAuthCodeResponse.StatusCode -eq 200) {
                $tokens = $redeemAuthCodeResponse.Content | ConvertFrom-Json
                $script:tokenLastRefreshTime = (Get-Date)
                $Script:RefreshToken = $tokens.refresh_token
                $Script:Token = $tokens.access_token
            } 
            else {
                Write-Host "Unable to redeem the authorization code for an access token." -ForegroundColor Red
                exit
            }
        }
    }
    #return $Script:Token
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

        [Parameter(Mandatory = $true)]
        [ValidatePattern("^([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_\-\+\/=]*)")]
        [string]$AccessToken,

        [Parameter(Mandatory = $false)]
        [int]$ExpectedStatusCode = 200,

        [Parameter(Mandatory = $true)]
        [string]$GraphApiUrl
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
        $graphApiRequestParams = @{
            Uri             = "$GraphApiUrl/$Endpoint/$($Query.TrimStart("/"))"
            Header          = @{ Authorization = "Bearer $AccessToken" }
            Method          = $Method
            ContentType     = $ContentType
            UseBasicParsing = $true
            ErrorAction     = "Stop"
        }

        if (-not([System.String]::IsNullOrEmpty($Body))) {
            Write-Log "Body: $Body" -Level DEBUG
            $graphApiRequestParams.Add("Body", $Body)
        }

        Write-Log "Graph API uri called: $($graphApiRequestParams.Uri)" -Level DEBUG
        $Global:graphApiResponse = Invoke-WebRequestWithProxyDetection -ParametersObject $graphApiRequestParams
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

    try {
        Invoke-WebRequest @params
    } 
    catch {
        $response = $_.Exception.Response
        $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
        $responseContent = ($reader.ReadToEnd() | ConvertFrom-Json)
        Write-Log "Response Content: $($responseContent.error.message)" -Level DEBUG
        #Write-VerboseErrorInformation
        return [PSCustomObject]@{
            ErrorCode    = $responseContent.error.code
            ErrorMessage   = $responseContent.error.message
            StatusCode = $response.StatusCode
            Successful = $false
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

function Write-VerboseErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Verbose"
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
        #$Script:GraphScope = "$($cloudService.graphApiEndpoint)/.default"
        $Script:GraphScope = "$($Script:GraphScope).default"
        if ([System.String]::IsNullOrEmpty($OAuthCertificate)) {
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($OAuthClientSecret)
            $Secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            $Script:applicationInfo.Add("AppSecret", $Secret)
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
                exit
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
            #Write-Host ""
            Write-Log "Unable to fetch an OAuth token. Please review the error message below and re-run the script:" -Level ERROR
            Write-Log $oAuthReturnObject.ExceptionMessage -Level ERROR
            exit
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
        #$Script:GraphScope = "$($cloudService.GraphApiEndpoint)//$($Scope)"
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

function SearchMailbox {
    param(
        [string]$uriQuery
    )
    Write-Log "Performing search against the mailbox..." -Level INFO
    $Script:SearchResults = New-Object System.Collections.ArrayList
    #perform the search against each folder
    foreach($MailboxFolder in $Global:searchFolders) {
        $Global:folderSearchResults = New-Object System.Collections.ArrayList
        $Script:currentFolderName = $MailboxFolder.displayName
        Write-Log "Searching folder: $($MailboxFolder.displayName)" -Level INFO
        Write-Log "Processing folder: $($MailboxFolder.id)" -level DEBUG
        $itemsInFolder = 0
        if($Archive){
            #Check to see if the folder is in the main archive or an aux archive
            $Uri = "admin/exchange/mailboxes/$($Script:userMailbox)/folders/$($MailboxFolder.id)"
            $FolderCheckParams = @{
                GraphApiUrl     = $cloudService.graphApiEndpoint
                AccessToken     = $Script:Token
                Endpoint        = "beta"
                Query           = $Uri
            }
            Write-Log "Checking the archive location for folder: $($MailboxFolder.displayName)" -Level DEBUG
            $Global:FolderCheck = Invoke-GraphApiRequest @FolderCheckParams
            if($FolderCheck.StatusCode -eq 308){
                #Modify the URL using the aux archive guid and the folder id for the folder within the aux archive mailbox
                $mailboxGuid = $FolderCheck.Response.ErrorMessage -match 'MBX:([a-f0-9\-]+)@' | ForEach-Object { $matches[1] }
                $folderValue = $FolderCheck.Response.ErrorMessage -match "folders\('([^']+)'\)" | ForEach-Object { $matches[1] }
                #$uriQuery = "/users/MBX:$($mailboxGuid)@9101fc97-5be5-4438-a1d7-83e051e52057/mailFolders/$($folderValue)"
                Write-Log "Checking auxiliary archive mailbox $($mailboxGuid) for items in $($MailboxFolder.displayName)" -Level DEBUG
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
        
        if(-not($UriFilter)) {
            $UriFilter = CreateSearchQuery
        }

        # Finalize the Uri with the final filter/search settings
        Write-Log ([string]::Format("Performing query using: {0}", $Uri)) -Level DEBUG
        
        # Search the mailbox for items
        $SearchParams = @{
            GraphApiUrl     = $cloudService.graphApiEndpoint
            Query           =  "$($Uri)?$UriFilter"
            AccessToken     = $Script:Token
            Endpoint        = "beta"
        }
            
        $SearchItems = Invoke-GraphApiRequest @SearchParams
        foreach($Result in $SearchItems.Content.Value){
            $Global:folderSearchResults.Add([PSCustomObject]@{mailbox=$mailboxName;id=$Result.id; folder=$MailboxFolder.displayName.Split('\')[-1]; internetMessageId=$Result.internetMessageId;subject=$Result.subject;receivedDateTime=$Result.receivedDateTime;from=$Result.from.emailaddress.address}) | Out-Null
            #$itemsInFolder++
        }
        if($global:folderSearchResults.count -gt 0){
            #exit
        }
        while($null -ne $SearchItems.Content.'@odata.nextLink'){
            $Query = $SearchItems.Content.'@odata.nextLink'.Substring($SearchItems.Content.'@odata.nextLink'.IndexOf("user"))
            $SearchItems = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -AccessToken $Script:Token -Query $Query
            foreach($Result in $SearchItems.Content.Value){
                $Global:folderSearchResults.Add([PSCustomObject]@{mailbox=$mailboxName;id=$Result.id; folder=$MailboxFolder.displayName.Split('\')[-1]; internetMessageId=$Result.internetMessageId;subject=$Result.subject;receivedDateTime=$Result.receivedDateTime;from=$Result.from.emailaddress.address}) | Out-Null
                #$itemsInFolder++
            }
        }
        Write-Log ([string]::Format("Found {0} items in the {1} folder.", $Global:folderSearchResults.Count, $MailboxFolder.displayName)) -Level INFO
        #Add the folder results to the total list of results
        $Script:SearchResults.AddRange($Global:folderSearchResults)

        #Delete items now to ensure correct mailbox using batches
        if($DeleteContent -and $global:folderSearchResults.count -gt 0){
            Write-Log "Deleting $($Global:folderSearchResults.Count) items from $($MailboxFolder.displayName)..." -Level WARN
            [int]$itemsDeleted = 0
            # Make sure the results are not less than the batch size
            if($Global:folderSearchResults.count -lt $BatchSize){
                $BatchSize = $Global:folderSearchResults.Count
            }
            $Query = "`$batch"
            # Loop thru the results creating batches to delete
            while($itemsDeleted -lt $Global:folderSearchResults.Count){
                # Make sure the batch size is not greater than the items left to process
                if(($Global:folderSearchResults.Count - $itemsDeleted) -lt $BatchSize){
                    $BatchSize = $Global:folderSearchResults.Count - $itemsDeleted
                }
                #region CreateBatch
                $requests = New-Object System.Collections.ArrayList
                for($x=0; $x -lt $BatchSize; $x++){
                    if($HardDelete){
                        $Method = "POST"
                        $Url = "/users/MBX:$($mailboxName)@$($OAuthTenantId)/messages/$($Global:folderSearchResults[$itemsDeleted].id)/permanentDelete"
                    }
                    else {
                        $Method = "DELETE"
                        $Url = "/users/MBX:$($mailboxName)@$($OAuthTenantId)/messages/$($Global:folderSearchResults[$itemsDeleted].id)"
                    }
                    $request = @{
                        Id          = $x+1
                        Method      = $Method
                        Url         = $Url
                    }
                    $requests.Add($request) | Out-Null
                    $itemsDeleted++
                }
                $batchRequest = @{
                    Requests = $requests
                } | ConvertTo-Json -Depth 6
                #endregion
                Write-Log "Sending batch delete request ($BatchSize items, total deleted so far: $itemsDeleted)" -Level DEBUG
                Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -Query $Query -AccessToken $Script:Token -Method POST -Body $batchRequest | Out-Null
            }
            Write-Log "Delete complete for folder $($MailboxFolder.displayName). $itemsDeleted items processed." -Level INFO
        }
    }      
}    
function CreateSearchQuery {
        if([string]::IsNullOrEmpty($MessageBody)) {
            if(-not([string]::IsNullOrEmpty($Subject))) {
                $UriFilter = "filter=contains(subject,`'$Subject`')&`$top=500"
            }
            if(-not([string]::IsNullOrEmpty($CreatedBefore))) {
                $TempStartDate = [datetime]$CreatedBefore
                $TempStartDate = $TempStartDate.ToUniversalTime()
                $SearchStartDate = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $TempStartDate
                if($UriFilter -like '*filter*'){
                    $UriFilter = $UriFilter.Replace('filter=', "filter=receivedDateTime le $($SearchStartDate) and ")
                }
                else {
                    $UriFilter = "filter=receivedDateTime le $($SearchStartDate)&`$top=500"
                }
            }
            if(-not([string]::IsNullOrEmpty($CreatedAfter))){
                $TempEndDate = [datetime]$CreatedAfter
                $TempEndDate = $TempEndDate.ToUniversalTime()
                $SearchEndDate = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $TempEndDate
                if($UriFilter -like '*filter*'){
                    $UriFilter = $UriFilter.Replace('filter=', "filter=receivedDateTime ge $($SearchEndDate) and ")
                }
                else {
                    $UriFilter = "filter=receivedDateTime ge $($SearchEndDate)&`$top=500"
                }
            }
            if(-not([string]::IsNullOrEmpty($Sender))){
                if($UriFilter -like '*filter*'){
                    $UriFilter = $UriFilter.Replace('filter=', "filter=(from/emailAddress/address) eq `'$Sender`' and ")
                }
                else {
                    $UriFilter = "filter=(from/emailAddress/address) eq `'$Sender`'&`$top=500"
                }
            }
        }
        else {
            # Build the search query based on specified parameters
            Write-Log "Creating a query using the search function." -Level DEBUG
            $UriFilter = "`$search=`"body:$MessageBody`"&`$top=25"

            if(-not([string]::IsNullOrEmpty($Sender))){
                if($UriFilter -like '*search*'){
                    $UriFilter = $UriFilter.Replace('search="', "search=`"from:$Sender` AND ")
                }
                else{
                    $UriFilter = "`$search=`"from:$Sender`"&`$top=25"
                }
            }
            if(-not([string]::IsNullOrEmpty($Subject))){
                if($UriFilter -like '*search*'){
                    $UriFilter = $UriFilter.Replace('search="', "search=`"subject:$Subject` AND ")
                }
                else{
                    $UriFilter = "`$search=`"subject:$Subject`"&`$top=1000&`$select=id,parentfolderid,receivedDateTime,subject,from"#&`$from=$PageNumber"
                }
            }
            if(-not([string]::IsNullOrEmpty($CreatedBefore))){
                $TempStartDate = [datetime]$CreatedBefore
                $TempStartDate = $TempStartDate.ToUniversalTime()
                $SearchBeforeDate = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $TempStartDate
                if($UriFilter -like '*search*'){
                    $UriFilter = $UriFilter.Replace('search="', "search=`"received<=$SearchBeforeDate AND ")
                }
                else{
                    $UriFilter = "`$search=`"received<=$SearchBeforeDate`"&`$top=25"
                }
            }
            if(-not([string]::IsNullOrEmpty($CreatedAfter))){
                $TempStartDate = [datetime]$CreatedAfter
                $TempStartDate = $TempStartDate.ToUniversalTime()
                $SearchAfterDate = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $TempStartDate
                if($UriFilter -like '*search*'){
                    $UriFilter = $UriFilter.Replace('search="', "search=`"received>=$SearchAfterDate AND ")
                }
                else{
                    $UriFilter = "`$search=`"received>=$SearchAfterDate`"&`$top=25"
                }
            }
        }
        return $UriFilter    
}
    
function GetFolderList{
    Write-Log "Getting a list of folders in the mailbox..." -Level INFO
    $Global:folderList = New-Object System.Collections.ArrayList
    [string]$Query = "users/$($Script:userMailbox)/mailFolders/delta"
    $params = @{
        AccessToken         = $Script:Token
        GraphApiUrl         = $cloudService.graphApiEndpoint
        Query               = $Query
    }

    $Global:FolderResults = Invoke-GraphApiRequest @params

    foreach($Result in $FolderResults.Content.Value){
        $Global:folderList.Add($Result) | Out-Null
    }
        
    while($null -ne $FolderResults.Content.'@odata.nextLink'){
        $Query = $FolderResults.Content.'@odata.nextLink'.Substring($FolderResults.Content.'@odata.nextLink'.IndexOf("user"))
        $FolderResults = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -AccessToken $Script:Token -Query $Query
        foreach($Result in $FolderResults.Content.Value){
            $Global:folderList.Add($Result) | Out-Null
        }
    }
    Write-Log "Folder enumeration complete. $($Global:folderList.Count) folders found." -Level INFO
    BuildFolderListTree
}
function BuildFolderListTree{
    Write-Log "Building folder tree..." -Level INFO
    $Script:folderListTree = New-Object System.Collections.ArrayList
    $foundFolders = 0
    
    #first folders under MsgFolderRoot
    $Global:folderList | Where-Object {
        ($_.parentFolderId -notin ($Global:folderList | ForEach-Object { $_.id }))
    } | ForEach-Object { 
        $_.displayName = "\$($_.displayName)"
        $Script:folderListTree.Add($_) | Out-Null
        $foundFolders++
    }
    #loop through the folder list to find all folders under the root
    foreach($folder in $Global:folderlist){
        if($folder.parentFolderId -eq $rootFolderId.id){
            $folder.displayName = "\$($folder.displayName)"
            $Script:folderListTree.Add($folder) | Out-Null
            $foundFolders++
        }
    }

    #loop through until all folders are found
    while($foundFolders -lt $Global:folderList.Count){
        foreach($folder in $Global:folderlist){
            foreach($treeFolder in $folderListTree){
                if($folder.parentFolderId -eq $treeFolder.id){
                    $folder.displayName = "$($treeFolder.displayName)\$($folder.displayName)"
                    $folderListTree.Add($folder) | Out-Null
                    $foundFolders++
                    break
                }
            }
        }
    }
}

function GetRecoverableItemsFolderList{
    Write-Log "Getting a list of folders in the recoverable items..." -Level INFO
    $Global:folderList = New-Object System.Collections.ArrayList
    [string]$Query = "users/$($Script:userMailbox)/mailFolders/RecoverableItemsRoot/childfolders/?includeHiddenFolders=true"

    $params = @{
        AccessToken         = $Script:Token
        GraphApiUrl         = $cloudService.graphApiEndpoint
        Query               = $Query
    }
    $Global:FolderResults = Invoke-GraphApiRequest @params

    foreach($Result in $FolderResults.Content.Value){
        $Global:folderList.Add($Result) | Out-Null
    }    
    while($null -ne $FolderResults.Content.'@odata.nextLink'){
        $Query = $FolderResults.Content.'@odata.nextLink'.Substring($FolderResults.Content.'@odata.nextLink'.IndexOf("user"))
        $FolderResults = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -AccessToken $Script:Token -Query $Query
        foreach($Result in $FolderResults.Content.Value){
            $Global:folderList.Add($Result) | Out-Null
        }
    }

    #get subfolders
    $subfolderList = New-Object System.Collections.ArrayList
    foreach($folder in $Global:folderList){
        $Query = "users/$($Script:userMailbox)/mailFolders/$($folder.id)/childfolders/?includeHiddenFolders=true"
        $params = @{
            AccessToken         = $Script:Token
            GraphApiUrl         = $cloudService.graphApiEndpoint
            Query               = $Query
        }
        $FolderResults = Invoke-GraphApiRequest @params

        foreach($Result in $FolderResults.Content.Value){
            $subfolderList.Add($Result) | Out-Null
        }    
        while($null -ne $FolderResults.Content.'@odata.nextLink'){
            $Query = $FolderResults.Content.'@odata.nextLink'.Substring($FolderResults.Content.'@odata.nextLink'.IndexOf("user"))
            $FolderResults = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -AccessToken $Script:Token -Query $Query
            foreach($Result in $FolderResults.Content.Value){
                $subfolderList.Add($Result) | Out-Null
            }
        }
    }
    foreach($subfolder in $subfolderList){
        $Global:folderList.Add($subfolder) | Out-Null
    }
    BuildFolderListTree
    Write-Log "Recoverable items enumeration complete. $($Global:folderList.Count) folders found." -Level INFO
}

#Get parameters and pass to obtain an OAuth token
$cloudService = Get-CloudServiceEndpoint $AzureEnvironment
$azureADEndpoint = $cloudService.AzureADEndpoint
$Script:applicationInfo = @{
    "TenantID" = $OAuthTenantId
    "ClientID" = $OAuthClientId
}
$Script:GraphScope = "$($cloudService.GraphApiEndpoint)/"
Get-OAuthToken -AppScope $Scope

[string]$Endpoint = "beta"

#Check if archive mailbox and obtain mailbox connection settings

    
[string]$Query = "users/$Mailbox/settings/exchange"
$params = @{
    AccessToken         = $Script:Token
    GraphApiUrl         = $cloudService.graphApiEndpoint
    Query               = $Query
    Endpoint            = $Endpoint
}
$Global:MailboxSettings = Invoke-GraphApiRequest @params
$Script:archiveMailbox = $MailboxSettings.Content.inPlaceArchiveMailboxId
$Script:primaryMailbox = $MailboxSettings.content.primaryMailboxId

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

#region BuildSearchFolderList
#Determine the folder to search based on include/exclude lists
Write-Log "Determining folders to search..." -Level INFO
$Global:searchFolders = New-Object System.Collections.ArrayList

if([string]::IsNullOrEmpty($IncludeFolderList)){
    #If no include list is specified, search all folders under the desired root
    $Global:searchFolders = $Script:folderListTree
}
else {
    Write-Log "Building folder search list from include list..." -Level INFO
    #Add folders that match the include list
    if($ProcessSubfolders){
        foreach($folder in $IncludeFolderList){
            $Global:searchFolders.Add(($Script:folderListTree | Where-Object { $_.displayName.split('\')[-1] -eq $folder })) | Out-Null
            $subfolders = ($Script:folderListTree | Where-Object { $_ -match "\\$($folder)($|\\)" })
            foreach($subfolder in $subfolders){
                $Global:searchFolders.Add($subfolder) | out-null
            }
        }
    }
    else {
        foreach($folder in $IncludeFolderList){
            $Global:searchFolders.Add(($Script:folderListTree | Where-Object { $_.displayName.split('\')[-1] -eq $folder })) | Out-Null
            if($Archive){
                $subfolders = ($Script:folderListTree | Where-Object { $_ -match "\\$($folder)($|\\)" })
                $auxSubfolders = $subfolders | Where-Object {
                    $parts = $_ -split '\\'
                    $rootFolder = $parts[-2]
                    $lastPart = $parts[-1]
                    # Check if last part matches pattern: $rootFolder\$rootFolder_YYYY (Created on
                    $lastPart -match "^$([regex]::Escape($rootFolder))_\d{4}\s+\(Created on"
                }
                foreach($subfolder in $auxSubfolders){
                    $Global:searchFolders.Add($subfolder) | out-null
                }

            }
        }
    }
}

#Remove folders that match the exclude list and includes all subfolders
$removeFolderList = New-Object System.Collections.ArrayList
if($ExcludeFolderList){
    Write-Log "Removing excluded folders from the list..." -Level INFO
    foreach ($exclude in $ExcludeFolderList) {
        $removeFolderList.Add(($Global:searchFolders | Where-Object { $_.displayName.split('\')[-1] -eq $exclude})) | out-Null
    }
    if($removeFolderList.Count -gt 0){
        foreach($r in $removeFolderList){
            $Global:searchFolders.Remove($r) | Out-Null
        }
    }
}

Write-Log "Final list of folders to be searched ($($Global:searchFolders.Count) folders):" -Level INFO
$Global:searchFolders | ForEach-Object { Write-Log "  $($_.displayName)" -Level DEBUG }
$Global:searchFolders | Format-Table displayName
#endregion

#Initiate the search
#$Global:testFolderAccess = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -Query "/admin/exchange/mailboxes/$Script:userMailbox/mailFolders" -AccessToken $Script:Token -Method GET -Endpoint beta

$Script:searchResultsCsvPath = "$($OutputPath)\SearchResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
Write-Log "Search results will be saved to: $($Script:searchResultsCsvPath)" -Level INFO
SearchMailbox -uriQuery "/users/$Script:userMailbox/mailFolders"

Write-Log ([string]::Format("Search complete. {0} item(s) found in total.", $Script:SearchResults.Count)) -Level INFO
$Script:SearchResults | Export-Csv -Path $Script:searchResultsCsvPath -NoTypeInformation -Append
Write-Log "Results exported to: $($Script:searchResultsCsvPath)" -Level INFO
Write-Log "Script completed. Log file: $($Script:LogFile)" -Level INFO
