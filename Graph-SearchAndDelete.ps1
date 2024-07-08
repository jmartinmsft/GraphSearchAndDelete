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
# Version 24.07.08.1516

param (
    [Parameter(Position=0,Mandatory=$True,HelpMessage="The Mailbox parameter specifies the mailbox to be accessed.")]
    [ValidateNotNullOrEmpty()] 
    [string]$Mailbox,
	
    [Parameter(Mandatory=$False,HelpMessage="The ProcessSubfolders parameter is a switch to enable searching the subfolders of any specified folder.")] 
    [switch]$ProcessSubfolders,
	
    [Parameter(Mandatory=$False,HelpMessage="The IncludeFolderList parameter specifies the folder(s) to be searched (if not present, then the Inbox folder will be searched).  Any exclusions override this list.")] 
    $IncludeFolderList=$null,
    
    [Parameter(Mandatory=$False,HelpMessage="The ExcludeFolderList parameter specifies the folder(s) to be excluded (these folders will not be searched).")] 
    $ExcludeFolderList=$null,

    [Parameter(Mandatory=$false,HelpMessage="The SearchDumpster parameter is a switch to search the recoverable items.")] 
    [switch]$SearchDumpster,
    
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
    
    [Parameter(Mandatory=$False,HelpMessage="The MessageId parameter specified the MessageId used by the search.")] 
    [string]$MessageId,    
    
    [Parameter(Mandatory=$False,HelpMessage="The DeleteContent parameter is a switch to delete the items found in the search results (moved to Deleted Items).")]
    [switch]$DeleteContent,
	
    [ValidateSet("Global", "USGovernmentL4", "USGovernmentL5", "ChinaCloud")]
    [Parameter(Mandatory = $false)]
    [string]$AzureEnvironment = "Global",

    [Parameter(Mandatory=$false, HelpMessage="The PermissionType parameter specifies whether the app registrations uses delegated or application permissions")] [ValidateSet('Application','Delegated')]
    [string]$PermissionType,
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthClientId parameter is the Azure Application Id that this script uses to obtain the OAuth token.  Must be registered in Azure AD.")] 
    [string]$OAuthClientId = "",
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthTenantId parameter is the tenant Id where the application is registered (Must be in the same tenant as mailbox being accessed).")] 
    [string]$OAuthTenantId = "",
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthRedirectUri parameter is the redirect Uri of the Azure registered application.")] 
    [string]$OAuthRedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient",
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthClientSecret parameter is the the secret for the registered application.")] 
    [SecureString]$OAuthClientSecret = "",
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthCertificate parameter is the certificate for the registered application. Certificate auth requires MSAL libraries to be available.")] 
    [string]$OAuthCertificate = $null,
  
    [Parameter(Mandatory=$False,HelpMessage="The CertificateStore parameter specifies the certificate store where the certificate is loaded.")] [ValidateSet("CurrentUser", "LocalMachine")]
     [string] $CertificateStore = $null,
    
    [ValidateScript({ Test-Path $_ })] [Parameter(Mandatory = $true, HelpMessage="The OutputPath parameter specifies the path for the EWS usage report.")] [string] $OutputPath,

    [Parameter(Mandatory=$False,HelpMessage="The ThrottlingDelay parameter specifies the throttling delay (time paused between sending EWS requests) - note that this will be increased automatically if throttling is detected")]
    [int]$ThrottlingDelay = 0
)

begin {
    function Write-VerboseLog ($Message) {
        $Script:Logger = $Script:Logger | Write-LoggerInstance $Message
    }

    function Write-HostLog ($Message) {
        $Script:Logger = $Script:Logger | Write-LoggerInstance $Message
    }

    function Write-Host {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Proper handling of write host with colors')]
        [CmdletBinding()]
        param(
            [Parameter(Position = 1, ValueFromPipeline)]
            [object]$Object,
            [switch]$NoNewLine,
            [string]$ForegroundColor
        )
        process {
            $consoleHost = $host.Name -eq "ConsoleHost"
    
            if ($null -ne $Script:WriteHostManipulateObjectAction) {
                $Object = & $Script:WriteHostManipulateObjectAction $Object
            }
    
            $params = @{
                Object    = $Object
                NoNewLine = $NoNewLine
            }
    
            if ([string]::IsNullOrEmpty($ForegroundColor)) {
                if ($null -ne $host.UI.RawUI.ForegroundColor -and
                    $consoleHost) {
                    $params.Add("ForegroundColor", $host.UI.RawUI.ForegroundColor)
                }
            } elseif ($ForegroundColor -eq "Yellow" -and
                $consoleHost -and
                $null -ne $host.PrivateData.WarningForegroundColor) {
                $params.Add("ForegroundColor", $host.PrivateData.WarningForegroundColor)
            } elseif ($ForegroundColor -eq "Red" -and
                $consoleHost -and
                $null -ne $host.PrivateData.ErrorForegroundColor) {
                $params.Add("ForegroundColor", $host.PrivateData.ErrorForegroundColor)
            } else {
                $params.Add("ForegroundColor", $ForegroundColor)
            }
    
            Microsoft.PowerShell.Utility\Write-Host @params
    
            if ($null -ne $Script:WriteHostDebugAction -and
                $null -ne $Object) {
                &$Script:WriteHostDebugAction $Object
            }
        }
    }
    
    function SetProperForegroundColor {
        $Script:OriginalConsoleForegroundColor = $host.UI.RawUI.ForegroundColor
    
        if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.WarningForegroundColor) {
            Write-Verbose "Foreground Color matches warning's color"
    
            if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
                $Host.UI.RawUI.ForegroundColor = "Gray"
            }
        }
    
        if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.ErrorForegroundColor) {
            Write-Verbose "Foreground Color matches error's color"
    
            if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
                $Host.UI.RawUI.ForegroundColor = "Gray"
            }
        }
    }
    
    function RevertProperForegroundColor {
        $Host.UI.RawUI.ForegroundColor = $Script:OriginalConsoleForegroundColor
    }
    
    function SetWriteHostAction ($DebugAction) {
        $Script:WriteHostDebugAction = $DebugAction
    }
    
    function SetWriteHostManipulateObjectAction ($ManipulateObject) {
        $Script:WriteHostManipulateObjectAction = $ManipulateObject
    }
    
    function Write-Verbose {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'In order to log Write-Verbose from Shared functions')]
        [CmdletBinding()]
        param(
            [Parameter(Position = 1, ValueFromPipeline)]
            [string]$Message
        )
    
        process {
    
            if ($null -ne $Script:WriteVerboseManipulateMessageAction) {
                $Message = & $Script:WriteVerboseManipulateMessageAction $Message
            }
    
            Microsoft.PowerShell.Utility\Write-Verbose $Message
    
            if ($null -ne $Script:WriteVerboseDebugAction) {
                & $Script:WriteVerboseDebugAction $Message
            }
    
            # $PSSenderInfo is set when in a remote context
            if ($PSSenderInfo -and
                $null -ne $Script:WriteRemoteVerboseDebugAction) {
                & $Script:WriteRemoteVerboseDebugAction $Message
            }
        }
    }
    
    function SetWriteVerboseAction ($DebugAction) {
        $Script:WriteVerboseDebugAction = $DebugAction
    }
    
    function SetWriteRemoteVerboseAction ($DebugAction) {
        $Script:WriteRemoteVerboseDebugAction = $DebugAction
    }
    
    function SetWriteVerboseManipulateMessageAction ($DebugAction) {
        $Script:WriteVerboseManipulateMessageAction = $DebugAction
    }
    
    function Write-Warning {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'In order to log Write-Warning from Shared functions')]
        [CmdletBinding()]
        param(
            [Parameter(Position = 1, ValueFromPipeline)]
            [string]$Message
        )
        process {
    
            if ($null -ne $Script:WriteWarningManipulateMessageAction) {
                $Message = & $Script:WriteWarningManipulateMessageAction $Message
            }
    
            Microsoft.PowerShell.Utility\Write-Warning $Message
    
            # Add WARNING to beginning of the message by default.
            $Message = "WARNING: $Message"
    
            if ($null -ne $Script:WriteWarningDebugAction) {
                & $Script:WriteWarningDebugAction $Message
            }
    
            # $PSSenderInfo is set when in a remote context
            if ($PSSenderInfo -and
                $null -ne $Script:WriteRemoteWarningDebugAction) {
                & $Script:WriteRemoteWarningDebugAction $Message
            }
        }
    }
    
    function SetWriteWarningAction ($DebugAction) {
        $Script:WriteWarningDebugAction = $DebugAction
    }
    
    function SetWriteRemoteWarningAction ($DebugAction) {
        $Script:WriteRemoteWarningDebugAction = $DebugAction
    }
    
    function SetWriteWarningManipulateMessageAction ($DebugAction) {
        $Script:WriteWarningManipulateMessageAction = $DebugAction
    }

    function Get-NewLoggerInstance {
        [CmdletBinding()]
        param(
            [string]$LogDirectory = (Get-Location).Path,
    
            [ValidateNotNullOrEmpty()]
            [string]$LogName = "Script_Logging",
    
            [bool]$AppendDateTime = $true,
    
            [bool]$AppendDateTimeToFileName = $true,
    
            [int]$MaxFileSizeMB = 10,
    
            [int]$CheckSizeIntervalMinutes = 10,
    
            [int]$NumberOfLogsToKeep = 10
        )
    
        $fileName = if ($AppendDateTimeToFileName) { "{0}_{1}.txt" -f $LogName, ((Get-Date).ToString('yyyyMMddHHmmss')) } else { "$LogName.txt" }
        $fullFilePath = [System.IO.Path]::Combine($LogDirectory, $fileName)
    
        if (-not (Test-Path $LogDirectory)) {
            try {
                New-Item -ItemType Directory -Path $LogDirectory -ErrorAction Stop | Out-Null
            } catch {
                throw "Failed to create Log Directory: $LogDirectory. Inner Exception: $_"
            }
        }
    
        return [PSCustomObject]@{
            FullPath                 = $fullFilePath
            AppendDateTime           = $AppendDateTime
            MaxFileSizeMB            = $MaxFileSizeMB
            CheckSizeIntervalMinutes = $CheckSizeIntervalMinutes
            NumberOfLogsToKeep       = $NumberOfLogsToKeep
            BaseInstanceFileName     = $fileName.Replace(".txt", "")
            Instance                 = 1
            NextFileCheckTime        = ((Get-Date).AddMinutes($CheckSizeIntervalMinutes))
            PreventLogCleanup        = $false
            LoggerDisabled           = $false
        } | Write-LoggerInstance -Object "Starting Logger Instance $(Get-Date)"
    }
    
    function Write-LoggerInstance {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
            [object]$LoggerInstance,
    
            [Parameter(Mandatory = $true, Position = 1)]
            [object]$Object
        )
        process {
            if ($LoggerInstance.LoggerDisabled) { return }
    
            if ($LoggerInstance.AppendDateTime -and
                $Object.GetType().Name -eq "string") {
                $Object = "[$([System.DateTime]::Now)] : $Object"
            }
    
            # Doing WhatIf:$false to support -WhatIf in main scripts but still log the information
            $Object | Out-File $LoggerInstance.FullPath -Append -WhatIf:$false
    
            #Upkeep of the logger information
            if ($LoggerInstance.NextFileCheckTime -gt [System.DateTime]::Now) {
                return
            }
    
            #Set next update time to avoid issues so we can log things
            $LoggerInstance.NextFileCheckTime = ([System.DateTime]::Now).AddMinutes($LoggerInstance.CheckSizeIntervalMinutes)
            $item = Get-ChildItem $LoggerInstance.FullPath
    
            if (($item.Length / 1MB) -gt $LoggerInstance.MaxFileSizeMB) {
                $LoggerInstance | Write-LoggerInstance -Object "Max file size reached rolling over" | Out-Null
                $directory = [System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)
                $fileName = "$($LoggerInstance.BaseInstanceFileName)-$($LoggerInstance.Instance).txt"
                $LoggerInstance.Instance++
                $LoggerInstance.FullPath = [System.IO.Path]::Combine($directory, $fileName)
    
                $items = Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)) -Filter "*$($LoggerInstance.BaseInstanceFileName)*"
    
                if ($items.Count -gt $LoggerInstance.NumberOfLogsToKeep) {
                    $item = $items | Sort-Object LastWriteTime | Select-Object -First 1
                    $LoggerInstance | Write-LoggerInstance "Removing Log File $($item.FullName)" | Out-Null
                    $item | Remove-Item -Force
                }
            }
        }
        end {
            return $LoggerInstance
        }
    }
    
    function Invoke-LoggerInstanceCleanup {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
            [object]$LoggerInstance
        )
        process {
            if ($LoggerInstance.LoggerDisabled -or
                $LoggerInstance.PreventLogCleanup) {
                return
            }
    
            Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)) -Filter "*$($LoggerInstance.BaseInstanceFileName)*" |
                Remove-Item -Force
        }
    }

    function TestInstalledModules {
        # Function to check if running as Administrator
        function IsAdmin {
            $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
            $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        }
        
        Write-Verbose "Checking for the MSAL.PS PowerShell module."
        if (-not (Get-InstalledModule -Name MSAL.PS -MinimumVersion 4.37.0.0 -ErrorAction SilentlyContinue)) {
            if (-not (IsAdmin)) {
                Write-Host "Administrator privileges required to install 'MSAL.PS' module. Re-run PowerShell or the script as Admin." -ForegroundColor Red
                exit
            }
        }
        else {
            Write-Host "Prerequisite not found: Attempting to install 'MSAL.PS' module..." -ForegroundColor Yellow
            try {
                Install-Module -Name MSAL.PS -MinimumVersion 4.37.0.0 -Repository PSGallery -Force
            }
            catch {
                Write-Host "Failed to install 'MSAL.PS' module. Please install it manually." -ForegroundColor Red
                exit
            }
        }
        
        # Check again for MSAL.PS module installation
        if(-not (Get-InstalledModule -Name MSAL.PS -MinimumVersion 4.37.0.0)) {
            Write-Host "Failed to install 'MSAL.PS' module. Please install it manually." -ForegroundColor Red
            exit
        }
    }
    

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
            Write-Verbose "Calling $($MyInvocation.MyCommand)"
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
                    break
                }
                "USGovernmentL4" {
                    $environmentName = "AzureUSGovernment"
                    $graphApiEndpoint = "https://graph.microsoft.us"
                    $exchangeOnlineEndpoint = "https://outlook.office365.us"
                    $autodiscoverSecureName = "https://autodiscover-s.office365.us"
                    $azureADEndpoint = "https://login.microsoftonline.us"
                    break
                }
                "USGovernmentL5" {
                    $environmentName = "AzureUSGovernment"
                    $graphApiEndpoint = "https://dod-graph.microsoft.us"
                    $exchangeOnlineEndpoint = "https://outlook-dod.office365.us"
                    $autodiscoverSecureName = "https://autodiscover-s-dod.office365.us"
                    $azureADEndpoint = "https://login.microsoftonline.us"
                    break
                }
                "ChinaCloud" {
                    $environmentName = "AzureChinaCloud"
                    $graphApiEndpoint = "https://microsoftgraph.chinacloudapi.cn"
                    $exchangeOnlineEndpoint = "https://partner.outlook.cn"
                    $autodiscoverSecureName = "https://autodiscover-s.partner.outlook.cn"
                    $azureADEndpoint = "https://login.partner.microsoftonline.cn"
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
            }
        }
    }
    
    function Get-NewJsonWebToken {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$CertificateThumbprint,
    
            [ValidateSet("CurrentUser", "LocalMachine")]
            [Parameter(Mandatory = $false)]
            [string]$CertificateStore = "CurrentUser",
    
            [Parameter(Mandatory = $false)]
            [string]$Issuer,
    
            [Parameter(Mandatory = $false)]
            [string]$Audience,
    
            [Parameter(Mandatory = $false)]
            [string]$Subject,
    
            [Parameter(Mandatory = $false)]
            [int]$TokenLifetimeInSeconds = 3600,
    
            [ValidateSet("RS256", "RS384", "RS512")]
            [Parameter(Mandatory = $false)]
            [string]$SigningAlgorithm = "RS256"
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
            Write-Verbose "Calling $($MyInvocation.MyCommand)"
        }
        process {
            try {
                $certificate = Get-ChildItem Cert:\$CertificateStore\My\$CertificateThumbprint
                if ($certificate.HasPrivateKey) {
                    $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
                    # Base64url-encoded SHA-1 thumbprint of the X.509 certificate's DER encoding
                    $x5t = [System.Convert]::ToBase64String($certificate.GetCertHash())
                    $x5t = ((($x5t).Replace("\+", "-")).Replace("/", "_")).Replace("=", "")
                    Write-Verbose "x5t is: $x5t"
                } else {
                    Write-Verbose "We don't have a private key for certificate: $CertificateThumbprint and so cannot sign the token"
                    return
                }
            } catch {
                Write-Verbose "Unable to import the certificate - Exception: $($Error[0].Exception.Message)"
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
                Write-Verbose "Issuer: $Issuer will be added to payload"
                $payload.Add("iss", $Issuer)
            }
    
            if (-not([System.String]::IsNullOrEmpty($Audience))) {
                Write-Verbose "Audience: $Audience will be added to payload"
                $payload.Add("aud", $Audience)
            }
    
            if (-not([System.String]::IsNullOrEmpty($Subject))) {
                Write-Verbose "Subject: $Subject will be added to payload"
                $payload.Add("sub", $Subject)
            }
    
            $headerJson = $header | ConvertTo-Json -Compress
            $payloadJson = $payload | ConvertTo-Json -Compress
    
            $headerBase64 = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($headerJson)).Split("=")[0].Replace("+", "-").Replace("/", "_")
            $payloadBase64 = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($payloadJson)).Split("=")[0].Replace("+", "-").Replace("/", "_")
    
            $signatureInput = [System.Text.Encoding]::ASCII.GetBytes("$headerBase64.$payloadBase64")
    
            Write-Verbose "Header (Base64) is: $headerBase64"
            Write-Verbose "Payload (Base64) is: $payloadBase64"
            Write-Verbose "Signature input is: $signatureInput"
    
            $signingAlgorithmToUse = switch ($SigningAlgorithm) {
                ("RS384") { [Security.Cryptography.HashAlgorithmName]::SHA384 }
                ("RS512") { [Security.Cryptography.HashAlgorithmName]::SHA512 }
                default { [Security.Cryptography.HashAlgorithmName]::SHA256 }
            }
            Write-Verbose "Signing the Json Web Token using: $SigningAlgorithm"
    
            $signature = $privateKey.SignData($signatureInput, $signingAlgorithmToUse, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
            $signature = [Convert]::ToBase64String($signature).Split("=")[0].Replace("+", "-").Replace("/", "_")
        }
        end {
            if ((-not([System.String]::IsNullOrEmpty($headerBase64))) -and
                (-not([System.String]::IsNullOrEmpty($payloadBase64))) -and
                (-not([System.String]::IsNullOrEmpty($signature)))) {
                Write-Verbose "Returning Json Web Token"
                return ("$headerBase64.$payloadBase64.$signature")
            } else {
                Write-Verbose "Unable to create Json Web Token"
                return
            }
        }
    }
    
    function Get-NewOAuthToken {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$TenantID,
    
            [Parameter(Mandatory = $true)]
            [string]$ClientID,
    
            [Parameter(Mandatory = $true)]
            [string]$Secret,
    
            [Parameter(Mandatory = $true)]
            [string]$Endpoint,
    
            [Parameter(Mandatory = $false)]
            [string]$TokenService = "oauth2/v2.0/token",
    
            [Parameter(Mandatory = $false)]
            [switch]$CertificateBasedAuthentication,
    
            [Parameter(Mandatory = $true)]
            [string]$Scope
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
            Write-Verbose "Calling $($MyInvocation.MyCommand)"
            $oAuthTokenCallSuccess = $false
            $exceptionMessage = $null
    
            Write-Verbose "TenantID: $TenantID - ClientID: $ClientID - Endpoint: $Endpoint - TokenService: $TokenService - Scope: $Scope"
            $body = @{
                scope      = $Scope
                client_id  = $ClientID
                grant_type = "client_credentials"
            }
    
            if ($CertificateBasedAuthentication) {
                Write-Verbose "Function was called with CertificateBasedAuthentication switch"
                $body.Add("client_assertion_type", "urn:ietf:params:oauth:client-assertion-type:jwt-bearer")
                $body.Add("client_assertion", $Secret)
            } else {
                Write-Verbose "Authentication is based on a secret"
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
                Write-Verbose "Now calling the Invoke-RestMethod cmdlet to create an OAuth token"
                $oAuthToken = Invoke-RestMethod @invokeRestMethodParams
                Write-Verbose "Invoke-RestMethod call was successful"
                $oAuthTokenCallSuccess = $true
            } catch {
                Write-Host "We fail to create an OAuth token - Exception: $($_.Exception.Message)" -ForegroundColor Red
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
            $createOAuthTokenParams = @{
                TenantID                       = $ApplicationInfo.TenantID
                ClientID                       = $ApplicationInfo.ClientID
                Endpoint                       = $AzureADEndpoint
                CertificateBasedAuthentication = (-not([System.String]::IsNullOrEmpty($ApplicationInfo.CertificateThumbprint)))
                Scope                          = $AuthScope
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
    
            $oAuthReturnObject = Get-NewOAuthToken @createOAuthTokenParams
            if ($oAuthReturnObject.Successful -eq $false) {
                Write-Host ""
                Write-Host "Unable to refresh EWS OAuth token. Please review the error message below and re-run the script:" -ForegroundColor Red
                Write-Host $oAuthReturnObject.ExceptionMessage -ForegroundColor Red
                exit
            }
            Write-Host "Obtained a new token" -ForegroundColor Green
            $Script:GraphToken = $oAuthReturnObject.OAuthToken
            $script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime
            return $oAuthReturnObject.OAuthToken.access_token
        }
        else {
            return $Script:Token
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
            Write-Verbose "Calling $($MyInvocation.MyCommand)"
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
                Write-Verbose "Body: $Body"
                $graphApiRequestParams.Add("Body", $Body)
            }
    
            Write-Verbose "Graph API uri called: $($graphApiRequestParams.Uri)"
            Write-Verbose "Method: $($graphApiRequestParams.Method) ContentType: $($graphApiRequestParams.ContentType)"
            $graphApiResponse = Invoke-WebRequestWithProxyDetection -ParametersObject $graphApiRequestParams
    
            if (($null -eq $graphApiResponse) -or
                ([System.String]::IsNullOrEmpty($graphApiResponse.StatusCode))) {
                Write-Verbose "Graph API request failed - no response"
            } elseif ($graphApiResponse.StatusCode -ne $ExpectedStatusCode) {
                Write-Verbose "Graph API status code: $($graphApiResponse.StatusCode) does not match expected status code: $ExpectedStatusCode"
            } else {
                Write-Verbose "Graph API request successful"
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
            }
        }
    }

    function Invoke-WebRequestWithProxyDetection {
        [CmdletBinding(DefaultParameterSetName = "Default")]
        param (
            [Parameter(Mandatory = $true, ParameterSetName = "Default")]
            [string]
            $Uri,
    
            [Parameter(Mandatory = $false, ParameterSetName = "Default")]
            [switch]
            $UseBasicParsing,
    
            [Parameter(Mandatory = $true, ParameterSetName = "ParametersObject")]
            [hashtable]
            $ParametersObject,
    
            [Parameter(Mandatory = $false, ParameterSetName = "Default")]
            [string]
            $OutFile
        )
    
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
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
        } catch {
            Write-VerboseErrorInformation
        }
    }

    function Confirm-ProxyServer {
        [CmdletBinding()]
        [OutputType([bool])]
        param (
            [Parameter(Mandatory = $true)]
            [string]
            $TargetUri
        )
    
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
        try {
            $proxyObject = ([System.Net.WebRequest]::GetSystemWebProxy()).GetProxy($TargetUri)
            if ($TargetUri -ne $proxyObject.OriginalString) {
                Write-Verbose "Proxy server configuration detected"
                Write-Verbose $proxyObject.OriginalString
                return $true
            } else {
                Write-Verbose "No proxy server configuration detected"
                return $false
            }
        } catch {
            Write-Verbose "Unable to check for proxy server configuration"
            return $false
        }
    }

    function WriteErrorInformationBase {
        [CmdletBinding()]
        param(
            [object]$CurrentError = $Error[0],
            [ValidateSet("Write-Host", "Write-Verbose")]
            [string]$Cmdlet
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

    function CreateOutputFile{
        # Create the output file
        $Script:OutputStream = New-Item -Path $OutputPath -Type file -Force -Name $($Script:FileName) -ErrorAction Stop -WarningAction Stop
        # Add the header to the csv file
        $strCSVHeader = $Script:csvOuput = "Sender,Subject,ReceivedDateTime,Folder,internetMessageId,id,"
        Add-Content $Script:OutputStream $strCSVHeader
    }
    

    function GetFolderList {
        Write-Host "Getting a list of mail folders in the mailbox." -ForegroundColor Cyan
        $FolderList = New-Object System.Collections.ArrayList
        $GetFolderListParams = @{
            AccessToken         = $Script:Token
            GraphApiUrl         = $cloudService.graphApiEndpoint
        }
        if($SearchDumpster) {
            $GetFolderListParams.Add("Query","users/$Mailbox/mailFolders/RecoverableItemsRoot/childfolders/?includeHiddenFolders=true")
        }
        else {
            $GetFolderListParams.Add("Query","users/$Mailbox/mailFolders/delta")
        }
        $FolderResults = Invoke-GraphApiRequest @GetFolderListParams
    
        foreach($Result in $FolderResults.Content.Value){
            $FolderList.Add($Result) | Out-Null
        }
        
        while($null -ne $FolderResults.Content.'@odata.nextLink'){
            $Query = $FolderResults.Content.'@odata.nextLink'.Substring($FolderResults.Content.'@odata.nextLink'.IndexOf("user"))
            $FolderResults = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -AccessToken $Script:Token -Query $Query
            foreach($Result in $FolderResults.Content.Value){
                $FolderList.Add($Result) | Out-Null
            }
        }
        
        if ($IncludeFolderList) {
            # We are searching specific folders
            $Script:SearchFolderList = New-Object System.Collections.ArrayList
            foreach ($includedFolder in $IncludeFolderList) {
                $folder = GetFolder -IncludeFolder $includedFolder
                if($folder) {
                    $Script:SearchFolderList.Add($folder) | Out-Null
                }
            }
            if($ProcessSubfolders){
                foreach($folder in $FolderList){
                    if($folder.parentFolderId -eq $Script:ParentFolder) {
                        $Script:SearchFolderList.Add($folder) | Out-Null
                    }
                }
            }
        }
        else {
            $Script:SearchFolderList = $FolderList
        }
        
        if($ExcludeFolderList) {
            foreach($excludedFolder in $ExcludeFolderList){
                $folder = GetFolder -IncludeFolder $excludedFolder
                if($folder) {
                    $Script:SearchFolderList.Remove($folder)
                }
            }
        }
    }
    
    function GetFolder{
        param (
        [Parameter(Mandatory=$true)] [string]$IncludeFolder
        )
        foreach($folder in $FolderList) {
            if($folder.DisplayName -eq $IncludeFolder) {
                $Script:ParentFolder = $folder.id
                return $folder
            }
        }
    }

    function SearchMailbox {
        Write-Host "Performing search against the mailbox..." -ForegroundColor Cyan
        $Script:SearchResults = New-Object System.Collections.ArrayList
        foreach($folder in $Script:SearchFolderList){
            $Script:Token = CheckTokenExpiry -Token ([ref]$Script:GraphToken) -ApplicationInfo $applicationInfo -AzureADEndpoint $azureADEndpoint -AuthScope $Script:GraphScope
            $Uri = "https://graph.microsoft.com/v1.0/users/$Mailbox/mailFolders/$($folder.id)/messages?"
            if(-not($UriFilter)) {
                $UriFilter = CreateSearchQuery
            }

            # Finalize the Uri with the final filter/search settings
            $Uri = $Uri + $UriFilter
            Write-Host ([string]::Format("Performing search against the {0} folder...", $folder.displayName)) -ForegroundColor Green
            Write-Verbose ([string]::Format("Performing query using: {0}", $Uri))
        
            # Search the mailbox for items
            $SearcHParams = @{
                GraphApiUrl     = $cloudService.graphApiEndpoint
                Query           = "users/$Mailbox/mailFolders/$($folder.id)/messages?$UriFilter"
                AccessToken     = $Script:Token
            }
            $SearchItems = Invoke-GraphApiRequest @SearcHParams
            foreach($Result in $SearchItems.Content.Value){
                $Script:SearchResults.Add($Result) | Out-Null
            }
            while($null -ne $SearchItems.Content.'@odata.nextLink'){
                $Query = $SearchItems.Content.'@odata.nextLink'.Substring($SearchItems.Content.'@odata.nextLink'.IndexOf("user"))
                $SearchItems = Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -AccessToken $Script:Token -Query $Query
                foreach($Result in $SearchItems.Content.Value){
                    $Script:SearchResults.Add($Result) | Out-Null
                }
            }
        }
    }
    
    function CreateSearchQuery {
        if($MessageBody -like $null) {
            if($null -ne $Subject) {
                $UriFilter = "filter=contains(subject,`'$Subject`')&`$top=500&`$from=$PageNumber"
            }
        
            if($null -notlike $CreatedBefore){
                $TempStartDate = [datetime]$CreatedBefore
                $TempStartDate = $TempStartDate.ToUniversalTime()
                $SearchStartDate = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $TempStartDate
                if($UriFilter -like '*filter*'){
                    $UriFilter = $UriFilter.Replace('filter=', "filter=receivedDateTime le $($SearchStartDate) and ")
                }
                else {
                    $UriFilter = "filter=receivedDateTime le $($SearchStartDate)&`$top=500&`$from=$PageNumber"
                }
            }
        
            if($null -notlike $CreatedAfter){
                $TempEndDate = [datetime]$CreatedAfter
                $TempEndDate = $TempEndDate.ToUniversalTime()
                $SearchEndDate = '{0:yyyy-MM-ddTHH:mm:ssZ}' -f $TempEndDate
                if($UriFilter -like '*filter*'){
                    $UriFilter = $UriFilter.Replace('filter=', "filter=receivedDateTime ge $($SearchEndDate) and ")
                }
                else {
                    $UriFilter = "filter=receivedDateTime ge $($SearchEndDate)&`$top=500&`$from=$PageNumber"
                }
            }
        
            if($null -ne $Sender) {
                if($UriFilter -like '*filter*'){
                    $UriFilter = $UriFilter.Replace('filter=', "filter=(from/emailAddress/address) eq `'$Sender`' and ")
                }
                else {
                    $UriFilter = "filter=(from/emailAddress/address) eq `'$Sender`'&`$top=100&`$from=$PageNumber"
                }
            }
        }
        else {
            # Build the search query based on specified parameters
            Write-Verbose "Creating a query using the search function."
            $UriFilter = "`$search=`"body:$MessageBody`"&`$top=25"
        
            if($null -notlike $SenderAddress){
                if($UriFilter -like '*search*'){
                    $UriFilter = $UriFilter.Replace('search="', "search=`"from:$SenderAddress` AND ")
                }
                else{
                    $UriFilter = "`$search=`"from:$SenderAddress`"&`$top=25"
                }
            }
        
            if($null -notlike $Subject){
                if($UriFilter -like '*search*'){
                    $UriFilter = $UriFilter.Replace('search="', "search=`"subject:$Subject` AND ")
                }
                else{
                    $UriFilter = "`$search=`"subject:$Subject`"&`$top=1000&`$select=id,parentfolderid,receivedDateTime,subject,from&`$from=$PageNumber"
                }
            }
        
            if($null -notlike $CreatedBefore){
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
        
            if($null -notlike $CreatedAfter){
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
    
    function BuildSearchReport {
        Write-Host "Creating report with the search results." -ForegroundColor Cyan
        foreach($result in $Script:SearchResults) {
            #$result | fl
            $FolderName = ($Script:SearchFolderList | Where-Object {$_.id -eq $result.parentFolderId}).displayName
            #$itemResult = New-Object PSObject -Property @{ InternetMessageId=$item.InternetMessageId; Sender=$item.Sender;ReceivedBy=$item.ReceivedBy;Id=$item.Id;ItemClass=$item.ItemClass;Subject=$item.Subject;DateTimeCreated=$item.DateTimeCreated;Folder=$folderPath;MailboxType=$Script:MailboxType};
            $Script:csvOuput = ($result.sender.emailaddress).address + "," + $result.Subject + "," + $result.receivedDateTime + "," + $FolderName + "," + $result.internetMessageId + "," + $result.id #($result.toRecipients.emailaddress).address + "," + 
            Add-Content $Script:OutputStream $Script:csvOuput
        }    
    }
    
}
process {}
end {
    if($PermissionType -eq "Delegated") {
        TestInstalledModules
    }
    
    $loggerParams = @{
        LogDirectory             = $OutputPath
        LogName                  = "GraphSearchAndDelete-$((Get-Date).ToString("yyyyMMddhhmmss"))-Debug"
        AppendDateTimeToFileName = $false
        ErrorAction              = "SilentlyContinue"
    }
    
    $Script:Logger = Get-NewLoggerInstance @loggerParams
    
    SetWriteHostAction ${Function:Write-HostLog}
    SetWriteVerboseAction ${Function:Write-VerboseLog}
    SetWriteWarningAction ${Function:Write-HostLog}
    
    $cloudService = Get-CloudServiceEndpoint $AzureEnvironment
    $Script:GraphScope = "$($cloudService.graphApiEndpoint)/.default"
    $azureADEndpoint = $cloudService.AzureADEndpoint
    $Script:FileName = "GraphSearchAndDelete-$((Get-Date).ToString("yyyyMMddhhmmss")).csv"

    CreateOutputFile

    $Script:applicationInfo = @{
        "TenantID" = $OAuthTenantId
        "ClientID" = $OAuthClientId
    }

    if ([System.String]::IsNullOrEmpty($OAuthCertificate)) {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($OAuthClientSecret)
        $Secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        $Script:applicationInfo.Add("AppSecret", $Secret)
    } else {
        $jwtParams = @{
            CertificateThumbprint = $OAuthCertificate
            CertificateStore      = $CertificateStore
            Issuer                = $OAuthClientId
            Audience              = "$azureADEndpoint/$OAuthTenantId/oauth2/v2.0/token"
            Subject               = $OAuthClientId
        }
        $jwt = Get-NewJsonWebToken @jwtParams

        if ($null -eq $jwt) {
            Write-Host "Unable to generate Json Web Token by using certificate: $CertificateThumbprint" -ForegroundColor Red
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
    $oAuthReturnObject = Get-NewOAuthToken @createOAuthTokenParams
    if ($oAuthReturnObject.Successful -eq $false) {
        Write-Host ""
        Write-Host "Unable to fetch an OAuth token for accessing EWS. Please review the error message below and re-run the script:" -ForegroundColor Red
        Write-Host $oAuthReturnObject.ExceptionMessage -ForegroundColor Red
        exit
    }
    $Script:GraphToken = $oAuthReturnObject.OAuthToken.access_token
    $Script:Token = $oAuthReturnObject.OAuthToken.access_token
    $Script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime

    GetFolderList
    $Script:SearchFolderList | ft displayName,totalItemCount

    SearchMailbox
    #$Script:SearchResults
    Write-Host ([string]::Format("{0} item(s) found in the search results...", $Script:SearchResults.Count)) -ForegroundColor Green

    BuildSearchReport

    if($DeleteContent) {
        $Script:Token = CheckTokenExpiry -Token ([ref]$Script:GraphToken) -ApplicationInfo $applicationInfo -AzureADEndpoint $azureADEndpoint -AuthScope $Script:GraphScope
        foreach($item in $Script:SearchResults) {
            Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -Query "users/$Mailbox/messages/$($item.id)" -AccessToken $Script:Token -Method DELETE | Out-Null
        }
    }
}

# SIG # Begin signature block
# MIInwQYJKoZIhvcNAQcCoIInsjCCJ64CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDuMqXKevuPMqPv
# Wc7ywG9zxtAQZjdLYCUGR1z+jDYFd6CCDXYwggX0MIID3KADAgECAhMzAAADrzBA
# DkyjTQVBAAAAAAOvMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMxMTE2MTkwOTAwWhcNMjQxMTE0MTkwOTAwWjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDOS8s1ra6f0YGtg0OhEaQa/t3Q+q1MEHhWJhqQVuO5amYXQpy8MDPNoJYk+FWA
# hePP5LxwcSge5aen+f5Q6WNPd6EDxGzotvVpNi5ve0H97S3F7C/axDfKxyNh21MG
# 0W8Sb0vxi/vorcLHOL9i+t2D6yvvDzLlEefUCbQV/zGCBjXGlYJcUj6RAzXyeNAN
# xSpKXAGd7Fh+ocGHPPphcD9LQTOJgG7Y7aYztHqBLJiQQ4eAgZNU4ac6+8LnEGAL
# go1ydC5BJEuJQjYKbNTy959HrKSu7LO3Ws0w8jw6pYdC1IMpdTkk2puTgY2PDNzB
# tLM4evG7FYer3WX+8t1UMYNTAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQURxxxNPIEPGSO8kqz+bgCAQWGXsEw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMTgyNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAISxFt/zR2frTFPB45Yd
# mhZpB2nNJoOoi+qlgcTlnO4QwlYN1w/vYwbDy/oFJolD5r6FMJd0RGcgEM8q9TgQ
# 2OC7gQEmhweVJ7yuKJlQBH7P7Pg5RiqgV3cSonJ+OM4kFHbP3gPLiyzssSQdRuPY
# 1mIWoGg9i7Y4ZC8ST7WhpSyc0pns2XsUe1XsIjaUcGu7zd7gg97eCUiLRdVklPmp
# XobH9CEAWakRUGNICYN2AgjhRTC4j3KJfqMkU04R6Toyh4/Toswm1uoDcGr5laYn
# TfcX3u5WnJqJLhuPe8Uj9kGAOcyo0O1mNwDa+LhFEzB6CB32+wfJMumfr6degvLT
# e8x55urQLeTjimBQgS49BSUkhFN7ois3cZyNpnrMca5AZaC7pLI72vuqSsSlLalG
# OcZmPHZGYJqZ0BacN274OZ80Q8B11iNokns9Od348bMb5Z4fihxaBWebl8kWEi2O
# PvQImOAeq3nt7UWJBzJYLAGEpfasaA3ZQgIcEXdD+uwo6ymMzDY6UamFOfYqYWXk
# ntxDGu7ngD2ugKUuccYKJJRiiz+LAUcj90BVcSHRLQop9N8zoALr/1sJuwPrVAtx
# HNEgSW+AKBqIxYWM4Ev32l6agSUAezLMbq5f3d8x9qzT031jMDT+sUAoCw0M5wVt
# CUQcqINPuYjbS1WgJyZIiEkBMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGaEwghmdAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAOvMEAOTKNNBUEAAAAAA68wDQYJYIZIAWUDBAIB
# BQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIMpgS51IWNzyTLcZW9ckHurP
# 1S9Ob8ECs552m1SfoEaWMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQBjTFBxg/Of0L8huYMuLfRCWHN/ZXkoBbzdwRqMk6QG2IRetZnBC0bq
# cJvtHn2KDchH4yCmce9vXd49EjsI/bC+zfuqtn1i887q0SQI8ij4wcnxcLB0V3zI
# sK3U1h8NAUFU7/0aLSmMVUjpGfSGlpOwMUSnOaAdzw+n/VBExS/Zx67OCbfcC2tt
# rMkbxJmrWdg+hiYcfuJ9LNahLVZD0EdyaV9dczjD+lrd3gUlrkLvvx1ixd0CSjFD
# vRSh/v9V6wslIT/wTUt1CxMHsQEuFD8Q8KS9Q3ezTH1qC/0rGHFeuJDEcfmeXUsp
# IyspFXK/Lq1yYz1mia2Hin0dyjp0uBNpoYIXKTCCFyUGCisGAQQBgjcDAwExghcV
# MIIXEQYJKoZIhvcNAQcCoIIXAjCCFv4CAQMxDzANBglghkgBZQMEAgEFADCCAVkG
# CyqGSIb3DQEJEAEEoIIBSASCAUQwggFAAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIIyWaItfPlONqqaqyeNVRdP5GLJPp820inLycP6BW6jUAgZmcvUt
# ZuAYEzIwMjQwNzA4MTkzNTA4LjAzNlowBIACAfSggdikgdUwgdIxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJ
# cmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBF
# U046ODZERi00QkJDLTkzMzUxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFNlcnZpY2WgghF4MIIHJzCCBQ+gAwIBAgITMwAAAd1dVx2V1K2qGwABAAAB3TAN
# BgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0y
# MzEwMTIxOTA3MDlaFw0yNTAxMTAxOTA3MDlaMIHSMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBP
# cGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjg2REYt
# NEJCQy05MzM1MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAqE4DlETqLnecdREfiWd8
# oun70m+Km5O1y1qKsLExRKs9LLkJYrYO2uJA/5PnYdds3aDsCS1DWlBltMMYXMrp
# 3Te9hg2sI+4kr49Gw/YU9UOMFfLmastEXMgcctqIBqhsTm8Um6jFnRlZ0owKzxpy
# OEdSZ9pj7v38JHu434Hj7GMmrC92lT+anSYCrd5qvIf4Aqa/qWStA3zOCtxsKAfC
# yq++pPqUQWpimLu4qfswBhtJ4t7Skx1q1XkRbo1Wdcxg5NEq4Y9/J8Ep1KG5qUuj
# zyQbupraZsDmXvv5fTokB6wySjJivj/0KAMWMdSlwdI4O6OUUEoyLXrzNF0t6t2l
# bRsFf0QO7HbMEwxoQrw3LFrAIS4Crv77uS0UBuXeFQq27NgLUVRm5SXYGrpTXtLg
# IqypHeK0tP2o1xvakAniOsgN2WXlOCip5/mCm/5hy8EzzfhtcU3DK13e6MMPbg/0
# N3zF9Um+6aOwFBCQrlP+rLcetAny53WcdK+0VWLlJr+5sa5gSlLyAXoYNY3n8pu9
# 4WR2yhNUg+jymRaGM+zRDucDn64HFAHjOWMSMrPlZbsEDjCmYWbbh+EGZGNXg1un
# 6fvxyACO8NJ9OUDoNgFy/aTHUkfZ0iFpGdJ45d49PqEwXQiXn3wsy7SvDflWJRZw
# BCRQ1RPFGeoYXHPnD5m6wwMCAwEAAaOCAUkwggFFMB0GA1UdDgQWBBRuovW2jI9R
# 2kXLIdIMpaPQjiXD8TAfBgNVHSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBf
# BgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3Bz
# L2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmww
# bAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0El
# MjAyMDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUF
# BwMIMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEALlTZsg0uBcgd
# ZsxypW5/2ORRP8rzPIsG+7mHwmuphHbP95o7bKjU6hz1KHK/Ft70ZkO7uSRTPFLI
# nUhmSxlnDoUOrrJk1Pc8SMASdESlEEvxL6ZteD47hUtLQtKZvxchmIuxqpnR8MRy
# /cd4D7/L+oqcJBaReCGloQzAYxDNGSEbBwZ1evXMalDsdPG9+7nvEXFlfUyQqdYU
# Q0nq6t37i15SBePSeAg7H/+Xdcwrce3xPb7O8Yk0AX7n/moGTuevTv3MgJsVe/G2
# J003l6hd1b72sAiRL5QYPX0Bl0Gu23p1n450Cq4GIORhDmRV9QwpLfXIdA4aCYXG
# 4I7NOlYdqWuql0iWWzLwo2yPlT2w42JYB3082XIQcdtBkOaL38E2U5jJO3Rh6Ets
# Oi+ZlQ1rOTv0538D3XuaoJ1OqsTHAEZQ9sw/7+91hSpomym6kGdS2M5//voMCFXL
# x797rNH3w+SmWaWI7ZusvdDesPr5kJV2sYz1GbqFQMEGS9iH5iOYZ1xDkcHpZP1F
# 5zz6oMeZuEuFfhl1pqt3n85d4tuDHZ/svhBBCPcqCqOoM5YidWE0TWBi1NYsd7jz
# zZ3+Tsu6LQrWDwRmsoPuZo6uwkso8qV6Bx4n0UKpjWwNQpSFFrQQdRb5mQouWiEq
# tLsXCN2sg1aQ8GBtDOcKN0TabjtCNNswggdxMIIFWaADAgECAhMzAAAAFcXna54C
# m0mZAAAAAAAVMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZp
# Y2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMy
# MjVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEA5OGmTOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51
# yMo1V/YBf2xK4OK9uT4XYDP/XE/HZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY
# 6GB9alKDRLemjkZrBxTzxXb1hlDcwUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9
# cmmvHaus9ja+NSZk2pg7uhp7M62AW36MEBydUv626GIl3GoPz130/o5Tz9bshVZN
# 7928jaTjkY+yOSxRnOlwaQ3KNi1wjjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDua
# Rr3tpK56KTesy+uDRedGbsoy1cCGMFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74
# kpEeHT39IM9zfUGaRnXNxF803RKJ1v2lIH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2
# K26oElHovwUDo9Fzpk03dJQcNIIP8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5
# TI4CvEJoLhDqhFFG4tG9ahhaYQFzymeiXtcodgLiMxhy16cg8ML6EgrXY28MyTZk
# i1ugpoMhXV8wdJGUlNi5UPkLiWHzNgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9Q
# BXpsxREdcu+N+VLEhReTwDwV2xo3xwgVGD94q0W29R6HXtqPnhZyacaue7e3Pmri
# Lq0CAwEAAaOCAd0wggHZMBIGCSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUC
# BBYEFCqnUv5kxJq+gpE8RjUpzxD/LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJl
# pxtTNRnpcjBcBgNVHSAEVTBTMFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9y
# eS5odG0wEwYDVR0lBAwwCgYIKwYBBQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUA
# YgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU
# 1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2Ny
# bC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIw
# MTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0w
# Ni0yMy5jcnQwDQYJKoZIhvcNAQELBQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/yp
# b+pcFLY+TkdkeLEGk5c9MTO1OdfCcTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulm
# ZzpTTd2YurYeeNg2LpypglYAA7AFvonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM
# 9W0jVOR4U3UkV7ndn/OOPcbzaN9l9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECW
# OKz3+SmJw7wXsFSFQrP8DJ6LGYnn8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4
# FOmRsqlb30mjdAy87JGA0j3mSj5mO0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3Uw
# xTSwethQ/gpY3UA8x1RtnWN0SCyxTkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPX
# fx5bRAGOWhmRaw2fpCjcZxkoJLo4S5pu+yFUa2pFEUep8beuyOiJXk+d0tBMdrVX
# VAmxaQFEfnyhYWxz/gq77EFmPWn9y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGC
# onsXHRWJjXD+57XQKBqJC4822rpM+Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdCPSWU
# 5nR0W2rRnj7tfqAxM328y+l7vzhwRNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEG
# ahC0HVUzWLOhcGbyoYIC1DCCAj0CAQEwggEAoYHYpIHVMIHSMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJl
# bGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNO
# Ojg2REYtNEJCQy05MzM1MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQA2I0cZZds1oM/GfKINsQ5yJKMWEKCBgzCB
# gKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBBQUA
# AgUA6jZ+STAiGA8yMDI0MDcwODIzMDMzN1oYDzIwMjQwNzA5MjMwMzM3WjB0MDoG
# CisGAQQBhFkKBAExLDAqMAoCBQDqNn5JAgEAMAcCAQACAgVyMAcCAQACAhE8MAoC
# BQDqN8/JAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEA
# AgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQEFBQADgYEAaxe1xMnAlxmvGz2T
# sEz6vmo+CLww38xUkkeM5qRqajCr03G9X4YY7WRpYCdJprquTh08NxwEljtpOzqA
# MKsQUWDIgOzi5YKUmGKuy7BZ2l3A3Qb0NQb274t3gzdElcFPmGJmGQ+J5YplnM9v
# V9fy/n8mFkWBCMUf4dKpuII/30AxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGlt
# ZS1TdGFtcCBQQ0EgMjAxMAITMwAAAd1dVx2V1K2qGwABAAAB3TANBglghkgBZQME
# AgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJ
# BDEiBCCf5z6ZB2whC4f+Njwf9OzQUMbWf3SizKItO+j/PKkVWDCB+gYLKoZIhvcN
# AQkQAi8xgeowgecwgeQwgb0EIGH/Di2aZaxPeJmce0fRWTftQI3TaVHFj5GI43rA
# MWNmMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAHd
# XVcdldStqhsAAQAAAd0wIgQgZu41esUTrU24iRy994CxZ4d8EmJAioj56mYTnSCx
# cHowDQYJKoZIhvcNAQELBQAEggIACOdfRBiR4c21PI6bNdZdsvgZXvPzN6nL4uk1
# sHty0hgZed+o7T4fGIWCr1dQwS5sSQwJdVI52QOZY6DiC4vF0DeKcoWnY6IAsJli
# vHzt9O/EW8oyH1n2qrr/DOLvypvQyCr9uxDxSlv/HYQZWpCnM3b2FyI4fpI/v3Qa
# 6nUsnZXJjMdFOsCDej4hgIRPGUwoR52H0AWnLXJVq8P8ZWHac4GgPgbr0l8NAbHP
# pAyDujRFGaQFEixydtPOLChslrGxz+PcDS7dYtxqQl8IWEIi/f/0ItOo2iU99Epq
# lPt0GysebooKDbCTOZb5Jdls03JbIBww1mOhNk+tJDxnlBT14geAi3qqpKn7d66+
# wczUCB5PJ6X/37oMQG6VOfOhK/YFwnAhEVmEieTRHRDogsfse18JL1JodjhktVaT
# 5n5gCUPCGjTPekyzyqj3pOKa1Rn19CdyHMLmJ1YkW2+gTdTdO/4iK99caqk4Ulc9
# 7xb97DmOtg5d0D+91Ehw6A0ZBRjVs9DlPwzfntzQp9XWyERfISnifHMrBwUQjbSt
# sZ4q7mJrWGeKpqaSYGfO9iGyiowLwCxC+UrUvVTjnY5w7ZE/ZdBkJ6sRxQjTTPhB
# YLLEKYbTa7z28qUK8Hq0MT7LxyjM/Co5kyBMcSR0R4yJBqDjCE/Cd752vnCl0q2+
# fzRWmJE=
# SIG # End signature block
