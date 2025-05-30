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
# Version 24.11.12.1014

param (
    [Parameter(Position=0,Mandatory=$True,HelpMessage="The Mailbox parameter specifies the mailbox to be accessed.")]
    [ValidateNotNullOrEmpty()] 
    [string]$Mailbox,
	
    [Parameter(Mandatory=$False,HelpMessage="The ProcessSubfolders parameter is a switch to enable searching the subfolders of the IncludeFolderList.")] 
    [switch]$IncludeSubfolders,
	
    [Parameter(Mandatory=$False,HelpMessage="The IncludeFolderList parameter specifies the folder(s) to be searched (if not present, then the Inbox folder will be searched).  Any exclusions override this list.")] 
    $IncludeFolderList=$null,
    
    [Parameter(Mandatory=$False,HelpMessage="The ExcludeFolderList parameter specifies the folder(s) to be excluded (these folders will not be searched).")] 
    $ExcludeFolderList=$null,

    [Parameter(Mandatory=$False,HelpMessage="The ExcludeSubfolders parameter is a switch to prevent searching the subfolders of the ExcludeFolderList.")] 
    [switch]$ExcludeSubfolders,

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

    [Parameter(Mandatory=$False, HelpMessage="The HardDelete parameter is a switch to hard-delete the items found in the search results (otherwise, they'll be moved to Deleted Items).")]
    [switch]$HardDelete,
	
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
    [SecureString]$OAuthClientSecret,
    
    [Parameter(Mandatory=$False,HelpMessage="The OAuthCertificate parameter is the certificate for the registered application. Certificate auth requires MSAL libraries to be available.")] 
    [string]$OAuthCertificate = $null,
  
    [Parameter(Mandatory=$False,HelpMessage="The CertificateStore parameter specifies the certificate store where the certificate is loaded.")] [ValidateSet("CurrentUser", "LocalMachine")]
     [string] $CertificateStore = $null,
    
    [ValidateScript({ Test-Path $_ })] [Parameter(Mandatory = $true, HelpMessage="The OutputPath parameter specifies the path for the EWS usage report.")] [string] $OutputPath,

    [Parameter(Mandatory=$False,HelpMessage="The ThrottlingDelay parameter specifies the throttling delay (time paused between sending EWS requests) - note that this will be increased automatically if throttling is detected")]
    [int]$ThrottlingDelay = 0,

    [Parameter(Mandatory=$false,HelpMessage="The BatchSize parameter specifies how many items to delete within a batch request.")]
    [ValidateRange(1,20)][int]$BatchSize=20
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
                    Write-Verbose "Adding the folder $($folder.displayName) to the list of folders to query."
                    $Script:SearchFolderList.Add($folder) | Out-Null
                }
            }
            # Find subfolders for the specified folders
            if($ProcessSubfolders){
                foreach($folder in $FolderList){
                    # Folder can be multiple levels deep subfolder
                    foreach($foundFolder in $Script:SearchFolderList){
                        if($folder.parentFolderId -eq $foundFolder.Id) {
                            Write-Verbose "Adding the folder $($folder.displayName) to the list of folders to query."
                            $Script:SearchFolderList.Add($folder) | Out-Null
                            break
                        }
                    }
                }
            }
        }
        else {
            $Script:SearchFolderList = $FolderList
        }
        $tempFolderList = New-Object System.Collections.ArrayList
        # Check for folders that need to be excluded from search queries
        if($ExcludeFolderList) {
            foreach($excludedFolder in $ExcludeFolderList){
                $folder = GetFolder -IncludeFolder $excludedFolder
                if($folder) {
                    Write-Verbose "Adding the $($folder.displayName) to list of folders to be removed."
                    $tempFolderList.Add($folder) | Out-Null
                    if($ExcludeSubfolders) {
                        # Remove any subfolders of the excluded folder from the list
                        foreach($foundFolder in $Script:SearchFolderList){
                            foreach($tempFolder in $tempFolderList) {
                                if($foundFolder.parentFolderId -eq $tempFolder.Id){
                                    Write-Verbose "Adding the $($foundFolder.displayName) to list of folders to be removed."
                                    $tempFolderList.Add($foundFolder) | Out-Null
                                    break
                                }
                            }
                        }
                    }
                }
            }
            # Remove the excluded folders from the list
            foreach($folder in $tempFolderList) {
                $Script:SearchFolderList.Remove($folder) | Out-Null
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
            $SearchParams = @{
                GraphApiUrl     = $cloudService.graphApiEndpoint
                Query           = "users/$Mailbox/mailFolders/$($folder.id)/messages?$UriFilter"
                AccessToken     = $Script:Token
            }
            $SearchItems = Invoke-GraphApiRequest @SearchParams
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
        if([string]::IsNullOrEmpty($MessageBody)) {
            if(-not([string]::IsNullOrEmpty($Subject))) {
                $UriFilter = "filter=contains(subject,`'$Subject`')&`$top=500&`$from=$PageNumber"
            }
            if(-not([string]::IsNullOrEmpty($CreatedBefore))) {
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
            if(-not([string]::IsNullOrEmpty($CreatedAfter))){
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
            if(-not([string]::IsNullOrEmpty($Sender))){
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
                    $UriFilter = "`$search=`"subject:$Subject`"&`$top=1000&`$select=id,parentfolderid,receivedDateTime,subject,from&`$from=$PageNumber"
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
    
    function BuildSearchReport {
        Write-Host "Creating report with the search results." -ForegroundColor Cyan
        foreach($result in $Script:SearchResults) {
            $FolderName = ($Script:SearchFolderList | Where-Object {$_.id -eq $result.parentFolderId}).displayName
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
    $Script:SearchFolderList | Format-Table displayName,totalItemCount

    SearchMailbox
    Write-Host ([string]::Format("{0} item(s) found in the search results...", $Script:SearchResults.Count)) -ForegroundColor Green

    BuildSearchReport

    if($DeleteContent) {
        $Script:Token = CheckTokenExpiry -Token ([ref]$Script:GraphToken) -ApplicationInfo $applicationInfo -AzureADEndpoint $azureADEndpoint -AuthScope $Script:GraphScope
        [int]$itemsDeleted = 0
        # Make sure the results are not less than the batch size
        if($Script:SearchResults.count -lt $BatchSize){
            $BatchSize = $Script:SearchResults.Count
        }
        $Query = "`$batch"
        # Loop thru the results creating batches to delete
        while($itemsDeleted -lt $Script:SearchResults.Count){
            # Make sure the batch size is not greater than the items left to process
            if(($Script:SearchResults.Count - $itemsDeleted) -lt $BatchSize){
                $BatchSize = $Script:SearchResults.Count - $itemsDeleted
            }
            #region CreateBatch
            $requests = New-Object System.Collections.ArrayList
            for($x=0; $x -lt $BatchSize; $x++){
                if($HardDelete){
                    $Method = "POST"
                    $Url = "/users/$($Mailbox)/messages/$($Script:SearchResults[$itemsDeleted].id)/permanentDelete"
                }
                else {
                    $Method = "DELETE"
                    $Url = "/users/$($Mailbox)/messages/$($Script:SearchResults[$itemsDeleted].id)"
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
            Write-Verbose "Sending request for the next batch of deletions."
            Invoke-GraphApiRequest -GraphApiUrl $cloudService.graphApiEndpoint -Query $Query -AccessToken $Script:Token -Method POST -Body $batchRequest | Out-Null
        }
    }
}
