<#
#NOTE - Disclaimer
#Following programming examples is for illustration only, without warranty either expressed or implied,
#including, but not limited to, the implied warranties of merchantability and/or fitness for a particular purpose. 
#This sample code assumes that you are familiar with the programming language being demonstrated and the tools 
#used to create and debug procedures. This sample code is provided for the purpose of illustration only and is 
#not intended to be used in a production environment. 

  .SYNOPSIS
  Use this PowerShell script to test POP3 connections to O365 mailbox using OAuth .
  Please install MSAL.PS Powershell module as prerequisite. 
   (https://github.com/AzureAD/MSAL.PS)Install-Module -Name MSAL.PS
  
  Referce article with more insides 
 
  https://techcommunity.microsoft.com/t5/exchange-team-blog/announcing-oauth-2-0-client-credentials-flow-support-for-pop-and/ba-p/3562963
  https://github.com/DanijelkMSFT/ThisandThat/blob/main/Get-IMAPAccessToken.ps1

  .DESCRIPTION
  The function helps admins to test their POP3 OAuth Azure Application, 
  with Interactive user login und providing or the lately released client credential flow
  using the right formatting for the XOAuth2 login string.
  After successful logon, a simple POP3 message counts on folders is done, in addition it also allows to 
  test shared mailbox acccess for users if fullaccess has been provided. 
  
  Using Windows Powershell allows MSAL to cache the access+refresh token on disk for further executions for interactive login scenario.
  It´s a simple proof of concept with no further error managment.

  .PARAMETER tenantID
  Specifies the target tenant.

  .PARAMETER clientId
  Specifies the ClientID/ApplicationID of the registered Azure AD Application with needed POP3 Graph permissions

  .PARAMETER clientsecret
  Specifies the ClientSecret configured in the Azure AD Application for client credential flow

  .PARAMETER clientcertificate
  Specifies the ClientCertificate Thumbprint configured in the Azure AD Application for client credential flow

  .PARAMETER targeMailbox
  Specifies the primary emailaddress of the targetmailbox which should be accessed by service principal which has fullaccess to for client credential flow

  .PARAMETER redirectUri
  Specifies the redirectUri of the registered Azure AD Application for authorization code flow (interactive flow)

  .PARAMETER LoginHint
  Specifies the Userprincipalname of the logging in user for authorization code flow (interactive flow)

  .PARAMETER SharedMailbox (optinal)
  Specifies the primary emailaddress of the Sharedmailbox logged in user has fullaccess to for authorization code flow (interactive flow)

  .EXAMPLE
  PS> .\Get-POP3AccessToken.ps1 -tenantID "" -clientId "" -redirectUri "https://localhost" -LoginHint "user@contoso.com"

  .EXAMPLE
  PS> .\Get-POP3AccessToken.ps1 -tenantID "" -clientId "" -redirectUri "https://localhost" -LoginHint "user@contoso.com" -SharedMailbox "SharedMailbox@contoso.com"

  .EXAMPLE
  PS> .\Get-POP3AccessToken.ps1 -tenantID "" -clientId "" -redirectUri "https://localhost" -LoginHint "user@contoso.com" -Verbose

  .EXAMPLE
  PS> .\Get-POP3AccessToken.ps1 -tenantID "" -clientId "" -clientsecret '' -targetMailbox "TargetMailbox@contoso.com"

  .EXAMPLE
  PS> .\Get-POP3AccessToken.ps1 -tenantID "" -clientId "" -clientcertificate '' -targetMailbox "TargetMailbox@contoso.com" 

  .EXAMPLE
  PS> .\Get-POP3AccessToken.ps1 -tenantID "" -clientId "" -clientsecret '' -targetMailbox "TargetMailbox@contoso.com" -Verbose

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][string]$tenantID,
    [Parameter(Mandatory = $true)][String]$clientId,
    
    [Parameter(Mandatory = $true,ParameterSetName="authorizationcode")][String]$redirectUri,
    [Parameter(Mandatory = $true,ParameterSetName="authorizationcode")][String]$LoginHint,
    [Parameter(Mandatory = $false,ParameterSetName="authorizationcode")][String]$SharedMailbox,

    [Parameter(Mandatory = $true,ParameterSetName="clientcredentialsSecret")][String]$clientsecret,
    [Parameter(Mandatory = $true,ParameterSetName="clientcredentialsCertificate")][String]$clientcertificate,

    [Parameter(Mandatory = $true,ParameterSetName="clientcredentialsSecret")]
    [Parameter(Mandatory = $true,ParameterSetName="clientcredentialsCertificate")]
    [String]$targetMailbox
)

function Test-POP3XOAuth2Connectivity {
# get Accesstoken via user authentication and store Access+Refreshtoken for next attempts
if ( $redirectUri ){
    $MsftPowerShellClient = New-MsalClientApplication -ClientId $clientID -TenantId $tenantID -RedirectUri $redirectURI  | Enable-MsalTokenCacheOnDisk -PassThru
    try {
        $authResult = $MsftPowerShellClient | Get-MsalToken -LoginHint LoginHint -Scopes 'https://outlook.office365.com/.default'
		}
	catch  {
        Write-Host "Ran into an exception while getting accesstoken user grant flow" -ForegroundColor Red
        $_.Exception.Message
        $_.FullyQualifiedErrorId
        break
    }
}

if ( $clientsecret ){
    $SecuredclientSecret = ConvertTo-SecureString $clientsecret -AsPlainText -Force
    $MsftPowerShellClient = New-MsalClientApplication -ClientId $clientID -TenantId $tenantID -ClientSecret $SecuredclientSecret 
    try {
    	$authResult = $MsftPowerShellClient | Get-MsalToken -Scopes 'https://outlook.office365.com/.default'
    }
    catch  {
        Write-Host "Ran into an exception while getting accesstoken using clientsecret" -ForegroundColor Red
        $_.Exception.Message
        $_.FullyQualifiedErrorId
        break
    }
}


if ( $clientcertificate ){
    $ClientCert = Get-ChildItem "cert:\currentuser\my\$clientcertificate"
    $MsftPowerShellClient = New-MsalClientApplication -ClientId $clientID -TenantId $tenantID -ClientCertificate $ClientCert
    try {
        $authResult = $MsftPowerShellClient | Get-MsalToken -Scopes 'https://outlook.office365.com/.default'
    }
    catch  {
        Write-Host "Ran into an exception while getting accesstoken using certificate" -ForegroundColor Red
        $_.Exception.Message
        $_.FullyQualifiedErrorId
        break
    }
}



$accessToken = $authResult.AccessToken
$username = $authResult.Account.Username

# build authentication string with accesstoken and username like documented here
# https://docs.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth#authenticate-connection-requests

# in the case if client credential usage we need to add the target mailbox like shared mailbox access
if ( $targetMailbox) { $SharedMailbox = $targetMailbox }

if ( $SharedMailbox ) {
    $b="user=" + $SharedMailbox + "$([char]0x01)auth=Bearer " + $accessToken + "$([char]0x01)$([char]0x01)"
	write-verbose "JWT Token --   $accessToken"
    Write-Host "Accessing Sharedmailbox - $SharedMailbox - with Accesstoken of User $userName." -ForegroundColor DarkGreen
} else {
        $b="user=" + $userName + "$([char]0x01)auth=Bearer " + $accessToken + "$([char]0x01)$([char]0x01)"
        }

$Bytes = [System.Text.Encoding]::ASCII.GetBytes($b)
$POPIMAPLogin =[Convert]::ToBase64String($Bytes)

Write-Verbose "SASL XOAUTH2 login string $POPIMAPLogin"

# connecting to Office 365 POP3 Service
Write-Host "Connect to Office 365 POP3 Service." -ForegroundColor DarkGreen
$ComputerName = 'Outlook.office365.com'
$Port = '995'
    try {
        $TCPConnection = New-Object System.Net.Sockets.Tcpclient($($ComputerName), $Port)
        $TCPStream = $TCPConnection.GetStream()
        try {
            $SSLStream  = New-Object System.Net.Security.SslStream($TCPStream)
            $SSLStream.ReadTimeout = 5000
            $SSLStream.WriteTimeout = 5000     
            $CheckCertRevocationStatus = $true
            $SSLStream.AuthenticateAsClient($ComputerName,$null,[System.Security.Authentication.SslProtocols]::Tls12,$CheckCertRevocationStatus)
        }
        catch  {
            Write-Host "Ran into an exception while negotating SSL connection. Exiting." -ForegroundColor Red
            $_.Exception.Message
            break
        }
    }
    catch  {
    Write-Host "Ran into an exception while opening TCP connection. Exiting." -ForegroundColor Red
    $_.Exception.Message
    break
    }    

    # continue if connection was successfully established
    $SSLstreamReader = new-object System.IO.StreamReader($sslStream)
    $SSLstreamWriter = new-object System.IO.StreamWriter($sslStream)
    $SSLstreamWriter.AutoFlush = $true
    $SSLstreamReader.ReadLine()

    Write-Host "Authenticate using XOAuth2." -ForegroundColor DarkGreen
    # authenticate and check for results
    #$command = "AUTH XOAUTH2 {0}" -f $POPIMAPLogin
	$command = "AUTH XOAUTH2"
    Write-Verbose "Executing command -- $command"
    $SSLstreamWriter.WriteLine($command) 
    #respose might take longer sometimes
    while (!$ResponseStr ) { 
        try { $ResponseStr = $SSLstreamReader.ReadLine() } catch { }
    }
#Write-Verbose $ResponseStr
    if ( $ResponseStr -like "*+*") 
    {
        $ResponseStr
	} else {
        Write-host "ERROR during authentication $ResponseStr" -Foregroundcolor Red
    }
	
		Write-Verbose "Passing XOAUTH2 formatted token"
		$SSLstreamWriter.WriteLine($POPIMAPLogin) 
		#respose might take longer sometimes
    while (!$ResponseStr2 ) { 
        try { $ResponseStr2 = $SSLstreamReader.ReadLine() } catch { }
    }
	
	Write-Verbose $ResponseStr2

    if ( $ResponseStr2 -like "*+OK*") 
    {
        $ResponseStr
        Write-Host "Getting list of messages as authentication was successfull." -ForegroundColor DarkGreen
        $command = 'LIST'
        Write-Verbose "Executing command -- $command"
        $SSLstreamWriter.WriteLine($command) 

        $done = $false
        $str = $null
        while (!$done ) {
            $str = $SSLstreamReader.ReadLine()
            if ($str -like "*.") { $str ; $done = $true } 
            elseif ($str -like "* BAD User is authenticated but not connected.") { $str; "Causing Error: POP3 protcol access to mailbox is disabled or permission not granted for client credential flow. Please enable POP3 protcol access or grant fullaccess to service principal."; $done = $true} 
            else { $str }
        }

        Write-Host "Logout and cleanup sessions." -ForegroundColor DarkGreen
        $command = 'QUIT'
        Write-Verbose "Executing command -- $command"
        $SSLstreamWriter.WriteLine($command) 
        $SSLstreamReader.ReadLine()

    } else {
        Write-host "ERROR during authentication $ResponseStr2" -Foregroundcolor Red
    }

    # Session cleanup
    if ($SSLStream) {
        $SSLStream.Dispose()
    }
    if ($TCPStream) {
        $TCPStream.Dispose()
    }
    if ($TCPConnection) {
        $TCPConnection.Dispose()
    }
}

#check for needed msal.ps module
if ( !(Get-Module msal.ps -ListAvailable) ) { Write-Host "MSAL.PS module not installed, please check it out here https://www.powershellgallery.com/packages/MSAL.PS/" -ForegroundColor Red; break}

# execute function
Test-POP3XOAuth2Connectivity
