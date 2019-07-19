function Connect-EwsExchange
{
<#
	.SYNOPSIS
		Establish a connection to Exchange using EWS.
	
	.DESCRIPTION
		Establish a connection to Exchange using EWS.
		The session is stored in the module scope and will automatically be used for subsequent requests.
	
	.PARAMETER Mailbox
		The email address to connect to.
		The correct server to contact will be determined using Auto-Discover based on this address.
	
	.PARAMETER Credential
		Alternative credentials to use for connecting to the server.
	
	.PARAMETER Impersonate
		Email address of the user to impersonate.
		Requires the highly sensitive "Impersonate" exchange privilege.
	
	.PARAMETER Version
		The exchange server version to connect to.
		Defaults to Exchange2013_SP1, only change for legacy server.
		This governs the compatibility mode and higher versions have greater performance.
	
	.EXAMPLE
		PS C:\> Connect-EwsExchange -Mailbox 'max.mustermann@contoso.com'
	
		Connect to Max' mailbox in the contoso domain.
#>
	[CmdletBinding()]
	param (
		[string]
		$Mailbox,
		
		[PSCredential]
		$Credential = [Management.Automation.PSCredential]::Empty,
		
		[string]
		$Impersonate,
		
		[Microsoft.Exchange.WebServices.Data.ExchangeVersion]
		$Version = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
	)
	
	begin
	{
		if (-not $Mailbox)
		{
			try
			{
				$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
				$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
				$aceuser = [ADSI]$sidbind
				$Mailbox = $aceuser.mail.ToString()
			}
			catch
			{
				Stop-PSFFunction -String 'Connect-EwsExchange.FailedAutodetect.Email' -StringValues $windowsIdentity -EnableException $true -Cmdlet $PSCmdlet -ErrorRecord $_
				return
			}
		}
	}
	process
	{
		#region Setting up the service
		Write-PSFMessage -String 'Connect-EwsExchange.ConnectionStart' -StringValues $Mailbox, $Version
		$exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($Version, [System.TimeZoneInfo]::Local)
		$exchangeService.Timeout = Get-PSFConfigValue -FullName 'MsgToEml.Connect.Timeout'
		$exchangeService.UseDefaultCredentials = $true
		if ($Credential -ne [Management.Automation.PSCredential]::Empty)
		{
			$exchangeService.UseDefaultCredentials = $false
			$exchangeService.Credentials = $Credential.GetNetworkCredential()
			Write-PSFMessage -String 'Connect-EwsExchange.AuthenticatingAs' -StringValues $Credential.UserName
		}
		try
		{
			Write-PSFMessage -String 'Connect-EwsExchange.AccessingMailbox' -StringValues $Mailbox
			$exchangeService.AutodiscoverUrl($Mailbox)
		}
		catch
		{
			Stop-PSFFunction -String 'Connect-EwsExchange.FailedAutodetect' -StringValues $Mailbox -EnableException $true -Cmdlet $PSCmdlet -ErrorRecord $_
			return
		}
		if ($Impersonate)
		{
			Write-PSFMessage -String 'Connect-EwsExchange.Impersonating' -StringValues $Impersonate
			$exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Impersonate)
		}
		#endregion Setting up the service
		
		#region Connection Test
		$testFolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root
		try
		{
			$baseFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $testFolder)
			Write-PSFMessage -String 'Connect-EwsExchange.ConnectionSuccess' -StringValues $Mailbox
		}
		catch
		{
			Stop-PSFFunction -String 'Connect-EwsExchange.ConnectionFailed' -StringValues $Mailbox -EnableException $true -Cmdlet $PSCmdlet -ErrorRecord $_
			return
		}
		#endregion Connection Test
		
		$script:EwsService = $exchangeService
	}
}