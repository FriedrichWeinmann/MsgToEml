function Connect-Outlook
{
<#
	.SYNOPSIS
		Establishes a COM binding to outlook.
	
	.DESCRIPTION
		Establishes a COM binding to outlook.
		Requires the outlook application to be installed on the current machine.
		Attaches to running outlook if already running, otherwise establishes a new session.
	
	.EXAMPLE
		PS C:\> Connect-Outlook
	
		Establishes a COM binding to outlook.
#>
	[CmdletBinding()]
	Param (
	
	)
	
	process
	{
		if (-not $script:Outlook)
		{
			if (Get-Process outlook -ErrorAction Ignore)
			{
				Write-PSFMessage -String 'Connect-Outlook.Existing'
				try { $script:Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") }
				catch { Stop-PSFFunction -String 'Connect-Outlook.Existing.Failed' -ErrorRecord $_ -EnableException $true -Cmdlet $PSCmdlet }
			}
			else
			{
				Write-PSFMessage -String 'Connect-Outlook.NewComObject'
				try { $script:Outlook = New-Object -ComObject Outlook.Application }
				catch { Stop-PSFFunction -String 'Connect-Outlook.NewComObject.Failed' -ErrorRecord $_ -EnableException $true -Cmdlet $PSCmdlet }
			}
		}
	}
}