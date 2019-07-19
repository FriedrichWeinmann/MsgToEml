function Assert-OutlookConnected
{
<#
	.SYNOPSIS
		Asserts Outlook has been connected before trying to run commands against it.
	
	.DESCRIPTION
		Asserts Outlook has been connected before trying to run commands against it.
	
	.PARAMETER Cmdlet
		The PSCmdetlet variable of the calling command
	
	.EXAMPLE
		PS C:\> Assert-OutlookConnected -Cmdlet $Cmdlet
	
		Asserts Outlook has been connected before trying to run commands against it.
#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.Management.Automation.PSCmdlet]
		$Cmdlet
	)
	
	process
	{
		if ($script:Outlook) { return }
		
		Write-PSFMessage -Level Warning -String 'Assert-OutlookConnected.Failed' -StringValues $Cmdlet.MyInvocation.MyCommand.Name -FunctionName $Cmdlet.MyInvocation.MyCommand.Name -Line (Get-PSCallstack)[1].ScriptLineNumber
		$exception = New-Object System.InvalidOperationException('Not connected to Outlook yet. Use Assert-OutlookConnected to connect to Outlook first')
		$record = New-Object System.Management.Automation.ErrorRecord($exception, 'NotConnected', 'ConnectionError', $null)
		$Cmdlet.ThrowTerminatingError($record)
	}
}