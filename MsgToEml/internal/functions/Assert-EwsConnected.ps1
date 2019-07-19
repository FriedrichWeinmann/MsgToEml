function Assert-EwsConnected
{
<#
	.SYNOPSIS
		Asserts EWS has been connected before trying to run commands against it.
	
	.DESCRIPTION
		Asserts EWS has been connected before trying to run commands against it.
	
	.PARAMETER Cmdlet
		The PSCmdetlet variable of the calling command
	
	.EXAMPLE
		PS C:\> Assert-EwsConnected -Cmdlet $Cmdlet
	
		Asserts EWS has been connected before trying to run commands against it.
#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.Management.Automation.PSCmdlet]
		$Cmdlet
	)
	
	process
	{
		if ($script:EwsService) { return }
		
		Write-PSFMessage -Level Warning -String 'Assert-EwsConnected.Failed' -StringValues $Cmdlet.MyInvocation.MyCommand.Name -FunctionName $Cmdlet.MyInvocation.MyCommand.Name -Line (Get-PSCallstack)[1].ScriptLineNumber
		$exception = New-Object System.InvalidOperationException('Not connected to EWS yet. Use Connect-EwsExchange to connect to an Exchange server first')
		$record = New-Object System.Management.Automation.ErrorRecord($exception, 'NotConnected', 'ConnectionError', $null)
		$Cmdlet.ThrowTerminatingError($record)
	}
}