function Convert-MsgToEml
{
<#
	.SYNOPSIS
		Converts MSG files to EML format.
	
	.DESCRIPTION
		Converts MSG files to EML format.
		This is done by:
		- Loading the MSG file into Outlook's Drafts folder (Import-Msg)
		- Waiting until it is synchronized to Exchange
		- Exporting the EML data from the exchange mailbox using Exchange Web Services (EWS) (Export-Eml)
		In order for this command to succeed, both a local Outlook with Exchange connection and EWS access are needed.
		It will automatically try to connect on first use.
		To manually connect, use:
		- Connect-Outlook
		- Connect-EwsExchange
	
	.PARAMETER Path
		Path to the .msg files to convert.
		The input files are NOT deleted.
	
	.PARAMETER OutPath
		Folder to store the results in.
		Defaults to the current path.
	
	.PARAMETER Timeout
		Seconds per message to wait for mail synchronization between Outlook and Exchange.
	
	.PARAMETER EnableException
		This parameters disables user-friendly warnings and enables the throwing of exceptions.
		This is less user friendly, but allows catching exceptions in calling scripts.
	
	.EXAMPLE
		PS C:\> Get-ChildItem *.msg | Convert-MsgToEml
	
		Converts all .msg files in the current folder to .eml files.
#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[Alias('FullName')]
		[string[]]
		$Path,
		
		[PsfValidateScript({ Test-Path $_ -PathType Container }, ErrorString = 'MsgToEml.Validate.Container')]
		[string]
		$OutPath = '.',
		
		[int]
		$Timeout = 30,
		
		[switch]
		$EnableException
	)
	
	begin
	{
		if (-not $script:EwsService) { Connect-EwsExchange }
		if (-not $script:Outllok) { Connect-Outlook }
		
		Assert-OutlookConnected -Cmdlet $PSCmdlet
		Assert-EwsConnected -Cmdlet $PSCmdlet
		
		$ewsDraftFolder = Get-EwsFolder -SearchBase Drafts -Name ''
	}
	process
	{
		:main foreach ($fileItem in $Path)
		{
			if (-not (Test-Path $fileItem))
			{
				Stop-PSFFunction -String 'Convert-MsgToEml.Path.NotFound' -StringValues $fileItem -EnableException $EnableException -Cmdlet $PSCmdlet -Continue
			}
			if ((Get-Item $fileItem).Extension -ne ".msg")
			{
				Stop-PSFFunction -String 'Convert-MsgToEml.Path.NotMsg' -StringValues $fileItem -EnableException $EnableException -Cmdlet $PSCmdlet -Continue
			}
			
			Write-PSFMessage -String 'Convert-MsgToEml.Importing' -StringValues $fileItem
			$outlookItem = $fileItem | Import-Msg -Folder Drafts
			
			$ewsMail = $null
			$startTime = Get-Date
			Write-PSFMessage -String 'Convert-MsgToEml.WaitingForSync' -StringValues $fileItem
			do
			{
				$ewsMail = Get-EwsMail -Folder $ewsDraftFolder -Subject $outlookItem.Subject
				if (-not $ewsMail -and ((Get-Date) -lt $startTime.AddSeconds($Timeout)))
				{
					Stop-PSFFunction -String 'Convert-MsgToEml.Convert.TimedOut' -StringValues $outlookItem.Subject -EnableException $EnableException -Continue -Cmdlet $PSCmdlet -ContinueLabel main
				}
			}
			until ($ewsMail)
			Write-PSFMessage -String 'Convert-MsgToEml.Exporting' -StringValues $outlookItem.Subject, $OutPath
			$ewsMail | Export-Eml -Path $OutPath
			$null = $ewsMail.Delete('HardDelete')
		}
	}
}