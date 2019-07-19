function Get-EwsFolder
{
<#
	.SYNOPSIS
		Search for folders in EWS.
	
	.DESCRIPTION
		Search for folders in EWS.
		Performs a wildcard pattern matching of the folder displayname.
	
	.PARAMETER SearchBase
		The base folder to search from.
		Defaults to the root folder of the mailbox (NOT the inbox)
	
	.PARAMETER Name
		The name-pattern to search for.
		Specify an empty string to only receive the searchbase.
		Defaults to "*"
	
	.PARAMETER PageSize
		The pagesize used when executing the query.
		Cannot be larger than the maximum configured on the server.
		Defaults to the setting stored in EWSAttachmentEncryption.Operations.PageSize
	
	.EXAMPLE
		PS C:\> Get-EwsFolder
	
		Lists all folders in the connected mailbox.
	
	.EXAMPLE
		PS C:\> Get-EwsFolder -SearchBase Inbox -Name ''
	
		Returns just the inbox folder.
#>
	[OutputType([Microsoft.Exchange.WebServices.Data.Folder])]
	[CmdletBinding()]
	Param (
		[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]
		$SearchBase = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,
		
		[AllowEmptyString()]
		[string]
		$Name = "*",
		
		[int]
		$PageSize = (Get-PSFConfigValue -FullName 'MsgToEml.Operations.PageSize')
	)
	
	begin
	{
		Assert-EwsConnected -Cmdlet $PSCmdlet
		try
		{
			Write-PSFMessage -String 'Get-EwsFolder.ConnectingSearchBase' -StringValues $SearchBase
			$baseFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:EwsService, $SearchBase)
		}
		catch
		{
			Stop-PSFFunction -String 'Get-EwsFolder.ConnectionFailed' -StringValues $SearchBase -ErrorRecord $_ -EnableException $true
			return
		}
		$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+Exists -ArgumentList @(
			[Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName
		)
	}
	process
	{
		if (Test-PSFFunctionInterrupt) { return }
		
		if (-not $Name) { return $baseFolder }
		if ($baseFolder.DisplayName -like $Name) { $baseFolder }
		
		$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView($PageSize, 0)
		$folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
		
		do
		{
			Write-PSFMessage -String 'Get-EwsFolder.LoadingFolder'
			$folders = $script:EwsService.FindFolders($baseFolder.Id, $searchFilter, $folderView)
			$folderView.Offset = $folders.NextPageOffset
			# Client Side Filtering, since EWS does not support proper wildcard filtering
			$folders.Folders | Where-Object DisplayName -Like $Name
		}
		while ($folders.MoreAvailable)
	}
}