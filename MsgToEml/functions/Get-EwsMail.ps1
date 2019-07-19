function Get-EwsMail
{
<#
	.SYNOPSIS
		Retrieves email objects.
	
	.DESCRIPTION
		Searches EWS folders for matching emails.
	
	.PARAMETER Folder
		The folder to search.
		Defaults to the inbox.
	
	.PARAMETER Subject
		A subject line to filter by.
	
	.PARAMETER Before
		Only emails received before this will be returned.
	
	.PARAMETER After
		Only emails received after this will be returned.
	
	.PARAMETER HasAttachment
		Setting this filters emails by whether they have an attachment.
		Note: Inline attachments (such as pictures that are part of the mail body) don't enable this flag.
	
	.PARAMETER PageSize
		The pagesize used when executing the query.
		Cannot be larger than the maximum configured on the server.
		Defaults to the setting stored in EWSAttachmentEncryption.Operations.PageSize
	
	.EXAMPLE
		PS C:\> Get-EwsMail
	
		Returns all emails in the inbox folder
	
	.EXAMPLE
		PS C:\> Get-EwsFolder -SearchBase Inbox | Get-EwsMail
	
		Returns all emails in the inbox folder and all subfolders.
	
	.EXAMPLE
		PS C:\> Get-EwsMail -After "-7d" -HasAttachment
	
		Returns all emails received in the last 7 days that have an attachment
#>
	[OutputType([Microsoft.Exchange.WebServices.Data.Item])]
	[CmdletBinding()]
	param (
		[Parameter(ValueFromPipeline = $true)]
		$Folder,
		
		[string]
		$Subject = '*',
		
		[PSFDateTime]
		$Before,
		
		[PSFDateTime]
		$After,
		
		[switch]
		$HasAttachment,
		
		[int]
		$PageSize = (Get-PSFConfigValue -FullName 'MsgToEml.Operations.PageSize')
	)
	
	begin
	{
		Assert-EwsConnected -Cmdlet $PSCmdlet
		
		#region Filtering
		$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection
		
		#region Subject Filter
		if ($Subject -eq '*')
		{
			$filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+Exists -ArgumentList @(
				[Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject
			)
			$searchFilter.Add($filter)
		}
		elseif ($Subject.Contains("*"))
		{
			foreach ($segment in $Subject.Split("*"))
			{
				if (-not $segment) { continue }
				
				$filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring  -ArgumentList @(
					[Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject
					$segment
					[Microsoft.Exchange.WebServices.Data.ContainmentMode]::Substring
					[Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase
				)
				$searchFilter.Add($filter)
			}
		}
		else
		{
			$filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring -ArgumentList @(
				[Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject
				$Subject
				[Microsoft.Exchange.WebServices.Data.ContainmentMode]::FullString
				[Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase
			)
			$searchFilter.Add($filter)
		}
		#endregion Subject Filter
		
		#region Before
		if ($Before)
		{
			$filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan -ArgumentList @(
				[Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
				$Before.Value
			)
			$searchFilter.Add($filter)
		}
		#endregion Before
		
		#region After
		if ($After)
		{
			$filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan -ArgumentList @(
				[Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
				$After.Value
			)
			$searchFilter.Add($filter)
		}
		#endregion After
		
		#region HasAttachment
		# Could be set to -HasAttachment:$false in order to explicitly only find mails without attachment
		if (Test-PSFParameterBinding -ParameterName HasAttachment)
		{
			$filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo -ArgumentList @(
				[Microsoft.Exchange.WebServices.Data.ItemSchema]::HasAttachments
				$HasAttachment.ToBool()
			)
			$searchFilter.Add($filter)
		}
		#endregion HasAttachment
		#endregion Filtering
	}
	process
	{
		if (-not $Folder) { $Folder = Get-EwsFolder -SearchBase Inbox -Name '' }
		
		#region Process Folders to retrieve items
		foreach ($folderItem in $Folder)
		{
			if ($folderItem -isnot [Microsoft.Exchange.WebServices.Data.Folder])
			{
				if ($folderItem -eq 'Inbox') { $folderItem = Get-EwsFolder -SearchBase Inbox -Name '' }
				else { $folderItem = Get-EwsFolder -SearchBase Inbox -Name $folderItem }
			}
			# Need inner loop, since wildcard string or foldername collision can lead to more than one output item.
			foreach ($resolvedFolderItem in $folderItem)
			{
				Write-PSFMessage -String 'Get-EwsMail.RetrievingFromFolder' -StringValues $resolvedFolderItem.DisplayName
				$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView $PageSize, 0
				
				# Starting here we can guarantee resolvedItem is a single EWS Folder object
				do
				{
					$list = $resolvedFolderItem.FindItems($searchFilter, $view)
					
					$list.Items
					$view.Offset = $list.NextPageOffset
				}
				while ($list.MoreAvailable)
			}
		}
		#endregion Process Folders to retrieve items
	}
}