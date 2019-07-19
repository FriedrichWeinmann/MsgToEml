function Export-Eml
{
<#
	.SYNOPSIS
		Exports an EWS item as EML file.
	
	.DESCRIPTION
		Exports an EWS item as EML file.
		Tested against mail messages, but other types of items should work equally well, so long as they have a subject and a MimeContent.
	
	.PARAMETER Item
		The item to export.
	
	.PARAMETER Path
		The path to export to.
		Will use the subject as filename if a folder is specified.
		Expects the extension to be .eml if a filename is specified.
	
	.PARAMETER EnableException
		This parameters disables user-friendly warnings and enables the throwing of exceptions.
		This is less user friendly, but allows catching exceptions in calling scripts.
	
	.EXAMPLE
		PS C:\> $ewsItem | Export-Eml -Path '.'
	
		Exports the items stored in $ewsItem into the current folder, each named for its subject.
#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[Microsoft.Exchange.WebServices.Data.Item[]]
		$Item,
		
		[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
		[string]
		$Path,
		
		[switch]
		$EnableException
	)
	
	begin
	{
		$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet
		$propertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
		$propertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
		
		# To ensure no accidental super-scope lookup
		$fileName = $null
	}
	process
	{
		#region Path Resolution
		try { $resolvedPath = Resolve-PSFPath -Path $Path -Provider FileSystem -SingleItem -NewChild }
		catch
		{
			Stop-PSFFunction -String 'Export-Eml.ResolvePath.Failed' -StringValues $Path -EnableException $EnableException -ErrorRecord $_
			return
		}
		
		if ($resolvedPath -like "*.eml")
		{
			$folderPath = Split-Path -Path $resolvedPath
			$fileName = Split-Path -Path $resolvedPath -Leaf
			if (-not (Test-Path $folderPath))
			{
				Stop-PSFFunction -String 'Export-Eml.PathValidation.FolderNotExists' -StringValues $folderPath -EnableException $EnableException -ErrorRecord $_
				return
			}
		}
		else
		{
			$folderPath = $resolvedPath
			if (-not (Test-Path $folderPath))
			{
				Stop-PSFFunction -String 'Export-Eml.PathValidation.FolderNotExists' -StringValues $folderPath -EnableException $EnableException -ErrorRecord $_
				return
			}
		}
		#endregion Path Resolution
		
		foreach ($ewsItem in $Item)
		{
			$ewsItem.Load($propertySet)
			if ($fileName) { $exportPath = Join-Path $folderPath $fileName }
			else { $exportPath = Join-Path $folderPath "$($ewsItem.Subject).eml" }
			Write-PSFMessage -String 'Export-Eml.Exporting' -StringValues $ewsItem.Subject, $exportPath
			$ewsItem.MimeContent | Set-Content $exportPath -Encoding UTF8
		}
	}
}