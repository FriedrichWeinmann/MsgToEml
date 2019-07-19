function Import-Msg
{
<#
	.SYNOPSIS
		Imports a .msg file into Outlook.
	
	.DESCRIPTION
		Imports a .msg file into Outlook.
		Requires a live connection to Outlook by using Connect-Outlook.
		(Note: Outlook, the application, not Outlook.com)
	
	.PARAMETER Path
		The path to the file to import.
	
	.PARAMETER Folder
		The well-known folder in Outlook to import into.
		Note: At the moment, custom folders are NOT supported.
	
	.PARAMETER EnableException
		This parameters disables user-friendly warnings and enables the throwing of exceptions.
		This is less user friendly, but allows catching exceptions in calling scripts.
	
	.EXAMPLE
		PS C:\> Get-ChildItem *.msg | Import-Msg
	
		Imports all msg files in the current folder into the user's mailbox.
#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[Alias('FullName')]
		[string[]]
		$Path,
		
		[ValidateSet('AllPublicFolders', 'Calendar', 'Conflicts', 'Contacts', 'DeletedItems', 'Drafts', 'Inbox', 'Journal', 'Junk', 'LocalFailures', 'ManagedEmail', 'Notes', 'Outbox', 'RssFeeds', 'SentMail', 'ServerFailures', 'SuggestedContacts', 'SyncIssues', 'Tasks', 'ToDo')]
		[string]
		$Folder = 'Inbox',
		
		[switch]
		$EnableException
	)
	
	begin
	{
		Assert-OutlookConnected -Cmdlet $PSCmdlet
		
		#region Resolve Folder
		$folderMapping = @{
			DeletedItems	  = 3 # The Deleted Items folder.
			Outbox		      = 4 # The Outbox folder.
			SentMail		  = 5 # The Sent Mail folder.
			Inbox			  = 6 # The Inbox folder.
			Calendar		  = 9 # The Calendar folder.
			Contacts		  = 10 # The Contacts folder.
			Journal		      = 11 # The Journal folder.
			Notes			  = 12 # The Notes folder.
			Tasks			  = 13 # The Tasks folder.
			Drafts		      = 16 # The Drafts folder.
			AllPublicFolders  = 18 # The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.
			Conflicts		  = 19 # The Conflicts folder (subfolder of Sync Issues folder). Only available for an Exchange account.
			SyncIssues	      = 20 # The Sync Issues folder. Only available for an Exchange account.
			LocalFailures	  = 21 # The Local Failures folder (subfolder of Sync Issues folder). Only available for an Exchange account.
			ServerFailures    = 22 # The Server Failures folder (subfolder of Sync Issues folder). Only available for an Exchange account.
			Junk			  = 23 # The Junk E-Mail folder.
			RssFeeds		  = 25 # The RSS Feeds folder.
			ToDo			  = 28 # The To Do folder.
			ManagedEmail	  = 29 # The top-level folder in the Managed Folders group. For more information on Managed Folders, see Help in Outlook. Only available for an Exchange account.
			SuggestedContacts = 30 # The Suggested Contacts folder.
		}
		$namespace = $script:Outlook.GetNamespace('MAPI')
		$folderObject = $namespace.GetDefaultFolder($folderMapping[$Folder])
		Write-PSFMessage -String 'Import-Msg.Folder.ImportingInto' -StringValues $Folder
		#endregion Resolve Folder
	}
	process
	{
		foreach ($pathItem in $Path)
		{
			try { $resolvedPaths = Resolve-PSFPath -Path $pathItem -Provider FileSystem }
			catch { Stop-PSFFunction -String 'Import-Msg.PathResolution.Failed' -StringValues $pathItem -ErrorRecord $_ -EnableException $EnableException -Continue }
			
			foreach ($resolvedPath in $resolvedPaths)
			{
				Write-PSFMessage -String 'Import-Msg.Message.Importing' -StringValues $resolvedPath
				$mailItem = $namespace.OpenSharedItem($resolvedPath)
				$newItem = $mailItem.Move($folderObject)
				[pscustomobject]@{
					Subject = $newItem.Subject
					Size    = $newItem.Size
					ReceivedOn = $newItem.ReceivedTime
					From    = $newItem.From
					To	    = $newItem.To
				}
			}
		}
	}
}