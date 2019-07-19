# This is where the strings go, that are written by
# Write-PSFMessage, Stop-PSFFunction or the PSFramework validation scriptblocks
@{
	'Connect-EwsExchange.FailedAutodetect.Email' = 'Failed to auto-detect email address for {0}' # $windowsIdentity
	'Connect-EwsExchange.ConnectionStart'	     = 'Connecting to Exchange for {0} in language mode {1}' # $Mailbox, $Version
	'Connect-EwsExchange.AuthenticatingAs'	     = 'Authenticating under alternate credentials as: {0}' # $Credential.UserName
	'Connect-EwsExchange.AccessingMailbox'	     = 'Performing autodiscover as {0}' # $Mailbox
	'Connect-EwsExchange.FailedAutodetect'	     = 'Failed to perform autodiscovery for {0}' # $Mailbox
	'Connect-EwsExchange.Impersonating'		     = 'Impersonating user: {0}' # $Impersonate
	'Connect-EwsExchange.ConnectionSuccess'	     = 'Successfully connected to {0}' # $Mailbox
	'Connect-EwsExchange.ConnectionFailed'	     = 'Failed to connect to {0}' # $Mailbox
	
	'Connect-Outlook.Existing'				     = 'Connecting to running Outlook application' # 
	'Connect-Outlook.Existing.Failed'		     = 'Failed to connect to running Outlook application' # 
	'Connect-Outlook.NewComObject'			     = 'Starting a new Outlook application' # 
	'Connect-Outlook.NewComObject.Failed'	     = 'Failed to start a new Outlook application' # 
	
	'Convert-MsgToEml.Path.NotFound'			 = 'Input file not found: {0}' # $fileItem
	'Convert-MsgToEml.Path.NotMsg'			     = 'Input file is not a msg file: {0}' # $fileItem
	'Convert-MsgToEml.Importing'				 = 'Importing file into Outlook Drafts folder: {0}' # $fileItem
	'Convert-MsgToEml.WaitingForSync'		     = 'Waiting for message to sync to Exchange: {0}' # $fileItem
	'Convert-MsgToEml.Convert.TimedOut'		     = 'Conversion timed out: {0}' # $outlookItem.Subject
	'Convert-MsgToEml.Exporting'				 = 'Exporting email: "{0}" to {1}' # $outlookItem.Subject, $OutPath
	
	'Export-Eml.ResolvePath.Failed'			     = 'Could not resolve export path: {0}' # $Path
	'Export-Eml.PathValidation.FolderNotExists'  = 'Export folder does not exist: {0}' # $folderPath
	'Export-Eml.Exporting'					     = 'Exporting email "{0}" to {1}' # $ewsItem.Subject, $exportPath
	
	'Get-EwsFolder.ConnectingSearchBase'		 = 'Connecting to mailbox at base folder: {0}' # $SearchBase
	'Get-EwsFolder.ConnectionFailed'			 = 'Failed to connect to mailbox at base folder: {0}' # $SearchBase
	'Get-EwsFolder.LoadingFolder'			     = 'Performing folder search' # 
	
	'Get-EwsMail.RetrievingFromFolder'		     = 'Retrieving emails from {0}' # $resolvedFolderItem.DisplayName
	
	'Import-Msg.Folder.ImportingInto'		     = 'Destination folder for imports: {0}' # $Folder
	'Import-Msg.PathResolution.Failed'		     = 'Failed to resolve path: {0}' # $pathItem
	'Import-Msg.Message.Importing'			     = 'Importing msg: {0}' # $resolvedPath
	
	'Assert-EwsConnected.Failed'				 = 'Failed to execute {0} due to lacking an EWS connection. Please run Connect-EwsExchange first.' # $Cmdlet.MyInvocation.MyCommand.Name
	
	'Assert-OutlookConnected.Failed'			 = 'Failed to execute {0} due to lacking an Outlook connection. Please run Connect-Outlook first.' # $Cmdlet.MyInvocation.MyCommand.Name
	
	'Validate.Container'						 = 'Not an existing folder: {0}'
}