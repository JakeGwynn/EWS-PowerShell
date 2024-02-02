function Copy-MailFolder {
    [cmdletbinding(DefaultParameterSetName="BySourceFolderName")]
    param(
        [Parameter(Mandatory=$true, HelpMessage="The email address of the source mailbox.")]
        [string]$SourceMailbox,

        [Parameter(Mandatory=$true, HelpMessage="The email address of the target mailbox.")]
        [string]$TargetMailbox,

        [Parameter(Mandatory=$false, HelpMessage="The name of the folder in the source mailbox to copy.")]
        [string]$SourceFolderName,

        [Parameter(Mandatory=$false, HelpMessage="Specify 'Mailbox' for the main mailbox or 'Archive' for the archive mailbox.")]
        [ValidateSet("Mailbox", "Archive")]
        [string]$SourceMailboxLocation,

        [Parameter(Mandatory=$false, HelpMessage="The names of the parent folders in the source mailbox in a comma-separated list. Leave empty if the folder is in the root.")]
        [string[]]$SourceParentFolderNames = @(),

        [Parameter(Mandatory=$false, HelpMessage="The ID of the folder in the source mailbox.")]
        [string]$SourceFolderId,

        [Parameter(Mandatory=$false, HelpMessage="The name of the folder in the target mailbox to copy.")]
        [string]$TargetFolderName,

        [Parameter(Mandatory=$false, HelpMessage="Specify 'Mailbox' for the main mailbox or 'Archive' for the archive mailbox.")]
        [ValidateSet("Mailbox", "Archive")]
        [string]$TargetMailboxLocation,

        [Parameter(Mandatory=$false, HelpMessage="The names of the parent folders in the target mailbox. Leave empty if the folder is in the root.")]
        [string[]]$TargetParentFolderNames = @(),

        [Parameter(Mandatory=$false, HelpMessage="The ID of the folder in the target mailbox.")]
        [string]$TargetFolderId,

        [Parameter(Mandatory=$false, HelpMessage="The ID of the parent folder in the target mailbox.")]
        [string]$TargetParentFolderId,

        [Parameter(Mandatory=$true, HelpMessage="The directory to save the emails to. Leave empty to not save emails.")]
        [string]$EmailDirectory,

        [Parameter(Mandatory=$false, HelpMessage="The path to the EWS Managed API DLL file. Leave empty to use the default path.")]
        [string]$EWSManagedAPIPath,

        [Parameter(Mandatory=$false, HelpMessage="The App ID for the Azure AD app.")]
        [string]$AppId,

        [Parameter(Mandatory=$false, HelpMessage="The client secret for the Azure AD app.")]
        [string]$ClientSecret,

        [Parameter(Mandatory=$false, HelpMessage="Outputs the message to the console.")]
        [switch]$DebugMode,

        [Parameter(Mandatory=$true, HelpMessage="The xxx.onmicrosoft.com name of the tenant. Example: jakegwynndemo.onmicrosoft.com.")]
        [string]$TenantName

    )

    <#
            #Usage:

            #.\Copy-MailFolder.ps1 -SourceMailbox "SourceMailboxName" -TargetMailbox "TargetMailboxName" -TenantName "TenantName" -EmailDirectory "C:\Emails" -MailboxLocation "Mailbox" -AppId "YourAppId" -ClientSecret "YourClientSecret"

            $SourceMailbox = "jakegwynn@jakegwynndemo.com"
            $TargetMailbox = "JakeGwynnSyncedMailbox1@jakegwynndemo.com"
            $TenantName = "jakegwynndemo.onmicrosoft.com"
            $EmailDirectory = "C:\Temp\EmailMigrationTest"
            $EWSManagedAPIPath = "C:\users\jakeupwork\desktop\Microsoft.Exchange.WebServices.dll"
            $AppId = "YourAppId"
            $ClientSecret = "YourClientSecret"

            # 
            $SourceMailboxLocation = "Mailbox"
            $SourceFolderName = "Inbox"
            $SourceParentFolderNames = "Inbox", "SubFolder1", "SubFolder2"

            $TargetMailboxLocation = "Mailbox"
            $TargetFolderName = "Inbox"
            $TargetParentFolderNames = "Inbox", "SubFolder1", "SubFolder2"
            $TargetParentFolderId = "AAMkADM4MmIwM2RiLTM4YWEtNDdhYS05NDE2LTExZGIxYzM3YzRhZAAuAAAAAACkux9zxCLrSK/6RJvVvD4XAQBlTYZLcpDwSrHTrU605HmwAAHG2ZSFAAA="

            $SourceParentFolderId = "AQMkADVmNDI3ZGMwLTE4NDItNDc5MC1hYQA1YS0yNzQ3NGI4M2M5YjUALgAAA8AX31UDu25Nqw/jEpSIWKABAEeo+RTd9FBLreePhnqH6yAAAAIBDAAAAA=="
            $TargetParentFolderId = "AAMkADM4MmIwM2RiLTM4YWEtNDdhYS05NDE2LTExZGIxYzM3YzRhZAAuAAAAAACkux9zxCLrSK/6RJvVvD4XAQBlTYZLcpDwSrHTrU605HmwAAAAAAEMAAA="


            $Params = @{
                SourceMailbox = $SourceMailbox
                TargetMailbox = $TargetMailbox
                TenantName = $TenantName
                EmailDirectory = $EmailDirectory
                EWSManagedAPIPath = $EWSManagedAPIPath
                AppId = $AppId
                ClientSecret = $ClientSecret

                SourceFolderName = $SourceFolderName
                #SourceParentFolderNames = $SourceParentFolderNames
                SourceMailboxLocation = $SourceMailboxLocation

                TargetFolderName = $TargetFolderName
                #TargetParentFolderNames = $TargetParentFolderNames
                TargetMailboxLocation = $TargetMailboxLocation

                #TargetFolderName = $TargetFolderName
                #TargetParentFolderId = $TargetParentFolderId

                #TargetFolderId = $TargetFolderId
                #SourceFolderId = $SourceFolderId
            }
            .\Copy-MailFolder.ps1 @Params -DebugMode
            Copy-MailFolder @Params -DebugMode

            $SourceFolderId = "AAMkADVmNDI3ZGMwLTE4NDItNDc5MC1hYTVhLTI3NDc0YjgzYzliNQAuAAAAAADAF99VA7tuTasP4xKUiFigAQBHqPkU3fRQS63nj4Z6h+sgAAHoooedAAA="
            $TargetFolderId = "AAMkADM4MmIwM2RiLTM4YWEtNDdhYS05NDE2LTExZGIxYzM3YzRhZAAuAAAAAACkux9zxCLrSK/6RJvVvD4XAQBlTYZLcpDwSrHTrU605HmwAAHG2ZSFAAA="

            .\Copy-MailFolder.ps1 -SourceMailbox $SourceMailbox -TargetMailbox $TargetMailbox -DebugMode `
                -AppId $AppId -ClientSecret $ClientSecret -TenantName $TenantName `
                -EmailDirectory $EmailDirectory -EWSManagedAPIPath $EWSManagedAPIPath `
                -SourceFolderId $SourceFolderId `
                -TargetFolderId $TargetFolderId 
        #>

    ########################### End of Function ###########################

    ########################### Main Script ###########################

    $ScriptStartTime = Get-Date -Format "yyyy-MM-dd HH.mm.ss"

    if ($DebugMode) {
        Start-Transcript -Path "$EmailDirectory\Copy-MailFolder_Debug_$ScriptStartTime.log" -Append
    }

    $FullSourceFolderName = ($SourceParentFolderNames -join "\") + "\$SourceFolderName"
    $FullTargetFolderName = ($TargetParentFolderNames -join "\") + "\$TargetFolderName"

    <#
        # Import the EWS Managed API
        Log-Message "Importing EWS Managed API from $EWSManagedAPIPath"
        try {
            Add-Type -Path $EWSManagedAPIPath
        } catch {
            Log-Message "Error importing EWS Managed API from $EWSManagedAPIPath" -MessageType "Error" -Output
            $_ | Format-List
            throw
        }
    #>


    # Set the directory to save the emails to. Create it if it doesn't exist.
    Log-Message "Setting the directory to save the emails to"
    $EmailDirectory = if ($EmailDirectory) {
        if (!(Test-Path $EmailDirectory)) {
            New-Item -ItemType Directory -Force -Path $EmailDirectory
        }
        $EmailDirectory + "\"
    } else {
        ""
    }

    # Connect to EWS using OAuth
    Connect-EWS -TenantName $TenantName -AppId $AppId -ClientSecret $ClientSecret

    Log-Message "Getting the source folder"
    Log-Message "SourceMailbox: $SourceMailbox | SourceFolderName: $FullSourceFolderName | SourceFolderId: $SourceFolderId" -Output
    Log-Message "TargetMailbox: $TargetMailbox | TargetFolderName: $FullTargetFolderName | TagetFolderId: $TargetFolderId | TargetParentFolderId: $TargetParentFolderId" -Output

    # Get the source folder
    $SourceFolderIdObject = if ($SourceFolderId) {
        Log-Message "Getting source folder by ID"
        $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $SourceMailbox)
        New-Object Microsoft.Exchange.WebServices.Data.FolderId($SourceFolderId)
    } elseif ($SourceParentFolderNames) {
        Log-Message "Getting source folder by name"
        Get-FolderId -FolderName $SourceFolderName -MailboxUPN $SourceMailbox -MailboxLocation $SourceMailboxLocation -ParentFolderNames $SourceParentFolderNames
    } else {
        Log-Message "Getting source folder from root of $SourceMailboxLocation"
        Get-FolderId -FolderName $SourceFolderName -MailboxUPN $SourceMailbox -MailboxLocation $SourceMailboxLocation
    }

    try {
        $SourceFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $SourceFolderIdObject)
    } catch {
        Log-Message "Error finding source folder. Exiting Script `r`n" -MessageType "Error" -Output
        $_ | Format-List
        throw
    }

    # Get the target folder
    Log-Message "Getting the target folder"
    $TargetFolderIdObject = if ($TargetFolderId) {
        Log-Message "Getting target folder by ID"
        $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetMailbox)
        New-Object Microsoft.Exchange.WebServices.Data.FolderId($TargetFolderId)
    } elseif ($TargetParentFolderId) {
        Log-Message "Getting target folder by Parent Folder ID"
        Get-FolderId -FolderName $TargetFolderName -MailboxUPN $TargetMailbox -ParentFolderId $TargetParentFolderId
    } elseif ($TargetParentFolderNames) {
        Log-Message "Getting target folder by name"
        Get-FolderId -FolderName $TargetFolderName -MailboxUPN $TargetMailbox -MailboxLocation $TargetMailboxLocation -ParentFolderNames $TargetParentFolderNames
    } else {
        Log-Message "Getting target folder from root of $TargetMailboxLocation"
        Get-FolderId -FolderName $TargetFolderName -MailboxUPN $TargetMailbox -MailboxLocation $TargetMailboxLocation
    }

    # Create the target folder if it doesn't exist
    try {
        if ($TargetFolderIdObject -eq $null) {
            Log-Message "Target folder not found" -Output
            Log-Message "Creating target folder $TargetFolderName" -Output
            $TargetFolder = if ($TargetParentFolderId) {
                New-Folder -FolderName $TargetFolderName -MailboxUPN $TargetMailbox -ParentFolderId $TargetParentFolderId
            } elseif ($TargetParentFolderNames) {
                New-Folder -FolderName $TargetFolderName -MailboxUPN $TargetMailbox -MailboxLocation $TargetMailboxLocation -ParentFolderNames $TargetParentFolderNames
            } else {
                New-Folder -FolderName $TargetFolderName -MailboxUPN $TargetMailbox -MailboxLocation $TargetMailboxLocation
            }
        } else {
            $TargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $TargetFolderIdObject)
        }
    } catch {
        Log-Message "Error finding target folder. Exiting Script" -MessageType "Error" -Output
        $_ | Format-List
        throw
    }

    # Define the property set to include the categories and all other first class properties
    $PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
    $PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML
    $PropertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Categories)

    Export-ImportEmails -SourceFolder $SourceFolder -TargetFolder $TargetFolder -EmailDir $EmailDirectory -SourceMailbox $SourceMailbox -TargetMailbox $TargetMailbox

    Log-Message "Finished copying folder '$SourceFolderName' from '$SourceMailbox' to '$TargetMailbox'" -Output -MessageType "Success"

    Log-Message "Script finished. Log file is: $("$EmailDirectory\Copy-MailFolder_$ScriptStartTime.log")" -Output -MessageType "Success"

    if ($DebugMode) {
        Stop-Transcript
    }
}