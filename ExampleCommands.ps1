# Import the module from the USERPROFILE environment variable
Import-Module "$env:USERPROFILE\Documents\GitHub\EWS-PowerShell\EWS-PowerShell.psm1" -force

# AppId, ClientSecret, and TenantName are required to connect to EWS using OAuth
$AppId ="6de6ba97-xxxx-xxxx-xxxx-fe55607ed126"
$TenantName = "jakegwynndemo.onmicrosoft.com"

# Prompt for the client secret. This is will obfuscate the input, but the value will be stored in memory in plain text.
$ClientSecret = Read-Host "Enter the client secret of your Azure AD app registration" -AsSecureString

# Connect to EWS using OAuth. All of the commands have this built in, so it is generally not necessary to call this directly.
    Connect-EWS -AppId $AppId -TenantName $TenantName -ClientSecret $ClientSecret

# Get-EwsFolderId is a helper function that returns the FolderId object for a given folder in a mailbox
    $Folder = Get-EwsFolderId -Mailbox "jakegwynn@jakegwynndemo.com" -FolderName "Inbox" -MailboxLocation "Mailbox"
    $Folder = Get-EwsFolderId -Mailbox "jakegwynn@jakegwynndemo.com" -ParentFolderNames "Inbox","SubFolder1" -FolderName "SubFolder2" -MailboxLocation "Mailbox"
    $Folder = Get-EwsFolderId -Mailbox "jakegwynn@jakegwynndemo.com" -FolderName "Inbox" -ParentFolderId "AAMk..."

# Get-EwsFolders is a helper function that lists all folders in a mailbox or archive. It can also list all the sub-folders of a specific folder.
# Use the -PassThru switch to return the Folder objects in addition to writing them to the console.
    Get-EwsFolders -MailboxLocation "Mailbox" -Mailbox "jakegwynn@jakegwynndemo.com" -CsvExportPath "C:\Temp\Folders.csv"
    $EwsFolders = Get-EwsFolders -FolderId $Folder -Mailbox "jakegwynn@jakegwynndemo.com" -CsvExportPath "C:\Temp\Folders.csv" -PassThru

# Add-EwsMessageCategory is a helper function that adds a category to a message
    Add-EwsMessageCategory -Mailbox "jakegwynn@jakegwynndemo.com"  -Category "ExampleCategory1" -MessageId "AAMk..."

# New-EwsFolder is a helper function that creates a new folder in a mailbox
    New-EwsFolder -Mailbox "jakegwynn@jakegwynndemo.com" -ParentFolderId $Folder.Id -FolderName "NewFolder"
    New-EwsFolder -Mailbox "jakegwynn@jakegwynndemo.com" -ParentFolderNames "Inbox","SubFolder1" -FolderName "NewFolder"

# Remove-EwsMessage is a helper function that deletes a message from a mailbox
    Remove-EwsMessage -Mailbox "jakegwynn@jakegwynndemo.com" -DeleteMode HardDelete -MessageId "AAMk..."
    Remove-EwsMessage -Mailbox "jakegwynn@jakegwynndemo.com" -DeleteMode SoftDelete -MessageId "AAMk..."

# Set-EwsImpersonation is a helper function that sets the impersonation context for the current session
# This needs to be used any time you are changing the mailbox context to a different mailbox
    Set-EwsImpersonation -Mailbox "jakegwynn@jakegwynndemo.com"

# Copy-EwsMailFolder is a helper function that copies a folder from one mailbox to another
# Use the -DebugMode switch to write to get more detailed output
# Copies the source mailbox folder "Inbox\Subfolder1\SubFolder2" to the target archive folder "Inbox\Subfolder1\New-SubFolder2" 
    Copy-EwsMailFolder -SourceMailbox "jakegwynn@jakegwynndemo.com" -TargetMailbox "JakeGwynnSyncedMailbox1@jakegwynndemo.com" `
        -SourceFolderName "SubFolder2" -SourceParentFolderNames "Inbox","SubFolder1" -SourceMailboxLocation Mailbox `
        -TargetFolderName "New-SubFolder2" -TargetParentFolderNames "Inbox","SubFolder1" -TargetMailboxLocation Archive `
        -EmailDirectory "C:\Temp\EmailMigrationTest" -DebugMode

    # Copies the source folder with a specified folder ID to the target folder with a specified folder ID
    Copy-EwsMailFolder -SourceMailbox "jakegwynn@jakegwynndemo.com" -TargetMailbox "JakeGwynnSyncedMailbox1@jakegwynndemo.com" `
        -SourceFolderId "AQMkADVmNDI3ZGMwLTE4NDItNDc5MC1hYQA1YS0yNzQ3NGI4M2M5YjUALgAAA8AX31UDu25Nqw/jEpSIWKABAEeo+RTd9FBLreePhnqH6yAAAAIBDAAAAA==" `
        -TargetFolderId "AAMkADM4MmIwM2RiLTM4YWEtNDdhYS05NDE2LTExZGIxYzM3YzRhZAAuAAAAAACkux9zxCLrSK/6RJvVvD4XAQBlTYZLcpDwSrHTrU605HmwAAAAAAEMAAA=" `
        -EmailDirectory "C:\Temp\EmailMigrationTest" 

    # Copies the source archive folder "Inbox\Subfolder1\SubFolder2" to the target  folder "SubFolder2" with a specified parent folder ID
    Copy-EwsMailFolder -SourceMailbox "jakegwynn@jakegwynndemo.com" -TargetMailbox "JakeGwynnSyncedMailbox1@jakegwynndemo.com" `
        -SourceFolderName "SubFolder2" -SourceParentFolderNames "Inbox","SubFolder1" -SourceMailboxLocation Archive `
        -TargetFolderName "SubFolder2" -TargetParentFolderId "AAMkA..." `
        -MailboxLocation Archive -EmailDirectory "C:\Temp\EmailMigrationTest"
