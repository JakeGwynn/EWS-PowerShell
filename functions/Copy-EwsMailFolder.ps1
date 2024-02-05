function Copy-EwsMailFolder {
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

        [Parameter(Mandatory=$false, HelpMessage="Outputs the message to the console.")]
        [switch]$DebugMode
    )

    

    function Export-ImportEmails {
        param(
            [Parameter(Mandatory=$true, HelpMessage="The source folder object.")]
            [Microsoft.Exchange.WebServices.Data.Folder]$SourceFolderId,
    
            [Parameter(Mandatory=$true, HelpMessage="The target folder object.")]
            [Microsoft.Exchange.WebServices.Data.Folder]$TargetFolder,
    
            [Parameter(Mandatory=$true, HelpMessage="The directory to save the emails to temporarily while exporting.")]
            [string]$EmailDir,
    
            [Parameter(Mandatory=$true, HelpMessage="The email address of the source mailbox.")]
            [string]$SourceMailbox,
    
            [Parameter(Mandatory=$true, HelpMessage="The email address of the target mailbox.")]
            [string]$TargetMailbox
        )
    
        Connect-EWS -AppId $AppId -ClientSecret $ClientSecret -TenantName $TenantName
    
        # Define the item view
        $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
        $ItemView.PropertySet = $PropertySet
    
        $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $TotalEmails = $SourceFolder.TotalCount
    
        Log-Message "Processing folder: $($SourceFolder.DisplayName)" -Output
        Log-Message "Total emails in folder: $TotalEmails" -Output
    
        $InvalidCharacters = [IO.Path]::GetInvalidFileNameChars() -join ''
        $Replace = "[{0}]" -f [RegEx]::Escape($InvalidCharacters)
    
        $CopiedEmails = 0
        do {
            # Set Impersonation
            Log-Message "Processing batch $([math]::Ceiling($itemView.Offset / 1000) + 1)" -Output
            Set-EwsImpersonation -Mailbox $SourceMailbox
            $findResults = $SourceFolder.FindItems($itemView)
            foreach ($item in $findResults) {
                $FileName = $EmailDir + ($item.ID -replace $Replace) + ".eml"
                try {
                    Export-EmailToFile -FileName $FileName -Item $item -Mailbox $SourceMailbox
                    Import-EmailToMailbox -FileName $FileName -TargetFolderId $TargetFolder.Id -Mailbox $TargetMailbox
                    $CopyStatus.Add([PSCustomObject]@{
                        "ID" = $item.ID
                        "Status" = "Success"
                    }) | Out-Null
                } catch {
                    $CopyStatus.Add([PSCustomObject]@{
                        "ID" = $item.ID
                        "Status" = "Error"
                    }) | Out-Null

                    Log-Message "Error copying email:" -Output -MessageType "Error"
                    $_ | Format-List
                }

    
                $CopiedEmails++
                #Change progress bar to calculate seconds remaining based on how long it has taken so far. Same with percent complete
    
                $ElapsedSeconds = $Stopwatch.Elapsed.TotalSeconds
                $EstimatedTotalSeconds = $ElapsedSeconds / $CopiedEmails * $TotalEmails
                $EstimatedSecondsRemaining = $EstimatedTotalSeconds - $ElapsedSeconds
                
                Write-Progress -SecondsRemaining $EstimatedSecondsRemaining -Activity "Copying emails" -Status "Copied $CopiedEmails of $TotalEmails emails" -PercentComplete ($CopiedEmails / $TotalEmails * 100)
            }
            $itemView.Offset += $findResults.Items.Count
            Log-Message "Copied $CopiedEmails of $($SourceFolder.TotalCount) emails" -Output
        } while ($findResults.MoreAvailable)
    
        # Process subfolders
        try {
            Set-EwsImpersonation -Mailbox $SourceMailbox
            $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
            $findFolderResults = $Service.FindFolders($SourceFolder.Id, $folderView)
        } catch {}
    
        foreach ($subFolder in $findFolderResults) {
            $subDir = Join-Path -Path $EmailDir -ChildPath $subFolder.DisplayName
            if (-not (Test-Path -Path $subDir)) {
                New-Item -ItemType Directory -Path $subDir | Out-Null
            }
    
            Set-EwsImpersonation -Mailbox $TargetMailbox
            try{
                $TargetSubFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($Service)
                $TargetSubFolder.DisplayName = $subFolder.DisplayName
                $TargetSubFolder.Save($TargetFolder.Id)
            } catch {}
    
            Log-Message "Processing subfolder: $($subFolder.DisplayName)"
    
            $TarFolderId = Get-EwsFolderId -FolderName $subFolder.DisplayName -Mailbox $TargetMailbox -ParentFolderId $TargetFolder.Id
            $TarFolderObject = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $TarFolderId)
    
            Export-ImportEmails -SourceFolder $subFolder -TargetFolder $TarFolderObject -EmailDir $subDir -SourceMailbox $SourceMailbox -TargetMailbox $TargetMailbox
        }
    }

    

    $ScriptStartTime = Get-Date -Format "yyyy-MM-dd HH.mm.ss"

    # Create generic list to store the copy status of each item
    $CopyStatus = New-Object System.Collections.Generic.List[PSObject]

    if ($DebugMode) {
        Start-Transcript -Path "$EmailDirectory\Copy-EwsMailFolder_Debug_$ScriptStartTime.log" -Append
    }

    $FullSourceFolderName = ($SourceParentFolderNames -join "\") + "\$SourceFolderName"
    $FullTargetFolderName = ($TargetParentFolderNames -join "\") + "\$TargetFolderName"

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
        Set-EwsImpersonation -Mailbox $SourceMailbox
        New-Object Microsoft.Exchange.WebServices.Data.FolderId($SourceFolderId)
    } elseif ($SourceParentFolderNames) {
        Log-Message "Getting source folder by name"
        Get-EwsFolderId -FolderName $SourceFolderName -Mailbox $SourceMailbox -MailboxLocation $SourceMailboxLocation -ParentFolderNames $SourceParentFolderNames
    } else {
        Log-Message "Getting source folder from root of $SourceMailboxLocation"
        Get-EwsFolderId -FolderName $SourceFolderName -Mailbox $SourceMailbox -MailboxLocation $SourceMailboxLocation
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
        Set-EwsImpersonation -Mailbox $TargetMailbox
        New-Object Microsoft.Exchange.WebServices.Data.FolderId($TargetFolderId)
    } elseif ($TargetParentFolderId) {
        Log-Message "Getting target folder by Parent Folder ID"
        Get-EwsFolderId -FolderName $TargetFolderName -Mailbox $TargetMailbox -ParentFolderId $TargetParentFolderId
    } elseif ($TargetParentFolderNames) {
        Log-Message "Getting target folder by name"
        Get-EwsFolderId -FolderName $TargetFolderName -Mailbox $TargetMailbox -MailboxLocation $TargetMailboxLocation -ParentFolderNames $TargetParentFolderNames
    } else {
        Log-Message "Getting target folder from root of $TargetMailboxLocation"
        Get-EwsFolderId -FolderName $TargetFolderName -Mailbox $TargetMailbox -MailboxLocation $TargetMailboxLocation
    }

    # Create the target folder if it doesn't exist
    try {
        if ($TargetFolderIdObject -eq $null) {
            Log-Message "Target folder not found" -Output
            Log-Message "Creating target folder $TargetFolderName" -Output
            $TargetFolder = if ($TargetParentFolderId) {
                New-EwsFolder -FolderName $TargetFolderName -Mailbox $TargetMailbox -ParentFolderId $TargetParentFolderId
            } elseif ($TargetParentFolderNames) {
                New-EwsFolder -FolderName $TargetFolderName -Mailbox $TargetMailbox -MailboxLocation $TargetMailboxLocation -ParentFolderNames $TargetParentFolderNames
            } else {
                New-EwsFolder -FolderName $TargetFolderName -Mailbox $TargetMailbox -MailboxLocation $TargetMailboxLocation
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

    Log-Message "Script finished. Log file is: $("$EmailDirectory\Copy-EwsMailFolder_$ScriptStartTime.log")" -Output -MessageType "Success"

    $CopyStatus | Export-Csv -Path "$EmailDirectory\Copy-EwsMailFolder_$ScriptStartTime.csv" -NoTypeInformation

    if ($DebugMode) {
        Stop-Transcript
    }
}