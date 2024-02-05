function New-EwsFolder {
    [CmdletBinding(DefaultParameterSetName="MailboxRoot")]
    param(
        [Parameter(Mandatory=$true, ParameterSetName="MailboxRoot")]
        [switch]$MailboxRoot,

        [Parameter(Mandatory=$true, ParameterSetName="ArchiveRoot")]
        [switch]$ArchiveRoot,

        [Parameter(Mandatory=$true, ParameterSetName="ByParentFolderNames")]
        [string[]]$ParentFolderNames,

        [Parameter(Mandatory=$true, ParameterSetName="ByParentFolderNames", HelpMessage="Specify 'Mailbox' for the main mailbox or 'Archive' for the archive mailbox.")]
        [ValidateSet("Mailbox", "Archive")]
        [string]$MailboxLocation,

        [Parameter(Mandatory=$true)]
        [string]$FolderName,

        [Parameter(Mandatory=$true)]
        [string]$Mailbox
    )

    # Connect to EWS using OAuth if not already connected or near timeout. This is necessary because the connection is lost when the OAuth token expires.
    Connect-EWS -TenantName $TenantName -AppId $AppId -ClientSecret $ClientSecret

    Set-EwsImpersonation -Mailbox $Mailbox 

    # Determine the parent folder based on the parameter set
    $ParentFolderId = if ($PSCmdlet.ParameterSetName -eq "MailboxRoot") {
        New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $Mailbox)
    } elseif ($PSCmdlet.ParameterSetName -eq "ArchiveRoot") {
        New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $Mailbox)
    } else {
        $WellKnownFolderName = if ($MailboxLocation -eq "Mailbox") { "MsgFolderRoot" } else { "ArchiveMsgFolderRoot" }
        $RootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$WellKnownFolderName, $Mailbox)

        $ParentFolderId = $RootFolderId
        foreach ($ParentFolderName in $ParentFolderNames) {
            $ParentFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
            $ParentFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
            $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $ParentFolderName)
            $findFolderResults = $Service.FindFolders($ParentFolderId, $searchFilter, $ParentFolderView)

            if ($findFolderResults.TotalCount -gt 0) {
                $ParentFolderId = $findFolderResults.Folders[0].Id
                Log-Message "Found parent folder: $ParentFolderNamesUsed$ParentFolderName with folder ID: $ParentFolderId" -MessageType "Success" -Output

            } else {
                Log-Message "Creating parent folder: $ParentFolderNamesUsed$ParentFolderName under parent folder ID: $ParentFolderId"
                try {
                    $NewFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($Service)
                    $NewFolder.DisplayName = $ParentFolderName
                    $NewFolder.Save($ParentFolderId)
                    $ParentFolderId = $NewFolder.Id
                    Log-Message "Created parent folder: $ParentFolderNamesUsed$ParentFolderName successfully" -MessageType "Success" -Output
                } catch {
                    Log-Message "Error creating parent folder: $ParentFolderNamesUsed$ParentFolderName. Exiting Script." -MessageType "Error" -Output
                    $_ | Format-List
                }
            }
            $ParentFolderNamesUsed += $ParentFolderName + "\"
        }
        $ParentFolderId
    }

    # Create the new folder
    try {
        Log-Message "Creating folder: $ParentFolderNamesUsed$FolderName under parent folder ID: $ParentFolderId"
        $NewFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($Service)
        $NewFolder.DisplayName = $FolderName
        $NewFolder.Save($ParentFolderId)
        Log-Message "Created folder: $FolderName" -MessageType "Success" -Output
        return $NewFolder
    } catch {
        Log-Message "Error creating folder: $FolderName. Exiting script" -MessageType "Error" -Output
        $_ | Format-List
        throw
    }
}