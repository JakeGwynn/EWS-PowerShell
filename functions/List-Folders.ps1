function List-Folders {
    [CmdletBinding(DefaultParameterSetName="ByMailboxLocation")]
    param(
        [Parameter(Mandatory=$true, ParameterSetName="ByMailboxLocation")]
        [ValidateSet("Mailbox", "Archive")]
        [string]$MailboxLocation,

        [Parameter(Mandatory=$false)]
        [string]$MailboxUPN,

        [Parameter(Mandatory=$true, ParameterSetName="ById")]
        [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId,

        [Parameter(Mandatory=$false)]
        [int]$Depth = 0
    )

    Connect-EWS -AppId $AppId -ClientSecret $ClientSecret -TenantName $TenantName

    $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxUPN) 

    # Determine the root folder based on the parameter set
    $rootFolder = if ($PSCmdlet.ParameterSetName -eq "ByMailboxLocation") {
        # Translate the MailboxLocation parameter to the appropriate WellKnownFolderName
        $wellKnownFolderName = if ($MailboxLocation -eq "Mailbox") { "MsgFolderRoot" } else { "ArchiveMsgFolderRoot" }
        $rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wellKnownFolderName, $MailboxUPN)
        [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $rootFolderId)
    } else {
        [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $FolderId)
    }

    $fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
    $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow

    do {
        $findFolderResults = $service.FindFolders($rootFolder.Id, $fvFolderView)
        foreach ($folder in $findFolderResults.Folders) {
            $Spaces = " " * $Depth
            Write-Host $Spaces $folder.DisplayName
            if ($folder.ChildFolderCount -gt 0) {
                List-Folders -FolderId $folder.Id -Depth ($Depth + 3)
            }
        }
        $fvFolderView.Offset += $findFolderResults.Folders.Count
    } while ($findFolderResults.MoreAvailable)
}

# Call the function to list all folders
# List-Folders -MailboxLocation "Mailbox" -MailboxUPN "jakegwynn@jakegwynndemo.com"