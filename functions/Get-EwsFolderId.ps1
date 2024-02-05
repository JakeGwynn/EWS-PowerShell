function Get-EwsFolderId {
    [CmdletBinding(DefaultParameterSetName="ByMailboxLocation")]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderName,

        [Parameter(Mandatory=$false, ParameterSetName="ByMailboxLocation")]
        [string[]]$ParentFolderNames = @(),

        [Parameter(Mandatory=$true)]
        [string]$Mailbox,

        [Parameter(Mandatory=$true, ParameterSetName="ByMailboxLocation", HelpMessage="Specify 'Mailbox' for the main mailbox or 'Archive' for the archive mailbox.")]
        [ValidateSet("Mailbox", "Archive")]
        [string]$MailboxLocation,

        [Parameter(Mandatory=$false, ParameterSetName="ByParentFolderId")]
        [string]$ParentFolderId
    )

    function Get-FolderFromHierarchy {
        param(
            [Parameter(Mandatory=$true)]
            [string[]]$ParentFolderNames,
            [Parameter(Mandatory=$true)]
            [Microsoft.Exchange.WebServices.Data.Folder]$Root
        )
        # Connect to EWS using OAuth if not already connected or near timeout. This is necessary because the connection is lost when the OAuth token expires.
        Connect-EWS -TenantName $TenantName -AppId $AppId -ClientSecret $ClientSecret
    
        # Traverse the folder hierarchy
        $CurrentFolder = $Root
        foreach ($DisplayName in $ParentFolderNames) {
            if ($CurrentFolder -ne $null) {
                $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $DisplayName)
                $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
                $FindFolderResults = $Service.FindFolders($CurrentFolder.Id, $SearchFilter, $FolderView)
                if ($FindFolderResults.TotalCount -eq 0) {
                    Log-Message "Folder '$DisplayName' not found in folder hierarchy" -MessageType "Warning"
                }
                $CurrentFolder = $FindFolderResults.Folders[0]
            }
        }
        return $CurrentFolder
    }

    # Connect to EWS using OAuth if not already connected or near timeout. This is necessary because the connection is lost when the OAuth token expires.
    Connect-EWS -TenantName $TenantName -AppId $AppId -ClientSecret $ClientSecret

    # Set the impersonation context
    Set-EwsImpersonation -Mailbox $Mailbox 

    # Set the root folder
    $RootId = if ($PSCmdlet.ParameterSetName -eq "ByMailboxLocation") {
        # Set the mailbox location
        $WellKnownFolderName = if ($MailboxLocation -eq "Mailbox") { "MsgFolderRoot" } else { "ArchiveMsgFolderRoot" }
        New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$WellKnownFolderName, $Mailbox)
    } else {
        # Use the provided parent folder ID
        New-Object Microsoft.Exchange.WebServices.Data.FolderId($ParentFolderId)
    }
    $Root = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $RootId)

    # Traverse the folder hierarchy if parent folders are specified
    if ($ParentFolderNames) {
        $Root = Get-FolderFromHierarchy -ParentFolderNames $ParentFolderNames -Root $Root
    }

    # Find the target folder
    if ($Root -eq $null) {
        Log-Message "Parent folder not found"
    }
    else {
        $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $FolderName)
        $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
        $FindFolderResults = $Service.FindFolders($Root.Id, $SearchFilter, $FolderView)
        if ($FindFolderResults.TotalCount -eq 0) {
            Log-Message "Folder '$FolderName' not found"
        }
        return $FindFolderResults.Folders[0].Id
    }
}
