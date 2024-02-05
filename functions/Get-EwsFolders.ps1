function Get-EwsFolders {
    [CmdletBinding(DefaultParameterSetName="ByMailboxLocation")]
    param(
        [Parameter(Mandatory=$true, ParameterSetName="ByMailboxLocation")]
        [ValidateSet("Mailbox", "Archive")]
        [string]$MailboxLocation,

        [Parameter(Mandatory=$false)]
        [string]$Mailbox,

        [Parameter(Mandatory=$true, ParameterSetName="ById")]
        [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId,

        [Parameter(Mandatory=$false)]
        [string]$CsvExportPath,

        [Parameter(Mandatory=$false)]
        [int]$Depth = 0,

        [Parameter(Mandatory=$false)]
        [switch]$PassThru
    )

    Connect-EWS -AppId $AppId -ClientSecret $ClientSecret -TenantName $TenantName

    if ($Depth -eq 0) {
        Set-EwsImpersonation -Mailbox $Mailbox 
        
        # Create empty generic list to store the folder hierarchy
        $FolderHierarchy = New-Object System.Collections.Generic.List[PSObject]
    }

    # Determine the root folder based on the parameter set
    $rootFolder = if ($PSCmdlet.ParameterSetName -eq "ByMailboxLocation") {
        # Translate the MailboxLocation parameter to the appropriate WellKnownFolderName
        $wellKnownFolderName = if ($MailboxLocation -eq "Mailbox") { "MsgFolderRoot" } else { "ArchiveMsgFolderRoot" }
        $rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wellKnownFolderName, $Mailbox)
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
                Get-EwsFolders -FolderId $folder.Id -Depth ($Depth + 3)
            }

            $FolderHierarchy.Add([pscustomobject]@{
                FolderName = $folder.DisplayName
                FolderId = $folder.Id.UniqueId
                ParentFolderId = $folder.ParentFolderId.UniqueId
            })

            #$ChildFolders

            #Write-Host "Count: $($ChildFolders.Count)"
            #$ChildFolders

            #if ($ChildFolders.Count -eq 1) {
            #    $FolderHierarchy.Add($ChildFolders)
            #} elseif ($ChildFolders) {
            #    $FolderHierarchy.Add($ChildFolders)
            #}
        }
        $fvFolderView.Offset += $findFolderResults.Folders.Count
    } while ($findFolderResults.MoreAvailable)

    if ($CsvExportPath -ne $null -and $Depth -eq 0) {
        $FolderHierarchy | Export-Csv -Path $CsvExportPath -NoTypeInformation
    }
    
    if ($PassThru) {
        return $FolderHierarchy
    } 
}

# Call the function to list all folders
# Get-EwsFolders -MailboxLocation "Mailbox" -Mailbox "jakegwynn@jakegwynndemo.com"