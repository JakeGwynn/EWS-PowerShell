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