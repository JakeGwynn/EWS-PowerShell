function Export-ImportEmails {
    param(
        [Microsoft.Exchange.WebServices.Data.Folder]$SourceFolder,
        [Microsoft.Exchange.WebServices.Data.Folder]$TargetFolder,
        [string]$EmailDir,
        [string]$SourceMailbox,
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
        $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $SourceMailbox)
        $findResults = $SourceFolder.FindItems($itemView)
        foreach ($item in $findResults) {
            # $Email = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($Service, $item.Id)
            $FileName = $EmailDir + ($item.ID -replace $Replace) + ".eml"

            Save-EmailToFile -FileName $FileName -Item $item -MailboxUPN $SourceMailbox
            
            Import-EmailToMailbox -FileName $FileName -TargetFolderId $TargetFolder.Id -MailboxUPN $TargetMailbox

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
        $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $SourceMailbox)
        $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
        $findFolderResults = $Service.FindFolders($SourceFolder.Id, $folderView)
    } catch {}

    foreach ($subFolder in $findFolderResults) {
        $subDir = Join-Path -Path $EmailDir -ChildPath $subFolder.DisplayName
        if (-not (Test-Path -Path $subDir)) {
            New-Item -ItemType Directory -Path $subDir | Out-Null
        }

        $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetMailbox)
        try{
            $TargetSubFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($Service)
            $TargetSubFolder.DisplayName = $subFolder.DisplayName
            $TargetSubFolder.Save($TargetFolder.Id)
        } catch {}

        Log-Message "Processing subfolder: $($subFolder.DisplayName)"

        $TarFolderId = Get-FolderId -FolderName $subFolder.DisplayName -MailboxUPN $TargetMailbox -ParentFolderId $TargetFolder.Id
        $TarFolderObject = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $TarFolderId)

        Export-ImportEmails -SourceFolder $subFolder -TargetFolder $TarFolderObject -EmailDir $subDir -SourceMailbox $SourceMailbox -TargetMailbox $TargetMailbox
    }
}