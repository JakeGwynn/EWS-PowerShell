function Import-EmailToMailbox {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FileName,

        [Parameter(Mandatory=$true)]
        [Microsoft.Exchange.WebServices.Data.FolderId]$TargetFolderId,

        [Parameter(Mandatory=$true)]
        [string]$MailboxUPN
    )
    try {
        # Connect to EWS using OAuth if not already connected or near timeout. This is necessary because the connection is lost when the OAuth token expires.
        Connect-EWS -TenantName $TenantName -AppId $AppId -ClientSecret $ClientSecret

        # Set Impersonation
        $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxUPN)
        
        # Create the email message
        $EmailMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($Service)
        $MimeContent = New-Object Microsoft.Exchange.WebServices.Data.MimeContent
        $MimeContent.CharacterSet = "UTF-8"
        $MimeContent.Content = [System.IO.File]::ReadAllBytes($FileName)
        $EmailMessage.MimeContent = $MimeContent

        # Set the message flags to mark the message as read if it was read in the source mailbox
        if ($EmailMessage.IsRead -eq $true) {
            $PR_MESSAGE_FLAGS_msgflag_read = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3591, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
            $EmailMessage.SetExtendedProperty($PR_MESSAGE_FLAGS_msgflag_read, 1)    
        }
        $EmailMessage.Save($TargetFolderId)
        Remove-Item $FileName
    }
    catch {
        Log-Message "Error importing email to mailbox:" -Output -MessageType "Error"
        $_ | Format-List
    }
}