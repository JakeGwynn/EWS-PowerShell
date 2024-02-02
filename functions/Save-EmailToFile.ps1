function Save-EmailToFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FileName,

        [Parameter(Mandatory=$true)]
        [object]$Item,

        [Parameter(Mandatory=$true)]
        [string]$MailboxUPN
    )
    try {
        # Connect to EWS using OAuth if not already connected or near timeout. This is necessary because the connection is lost when the OAuth token expires.
        Connect-EWS -TenantName $TenantName -AppId $AppId -ClientSecret $ClientSecret

        # Set Impersonation
        $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxUPN)
        
        if ($DebugMode) {
            # Log-Message "Saving email to file: $($Item.Id)"
        }
        
        # Create a property set to include the MIME content
        $psPropset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
        $Item.Load($psPropset)

        # Add the isRead property to the property set if the email was read in the source mailbox
        if ($Item.IsRead -eq $true) {
            $PR_MESSAGE_FLAGS_msgflag_read = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3591, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
            $Item.SetExtendedProperty($PR_MESSAGE_FLAGS_msgflag_read, 1)    
        }

        # Save the MIME content to a file
        $Email = New-Object System.IO.FileStream($FileName, [System.IO.FileMode]::Create)
        $Email.Write($Item.MimeContent.Content, 0, $Item.MimeContent.Content.Length)
        $Email.Close()
    }
    catch {
        Log-Message "Error saving email to file: $_" -Output -MessageType "Error"
    }
}