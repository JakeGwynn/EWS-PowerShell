function Delete-Message {
    param(
        [Parameter(Mandatory=$true)]
        [string]$MessageId,
        [Parameter(Mandatory=$true)]
        [string]$MailboxUPN,
        [Parameter(Mandatory=$true, HelpMessage="Specify 'HardDelete' for a permanent delete or 'SoftDelete' for a recoverable delete.")]
        [ValidateSet("HardDelete", "SoftDelete")]
        [string]$DeleteMode
    )

    Connect-EWS -AppId $AppId -ClientSecret $ClientSecret -TenantName $TenantName

    # Set the impersonation context
    $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxUPN) 

    # Create the message ID object
    $ItemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId($MessageId)

    # Bind to the message
    $Message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($Service, $ItemId)

    # Delete the message
    $Message.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::$DeleteMode)
}

# Example usage
# Delete-Message -MailboxUPN "jakegwynn@jakegwynndemo.com" -DeleteMode HardDelete -MessageId "AAMkADVmNDI3ZGMwLTE4NDItNDc5MC1hYTVhLTI3NDc0YjgzYzliNQBGAAAAAADAF99VA7tuTasP4xKUiFigBwBHqPkU3fRQS63nj4Z6h+sgAAAAAAEMAABHqPkU3fRQS63nj4Z6h+sgAAHoo6dWAAA=" 