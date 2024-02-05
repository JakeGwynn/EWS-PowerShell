function Remove-EwsMessage {
    param(
        [Parameter(Mandatory=$true)]
        [string]$MessageId,
        [Parameter(Mandatory=$true)]
        [string]$Mailbox,
        [Parameter(Mandatory=$true, HelpMessage="Specify 'HardDelete' for a permanent delete or 'SoftDelete' for a recoverable delete.")]
        [ValidateSet("HardDelete", "SoftDelete")]
        [string]$DeleteMode
    )

    Connect-EWS -AppId $AppId -ClientSecret $ClientSecret -TenantName $TenantName

    # Set the impersonation context
    Set-EwsImpersonation -Mailbox $Mailbox 

    # Create the message ID object
    $ItemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId($MessageId)

    # Bind to the message
    $Message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($Service, $ItemId)

    # Delete the message
    $Message.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::$DeleteMode)
}

# Example usage
# Remove-EwsMessage -Mailbox "jakegwynn@jakegwynndemo.com" -DeleteMode HardDelete -MessageId "AAMkADVmNDI3ZGMwLTE4NDItNDc5MC1hYTVhLTI3NDc0YjgzYzliNQBGAAAAAADAF99VA7tuTasP4xKUiFigBwBHqPkU3fRQS63nj4Z6h+sgAAAAAAEMAABHqPkU3fRQS63nj4Z6h+sgAAHoo6dWAAA=" 