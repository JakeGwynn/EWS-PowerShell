function Add-EwsMessageCategory {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Mailbox,
        [Parameter(Mandatory=$true)]
        [string]$MessageId,
        [Parameter(Mandatory=$true)]
        [string]$Category
    )

    Connect-EWS -AppId $AppId -ClientSecret $ClientSecret -TenantName $TenantName

    Set-EwsImpersonation -Mailbox $Mailbox 

    # Create the ItemId object
    $ItemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId($MessageId)

    # Bind to the message
    $Message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $ItemId)

    # Update the category
    $Message.Categories.Add($Category)
    $Message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)

    Write-Host "Updated category to '$Category' for message with ID '$MessageId'."
}

# Example usage
# Update-MessageCategory -Mailbox "jakegwynn@jakegwynndemo.com" -Category "ExampleCategory1" -MessageId "AAMkADQ3YjY1YjJkLWRkNTItNGNjMy1hZDljLTFmNTFlMTlkOTc3OABGAAAAAABOcP6dsp+pR6aGc6APCrSqBwAY39mUpRTDSaV15xLLXzHeAAAAAAE0AAAY39mUpRTDSaV15xLLXzHeAAADct4TAAA=" 