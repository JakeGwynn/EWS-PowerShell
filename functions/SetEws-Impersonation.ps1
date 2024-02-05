function Set-EwsImpersonation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Mailbox
    )

    if ($Service.ImpersonatedUserId.Id -ne $Mailbox) {
        $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox)
    }
}