function New-ExchangeServiceObject {
    param(
        [Parameter(Mandatory=$true)]
        [string]$AccessToken
    )

    # Create Exchange Service Object  
    $Service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
    $Service.Url= new-object Uri("https://outlook.office365.com/EWS/Exchange.asmx")

    # Set the credentials of the service to the obtained OAuth token
    $Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($AccessToken)

    return $Service
}