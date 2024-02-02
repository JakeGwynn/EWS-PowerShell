function Connect-EWS {
    param(
        [Parameter(Mandatory=$false)]
        [string]$TenantName,

        [Parameter(Mandatory=$false)]
        [string]$AppId,

        [Parameter(Mandatory=$false)]
        [string]$ClientSecret
    )

    if (!$global:AuthTimeout -or ((Get-Date) - $global:AuthTimeout).TotalMinutes -ge 55) {

        Log-Message "Connecting to EWS using OAuth" -Output 

        # If the parameters are not passed, prompt the user for them
        if (!$TenantName) {
            $TenantName = Read-Host "Enter the tenant name (e.g. jakegwynndemo.onmicrosoft.com)"
        }
        if (!$AppId) {
            $AppId = Read-Host "Enter the application ID"
        }
        if (!$ClientSecret) {
            $ClientSecret = Read-Host "Enter the client secret"
        }

        $Scope = "https://outlook.office365.com/.default"
        $Url = "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token"

        # Create body
        $Body = @{
            client_id = $AppId
            client_secret = $ClientSecret
            scope = $Scope
            grant_type = 'client_credentials'
        }

        # Splat the parameters for Invoke-Restmethod for cleaner code
        $PostSplat = @{
            ContentType = 'application/x-www-form-urlencoded'
            Method = 'POST'
            Body = $Body
            Uri = $Url
        }
        
        try {
            # Request the token
            $Request = Invoke-RestMethod @PostSplat

            # Set the global AuthTimeout variable
            $global:AuthTimeout = Get-Date

            # Create the Exchange Service object with Oauth creds
            $global:Service = New-ExchangeServiceObject -AccessToken $Request.access_token
        } catch {
            Log-Message "Error connecting to EWS using OAuth" -MessageType "Error" -Output
            $_ | Format-List
            throw
        }
    }
}