function Get-GraphAccessToken() {
    <#
    .SYNOPSIS
        This gets the Access Token for the Microsoft Graph API.

    .DESCRIPTION
       The function gets the AppId and Cert from the KeyVault, creates a JWT, and calls the Azure AD endpoint to get the access token.

    .PARAMETER Prefix
        The prefix to use for the resources.

    .PARAMETER Name
        The name to use for the resources. Default is "mail-send".
        
    .EXAMPLE
        Get-GraphAccessToken -Prefix "contoso" -Name "mail-send"    

    .OUTPUTS
        Access Token.

    .LINK
        For more information about getting an access token with Microsoft Graph API, see: https://learn.microsoft.com/graph/auth-overview   
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $Prefix,
        [Parameter(Mandatory = $false)]
        [string]
        $Name = "mail-send"
    )

    $KeyVaultName = "$($Prefix)kv$($Name)".Replace('-', '')
    $AppName = "$Prefix-app-$Name"

    $appIdSecretName = $("$($AppName)-AppId").ToLowerInvariant()
    $appCertSecretName = $("$($AppName)-Cert").ToLowerInvariant()

    $Scope = "https://graph.microsoft.com/.default"
    $context = az account show | ConvertFrom-Json
    $tenantId = $context.tenantId

    Write-Host "Getting the AppId and Cert from KeyVault..."
    $secretClientId = az keyvault secret show --vault-name $KeyVaultName --name $appIdSecretName | ConvertFrom-Json
    $certValue = az keyvault secret show --vault-name $KeyVaultName --name $appCertSecretName | ConvertFrom-Json
    $certDetails = az keyvault certificate show --vault-name $KeyVaultName --name $appCertSecretName | ConvertFrom-Json

    Write-Host -MessageData:"Creating an in memory certificate..."
    $CertificateBase64String = [System.Convert]::FromBase64String("$($certValue.value)")
    $CertificateCollection = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2Collection
    $CertificateCollection.Import($CertificateBase64String, $null, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
    $rsaPrivateKey = $CertificateCollection[0].PrivateKey.ExportRSAPrivateKey()
    $hashValue = $certDetails.x509Thumbprint

    Write-Host "Creating a JWT..."
    $exp = ([DateTimeOffset](Get-Date).AddHours(1).ToUniversalTime()).ToUnixTimeSeconds()
    $nbf = ([DateTimeOffset](Get-Date).ToUniversalTime()).ToUnixTimeSeconds()

    $JWTHeader = @{
        alg = "RS256"
        typ = "JWT"
        x5t = $hashValue 
    }

    $JWTPayLoad = @{
        # What endpoint is allowed to use this JWT
        aud = "https://login.microsoftonline.com/$tenantId/oauth2/token"

        # Expiration timestamp
        exp = $exp 

        # Issuer = your application
        iss = $($secretClientId.value)

        # JWT ID: random guid
        jti = [guid]::NewGuid()

        # Not to be used before
        nbf = $nbf 

        # JWT Subject
        sub = $($secretClientId.value)
    }

    # Convert header and payload to base64
    $JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
    $EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)

    $JWTPayLoadToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))
    $EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)

    # Join header and Payload with "." to create a valid (unsigned) JWT
    $JWT = $EncodedHeader + "." + $EncodedPayload

    $PrivateKey = New-Object System.Security.Cryptography.RSACryptoServiceProvider
    $PrivateKey.ImportRSAPrivateKey($rsaPrivateKey, [ref]$true)
    # Get the private key object of your certificate

    # Define RSA signature and hashing algorithm
    $RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
    $HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256

    # Create a signature of the JWT
    $Signature = [Convert]::ToBase64String(
        $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT), $HashAlgorithm, $RSAPadding)
    ) -replace '\+', '-' -replace '/', '_' -replace '='

    # Join the signature to the JWT with "."
    $JWT = $JWT + "." + $Signature

    # Create a hash with body parameters
    $Body = @{
        client_id             = $AppId
        client_assertion      = $JWT
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        scope                 = $Scope
        grant_type            = "client_credentials"

    }

    $Url = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Use the self-generated JWT as Authorization
    $Header = @{
        Authorization = "Bearer $JWT"
    }

    # Splat the parameters for Invoke-Restmethod for cleaner code
    $PostSplat = @{
        ContentType = 'application/x-www-form-urlencoded'
        Method      = 'POST'
        Body        = $Body
        Uri         = $Url
        Headers     = $Header
    }
    
    Write-Host "Calling the Azure AD endpoint to get the access token..."
    $Request = Invoke-RestMethod @PostSplat
    Write-Output $request
}
