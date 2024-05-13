<#
    .SYNOPSIS
        Creates the App Registration, Resource Group, Key Vault, and Service Principal for the specified prefix and name with the correct permissions.

    .DESCRIPTION
        Create Resource Group and Key Vault
        Create an App Registration in Azure AD with Self Signed Certificate
        Add the following API Permissions:
        - Microsoft Graph
            - Mail.Send
            - Mail.ReadWrite
        Grant Role Assignment to the App Registration
        Adds the current user as a Key Vault Administrator

    .PARAMETER Prefix
        The prefix to use for the resources.

    .PARAMETER Name
        The name to use for the resources. Default is "mail-send".

    .PARAMETER Location
        The location to create the resources. Default is "uksouth".
    
    
    .EXAMPLE
        .\create-app-reg.ps1 -Prefix "contoso" -Name "mail-send" -Location "uksouth"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]
    $Prefix,
    [Parameter(Mandatory = $false)]
    [string]
    $Name = "mail-send",
    [Parameter(Mandatory = $false)]
    [string]
    $Location = "uksouth"
)

$ResourceGroupName = "$Prefix-rg-$Name"
$KeyVaultName = "$($Prefix)kv$($Name)".Replace('-', '')
$AppName = "$Prefix-app-$Name"

# Create Resource Group
Write-host "Create or Update Resource Group..."
az group create --name $ResourceGroupName --location $Location

# Create Key Vault
$existingKeyVaultId = az resource list --resource-type "Microsoft.KeyVault/vaults" --name "$KeyVaultName" --resource-group $ResourceGroupName --query "[].id" --output tsv
if (-not $existingKeyVaultId) { 
    Write-Host "Creating KeyVault"
    $keyVault = az keyvault create --name $KeyVaultName `
        --resource-group $ResourceGroupName `
        --enable-rbac-authorization $true `
        --enabled-for-deployment $true `
        --enabled-for-disk-encryption $true `
        --enabled-for-template-deployment $true | ConvertFrom-Json

    $existingKeyVaultId = $keyVault.id
}

# Add Current user as RBAC KeyVault Administrator
Write-Host "Adding Current User as KeyVault Administrator..."
$azContext = az account show | ConvertFrom-Json
$currentUserObjectId = az ad signed-in-user show --query id -o tsv
az role assignment create --role "Key Vault Administrator" --assignee-object-id $currentUserObjectId --assignee-principal-type $($azContext.user.type) --scope $existingKeyVaultId

# Create App Registration
Write-Host "Creating App Registration..."
$appRegistration = az ad app create --sign-in-audience AzureADMultipleOrgs --display-name $AppName | ConvertFrom-Json

# Create Service Principal
$servicePrincipalList = az ad sp list --spn $($appRegistration.appId) | ConvertFrom-Json
if ($servicePrincipalList.Length -eq 0) {
    Write-Host "Creating Service Principal..."
    $servicePrincipal = az ad sp create --id $($appRegistration.appId) | ConvertFrom-Json
}
else {
    $servicePrincipal = $servicePrincipalList[0]
}

$checks = 0;
do {
    try {
        $url = "https://graph.microsoft.com/v1.0/applications/$($appRegistration.id)" 
        $result = az rest --method GET --url $url | ConvertFrom-Json
    }
    catch {
        Write-Host "App not ready in MS Graph... Waiting 5 seconds and trying again..."
        $checks++
        Start-Sleep -seconds:5 
    }
} while (-not $result -or ($checks -gt 4))

# StoreApp ID in KeyVault
$appIdSecretName = $("$($AppName)-AppId").ToLowerInvariant()

# Create if not exist.
$secretCollection = az keyvault secret list --vault-name $KeyVaultName --query "[?name == '$appIdSecretName']" | ConvertFrom-Json
if ($secretCollection.Length -gt 0) {
    
    $secret = az keyvault secret show --vault-name $KeyVaultName --name $appIdSecretName | ConvertFrom-Json
}
if ($secret.value -ne $appRegistration.appId) {
    Write-Host "Adding AppID to KeyVault..."
    az keyvault secret set --vault-name $KeyVaultName --name $appIdSecretName --value $($appRegistration.appId) | Out-Null
}

# Create Certificate in KeyVault

#Get App Registration Certificate Details
$appCertSecretName = $("$($AppName)-Cert").ToLowerInvariant()
[int]$monthsValidityRequired = 9
[DateTime]$endDate = $(Get-Date).AddMonths($monthsValidityRequired)

Write-Host "Getting App Certificates stored in AppReg...."
$keyCredentials = $(az rest --method GET --url "https://graph.microsoft.com/v1.0/applications/$($appRegistration.id)`?`$select=keyCredentials" | ConvertFrom-Json).keyCredentials

Write-Host "Getting Certificate from KeyVault..."
$certificatecollection = az keyvault certificate list --include-pending $true --vault-name "$KeyVaultName" --query "[?name == '$appCertSecretName']" | ConvertFrom-Json
if ($certificatecollection.Length -gt 0) {
    $keyVaultCertificate = az keyvault certificate show --vault-name $KeyVaultName --name $appCertSecretName | ConvertFrom-Json
}
$create = $false
if ((-not $create) -and (-not $keyVaultCertificate)) {
    #No certificate found in KeyVault
    $create = $true
}

if ((-not $create) -and (-not $keyCredentials)) {
    #No App Cert
    $create = $true
}

if ((-not $create) -and ($keyVaultCertificate.x509ThumbprintHex -notmatch $appCertificate.customKeyIdentifier)) { 
    #Certificate in KeyVault does not match App Registration
    $create = $true
}

if ((-not $create) -and ($endDate -gt $appCertificate.endDateTime)) {
    #Certificate in KeyVault is not valid
    $create = $true
}

if ($create) {
    Write-Host "Create Certificate in KeyVault..."
    $policy = az keyvault certificate get-default-policy
    if ($PSVersionTable.Platform -eq "Win32NT") {
        $policy = $policy -replace '"', '\"'
    }
        
    #13 months because #12 throws an error as Key Credential end date is invalid.
    az keyvault certificate create --vault-name "$KeyVaultName" --name "$appCertSecretName" --policy "$policy" --validity 13 | Out-Null

    $certificatecollection = az keyvault certificate list --include-pending $true --vault-name "$KeyVaultName" --query "[?name == '$appCertSecretName']" | ConvertFrom-Json
    if ($certificatecollection.Length -gt 0) {
        $keyVaultCertificate = az keyvault certificate show --vault-name $KeyVaultName --name $appCertSecretName | ConvertFrom-Json
    }

    Write-Host "Updating $($AppRegistration.appId) App Registration key credentials..."
    az ad app update --id "$($AppRegistration.appId)" --key-type "AsymmetricX509Cert" --key-usage "Verify" --key-value "$($keyVaultCertificate.cer)" | Out-Null
}

# Set App Registration API
#MSGraph
Write-Host "Get MS Graph API Permissions..."
$msgraphAPISP = az ad sp list --query "[?appDisplayName == 'Microsoft Graph'].{appId:appId, id:id}" --all | ConvertFrom-Json
$mailSendPerm = az ad sp show --id $($msgraphAPISP.appId) --query "appRoles[?value=='Mail.Send']" | ConvertFrom-Json
$mailReadWritePerm = az ad sp show --id $($msgraphAPISP.appId) --query "appRoles[?value=='Mail.ReadWrite']" | ConvertFrom-Json
$userReadBasicPerm = az ad sp show --id $($msgraphAPISP.appId) --query "appRoles[?value=='User.ReadBasic.All']" | ConvertFrom-Json

#Check if already exist.
$existingPermissions = az ad app permission list --id $($appRegistration.appId) --query "[?resourceAppId == '$($msgraphAPISP.appId)'].resourceAccess | [].{id:id} | [?id=='$($mailSendPerm.id)']" | ConvertFrom-Json
if ($existingPermissions.Length -eq 0) {
    Write-Host "Adding Mail Send Permission to App Registration..."
    az ad app permission add --id $($appRegistration.appId) --api $($msgraphAPISP.appId) --api-permissions "$($mailSendPerm.id)=Role"
}

$existingPermissions = az ad app permission list --id $($appRegistration.appId) --query "[?resourceAppId == '$($msgraphAPISP.appId)'].resourceAccess | [].{id:id} | [?id=='$($mailReadWritePerm.id)']" | ConvertFrom-Json
if ($existingPermissions.Length -eq 0) {
    Write-Host "Adding Mail Read Write Permission to App Registration..."
    az ad app permission add --id $($appRegistration.appId) --api $($msgraphAPISP.appId) --api-permissions "$($mailReadWritePerm.id)=Role"
}

$existingPermissions = az ad app permission list --id $($appRegistration.appId) --query "[?resourceAppId == '$($msgraphAPISP.appId)'].resourceAccess | [].{id:id} | [?id=='$($userReadBasicPerm.id)']" | ConvertFrom-Json
if ($existingPermissions.Length -eq 0) {
    Write-Host "Adding User Read Basic Permission to App Registration..."
    az ad app permission add --id $($appRegistration.appId) --api $($msgraphAPISP.appId) --api-permissions "$($userReadBasicPerm.id)=Role"
}

#Grant Permissions.
Write-Host "Granting Admin Consent..."
az ad app permission admin-consent --id $appRegistration.appId