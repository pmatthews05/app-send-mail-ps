# Need to install PowerShell Module
# - ExchangeOnlineManagement 'Install-Module ExchangeOnlineManagement -AllowClobber -Scope Currentuser -Force'
# Connect-ExchangeOnline -ShowBanner:$false

param(
    [string]$Prefix = "cfcode",
    [string]$Name = "mail-send"
)

$KeyVaultName = "$($Prefix)kv$($Name)".Replace('-', '')
$AppName = "$Prefix-app-$Name"

$exchangeSetupJsonFile = Get-Content -path "$PSScriptRoot/../config/email-setup.json" -Raw | ConvertFrom-Json

$exchangeSetupJsonFile.sharedMailBoxes | ForEach-Object {
    $sharedMailbox = $PSItem
    $nameSplit = $sharedMailbox.name.Split("@")[0]

    $foundMailbox = Get-MailBox -Identity $($sharedMailbox.name) -ErrorAction SilentlyContinue
    if($null -eq $foundMailBox){
        Write-Host "Creating Mailbox for $($sharedMailbox.name)..."

        New-Mailbox `
            -Name $nameSplit `
            -DisplayName $($sharedMailbox.displayName) `
            -PrimarySmtpAddress "$($sharedMailbox.name)" `
            -Shared `
            -Confirm: $false
    }
    else{
        Write-Host "Mailbox for $($sharedMailbox.name) already exists..."
    }

    if($sharedMailbox.automaticReply){
        Write-Host "Setting Automatic Reply for $($sharedMailbox.name)..."
        Set-MailboxAutoReplyConfiguration `
        -Identity $nameSplit `
        -AutoReplyState Enabled `
        -ExternalAudience All `
        -InternalMessage $($sharedMailbox.internalMessage) `
        -ExternalMessage $($sharedMailbox.externalMessage) `
        -Confirm: $false
    }
    else{
        Write-Host "Disabling Automatic Reply for $($sharedMailbox.name)..."
        Set-MailboxAutoReplyConfiguration `
        -Identity $nameSplit `
        -AutoReplyState Disabled `
        -Confirm: $false
    }

    Write-Host "Setting mailbox properties for Shared Mailbox $($sharedMailbox.name)..."
    Set-Mailbox `
        -Identity $nameSplit `
        -MailTip $($sharedMailbox.automaticReplyMessage) `
        -HiddenFromAddressListsEnabled $true `
        -MaxSendSize 150MB `
        -Confirm: $false
}

$exchangeSetupJsonFile.mailEnabledSecurityGroups | ForEach-Object {
    $mailEnabledSecurityGroup = $PSItem
    $nameSplit = $mailEnabledSecurityGroup.name.Split("@")[0]

    $foundMailSecurityGroup = Get-DistributionGroup -Identity $($mailEnabledSecurityGroup.name) -ErrorAction SilentlyContinue
    if($null -eq $foundMailSecurityGroup){
        Write-Host "Creating Mail Enabled Security Group for $($mailEnabledSecurityGroup.name)..."
        New-DistributionGroup `
        -Name $nameSplit `
        -DisplayName $($mailEnabledSecurityGroup.displayName) `
        -Description $($mailEnabledSecurityGroup.description) `
        -RequireSenderAuthenticationEnabled $($mailEnabledSecurityGroup.allowExternalToEmail) `
        -ManagedBy $($mailEnabledSecurityGroup.Owners) `
        -Members $($mailEnabledSecurityGroup.Members) `
        -PrimarySmtpAddress $($mailEnabledSecurityGroup.name) `
        -Type "Security" `
        -Confirm: $false
    }

    Write-Host "Setting mail enabled security group properties for $($mailEnabledSecurityGroup.name)..."
    Set-DistributionGroup `
        -Identity $nameSplit `
        -DisplayName $($mailEnabledSecurityGroup.displayName) `
        -Description: ${$mailEnabledSecurityGroup.description} `
        -HiddenFromAddressListsEnabled $($mailEnabledSecurityGroup.hideFromAddressLists) `
        -RequireSenderAuthenticationEnabled $($mailEnabledSecurityGroup.allowExternalToEmail) `
        -PrimarySmtpAddress $($mailEnabledSecurityGroup.name) `
        -MailTip $($mailEnabledSecurityGroup.mailTip) `
        -Confirm: $false
}

$appId = az keyvault secret show --vault-name $KeyVaultName --name "$($AppName)-AppId" --query "value" -o tsv
$appDetails = az ad app show --id $appId | ConvertFrom-Json
$account = az account show | ConvertFrom-Json

$allApplicationAccessPolicies = Get-ApplicationAccessPolicy
$exchangeSetupJsonFile.applicationAccessPolicy | ForEach-Object {
    $appPolicy = $PSItem
    
    Write-Host "Getting the Distribution Group for $($appPolicy.policyScopeGroupId)..."
    $distributionGroup = Get-DistributionGroup -Identity $appPolicy.policyScopeGroupId
    Write-Host "Getting a regex version of the Application Access Policy Identifier..."
    $accessPolicyIdentifier = "$($account.tenantId)\$($appId):*;$($distributionGroup.ExternalDirectoryObjectId)"
    $foundPolicy = $allApplicationAccessPolicies | Where-Object {$_.Identity -like $accessPolicyIdentifier}

    if ($null -eq $foundPolicy) {
        Write-Information "Creating application access policy for $($appId) and policyScopeGroupId $($appPolicy.policyScopeGroupId)..."
        $description = $appPolicy.description -ne "" ? $appPolicy.description.Replace("##AppName##", $appDetails.DisplayName).Replace("##policyScopeGroupId##", $appPolicy.policyScopeGroupId) : $appPolicy.description
        New-ApplicationAccessPolicy `
            -AccessRight $($appPolicy.accessRight) `
            -AppId $($appDetails.AppId) `
            -PolicyScopeGroupId $($appPolicy.policyScopeGroupId) `
            -Description $description `
    }
    else {
        Write-Information "Application access policy for $($appDetails.DisplayName) and policyScopeGroupId $($appPolicy.policyScopeGroupId) already exists..."        
    }
}