# Need to install PowerShell Module
# - ExchangeOnlineManagement 'Install-Module ExchangeOnlineManagement -AllowClobber -Scope Currentuser -Force'
# Connect-ExchangeOnline -ShowBanner:$false

param(
    [Parameter(Mandatory = $true)]
    [string]
    $Prefix,
    [Parameter(Mandatory = $true)]
    [string]
    $TenantDomain,
    [Parameter(Mandatory = $false)]
    [string]
    $Name = "mail-send"
)

$KeyVaultName = "$($Prefix)kv$($Name)".Replace('-', '')
$AppName = "$Prefix-app-$Name"

$exchangeSetupJsonFile = Get-Content -path "$PSScriptRoot/../config/email-setup.json" -Raw | ConvertFrom-Json

$exchangeSetupJsonFile.sharedMailBoxes | ForEach-Object {
    $sharedMailbox = $PSItem
    
    $sharedMailboxEmail = "$($sharedMailbox.name)@$TenantDomain"

    $foundMailbox = Get-MailBox -Identity $sharedMailboxEmail -ErrorAction SilentlyContinue
    if ($null -eq $foundMailBox) {
        Write-Host "Creating Mailbox for $sharedMailboxEmail..."

        New-Mailbox `
            -Name $($sharedMailbox.name) `
            -DisplayName $($sharedMailbox.displayName) `
            -PrimarySmtpAddress "$sharedMailboxEmail" `
            -Shared `
            -Confirm: $false
    }
    else {
        Write-Host "Mailbox for $sharedMailboxEmail already exists..."
    }

    if ($sharedMailbox.automaticReply) {
        Write-Host "Setting Automatic Reply for $sharedMailboxEmail..."
        Set-MailboxAutoReplyConfiguration `
            -Identity $($sharedMailbox.name) `
            -AutoReplyState Enabled `
            -ExternalAudience All `
            -InternalMessage $($sharedMailbox.internalMessage) `
            -ExternalMessage $($sharedMailbox.externalMessage) `
            -Confirm: $false
    }
    else {
        Write-Host "Disabling Automatic Reply for $sharedMailboxEmail..."
        Set-MailboxAutoReplyConfiguration `
            -Identity $($sharedMailbox.name) `
            -AutoReplyState Disabled `
            -Confirm: $false
    }

    Write-Host "Setting mailbox properties for Shared Mailbox $sharedMailboxEmail..."
    Set-Mailbox `
        -Identity $($sharedMailbox.name) `
        -MailTip $($sharedMailbox.automaticReplyMessage) `
        -HiddenFromAddressListsEnabled $true `
        -MaxSendSize 150MB `
        -Confirm: $false
}

$exchangeSetupJsonFile.mailEnabledSecurityGroups | ForEach-Object {
    $mailEnabledSecurityGroup = $PSItem
    $mailEnabledSecurityGroupEmail = "$($mailEnabledSecurityGroup.name)@$TenantDomain"

    $foundMailSecurityGroup = Get-DistributionGroup -Identity $mailEnabledSecurityGroupEmail -ErrorAction SilentlyContinue
    if ($null -eq $foundMailSecurityGroup) {
        Write-Host "Creating Mail Enabled Security Group for $mailEnabledSecurityGroupEmail..."
        New-DistributionGroup `
            -Name $($mailEnabledSecurityGroup.name) `
            -DisplayName $($mailEnabledSecurityGroup.displayName) `
            -Description $($mailEnabledSecurityGroup.description) `
            -RequireSenderAuthenticationEnabled $($mailEnabledSecurityGroup.allowExternalToEmail) `
            -ManagedBy $($mailEnabledSecurityGroup.Owners) `
            -Members $($mailEnabledSecurityGroup.Members) `
            -PrimarySmtpAddress $mailEnabledSecurityGroupEmail `
            -Type "Security" `
            -Confirm: $false
    }

    Write-Host "Setting mail enabled security group properties for $mailEnabledSecurityGroupEmail..."
    Set-DistributionGroup `
        -Identity: $($mailEnabledSecurityGroup.name) `
        -DisplayName $($mailEnabledSecurityGroup.displayName) `
        -Description: $($mailEnabledSecurityGroup.description) `
        -HiddenFromAddressListsEnabled $($mailEnabledSecurityGroup.hideFromAddressLists) `
        -RequireSenderAuthenticationEnabled $($mailEnabledSecurityGroup.allowExternalToEmail) `
        -PrimarySmtpAddress $mailEnabledSecurityGroupEmail `
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
    $foundPolicy = $allApplicationAccessPolicies | Where-Object { $_.Identity -like $accessPolicyIdentifier }

    $groupIDDomain = "$($appPolicy.policyScopeGroupId)@$TenantDomain"
    if ($null -eq $foundPolicy) {
        Write-Host "Creating application access policy for $($appId) and policyScopeGroupId $groupIDDomain..."
        $description = $appPolicy.description -ne "" ? $appPolicy.description.Replace("##AppName##", $appDetails.DisplayName).Replace("##policyScopeGroupId##", $groupIDDomain) : $appPolicy.description
        New-ApplicationAccessPolicy `
            -AccessRight $($appPolicy.accessRight) `
            -AppId $($appDetails.AppId) `
            -PolicyScopeGroupId $groupIDDomain `
            -Description $description `
    
    }
    else {
        Write-Host "Application access policy for $($appDetails.DisplayName) and policyScopeGroupId $groupIDDomain already exists..."        
    }
}