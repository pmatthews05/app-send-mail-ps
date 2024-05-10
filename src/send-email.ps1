param(
    [string]$Prefix = "cfcode",
    [string]$Name = "mail-send",
    [string]$Location = "uksouth"
)

$ResourceGroupName = "$Prefix-rg-$Name"
$KeyVaultName = "$($Prefix)kv$($Name)".Replace('-', '')
$AppName = "$Prefix-app-$Name"
