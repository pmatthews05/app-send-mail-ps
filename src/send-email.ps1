param(
    [Parameter(Mandatory = $true)]
    [string]
    $Prefix,
    [Parameter(Mandatory = $true)]
    [string[]]
    $EmailTo,
    [Parameter(Mandatory = $true)]
    [string]
    $EmailSubject, 
    [Parameter(Mandatory = $true)]
    [string]
    $EmailBody,
    [Parameter(Mandatory = $true)]
    [ValidateSet("HTML", "Text")]
    [string]
    $EmailBodyContentType,
    [Parameter(Mandatory = $true)]
    [string]
    $EmailFrom,
    [Parameter(Mandatory = $false)]
    [string]
    $Name = "mail-send"
)

Write-Host "Loading required scripts..."
. $PSScriptRoot/log-in-app.ps1
. $PSScriptRoot/get-msgraphmessageparameters.ps1

$accessToken = Get-GraphAccessToken `
    -Prefix $Prefix `
    -Name $Name

Write-Host "Getting Graph Message Parameters..."
$graphParams = $(Get-MSGraphMessageParameters `
        -EmailTo $EmailTo `
        -EmailSubject $EmailSubject `
        -EmailBody $EmailBody `
        -EmailBodyContentType $EmailBodyContentType) | ConvertTo-Json -Depth 10 -Compress

$Header = @{
    Authorization = "Bearer $($accessToken.access_token)"
}

Write-Host "Creating Draft Email..."
$url = "https://graph.microsoft.com/v1.0/users/$EmailFrom/messages"

$PostSplat = @{
    ContentType = 'application/json'
    Method      = 'POST'
    Body        = $graphParams
    Uri         = $url
    Headers     = $Header
}

$messageRequest = Invoke-RestMethod @PostSplat

Write-Host "Sending Email..."

$url = "https://graph.microsoft.com/v1.0/users/$EmailFrom/messages/$($messageRequest.id)/send"

$PostSplat = @{
    ContentType = 'application/json'
    Method      = 'POST'
    Body        = $null
    Uri         = $url
    Headers     = $Header
}

Invoke-RestMethod @PostSplat