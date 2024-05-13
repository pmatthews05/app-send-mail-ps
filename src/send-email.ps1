param(
    [string]$Prefix = "cfcode",
    [string]$Name = "mail-send",
    [string]$Location = "uksouth",
    [string[]]$EmailTo = @("pmatthews.admin@beisdevpme5.onmicrosoft.com"),
    [string]$EmailSubject = "Sending email from PowerShell",
    [string]$EmailBody = "<b>Hello, world!</b>",
    [string]$EmailBodyContentType = "HTML",
    [string]$EmailFrom = "no-reply.mailsend@beisdevpme5.onmicrosoft.com"
)

. ./log-in-app.ps1

$accessToken = Get-GraphAccessToken `
    -Prefix $Prefix `
    -Name $Name

. ./get-msgraphmessageparameters.ps1

$graphParams = $(Get-MSGraphMessageParameters `
        -EmailTo $EmailTo `
        -EmailSubject $EmailSubject `
        -EmailBody $EmailBody `
        -EmailBodyContentType $EmailBodyContentType) | ConvertTo-Json -Depth 10 -Compress

$url = "https://graph.microsoft.com/v1.0/users/$EmailFrom/messages"

$PostSplat = @{
    ContentType = 'application/json'
    Method      = 'POST'
    Body        = $graphParams
    Uri         = $url
    Headers     = $Header
}

$messageRequest = Invoke-RestMethod @PostSplat

Write-output $messageRequest

$url = "https://graph.microsoft.com/v1.0/users/$EmailFrom/messages/$($messageRequest.id)/send"

$PostSplat = @{
    ContentType = 'application/json'
    Method      = 'POST'
    Body        = $null
    Uri         = $url
    Headers     = $Header
}

Invoke-RestMethod @PostSplat