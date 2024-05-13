<#
    .SYNOPSIS
        Sends an email using Microsoft Graph API.

    .DESCRIPTION
        The script sends an email using Microsoft Graph API. 
        It creates a draft email, adds attachments to the draft email, and then sends the email.

    .PARAMETER Prefix
        The prefix to use for the resources.

    .PARAMETER EmailTo
        An array of email addresses to which the message will be sent.
    
    .PARAMETER EmailSubject
        The subject of the message.

    .PARAMETER EmailBody
        The body of the message.

    .PARAMETER EmailBodyContentType
        The content type of the message body. This should be "Text" or "HTML".

    .PARAMETER EmailFrom
        The email address from which the message will be sent.

    .PARAMETER Attachments
        (Optional) An array of file paths to the attachments. 

    .PARAMETER Name
        The name to use for the resources. Default is "mail-send".
    
    .EXAMPLE
        ./src/send-email.ps1 `
            -Prefix "contoso" `
            -EmailTo @("no-reply.mailsend-sg@contoso.com") `
            -EmailSubject "Welcome to the team" `
            -EmailBody "<b>Welcome to the team</b>, Please review our policies found here https://contoso.com/policies" `
            -EmailBodyContentType "HTML" `
            -EmailFrom "no-reply.mailsend@contso.com" `
            -Attachments @( "..\files\under3mb.pdf", "..\files\over3mb.pdf")

    .LINK
        For more information about sending messages with Microsoft Graph API, see: 
        https://learn.microsoft.com/en-us/graph/api/user-post-messages
        https://learn.microsoft.com/en-us/graph/api/user-sendmail
        https://learn.microsoft.com/en-us/graph/api/resources/fileattachment
        https://learn.microsoft.com/en-us/graph/api/attachment-createuploadsession

#>
[CmdletBinding()]
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
    [string[]]
    $Attachments,
    [Parameter(Mandatory = $false)]
    [string]
    $Name = "mail-send"
)

# Note: I know this isn't the best way to do this. Used for demo purposes only.
Write-Host "Loading required scripts..."
. $PSScriptRoot/log-in-app.ps1
. $PSScriptRoot/get-msgraphmessageparameters.ps1
. $PSScriptRoot/add-largeattachmentstodraftmessage.ps1

$accessToken = Get-GraphAccessToken `
    -Prefix $Prefix `
    -Name $Name

$Header = @{
    Authorization = "Bearer $($accessToken.access_token)"
}

$attachmentsGreaterThan3MB = @()
$attachmentsLessThan3MB = @()

Write-Host "Check if Attachment(s) are greater than 3MB..."
$Attachments | ForEach-Object {
    if (-not $PSItem) { return }
    $fileSize = [math]::Round((Get-Item $_).length / 1MB, 2)

    if ($fileSize -gt 150) {
        Write-Error "Attachment $($_) is greater than 35840MB, please reduce the size of the attachment and try again."
    }
    elseif ($fileSize -gt 3) {
        $attachmentsGreaterThan3MB += $_
    }   
    else {
        $attachmentsLessThan3MB += $_
    }
}

# Note: You don't need to create a draft email before sending an email, you can just use https://graph.microsoft.com/v1.0/users/$EmailFrom/sendMail 
#       with the same parameters as the draft email.
#       Typically, you would create a draft email if you want to attach files to the email before sending it, if the files are larger than 3MBs.
#       Under 3MB, they can be added as attachments in the body parameters

Write-Host "Getting Graph Message Parameters..."
$graphParams = $(Get-MSGraphMessageParameters `
        -EmailTo $EmailTo `
        -EmailSubject $EmailSubject `
        -EmailBody $EmailBody `
        -EmailBodyContentType $EmailBodyContentType `
        -Attachments $attachmentsLessThan3MB) | ConvertTo-Json -Depth 10 -Compress

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

Add-LargeAttachmentsToDraftMessage `
    -MessageId $messageRequest.id `
    -Attachments $attachmentsGreaterThan3MB `
    -User $EmailFrom `
    -AccessToken $accessToken.access_token

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