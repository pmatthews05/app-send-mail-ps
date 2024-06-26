function Get-MSGraphMessageParameters {
    <#
    .SYNOPSIS
        This function generates parameters for a Microsoft Graph API message.

    .DESCRIPTION
        The function takes email addresses, subject, body, body content type.
        It generates a hash table of parameters for a Microsoft Graph API message.

    .PARAMETER EmailTo
        An array of email addresses to which the message will be sent.
   
    .PARAMETER EmailSubject
        The subject of the message.

    .PARAMETER EmailBody
        The body of the message.

    .PARAMETER EmailBodyContentType
        The content type of the message body. This should be "Text" or "HTML".

    .PARAMETER Attachments
        An array of file paths to the attachments. These need to be less than 3MB in size.
        
    .EXAMPLE
        $messageParameters = Get-MSGraphMessageParameters -EmailTo @("email1@example.com") -EmailSubject "Hello" -EmailBody "Hello, world!" -EmailBodyContentType "Text"

    .FUNCTIONALITY
        Microsoft Graph API
        Email

    .OUTPUTS
        Hash table. The hash table contains the following keys: subject, toRecipients, ccRecipients (optional), attachments (optional), body.

    .LINK
        For more information about sending messages with Microsoft Graph API, see: https://learn.microsoft.com/graph/api/user-sendmail
    #>
    param (
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
        [string]
        $EmailBodyContentType,
        [Parameter(Mandatory = $false)]
        [string[]]
        $Attachments
    )

    [array]$msgToRecipients = $EmailTo | ForEach-Object {
        @{
            emailAddress = @{address = $_ }
        }
    } 
    $message = @{subject = $EmailSubject }
    $message += @{toRecipients = $msgToRecipients }
 
    $message += @{body = @{contentType = $EmailBodyContentType; content = $EmailBody } }

    if ($Attachments) {
        [array]$msgAttachments = $Attachments | ForEach-Object {
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name          = ($_ -split [regex]::Escape([IO.Path]::DirectorySeparatorChar))[-1]
                contentBytes  = [Convert]::ToBase64String([IO.File]::ReadAllBytes($_))
            }
        }
        $message += @{attachments = $msgAttachments }
    }

    return $message
}