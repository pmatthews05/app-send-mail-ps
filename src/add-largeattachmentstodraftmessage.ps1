function Add-LargeAttachmentsToDraftMessage {
    <#
    .SYNOPSIS
        This function adds large attachments to a draft message in Microsoft Graph.

    .DESCRIPTION
        The function takes a message ID, an array of file paths to the attachments, and a user ID. 
        It splits each file into 4MB chunks and uploads each chunk to the draft message.

    .PARAMETER MessageId
        The ID of the draft message to which the attachments will be added.

    .PARAMETER Attachments
        An array of file paths to the attachments.

    .PARAMETER UserId
        The ID of the user who owns the draft message.
    
    .PARAMETER AccessToken
        The access token to authenticate the request.

    .EXAMPLE
        Add-LargeAttachmentsToDraftMessage -MessageId "messageId" -Attachments @("path/to/file1.txt", "path/to/file2.txt") -UserId "userId" -AccessToken $($accessToken.access_token)

    .FUNCTIONALITY
        Microsoft Graph API
        Email

    .OUTPUTS
        None. The function uploads the attachments to the draft message and does not return anything.

    .LINK
        For more information about adding attachments to messages with Microsoft Graph API, see: https://learn.microsoft.com/graph/api/message-post-attachments
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $MessageId,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]
        $Attachments,
        [Parameter(Mandatory = $true)]
        [string]
        $User,
        [Parameter(Mandatory = $true)]
        [string]
        $AccessToken
    )

    $Header = @{
        Authorization = "Bearer $AccessToken"
    }
    

    # GetFileSize in bytes
    $Attachments | ForEach-Object {
        # Get File Size in Bytes
        $attachment = $_
        $fileName = ($attachment -split [regex]::Escape([IO.Path]::DirectorySeparatorChar))[-1]
        $fileInBytes = [System.IO.File]::ReadAllBytes($attachment)
        $fileSize = $fileInBytes.Length
        $params = @{
            AttachmentItem = @{
                attachmentType = "file"
                name           = $fileName
                size           = $fileSize
            }
        }

        $url = "https://graph.microsoft.com/v1.0/users/$user/messages/$MessageId/attachments/createUploadSession"
       
        $PostSplat = @{
            ContentType = 'application/json'
            Method      = 'POST'
            Body        = $params | ConvertTo-Json -Depth 10
            Uri         = $url
            Headers     = $Header
        }

        $uploadSession = Invoke-RestMethod @PostSplat

        $headers = @{ 'Content-Type' = 'application/json' }
     
        ##Split the file up in 4MB chunks
        $partSizeBytes = 4 * 1024 * 1024 #327680
        $index = 0
        $start = 0
        $end = 0

        $maxloops = [Math]::Round([Math]::Ceiling($fileSize / $partSizeBytes))

        $uploadResult = $null
        while ($fileSize -gt ($end + 1)) {
            $counttries = 0;
            do {               
                $success = $false
                try {

                    $start = $index * $partSizeBytes
                    if (($start + $partSizeBytes - 1 ) -lt $fileSize) {
                        $end = ($start + $partSizeBytes - 1)
                    }
                    else {
                        $end = ($start + ($fileSize - ($index * $partSizeBytes)) - 1)
                    }
                    [byte[]]$body = $fileInBytes[$start..$end]
                    $headers = @{    
                        'Content-Range' = "bytes $start-$end/$fileSize"
                    }
                    Write-Verbose "bytes $start-$end/$fileSize | Index: $index and ChunkSize: $partSizeBytes"
                    $uploadResult = Invoke-WebRequest -Method Put -Uri $uploadSession.uploadUrl -Body $body -Headers $headers -SkipHeaderValidation 
                    $success = $true
                }
                catch {
                    if ($counttries -gt 4) {
                        Write-Error "attempted 5 times."
                        throw $_
                    }
                    Write-Warning "Issue uploading $attachment..."
                    Write-Warning "Error Message: $($_.Exception.Message)"
                    Write-Warning "Inner Exception: $($_.Exception)"
                             
                    if ($_.Exception.Response.StatusCode -eq 429) {
                        try {
                            $sleeptime = [double]($_.Exception.Response.Headers.GetValues('Retry-After')[0])
                        }
                        catch {
                            Write-Verbose "Request throttled by Retry-After not provided"
                            $sleeptime = 60
                        }
                        Write-Host "Too many requests. Waiting $sleeptime seconds..."
                        Start-Sleep -Seconds $sleeptime
                    }
                    else {
                        Write-Host "Attemping again in 5 seconds..."
                        Start-Sleep -Seconds 5
                    }
                    $counttries++
                }
            }
            while ($success -ne $true)
       
            $index++
            $percentageComplete = $([Math]::Ceiling($index / $maxloops * 100))
            Write-Host "Uploading $fileName - $percentageComplete % complete..."     
        }
    }
}