# Load config file
$config = Import-PowerShellDataFile -Path .\config.psd1

# Setup required variables
$baseUrl = $config.baseURL
$accessKey = $config.accessKey
$secretKey = $config.secretKey
$appId = $config.appId
$appKey = $config.appKey

$getAttachmentLogs = "/api/ttp/attachment/get-logs"
$findMessageId = "/api/message-finder/search"
$releaseMessage = "/api/gateway/hold-release"

# Generate request header values
$hdrDate = (Get-Date).ToUniversalTime().ToString("ddd, dd MMM yyyy HH:mm:ss UTC")
$requestId = [guid]::NewGuid().guid
$endTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ss+0000")
$startTime = (Get-Date).AddHours(-1).ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ss+0000")

# Create the HMAC SHA1 of the Base64 decoded secret key for the Authorization header
$sha = New-Object System.Security.Cryptography.HMACSHA1
$sha.key = [Convert]::FromBase64String($secretKey)
$sig = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($hdrDate + ":" + $requestId + ":" + $getAttachmentLogs + ":" + $appKey))
$sig = [Convert]::ToBase64String($sig)

# Create Headers
$headers = @{
    "Authorization" = "MC " + $accessKey + ":" + $sig;
    "x-mc-date"              = $hdrDate;
    "x-mc-app-id"            = $appId;
    "x-mc-req-id"            = $requestId;
    "Content-Type"           = "application/json"
}

# Create post body
$postBody = "{
                ""meta"": {
                    ""pagination"": {
                        ""pageSize"": 500
                    }
                },
                ""data"": [
                    {
                        ""from"": ""$startTime"",
                        ""route"": ""all"",
                        ""to"": ""$endTime"",
                        ""scanResult"": ""all""
                    }
                ]
            }"
# Send Request
$response = Invoke-RestMethod -Method Post -Headers $headers -Body $postBody -Uri ($baseUrl + $getAttachmentLogs)

$messageToRelease = @()

foreach ($messageId in $response.data.attachmentLogs.messageId) {
    # Create the HMAC SHA1 of the Base64 decoded secret key for the Authorization header
    $sha = New-Object System.Security.Cryptography.HMACSHA1
    $sha.key = [Convert]::FromBase64String($secretKey)
    $sig = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($hdrDate + ":" + $requestId + ":" + $findMessageId + ":" + $appKey))
    $sig = [Convert]::ToBase64String($sig)

    # Create Headers
    $headers = @{
        "Authorization" = "MC " + $accessKey + ":" + $sig;
        "x-mc-date"     = $hdrDate;
        "x-mc-app-id"   = $appId;
        "x-mc-req-id"   = $requestId;
        "Content-Type"  = "application/json"
    }
    Write-Output $messageId
    $postBody = "{
        ""meta"": {
            ""pagination"": {
                ""pageSize"": 500
            }
        },
        ""data"": [
            {
                ""messageId"": ""$messageId""
            }
        ]
    }"

    $response = Invoke-RestMethod -Method Post -Headers $headers -Body $postBody -Uri ($baseUrl + $findMessageId)
    $messageToRelease += $response.data.trackedEmails.id
}


# Loop over messages_to_release and release each message
foreach ($message in $messageToRelease) {
    # Create the HMAC SHA1 of the Base64 decoded secret key for the Authorization header
    $sha = New-Object System.Security.Cryptography.HMACSHA1
    $sha.key = [Convert]::FromBase64String($secretKey)
    $sig = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($hdrDate + ":" + $requestId + ":" + $releaseMessage + ":" + $appKey))
    $sig = [Convert]::ToBase64String($sig)

    # Create Headers
    $headers = @{
        "Authorization" = "MC " + $accessKey + ":" + $sig;
        "x-mc-date"     = $hdrDate;
        "x-mc-app-id"   = $appId;
        "x-mc-req-id"   = $requestId;
        "Content-Type"  = "application/json"
    }

    #Create post body
    $postBody = "{
        ""data"": [
            {
                ""id"": ""$message""
            }
        ]
    }"

    # Send Request
    $response = Invoke-RestMethod -Method Post -Headers $headers -Body $postBody -Uri $($baseUrl + $releaseMessage)
}
