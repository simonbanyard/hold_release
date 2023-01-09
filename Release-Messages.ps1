# Load config file
$config = Import-PowerShellDataFile -Path .\config.psd1

# Setup required variables
$baseUrl = $config.baseURL
$uri = "/api/gateway/get-hold-message-list"
$url = $baseUrl + $uri
$accessKey = $config.accessKey
$secretKey = $config.secretKey
$appId = $config.appId
$appKey = $config.appKey


# Generate request header values
$hdrDate = (Get-Date).ToUniversalTime().ToString("ddd, dd MMM yyyy HH:mm:ss UTC")
$requestId = [guid]::NewGuid().guid
$endTime = (Get-Date).ToUniversalTime().ToString("yyyy-mm-ddThh:mm:ss+0000")
$startTime = (Get-Date).AddHours(-1).ToUniversalTime().ToString("yyyy-mm-ddThh:mm:ss+0000")

# Create the HMAC SHA1 of the Base64 decoded secret key for the Authorization header
$sha = New-Object System.Security.Cryptography.HMACSHA1
$sha.key = [Convert]::FromBase64String($secretKey)
$sig = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($hdrDate + ":" + $requestId + ":" + $uri + ":" + $appKey))
$sig = [Convert]::ToBase64String($sig)

# Create Headers
$headers = @{"Authorization" = "MC " + $accessKey + ":" + $sig;
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
                        ""admin"": true,
                        ""start"": ""$startTime"",
                        ""end"": ""$endTime""
                    }
                ]
            }"

Write-Output $postBody

# Send Request
$response = Invoke-RestMethod -Method Post -Headers $headers -Body $postBody -Uri $url

# Initialise array to hold message id
$messageToRelease = @()

# Loop over response to get message ids to release
foreach ($item in $response.data) {
    # Replace $item.reasonCode with the reason that needs to be bulk released
    if ($item.reasonCode -contains "default_inbound_attachment_protect_definition") {
        $messageToRelease += $item.id
    }
}

# Set next request variables
$uri = "/api/gateway/hold-release"
$url = $baseUrl + $uri

# Loop over messages_to_release and release each message
foreach ($message in $messageToRelease) {
    # Create the HMAC SHA1 of the Base64 decoded secret key for the Authorization header
    $sha = New-Object System.Security.Cryptography.HMACSHA1
    $sha.key = [Convert]::FromBase64String($secretKey)
    $sig = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($hdrDate + ":" + $requestId + ":" + $uri + ":" + $appKey))
    $sig = [Convert]::ToBase64String($sig)

    # Create Headers
    $headers = @{"Authorization" = "MC " + $accessKey + ":" + $sig;
        "x-mc-date"              = $hdrDate;
        "x-mc-app-id"            = $appId;
        "x-mc-req-id"            = $requestId;
        "Content-Type"           = "application/json"
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
    $response = Invoke-RestMethod -Method Post -Headers $headers -Body $postBody -Uri $url
}
