# Load config file
$config = Get-Content -Path config.json | ConvertFrom-Json

# Setup required variables
$baseUrl = "https://" + $config.region + "-api.mimecast.com"
$uri = "/api/gateway/get-hold-message-list"
$url = $baseUrl + $uri
$accessKey = $config.access_key
$secretKey = $config.secret_key
$appId = $config.app_id
$appKey = $config.app_key


# Generate request header values
$hdrDate = (Get-Date).ToUniversalTime().ToString("ddd, dd MMM yyyy HH:mm:ss UTC")
$requestId = [guid]::NewGuid().guid

# Create the HMAC SHA1 of the Base64 decoded secret key for the Authorization header
$sha = New-Object System.Security.Cryptography.HMACSHA1
$sha.key = [Convert]::FromBase64String($secretKey)
$sig = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($hdrDate + ":" + $requestId + ":" + $uri + ":" + $appKey))
$sig = [Convert]::ToBase64String($sig)

# Create Headers
$headers = @{"Authorization" = "MC " + $accessKey + ":" + $sig;
                "x-mc-date" = $hdrDate;
                "x-mc-app-id" = $appId;
                "x-mc-req-id" = $requestId;
                "Content-Type" = "application/json"}

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
                        ""start"": ""2022-12-05T14:49:18+0000"",
                        ""end"": ""2022-12-10T14:49:18+0000""
                    }
                ]
            }"

# Send Request
$response = Invoke-RestMethod -Method Post -Headers $headers -Body $postBody -Uri $url

# Initialise array to hold message id
$message_to_release = @()

# Loop over response to get message ids to release
foreach ($item in $response.data) {
    # Replace $item.reasonCode with the reason that needs to be bulk released
    if ($item.reasonCode -contains "default_inbound_attachment_protect_definition") {
        $message_to_release += $item.id
    }
}

# Set next request variables
$uri = "/api/gateway/hold-release"
$url = $baseUrl + $uri

# Loop over messages_to_release and release each message
foreach ($message in $message_to_release) {
    # Create the HMAC SHA1 of the Base64 decoded secret key for the Authorization header
    $sha = New-Object System.Security.Cryptography.HMACSHA1
    $sha.key = [Convert]::FromBase64String($secretKey)
    $sig = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($hdrDate + ":" + $requestId + ":" + $uri + ":" + $appKey))
    $sig = [Convert]::ToBase64String($sig)

    # Create Headers
    $headers = @{"Authorization" = "MC " + $accessKey + ":" + $sig;
    "x-mc-date" = $hdrDate;
    "x-mc-app-id" = $appId;
    "x-mc-req-id" = $requestId;
    "Content-Type" = "application/json"}

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
