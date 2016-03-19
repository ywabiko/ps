#
# PowerShell OAuth2 Sample (FitBit)
#
# This code demonstrates how to access FitBit API via OAuth2 to retrieve your personal heartrate data.
# cf. https://dev.fitbit.com/

$DebugPreference = "continue"; #Enable debug messages to be sent to console
Set-Variable -Option Constant TokenRegkeyPath -Value "HKCU:\Software\FitBitSample"
Set-Variable -Option Constant CallbackUri     -Value "YOUR_CALLBACK_URI"
Set-Variable -Option Constant ClientId        -Value "YOUR_CLIENT_ID";
Set-Variable -Option Constant AppSecret       -Value "YOUR_APP_SECRET";
Set-Variable -Option Constant AuthUrl         -Value "https://www.fitbit.com/oauth2/authorize";
Set-Variable -Option Constant TokenUrl        -Value "https://api.fitbit.com/oauth2/token";
Set-Variable -Option Constant Scope           -Value  @("heartrate");

function Get-MyToken {
    [CmdletBinding()]
    param (
	[Parameter(Mandatory = $false)]
	[string] $Path = $TokenRegkeyPath,
	[Parameter(Mandatory = $false)]
	[string] $TokenType = "AccessToken" # AccessToken | RefreshToken
    )

    try
    {
	(Get-ItemProperty -Path $Path -Name $TokenType -ErrorAction Stop).$TokenType
    }
    catch
    {
	return $false;
    }
}

function Store-MyToken {
    [CmdletBinding()]
    param (
	[Parameter(Mandatory = $false)]
	[string] $Path = $TokenRegkeyPath,
	[Parameter(Mandatory = $false)]
	[string] $TokenType = "AccessToken", # AccessToken | RefreshToken
	[Parameter(Mandatory = $true)]
	[string] $TokenValue
    )

    if (!(Test-Path $Path))
    {
	New-Item $Path
    }
    
    try
    {
	New-ItemProperty -Path $Path -Name $TokenType -Value $TokenValue -PropertyType String -Force -ErrorAction Stop | Out-Null
    }
    catch
    {
	return $false;
    }
}

function Initiate-OAuth2 {
    [CmdletBinding()]
    param (
    )
    $scopeString = $Scope -join "+";

    ###
    ### Get an auth code
    ###

    $requestUrl = "{0}?client_id={1}&response_type=code&scope={2}&redirect_uri={3}" `
      -f $AuthUrl, $ClientId, $scopeString, $CallbackUri;
    "Request URL is: {0}" -f $requestUrl | Write-Debug

    $IE = New-Object -ComObject InternetExplorer.Application;
    $IE.Navigate($requestUrl);
    $IE.Visible = $true;

    # TODO: timeout/bailout to deal with cancellation by user...
    while ($IE.LocationUrl -notmatch "\?code=") {
	Start-Sleep -s 1;
    }

    "Received redirection: URL: {0}" -f $IE.LocationUrl | Write-Debug
    "{0}" -f (((($IE.LocationUrl -split '`\`?')[-1]) -split '&') -join "`r`n") | Write-Debug

    [Void]($IE.LocationUrl -match "code=([^&]+)");
    $authCode = $Matches[1];
    $IE.Quit();
    "Received AuthCode: {0}" -f $authCode | Write-Debug

    ###
    ### Get an access token
    ###

    $appSecretStr = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($ClientId):$($AppSecret)"));
    $headers = @{"Authorization" = "Basic $appSecretStr"};
    $body = @{
	code = $authCode;
	grant_type = "authorization_code";
	client_id = $ClientId;
	redirect_uri = $CallbackUri;
    };
    $bodyArray = @();
    foreach ($b in $body.Keys)
    {
	$bodyArray += "{0}={1}" -f $b, [System.Uri]::EscapeDataString($body.$b);
    }
    $bodyStr = $bodyArray -join "&";
    "### bodyStr = $bodyStr" | Write-Debug;
    "### appSecretStr = $appSecretStr" | Write-Debug;

    $result = Invoke-RestMethod -Method Post -Uri $TokenUrl -ContentType "application/x-www-form-urlencoded" `
      -Headers $headers -Body $bodyStr -Verbose:$true -Debug:$true -ErrorAction Stop

    #$result | Format-Table  # too long to human
    @{expires_in = $result.expires_in; token_type = $result.token_type;} | Format-Table;

    $result.refresh_token | Out-File refresh_token.txt
    $result.access_token  | Out-File access_token.txt
    
    $accessToken = $result.access_token;
    $refreshToken = $result.refresh_token;

    ###
    ### Save Tokens
    ###

    Store-MyToken -TokenValue $accessToken  -TokenType AccessToken 
    Store-MyToken -TokenValue $refreshToken -TokenType RefreshToken

    "Access token is: {0}" -f $accessToken | Write-Debug
}

function Refresh-OAuth2 {
    [CmdletBinding()]
    param (
    )
    $scopeString = $Scope -join "+";

    $refreshToken = Get-MyToken -TokenType RefreshToken

    ###
    ### Refresh the access token
    ###

    $appSecretStr = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($ClientId):$($AppSecret)"));
    $headers = @{"Authorization" = "Basic $appSecretStr"};
    $body = @{
	grant_type = "refresh_token";
	refresh_token = $refreshToken;
    };
    $bodyArray = @();
    foreach ($b in $body.Keys)
    {
	$bodyArray += "{0}={1}" -f $b, [System.Uri]::EscapeDataString($body.$b);
    }
    $bodyStr = $bodyArray -join "&";
    "### bodyStr = $bodyStr" | Write-Debug;
    "### appSecretStr = $appSecretStr" | Write-Debug;

    try
    {
        $result = Invoke-RestMethod -Method Post -Uri $TokenUrl -ContentType "application/x-www-form-urlencoded" `
          -Headers $headers -Body $bodyStr -Verbose:$true -Debug:$true -ErrorAction Stop
    }
    catch
    {
	return $false;
    }

    @{expires_in = $result.expires_in; token_type = $result.token_type;} | Format-Table;

    $result.refresh_token | Out-File refresh_token.txt
    $result.access_token  | Out-File access_token.txt
    
    $accessToken = $result.access_token;
    $refreshToken = $result.refresh_token;

    ###
    ### Save Tokens
    ###

    Store-MyToken -TokenValue $accessToken  -TokenType AccessToken
    Store-MyToken -TokenValue $refreshToken -TokenType RefreshToken

    "Access token is: {0}" -f $accessToken | Write-Debug
    $true;
}

function Query-OAuth2 {
    [CmdletBinding()]
    param (
	[Parameter(Mandatory = $true)]
	[string] $Uri
    )

    $accessToken = Get-MyToken -TokenType AccessToken;
    $result = Invoke-RestMethod -Method Get -Uri $Uri -Headers @{"Authorization" = "Bearer $accessToken"}
    $result | ConvertTo-Json
}


###
### Main
###

###
### Try to refresh first. If this attempt failed, then initiate access token retrieval process.
###
$refreshed = Refresh-OAuth2 

if (!$refreshed)
{
    Initiate-OAuth2
    $refreshed = Refresh-OAuth2

    if (!$refreshed)
    {
        "Fatal error. Giving up" | Write-Host;
        return $false;
    }
}

###
### Query activities (samples)
###
$HRUri1min = "https://api.fitbit.com/1/user/-/activities/heart/date/{0}/1d.json" -f "today";
Query-OAuth2 -Uri $HRUri1min;
"HRUri1min: activeties-heart: dateTime={0} value={1}" -f $result.'activities-heart'.dateTime, $result.'activities-heart'.value | Write-Output

$HRUri = "https://api.fitbit.com/1/user/-/activities/heart/date/{0}/1d/{1}/time/18:00/18:39.json" -f "today", "1";
Query-OAuth2 -Uri $HRUri;
"HRUri(intra): activeties-heart: dateTime={0} value={1}" -f $result.'activities-heart'.dateTime, $result.'activities-heart'.value | Write-Output
$result.'activities-heart-intraday'.dataset | format-table;

