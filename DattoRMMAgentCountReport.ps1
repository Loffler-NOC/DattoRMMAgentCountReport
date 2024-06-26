# Specify security protocols
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function New-AemApiAccessToken {
    param
    (
        [string]$apiUrl,
        [string]$apiKey,
        [string]$apiSecretKey
    )

    # Convert password to secure string
    $securePassword = ConvertTo-SecureString -String 'public' -AsPlainText -Force

    # Define parameters for Invoke-WebRequest cmdlet
    $params = @{
        Credential  =	New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ('public-client', $securePassword)
        Uri         =	'{0}/auth/oauth/token' -f $apiUrl
        Method      =	'POST'
        ContentType = 'application/x-www-form-urlencoded'
        Body        = 'grant_type=password&username={0}&password={1}' -f $apiKey, $apiSecretKey
    }
	
    # Request access token
    try { (Invoke-WebRequest @params -UseBasicParsing | ConvertFrom-Json).access_token }
    catch { $_.Exception }
}

function New-AemApiRequest {
    param 
    (
        [string]$apiUrl,
        [string]$apiAccessToken,
        [string]$apiMethod,
        [string]$apiRequest,
        [string]$apiRequestBody
    )

    # Define parameters for Invoke-WebRequest cmdlet
    $params = @{
        Uri         =	'{0}/api{1}' -f $apiUrl, $apiRequest
        Method      =	$apiMethod
        ContentType	= 'application/json'
        Headers     =	@{
            'Authorization'	=	'Bearer {0}' -f $apiAccessToken
        }
    }

    # Add body to parameters if present
    If ($apiRequestBody) { $params.Add('Body', $apiRequestBody) }

    # Make request
    try { (Invoke-WebRequest @params -UseBasicParsing).Content }
    catch { $_.Exception }
}

# Define parameters
$params = @{
    apiUrl       =	$env:RMMapiUrl
    apiKey       =	$env:RMMapiKey
    apiSecretKey =	$env:RMMapiSecretKey
    apiMethod    =	'GET'
    apiRequest   =	'/v2/account/sites'
}

# Call New-AemApiAccessToken function using defined parameters 
$apiAccessToken = New-AemApiAccessToken @params

# Call New-AemApiRequest function using defined parameters
$sites = New-AemApiRequest @params -ApiAccessToken $apiAccessToken
Write-Output "Begin params"
Write-Output $params
Write-Output "End params"
Write-Output "Begin raw"
Write-Output $sites
Write-Output "End raw"

$sites = New-AemApiRequest @params -ApiAccessToken $apiAccessToken | ConvertFrom-Json
Write-Output "Begin raw"
Write-Output $sites
Write-Output "End raw"

$sitesData = @()


foreach ($site in $sites.sites) {
    $siteData = [PSCustomObject]@{
        Name        = $site.name
        DeviceCount = $site.devicesStatus.numberOfDevices
    }
    $sitesData += $siteData
}

#filter out sites we don't want
$sitesData = $sitesData | Where-Object { 
        $_.Name -ne "411" -and `
        $_.Name -ne "Managed" -and `
        $_.Name -ne "OnDemand" -and `
        $_.Name -ne "Deleted Devices" -and `
        $_.Name -ne "LITSLAB" -and `
        $_.Name -ne "z_Agent Removal Full"
}

# Identify sites with names ending in "- Police"
$policeSites = $sitesData | Where-Object { $_.Name -like "* - Police" }

#Combine police site count into normal site count
foreach ($policeSite in $policeSites) {
    $policeSiteNameWithoutSuffix = $policeSite.Name -replace ' - Police$', ''
    $matchingSite = $sitesData | Where-Object { $_.Name -eq $policeSiteNameWithoutSuffix }
    $matchingSite.DeviceCount = $matchingSite.DeviceCount + $policeSite.DeviceCount
}

#Remove police sites
$sitesData = $sitesData | Where-Object { $_.Name -notlike "* - Police" }

#sort the results
$sitesData = $sitesData | Sort-Object Name

# Export data to CSV
$csvFilePath = ".\sites_data.csv"

#export to path
$sitesData | Export-Csv -Path $csvFilePath -NoTypeInformation

#Email the report

# Install Mailozaurr
Get-PackageProvider -Name NuGet -Force
Install-Module -Name Mailozaurr -AllowClobber -Force
Import-Module -Name Mailozaurr -Force

# Send email with CSV attachment
$smtpServer = $env:SMTPServer
$smtpPort = $env:SMTPPort
$from = $env:EmailSendFromAddress
$to = $env:DRMMAgentCountToEmail
$cc = $env:DRMMAgentCountCCEmail
$SMTPUsername = $env:SMTPEmailUsername
$SMTPPassword = $env:SMTPEmailPassword
[securestring]$secStringPassword = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
[pscredential]$EmailCredential = New-Object System.Management.Automation.PSCredential ($SMTPUsername, $secStringPassword)
$subject = "Datto Agent Count Report"
$body = @"
Please find attached the Datto RMM Agent Report CSV file.
If you have questions do not reply to this message, please send a message to the NOC in NOC-Toolkit or email $env:NOCEmail.
"@
$attachment = $csvFilePath

Send-EmailMessage `
    -SmtpServer $smtpServer `
    -Port $smtpPort `
    -From $from `
    -To $to `
    -Cc $cc `
    -Credential $EmailCredential `
    -Subject $subject `
    -Body $body `
    -Attachments $attachment
