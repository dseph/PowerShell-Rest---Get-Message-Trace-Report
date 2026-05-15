# PowerShell-RestMessageTraceReport.ps1
#
# This script demonstrates how to retrieve message trace data via REST API calls to the reporting web service.
# It uses client credentials flow for authentication and queries message trace data for a specified date range. 
#
# Endpoints:
#   Commercial and GCC: https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace
#   GCCH:               https://reports.office365.us/ecp/reportingwebservice/reporting.svc/MessageTrace
# 
# Reference: https://learn.microsoft.com/en-us/exchange/monitoring/trace-an-email-message/graph-api-message-trace
# Note: The script is designed for GCCH but can be adapted for Commercial/GCC by changing the endpoints and token URL.  
# Required Azure permissions: MessageTrace.Read.All (application permission with admin consent)
 
# Usage:
#   1) Update the configuration section with your tenant/app details and desired date range.
#   2) Run the script in PowerShell. Output will be logged to c:\temp\msgtrace_log.txt and the full JSON response saved to c:\temp\msgtrace_output.json.
#   3) Review the log and output files for results and troubleshooting. 
#
# Required permissions:
#   MessageTrace.Read.All   - Admin consent needs to be is granted. 
#
# Note: A redirect is not needed.
#
# To decode a token go here: https://jwt.ms/

<# 
.SYNOPSIS
    Retrieve message trace data from the reporting web service using raw REST calls. This script:
 
  - Links:
    - https://learn.microsoft.com/en-us/previous-versions/office/developer/o365-enterprise-developers/jj984342(v=office.15)
    - https://learn.microsoft.com/en-us/previous-versions/office/developer/o365-enterprise-developers/jj984328(v=office.15)
#>

<# 
param(
    [ValidateSet("Commercial","GCC","GCCH")]
    [string]$Cloud = "GCCH",

    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,

    [string]$StartDate = "2026-05-10T00:00:00Z",
    [string]$EndDate   = "2026-05-11T23:59:59Z"
)
#>

# Testing Settings - Start ----------------------

$Cloud = "Commercial" # Only these are allowed: Commercial, GCC, GCCH
$TenantId     = "dd55b8f6-xxxxxxxxxxxxxxxxxxxxx"    # TODO: Update with your tenant ID
$ClientId     = "7a178bf3-xxxxxxxxxxxxxxxxxxxxx"    # TODO: Update with your app registration's client ID
$ClientSecret = "8vT8Q~xxxxxxxxxxxxxxxxxxxxxxx"     # TODO: Update with your app registration's client secret
$StartDate = "2026-04-19T00:00:00Z"                 # TODO: Update with your desired start date/time (ISO 8601 format)
$EndDate   = "2026-04-20T23:59:59Z"                 # TODO: Update with your desired end date/time (ISO 8601 format)
# Testing Settings  - End ---------------------- 
 

# =============================
# CLOUD CONFIG
# =============================

switch ($Cloud)
{
    "Commercial" {
        $TokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        $BaseUrl       = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc"
    }
    "GCC" {
        $TokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        $BaseUrl       = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc"
    }
    "GCCH" {
        $TokenEndpoint = "https://login.microsoftonline.us/$TenantId/oauth2/v2.0/token"
        $BaseUrl       = "https://reports.office365.us/ecp/reportingwebservice/reporting.svc"
    }
}


# =============================
# BUILD QUERY
# =============================

$Query = "$BaseUrl/MessageTrace`?\$filter=StartDate eq datetime'$StartDate' and EndDate eq datetime'$EndDate'"

$LogFile = "c:\temp\msgtrace_log.txt"

# =============================
# LOG FUNCTION
# =============================

function Write-Log {
    param($msg)
    $msg | Out-File -FilePath $LogFile -Append
    Write-Host $msg
}

# =============================
# AUTH (RAW - NO MSAL)
# =============================

Write-Log "Cloud: $Cloud"
Write-Log "Token Endpoint: $TokenEndpoint"
Write-Log "Base URL: $BaseUrl"

$tokenBody = @{
    client_id     = $ClientId
    scope         = "https://outlook.office365.com/.default"
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
}

$tokenResponse = Invoke-RestMethod -Method POST -Uri $TokenEndpoint `
    -Body $tokenBody -ContentType "application/x-www-form-urlencoded"

$accessToken = $tokenResponse.access_token

$accessToken

Write-Log "Token acquired"
Write-Log "Token length: $($accessToken.Length)"

# =============================
# CALL MESSAGE TRACE
# =============================

$headers = @{
    Authorization = "Bearer $accessToken"
    Accept        = "application/json"
}

Write-Log "Executing query..."
Write-Log $Query

try {
    $response = Invoke-RestMethod -Method GET -Uri $Query -Headers $headers

    Write-Log "Success"

    # Save full JSON
    $response | ConvertTo-Json -Depth 10 | Out-File "c:\temp\msgtrace_output.json"

    # Display key fields
    $response.value | Select SenderAddress, RecipientAddress, Subject, Status

}
catch {
    Write-Log "ERROR"
    Write-Log $_

    if ($_.Exception.Response) {
        #$reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        #$errorBody = $reader.ReadToEnd()
        #Write-Log "Error Body:"
        #Write-Log $errorBody
    }
}
 
