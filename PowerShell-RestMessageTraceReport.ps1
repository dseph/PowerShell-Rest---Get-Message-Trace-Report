# PowerShell-RestMessageTraceReport.ps1

<# 

.SYNOPSIS
---------
This script demonstrates how to retrieve message trace data via REST API calls to the reporting web service.
It uses client credentials flow for authentication and queries message trace data for a specified date range. 
 
  - Links:
    - https://learn.microsoft.com/en-us/previous-versions/office/developer/o365-enterprise-developers/jj984342(v=office.15)
    - https://learn.microsoft.com/en-us/previous-versions/office/developer/o365-enterprise-developers/jj984328(v=office.15)
    - https://learn.microsoft.com/en-us/exchange/monitoring/trace-an-email-message/graph-api-message-trace

.ENDPOINTS
----------
    Commercial and GCC: https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace
    GCCH:               https://reports.office365.us/ecp/reportingwebservice/reporting.svc/MessageTrace


 .USSAGE
 -------
   1) Update the configuration section with your tenant/app details and desired date range.
   2) Run the script in PowerShell. Output will be logged to c:\temp\msgtrace_log.txt and the full JSON response saved to c:\temp\msgtrace_output.json.
   3) Review the log and output files for results and troubleshooting. 
 
.PERMISSIONS
---------------------
Application prmissions in Purview:
    MessageTrace.Read.All   - Admin consent needs to be is granted. 

If you encounter an error stating "No permission to access the report for the organization," it typically indicates that the service principal 
needs admin permissions.

This error typically occurs in the Microsoft Exchange Online Reporting Web Service when an API app lacks the correct administrator roles. 
Even if API  permissions are granted, your service principal or app account must be assigned an Exchange Admin, Global Reader, or 
Security Reader role to successfully pull the data.
To resolve this issue, you must assign an Exchange Admin role to the service principal object or the account making the API calls. You can 
do this by following these steps:  
1.	Navigate to Entra ID: Go to the Microsoft Entra admin center and sign in with an administrative account.  
2.	Assign a Role:
o	Navigate to Roles and administrators.
o	Search for Global Reader, Security Reader, or Exchange Administrator.
o	Add your Microsoft Entra ID (Azure AD) application as a member/assign the role to the service principal.  
3.	Verify API Permissions: Ensure your app registration has either ReportingWebService.Read (Delegated) or ReportingWebService.Read.All (Application) 
    granted under the Office 365 Exchange Online API. [1, 2]
4.	Get a Fresh Token: Because permission changes can take time to propagate or require a new login, generate a fresh authentication token before 
    retrying your script or tool. 

Note: A redirect is not needed.

Note: To decode a token go here: https://jwt.ms/
 
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
        $Scopes = "https://outlook.office365.com/.default"
    }
    "GCC" {
        $TokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        $BaseUrl       = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc"
        $Scopes = "https://outlook.office365.com/.default"
    }
    "GCCH" {
        $TokenEndpoint  = "https://login.microsoftonline.us/$TenantId/oauth2/v2.0/token"
        $BaseUrl        = "https://outlook.office365.us/ecp/reportingwebservice/reporting.svc"
        $Scopes         = "https://outlook.office365.us/.default"
    }
}


# =============================
# BUILD QUERY
# =============================

$Query = "$BaseUrl/MessageTrace`?\$filter=StartDate eq datetime'$StartDate' and EndDate eq datetime'$EndDate'"
#$Query = "$BaseUrl/MessageTrace?`$top=5"

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
    scope         = $Scopes
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
 
