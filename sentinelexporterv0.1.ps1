# Import required modules
Import-Module AzureAD
Import-Module AzureADPreview
Import-Module Microsoft.Graph.Security
Import-Module ImportExcel

# Define Azure AD Application credentials and Graph API endpoint
$clientId = "<App-Client-ID>"
$clientSecret = "<App-Client-Secret>"
$tenantId = "<Tenant-ID>"
$resource = "https://graph.microsoft.com"

# Authenticate with Azure AD and get an access token
$token = Get-AzureADToken -ClientApplicationId $clientId -ClientSecret $clientSecret -TenantId $tenantId -Resource $resource

# Set the access token
$accessToken = $token.AccessToken

# Define Microsoft Sentinel query
$query = "SecurityIncidents | project IncidentName, Severity, Status, IncidentNumber, CreatedDateTime, ResolvedDateTime"

# Define the Graph API request headers
$headers = @{
    'Authorization' = "Bearer $accessToken"
}

# Make a request to Microsoft Sentinel using Graph API
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/security/alerts/query" -Headers $headers -Method POST -ContentType 'application/json' -Body @{
    'query' = $query
}

# Export the data to an Excel file
$exportPath = "C:\Path\To\Export\SecurityIncidents.xlsx" 
$workSheetName = "SecurityIncidents"  

Write-Host "Security incidents exported to $exportPath"
