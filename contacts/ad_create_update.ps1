# ==============================================================================
# Description:
#   Retrieves user data from Active Directory and synchronizes it with the
#   ITFlow API. For each AD user, the script determines an ITFlow client based
#   on the OU (found in the DistinguishedName) and then either creates or
#   updates the contact record in ITFlow.
#
# Requirements:
#   - RSAT / ActiveDirectory PowerShell module
#   - Valid ITFlow API endpoint and API key
#
# Customize the configuration section below to suit your environment.
# ==============================================================================

# -----------------------------
# Configuration
# -----------------------------
$ITFlowApiKey = "YOUR_ITFLOW_API_KEY"                          # Replace with your ITFlow API key
$ITFlowApiUrl = "https://itflow.company.com/api/v1"   # Replace with your ITFlow API base URL

# Define OU-to-client mapping.
# Keys are substrings you expect in the AD DistinguishedName.
# Values are the corresponding ITFlow Client IDs.
$OUClientMapping = @{
    "OU=Sales"   = 101;   # ITFlow client ID for Sales
    "OU=Support" = 102;   # ITFlow client ID for Support
    "OU=IT"      = 103;   # ITFlow client ID for IT
    # Add additional OU mappings as needed.
}

# Define the AD search base (adjust to your domain structure)
$ADSearchBase = "DC=yourdomain,DC=com"

# Define log file path
$LogFilePath = ".\Sync-ADContactsToITFlow.log"

# -----------------------------
# Logging Function
# -----------------------------
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"

    # Write to console
    Write-Output $logMessage

    # Append to log file
    Add-Content -Path $LogFilePath -Value $logMessage
}

# -----------------------------
# Functions
# -----------------------------

function Get-ITFlowContact {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Email
    )
    $uri = "$ITFlowApiUrl/contacts?email=$($Email)"
    Write-Log "Retrieving ITFlow contact for email '$Email' using URI: $uri" "INFO"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers @{ "X-API-Key" = $ITFlowApiKey }
        Write-Log "Successfully retrieved contact for '$Email'" "INFO"
        return $response
    }
    catch {
        Write-Log "Error retrieving ITFlow contact for email '$Email': $_" "ERROR"
        return $null
    }
}

function Create-ITFlowContact {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$ContactData
    )
    $uri = "$ITFlowApiUrl/contacts"
    Write-Log "Creating ITFlow contact for '$($ContactData.email)' with data: $(ConvertTo-Json $ContactData)" "INFO"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post `
                     -Headers @{ "X-API-Key" = $ITFlowApiKey } `
                     -Body ($ContactData | ConvertTo-Json -Depth 5) `
                     -ContentType "application/json"
        Write-Log "Created contact: $($ContactData.name)" "INFO"
        return $response
    }
    catch {
        Write-Log "Error creating ITFlow contact for '$($ContactData.email)': $_" "ERROR"
        return $null
    }
}

function Update-ITFlowContact {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ContactId,
        [Parameter(Mandatory)]
        [hashtable]$ContactData
    )
    $uri = "$ITFlowApiUrl/contacts/$ContactId"
    Write-Log "Updating ITFlow contact (ID: $ContactId) for '$($ContactData.email)' with data: $(ConvertTo-Json $ContactData)" "INFO"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Put `
                     -Headers @{ "X-API-Key" = $ITFlowApiKey } `
                     -Body ($ContactData | ConvertTo-Json -Depth 5) `
                     -ContentType "application/json"
        Write-Log "Updated contact: $($ContactData.name)" "INFO"
        return $response
    }
    catch {
        Write-Log "Error updating ITFlow contact for '$($ContactData.email)': $_" "ERROR"
        return $null
    }
}

function Get-ClientIdFromOU {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DistinguishedName
    )
    Write-Log "Determining client ID for DN: $DistinguishedName" "INFO"
    foreach ($ouKey in $OUClientMapping.Keys) {
        if ($DistinguishedName -like "*$ouKey*") {
            Write-Log "Matched OU '$ouKey' to client ID: $($OUClientMapping[$ouKey])" "INFO"
            return $OUClientMapping[$ouKey]
        }
    }
    Write-Log "No client mapping found for DN: $DistinguishedName" "WARNING"
    return $null
}

# -----------------------------
# Main Script Execution
# -----------------------------

Write-Log "Script execution started." "INFO"

# Ensure the ActiveDirectory module is available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Log "The ActiveDirectory module is not available. Please install RSAT tools." "ERROR"
    exit 1
}
Import-Module ActiveDirectory
Write-Log "ActiveDirectory module imported successfully." "INFO"

# Retrieve Active Directory users.
try {
    Write-Log "Querying Active Directory users from search base: $ADSearchBase" "INFO"
    $ADUsers = Get-ADUser -Filter { Enabled -eq $true } -SearchBase $ADSearchBase `
                -Properties DisplayName, EmailAddress, Title, Department, DistinguishedName
    Write-Log "Retrieved $($ADUsers.Count) users from Active Directory." "INFO"
}
catch {
    Write-Log "Error retrieving AD users: $_" "ERROR"
    exit 1
}

foreach ($adUser in $ADUsers) {
    Write-Log "Processing AD user: $($adUser.DisplayName)" "INFO"
    # Determine the ITFlow Client based on the user's OU
    $clientId = Get-ClientIdFromOU -DistinguishedName $adUser.DistinguishedName
    if (-not $clientId) {
        Write-Log "No client mapping found for user '$($adUser.DisplayName)'. Skipping..." "WARNING"
        continue
    }
    
    # Build the contact data for ITFlow
    $contactData = @{
        name       = $adUser.DisplayName
        email      = $adUser.EmailAddress
        title      = $adUser.Title
        department = $adUser.Department
        clientId   = $clientId
        # Add any additional fields required by your ITFlow API here.
    }
    
    if (-not $contactData.email) {
        Write-Log "No email address found for '$($adUser.DisplayName)'. Skipping contact." "WARNING"
        continue
    }

    # Check if this contact already exists in ITFlow (lookup by email)
    $existingContact = Get-ITFlowContact -Email $contactData.email

    if ($existingContact -and $existingContact.id) {
        Write-Log "Contact exists in ITFlow (ID: $($existingContact.id)). Proceeding with update." "INFO"
        Update-ITFlowContact -ContactId $existingContact.id -ContactData $contactData
    }
    else {
        Write-Log "No existing contact found for '$($adUser.EmailAddress)'. Proceeding with creation." "INFO"
        Create-ITFlowContact -ContactData $contactData
    }
}

Write-Log "Script execution completed." "INFO"
