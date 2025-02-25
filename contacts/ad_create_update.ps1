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
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers @{ "X-API-Key" = $ITFlowApiKey }
        return $response
    }
    catch {
        Write-Warning "Error retrieving ITFlow contact for email '$Email': $_"
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
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post `
                     -Headers @{ "X-API-Key" = $ITFlowApiKey } `
                     -Body ($ContactData | ConvertTo-Json -Depth 5) `
                     -ContentType "application/json"
        Write-Output "Created contact: $($ContactData.name)"
        return $response
    }
    catch {
        Write-Warning "Error creating ITFlow contact for '$($ContactData.email)': $_"
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
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Put `
                     -Headers @{ "X-API-Key" = $ITFlowApiKey } `
                     -Body ($ContactData | ConvertTo-Json -Depth 5) `
                     -ContentType "application/json"
        Write-Output "Updated contact: $($ContactData.name)"
        return $response
    }
    catch {
        Write-Warning "Error updating ITFlow contact for '$($ContactData.email)': $_"
        return $null
    }
}

function Get-ClientIdFromOU {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DistinguishedName
    )
    foreach ($ouKey in $OUClientMapping.Keys) {
        if ($DistinguishedName -like "*$ouKey*") {
            return $OUClientMapping[$ouKey]
        }
    }
    return $null
}

# -----------------------------
# Main Script Execution
# -----------------------------

# Ensure the ActiveDirectory module is available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Error "The ActiveDirectory module is not available. Please install RSAT tools."
    exit 1
}
Import-Module ActiveDirectory

# Retrieve Active Directory users.
# (Use Get-ADUser or Get-ADContact as needed. Adjust the properties as required.)
try {
    $ADUsers = Get-ADUser -Filter { Enabled -eq $true } -SearchBase $ADSearchBase `
                -Properties DisplayName, EmailAddress, Title, Department, DistinguishedName
}
catch {
    Write-Error "Error retrieving AD users: $_"
    exit 1
}

foreach ($adUser in $ADUsers) {
    # Determine the ITFlow Client based on the user's OU
    $clientId = Get-ClientIdFromOU -DistinguishedName $adUser.DistinguishedName
    if (-not $clientId) {
        Write-Warning "No client mapping found for '$($adUser.DisplayName)' (DN: $($adUser.DistinguishedName)). Skipping..."
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
        Write-Warning "No email address found for '$($adUser.DisplayName)'. Skipping contact."
        continue
    }

    # Check if this contact already exists in ITFlow (lookup by email)
    $existingContact = Get-ITFlowContact -Email $contactData.email

    if ($existingContact -and $existingContact.id) {
        # Update the existing contact
        Update-ITFlowContact -ContactId $existingContact.id -ContactData $contactData
    }
    else {
        # Create a new contact
        Create-ITFlowContact -ContactData $contactData
    }
}
