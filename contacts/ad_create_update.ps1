# Import the ActiveDirectory module if not already imported
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "ActiveDirectory module is required. Please install it."
    exit
}

# ITFlow API settings
$itflowUrl = "https://your-itflow-instance.example.com/api/v1"   # Update as needed
$apiKey = "YOUR_API_KEY_HERE"  # Replace with your API key

# Set the AD search base for contacts (update for your environment)
$contactsOU = "OU=Contacts,DC=example,DC=com"

# Define log file and exclusion file paths
$logFile = "C:\Logs\ITFlowContactSync.log"
$exclusionFile = "C:\Logs\ITFlowContactExclusions.txt"

# Ensure log directory exists
$logDir = Split-Path -Path $logFile -Parent
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

# Function to get the exclusion list (contacts not to process)
function Get-ExcludedContacts {
    if (Test-Path $exclusionFile) {
        return Get-Content $exclusionFile
    }
    return @()
}

# Function to add a contact email to the exclusion list
function Add-ExcludedContact {
    param([string]$contactEmail)
    if (-not ((Get-ExcludedContacts) -contains $contactEmail)) {
        Add-Content -Path $exclusionFile -Value $contactEmail
    }
}

# Centralized logging function
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage
    Add-Content -Path $logFile -Value $logMessage
}

Write-Log "Script started."

# Function to query ITFlow for a specific contact by email
function Get-ITFlowContactByEmail {
    param(
         [string]$Email,
         [string]$ApiUrl,
         [string]$ApiKey
    )
    $uri = "$ApiUrl/contacts/read.php?api_key=$ApiKey&contact_email=$Email"
    try {
         $response = Invoke-RestMethod -Method Get -Uri $uri
         Write-Log "Retrieved ITFlow contact data for $Email."
         return $response.data
    } catch {
         Write-Log "Error fetching ITFlow contact for ${Email}: $_" "ERROR"
         return $null
    }
}

# Function to get all clients from ITFlow
function Get-ITFlowClients {
    param (
        [string]$ApiUrl,
        [string]$ApiKey
    )
    $uri = "$ApiUrl/clients/read.php?api_key=$ApiKey"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Get
        Write-Log "Successfully retrieved clients from ITFlow."
        return $response.data
    } catch {
        Write-Log "Error fetching ITFlow clients: $_" "ERROR"
        return $null
    }
}

# Get AD contacts from the specified OU
try {
    $adContacts = Get-ADUser -Filter * -SearchBase $contactsOU -Properties displayName, title, department, mail, telephoneNumber, mobile, ipPhone, DistinguishedName
    Write-Log "Retrieved $($adContacts.Count) AD contacts from $contactsOU."
} catch {
    Write-Log "Error querying AD contacts: $_" "ERROR"
    exit
}

# Initialize arrays for processed contacts
$existingContacts = @()
$newContacts = @()

# Process each AD contact
foreach ($adContact in $adContacts) {
    if ($adContact.mail) {
        $itflowContactData = Get-ITFlowContactByEmail -Email $adContact.mail -ApiUrl $itflowUrl -ApiKey $apiKey
        if ($itflowContactData -and (($itflowContactData -is [array] -and $itflowContactData.Count -gt 0) -or ($itflowContactData -isnot [array]))) {
            # If the contact exists in ITFlow, update its details.
            if ($itflowContactData -is [array]) {
                $itflowContact = $itflowContactData[0]
            } else {
                $itflowContact = $itflowContactData
            }
            
            # Use AD's ipPhone for the contact extension (if available)
            $extension = if ($adContact.ipPhone) { $adContact.ipPhone } else { "" }
            
            # Process phone numbers: replace a leading +61 with 00
            $phoneNumber = if ($adContact.telephoneNumber -match "^\+61") { $adContact.telephoneNumber -replace "^\+61", "00" } else { $adContact.telephoneNumber }
            $mobileNumber = if ($adContact.mobile -match "^\+61") { $adContact.mobile -replace "^\+61", "00" } else { $adContact.mobile }
            
            # Build JSON payload for update (following ITFlow sample format)
            $body = @"
{
    "api_key" : "$apiKey",
    "contact_id" : "$($itflowContact.contact_id)",
    "contact_name" : "$($adContact.displayName)",
    "contact_title" : "$($adContact.title)",
    "contact_department" : "$($adContact.department)",
    "contact_email" : "$($adContact.mail)",
    "contact_phone" : "$phoneNumber",
    "contact_extension" : "$extension",
    "contact_mobile" : "$mobileNumber",
    "contact_notes" : "",
    "contact_auth_method" : "",
    "contact_important" : "",
    "contact_billing" : "",
    "contact_technical" : "",
    "contact_location_id" : "",
    "client_id" : "$($itflowContact.contact_client_id)"
}
"@
            $uri = "$itflowUrl/contacts/update.php"
            try {
                Write-Log "Updating contact with payload: $body"
                Invoke-RestMethod -Uri $uri -Method Post -Body $body -ContentType "application/json" | Out-Null
                Write-Log "Updated contact: $($adContact.displayName)"
                $existingContacts += $adContact
            } catch {
                Write-Log "Error updating contact $($adContact.displayName): $_" "ERROR"
            }
        } else {
            # If the contact does not exist in ITFlow, add to new contacts list.
            Write-Log "New contact found: $($adContact.displayName)"
            $newContacts += $adContact
        }
    } else {
        Write-Log "AD contact '$($adContact.displayName)' has no email address. Skipping." "WARNING"
    }
}

# Retrieve ITFlow clients for allocation
$clients = Get-ITFlowClients -ApiUrl $itflowUrl -ApiKey $apiKey
if (-not $clients) {
    Write-Log "Failed to retrieve ITFlow clients. New contacts will not be created." "ERROR"
    exit
}

# Process each new contact
foreach ($newContact in $newContacts) {
    # Check exclusion list so that we don't repeatedly prompt about the same contact.
    $excludedList = Get-ExcludedContacts
    if ($newContact.mail -and ($excludedList -contains $newContact.mail)) {
        Write-Log "Skipping new contact $($newContact.displayName) as it is in the exclusion list."
        continue
    }

    Write-Log "Processing new contact: $($newContact.displayName), $($newContact.mail)"
    $clientId = $null

    # Always prompt for client allocation (no auto-assignment)
    Write-Host "Allocate contact '$($newContact.displayName)' ($($newContact.mail)) to a client. Select an option:"
    Write-Host "0. Do not create this contact"
    for ($i = 0; $i -lt $clients.Count; $i++) {
        Write-Host "$($i + 1). $($clients[$i].client_name)"
    }
    do {
        $selection = Read-Host "Enter the number of the client (0-$($clients.Count))"
        $valid = $selection -match '^\d+$' -and [int]$selection -ge 0 -and [int]$selection -le $clients.Count
        if (-not $valid) {
            Write-Host "Invalid selection. Please enter a number between 0 and $($clients.Count)."
        }
    } while (-not $valid)

    if ([int]$selection -eq 0) {
        Write-Log "User opted to not create the contact: $($newContact.displayName)"
        if ($newContact.mail) { Add-ExcludedContact -contactEmail $newContact.mail }
        continue
    } else {
        $clientId = $clients[$selection - 1].client_id
        Write-Log "User selected client ID: $clientId for contact: $($newContact.displayName)"
    }

    # Process phone numbers: replace +61 with 00
    $phoneNumber = if ($newContact.telephoneNumber -match "^\+61") { $newContact.telephoneNumber -replace "^\+61", "00" } else { $newContact.telephoneNumber }
    $mobileNumber = if ($newContact.mobile -match "^\+61") { $newContact.mobile -replace "^\+61", "00" } else { $newContact.mobile }
    
    # Get extension from AD's ipPhone attribute
    $extension = if ($newContact.ipPhone) { $newContact.ipPhone } else { "" }

    # Build JSON payload for new contact creation (using ITFlow sample format)
    $body = @"
{
    "api_key" : "$apiKey",
    "contact_name" : "$($newContact.displayName)",
    "contact_title" : "$($newContact.title)",
    "contact_department" : "$($newContact.department)",
    "contact_email" : "$($newContact.mail)",
    "contact_phone" : "$phoneNumber",
    "contact_extension" : "$extension",
    "contact_mobile" : "$mobileNumber",
    "contact_notes" : "",
    "contact_auth_method" : "local",
    "contact_important" : "0",
    "contact_billing" : "1",
    "contact_technical" : "0",
    "contact_location_id" : "0",
    "client_id" : "$clientId"
}
"@
    $uri = "$itflowUrl/contacts/create.php"
    try {
        Write-Log "Creating contact with payload: $body"
        Invoke-RestMethod -Uri $uri -Method Post -Body $body -ContentType "application/json" | Out-Null
        Write-Log "Created contact: $($newContact.displayName) for client ID: $clientId"
    } catch {
        Write-Log "Error creating contact $($newContact.displayName): $_" "ERROR"
    }
}

Write-Log "Script completed successfully."
