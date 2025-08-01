# ddlcreate_new_exchange_online_runbook_enhanced.ps1
# Creates/Updates Dynamic Distribution Groups in Exchange Online
# Enhanced version with comprehensive validation, error handling, and security improvements

param(
    [Parameter(Mandatory=$false)]
    [string]$OrganizationDomain = "yourtenant.onmicrosoft.com",
    
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf = $false,
    
    [Parameter(Mandatory=$false)]
    [string]$LogPath = "",
    
    [Parameter(Mandatory=$false)]
    [int]$MaxRetries = 3
)

# Initialize logging
if ($LogPath) {
    $logFile = Join-Path $LogPath "DDG_Automation_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Start-Transcript -Path $logFile -Append
}

# Function to write output with timestamp
function Write-TimestampedOutput {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $output = "[$timestamp] $Message"
    Write-Output $output
}

# Function to validate department name format
function Test-DepartmentFormat {
    param([string]$DepartmentName)
    
    # Check minimum length (format: "12345 Name - USA" = minimum 13 chars)
    if ($DepartmentName.Length -lt 13) {
        return @{
            IsValid = $false
            Error = "Department name too short (minimum 13 characters required)"
        }
    }
    
    # Validate format pattern (5 digits, space, name, space, dash, space, 3 letters)
    if ($DepartmentName -notmatch '^[0-9]{5}\s+.+\s+-\s+[A-Z]{3}$') {
        return @{
            IsValid = $false
            Error = "Invalid format. Expected: '12345 Department Name - USA'"
        }
    }
    
    return @{
        IsValid = $true
        Error = $null
    }
}

# Function to parse department components safely
function Get-DepartmentComponents {
    param([string]$DepartmentName)
    
    try {
        # Extract department number (first 5 characters)
        $deptNum = $DepartmentName.Substring(0, 5)
        if ($deptNum -notmatch '^[0-9]{5}$') {
            throw "Invalid department number: must be exactly 5 digits"
        }
        
        # Extract country code (last 3 characters)
        $countryCode = $DepartmentName.Substring($DepartmentName.Length - 3, 3)
        if ($countryCode -notmatch '^[A-Z]{3}$') {
            throw "Invalid country code: must be exactly 3 uppercase letters"
        }
        
        # Extract department name (everything between number and country, excluding " - ")
        $startPos = 6  # After "12345 "
        $endPos = $DepartmentName.Length - 7  # Before " - USA"
        $nameLength = $endPos - $startPos
        
        if ($nameLength -le 0) {
            throw "Empty department name detected"
        }
        
        $deptName = $DepartmentName.Substring($startPos, $nameLength).Trim()
        
        if ([string]::IsNullOrWhiteSpace($deptName)) {
            throw "Department name is empty after parsing"
        }
        
        return @{
            Success = $true
            Number = $deptNum
            Name = $deptName
            CountryCode = $countryCode
            Error = $null
        }
    }
    catch {
        return @{
            Success = $false
            Number = $null
            Name = $null
            CountryCode = $null
            Error = $_.Exception.Message
        }
    }
}

# Function to create recipient filter with enhanced filtering
function New-RecipientFilter {
    param(
        [string]$DepartmentNumber,
        [string]$CountryCode
    )
    
    return "((RecipientTypeDetails -eq 'UserMailbox' -or RecipientTypeDetails -eq 'MailUser') " +
           "-and Department -like '$DepartmentNumber*' " +
           "-and Department -like '*$CountryCode' " +
           "-and Name -notlike 'SystemMailbox*' " +
           "-and Name -notlike 'CAS_*' " +
           "-and Name -notlike 'HealthMailbox*' " +
           "-and Name -notlike 'DiscoverySearchMailbox*' " +
           "-and RecipientTypeDetails -ne 'SharedMailbox' " +
           "-and RecipientTypeDetails -ne 'RoomMailbox' " +
           "-and RecipientTypeDetails -ne 'EquipmentMailbox' " +
           "-and RecipientTypeDetails -ne 'PublicFolder' " +
           "-and RecipientTypeDetails -ne 'ArbitrationMailbox')"
}

# Function to safely execute Exchange operations with retry logic
function Invoke-ExchangeOperation {
    param(
        [scriptblock]$Operation,
        [string]$OperationName,
        [int]$MaxRetries = 3
    )
    
    $attempt = 1
    while ($attempt -le $MaxRetries) {
        try {
            $result = & $Operation
            return @{
                Success = $true
                Result = $result
                Error = $null
            }
        }
        catch {
            Write-TimestampedOutput "  Attempt $attempt failed for $OperationName : $($_.Exception.Message)"
            if ($attempt -eq $MaxRetries) {
                return @{
                    Success = $false
                    Result = $null
                    Error = $_.Exception.Message
                }
            }
            $attempt++
            Start-Sleep -Seconds (2 * $attempt)  # Exponential backoff
        }
    }
}

# Main script execution starts here
Write-TimestampedOutput "========================================================================================"
Write-TimestampedOutput "Exchange Online DDG Creation Runbook Started (Enhanced Version)"
Write-TimestampedOutput "Parameters: Organization=$OrganizationDomain, WhatIf=$WhatIf, MaxRetries=$MaxRetries"
Write-TimestampedOutput "========================================================================================"

# Connect to Exchange Online with retry logic
$connectionResult = Invoke-ExchangeOperation -OperationName "Exchange Connection" -Operation {
    Connect-ExchangeOnline -ManagedIdentity -Organization $OrganizationDomain
} -MaxRetries $MaxRetries

if (-not $connectionResult.Success) {
    Write-TimestampedOutput "FATAL ERROR: Failed to connect to Exchange Online after $MaxRetries attempts: $($connectionResult.Error)"
    if ($LogPath) { Stop-Transcript }
    exit 1
}

Write-TimestampedOutput "Successfully connected to Exchange Online: $OrganizationDomain"

# Get unique departments from Exchange recipients with error handling
Write-TimestampedOutput "Retrieving departments from Exchange recipients..."

$getDepartmentsResult = Invoke-ExchangeOperation -OperationName "Get Departments" -Operation {
    Get-Recipient -ResultSize Unlimited | 
        Where-Object {
            $_.Department -ne "" -and 
            $_.Department -ne $null -and
            $_.RecipientTypeDetails -match "UserMailbox|MailUser"
        } | 
        Select-Object -Unique Department
} -MaxRetries $MaxRetries

if (-not $getDepartmentsResult.Success) {
    Write-TimestampedOutput "FATAL ERROR: Failed to retrieve departments: $($getDepartmentsResult.Error)"
    Disconnect-ExchangeOnline -Confirm:$false
    if ($LogPath) { Stop-Transcript }
    exit 1
}

$depts = $getDepartmentsResult.Result

if ($depts.Count -eq 0) {
    Write-TimestampedOutput "No departments found with valid recipients. Exiting gracefully."
    Disconnect-ExchangeOnline -Confirm:$false
    if ($LogPath) { Stop-Transcript }
    exit 0
}

Write-TimestampedOutput "Found $($depts.Count) unique departments to process"

# Initialize comprehensive counters for reporting
$stats = @{
    Created = 0
    Updated = 0
    Skipped = 0
    ValidationErrors = 0
    ProcessingErrors = 0
    TotalProcessed = 0
}

$processedDepartments = @()

# Process each department with comprehensive error handling
foreach ($dept in $depts) {
    $deptname = $dept.Department
    $stats.TotalProcessed++
    
    Write-TimestampedOutput "Processing department ($($stats.TotalProcessed)/$($depts.Count)): $deptname"
    
    # Validate department format
    $formatValidation = Test-DepartmentFormat -DepartmentName $deptname
    if (-not $formatValidation.IsValid) {
        Write-TimestampedOutput "  VALIDATION ERROR: $($formatValidation.Error)"
        $stats.ValidationErrors++
        $processedDepartments += @{
            Department = $deptname
            Status = "ValidationError"
            Error = $formatValidation.Error
        }
        continue
    }
    
    # Parse department components safely
    $components = Get-DepartmentComponents -DepartmentName $deptname
    if (-not $components.Success) {
        Write-TimestampedOutput "  PARSING ERROR: $($components.Error)"
        $stats.ValidationErrors++
        $processedDepartments += @{
            Department = $deptname
            Status = "ParsingError"
            Error = $components.Error
        }
        continue
    }
    
    Write-TimestampedOutput "  Parsed - Number: $($components.Number), Name: '$($components.Name)', Country: $($components.CountryCode)"
    
    # Construct group details
    $groupName = $components.Number + $components.CountryCode
    $displayName = $components.Name + ' - ' + $components.CountryCode
    
    # Create recipient filter
    $recipientFilter = New-RecipientFilter -DepartmentNumber $components.Number -CountryCode $components.CountryCode
    
    Write-TimestampedOutput "  Target Group - Name: $groupName, Display: $displayName"
    
    if ($WhatIf) {
        Write-TimestampedOutput "  WHAT-IF: Would process group $groupName"
        continue
    }
    
    try {
        # Check if ANY recipient exists with this name
        $existingRecipientResult = Invoke-ExchangeOperation -OperationName "Check Existing Recipient" -Operation {
            Get-Recipient -Identity $groupName -ErrorAction SilentlyContinue
        } -MaxRetries 2
        
        $existingRecipient = $existingRecipientResult.Result
        
        if ($existingRecipient) {
            Write-TimestampedOutput "  Found existing recipient: $groupName ($($existingRecipient.RecipientTypeDetails))"
            
            # Check if it's specifically a Dynamic Distribution Group
            $existingDDGResult = Invoke-ExchangeOperation -OperationName "Check Existing DDG" -Operation {
                Get-DynamicDistributionGroup -Identity $groupName -ErrorAction SilentlyContinue
            } -MaxRetries 2
            
            $existingDDG = $existingDDGResult.Result
            
            if ($existingDDG) {
                # It's a Dynamic group - update it
                Write-TimestampedOutput "  Updating existing DDG: $groupName"
                
                $updateResult = Invoke-ExchangeOperation -OperationName "Update DDG" -Operation {
                    Set-DynamicDistributionGroup -Identity $groupName `
                        -DisplayName $displayName `
                        -RecipientFilter $recipientFilter `
                        -ErrorAction Stop
                } -MaxRetries $MaxRetries
                
                if ($updateResult.Success) {
                    Write-TimestampedOutput "  SUCCESS: Updated DDG $groupName"
                    $stats.Updated++
                    $processedDepartments += @{
                        Department = $deptname
                        Status = "Updated"
                        GroupName = $groupName
                        DisplayName = $displayName
                    }
                } else {
                    Write-TimestampedOutput "  ERROR: Failed to update DDG $groupName : $($updateResult.Error)"
                    $stats.ProcessingErrors++
                    $processedDepartments += @{
                        Department = $deptname
                        Status = "UpdateError"
                        Error = $updateResult.Error
                    }
                }
            } else {
                # It's some other type of recipient - skip it
                Write-TimestampedOutput "  SKIPPING: $groupName already exists as $($existingRecipient.RecipientTypeDetails)"
                $stats.Skipped++
                $processedDepartments += @{
                    Department = $deptname
                    Status = "Skipped"
                    Reason = "Exists as $($existingRecipient.RecipientTypeDetails)"
                }
            }
        } else {
            # Nothing exists with this name - safe to create
            Write-TimestampedOutput "  Creating new DDG: $groupName"
            
            $createResult = Invoke-ExchangeOperation -OperationName "Create DDG" -Operation {
                New-DynamicDistributionGroup -Name $groupName `
                    -Alias $groupName `
                    -DisplayName $displayName `
                    -RecipientFilter $recipientFilter `
                    -IncludedRecipients "MailboxUsers, MailUsers" `
                    -ErrorAction Stop
            } -MaxRetries $MaxRetries
            
            if ($createResult.Success) {
                Write-TimestampedOutput "  SUCCESS: Created DDG $groupName"
                $stats.Created++
                $processedDepartments += @{
                    Department = $deptname
                    Status = "Created"
                    GroupName = $groupName
                    DisplayName = $displayName
                }
            } else {
                Write-TimestampedOutput "  ERROR: Failed to create DDG $groupName : $($createResult.Error)"
                $stats.ProcessingErrors++
                $processedDepartments += @{
                    Department = $deptname
                    Status = "CreateError"
                    Error = $createResult.Error
                }
            }
        }
    }
    catch {
        Write-TimestampedOutput "  UNEXPECTED ERROR processing department '$deptname': $($_.Exception.Message)"
        $stats.ProcessingErrors++
        $processedDepartments += @{
            Department = $deptname
            Status = "UnexpectedError"
            Error = $_.Exception.Message
        }
    }
}

# Generate comprehensive summary report
Write-TimestampedOutput "========================================================================================"
Write-TimestampedOutput "EXCHANGE ONLINE DDG AUTOMATION SUMMARY REPORT"
Write-TimestampedOutput "========================================================================================"
Write-TimestampedOutput "EXECUTION RESULTS:"
Write-TimestampedOutput "  ✓ Created: $($stats.Created) new Dynamic Distribution Groups"
Write-TimestampedOutput "  ✓ Updated: $($stats.Updated) existing Dynamic Distribution Groups"
Write-TimestampedOutput "  ⚠ Skipped: $($stats.Skipped) (name conflicts with other recipient types)"
Write-TimestampedOutput "  ✗ Validation Errors: $($stats.ValidationErrors) (format/parsing failures)"
Write-TimestampedOutput "  ✗ Processing Errors: $($stats.ProcessingErrors) (Exchange operation failures)"
Write-TimestampedOutput ""
Write-TimestampedOutput "STATISTICS:"
Write-TimestampedOutput "  Total Departments Found: $($depts.Count)"
Write-TimestampedOutput "  Total Departments Processed: $($stats.TotalProcessed)"
$successCount = $stats.Created + $stats.Updated
$totalErrors = $stats.ValidationErrors + $stats.ProcessingErrors
$successRate = if ($stats.TotalProcessed -gt 0) { [math]::Round(($successCount / $stats.TotalProcessed) * 100, 2) } else { 0 }
$errorRate = if ($stats.TotalProcessed -gt 0) { [math]::Round(($totalErrors / $stats.TotalProcessed) * 100, 2) } else { 0 }
Write-TimestampedOutput "  Success Rate: $successRate% ($successCount/$($stats.TotalProcessed))"
Write-TimestampedOutput "  Error Rate: $errorRate% ($totalErrors/$($stats.TotalProcessed))"

# Detailed error reporting if errors occurred
if ($totalErrors -gt 0) {
    Write-TimestampedOutput ""
    Write-TimestampedOutput "DETAILED ERROR REPORT:"
    $errorDepartments = $processedDepartments | Where-Object { $_.Status -like "*Error*" }
    foreach ($errorDept in $errorDepartments) {
        Write-TimestampedOutput "  • $($errorDept.Department) - $($errorDept.Status): $($errorDept.Error)"
    }
}

# Success summary if any groups were created/updated
if ($successCount -gt 0) {
    Write-TimestampedOutput ""
    Write-TimestampedOutput "SUCCESSFUL OPERATIONS:"
    $successDepartments = $processedDepartments | Where-Object { $_.Status -in @("Created", "Updated") }
    foreach ($successDept in $successDepartments) {
        Write-TimestampedOutput "  • $($successDept.Status): $($successDept.GroupName) ($($successDept.DisplayName))"
    }
}

Write-TimestampedOutput "========================================================================================"
Write-TimestampedOutput "Runbook Completed at $(Get-Date)"
Write-TimestampedOutput "Total Execution Time: $((Get-Date) - $startTime)"
Write-TimestampedOutput "========================================================================================"

# Clean disconnect from Exchange Online
try {
    Disconnect-ExchangeOnline -Confirm:$false
    Write-TimestampedOutput "Successfully disconnected from Exchange Online"
} catch {
    Write-TimestampedOutput "Warning: Error during Exchange Online disconnect: $($_.Exception.Message)"
}

# Stop logging if enabled
if ($LogPath) {
    Stop-Transcript
    Write-Output "Log file saved to: $logFile"
}

# Exit with appropriate code
$exitCode = if ($totalErrors -eq 0) { 0 } else { 1 }
exit $exitCode