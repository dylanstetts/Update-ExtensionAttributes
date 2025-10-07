[CmdletBinding(DefaultParameterSetName = 'SingleUser')]
param(
    [Parameter(Mandatory, ParameterSetName = 'SingleUser')]
    [string] $UserIdOrUpn,  # Can be userId (GUID) or UPN

    # Example: @{ extensionAttribute1 = 'VAL1'; extensionAttribute10 = $null }
    [Parameter(Mandatory, ParameterSetName = 'SingleUser')]
    [hashtable] $ExtensionAttributes,

    # Path to CSV file for bulk processing
    [Parameter(Mandatory, ParameterSetName = 'BulkCsv')]
    [string] $CsvPath,

    # Optional: skip interactive auth if you already have sessions
    [switch] $SkipConnect
)

# ---- Helper: map Graph ext attrs -> EXO CustomAttribute params ----
function Convert-ToExoCustomAttributeParams {
    param([hashtable] $Attributes)

    $exoParams = @{}
    foreach ($k in $Attributes.Keys) {
        if ($k -match '^extensionAttribute([1-9]|1[0-5])$') {
            # FIX: use -replace to extract the numeric suffix
            $num = [int]($k -replace 'extensionAttribute', '')
            $exoParams["CustomAttribute$num"] = $Attributes[$k]
        } else {
            Write-Verbose "Ignoring non-supported key '$k' (only extensionAttribute1..15 are mapped)"
        }
    }
    return $exoParams
}

# ---- Helper: process CSV file and return array of user update objects ----
function Read-BulkUpdateCsv {
    param([string] $CsvPath)

    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found at path: $CsvPath"
    }

    try {
        $csvData = Import-Csv -Path $CsvPath -ErrorAction Stop
    } catch {
        throw "Failed to read CSV file: $($_.Exception.Message)"
    }

    if ($csvData.Count -eq 0) {
        throw "CSV file is empty or contains no data rows."
    }

    $results = @()
    $validExtensionAttributes = @('extensionAttribute1', 'extensionAttribute2', 'extensionAttribute3', 'extensionAttribute4', 'extensionAttribute5', 
                                 'extensionAttribute6', 'extensionAttribute7', 'extensionAttribute8', 'extensionAttribute9', 'extensionAttribute10',
                                 'extensionAttribute11', 'extensionAttribute12', 'extensionAttribute13', 'extensionAttribute14', 'extensionAttribute15')

    foreach ($row in $csvData) {
        if (-not $row.UserIdOrUpn) {
            Write-Warning "Skipping row with missing UserIdOrUpn: $($row | ConvertTo-Json -Compress)"
            continue
        }

        $extensionAttributes = @{}
        foreach ($prop in $row.PSObject.Properties) {
            if ($prop.Name -in $validExtensionAttributes -and $prop.Value) {
                # Handle special values: empty string or 'null' means clear the attribute
                if ($prop.Value -eq 'null' -or $prop.Value -eq '') {
                    $extensionAttributes[$prop.Name] = $null
                } else {
                    $extensionAttributes[$prop.Name] = $prop.Value
                }
            }
        }

        if ($extensionAttributes.Count -eq 0) {
            Write-Warning "Skipping user '$($row.UserIdOrUpn)' - no valid extension attributes found."
            continue
        }

        $results += @{
            UserIdOrUpn = $row.UserIdOrUpn
            ExtensionAttributes = $extensionAttributes
        }
    }

    if ($results.Count -eq 0) {
        throw "No valid user records found in CSV file."
    }

    Write-Host "Loaded $($results.Count) user(s) from CSV for bulk processing." -ForegroundColor Cyan
    return $results
}

# ---- Helper: update extension attributes for a single user ----
function Update-UserExtensionAttributes {
    param(
        [string] $UserIdOrUpn,
        [hashtable] $ExtensionAttributes
    )

    # ---- Build Graph PATCH body ----
    $graphBody = @{ onPremisesExtensionAttributes = @{} }
    foreach ($k in $ExtensionAttributes.Keys) {
        $graphBody.onPremisesExtensionAttributes[$k] = $ExtensionAttributes[$k]
    }
    $graphJson = $graphBody | ConvertTo-Json -Depth 5

    # ---- Attempt Graph update first ----
    $graphUri = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($UserIdOrUpn))"
    $graphFailedWithExternalServiceError = $false

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $graphUri -Body $graphJson -ContentType 'application/json' -ErrorAction Stop
        Write-Host "[Graph] Updated onPremisesExtensionAttributes for '$UserIdOrUpn' successfully." -ForegroundColor Green
        return $true
    }
    catch {
        $raw = $_
        # Try to extract a meaningful message
        $msg = $raw.ErrorDetails?.Message
        if (-not $msg) { $msg = $raw.Exception.Message }

        # The specific condition from your tenant:
        # "Unable to update the specified properties for objects that have originated within an external service."
        if ($msg -match 'originated within an external service') {
            $graphFailedWithExternalServiceError = $true
            Write-Warning "[Graph] Blocked by external-service authority for '$UserIdOrUpn'. Will attempt Exchange Online fallback."
        } else {
            Write-Error "[Graph] Update failed for '$UserIdOrUpn' with an unexpected error: $msg"
            return $false
        }
    }

    # ---- Fallback to Exchange Online (only if the specific error occurred) ----
    if ($graphFailedWithExternalServiceError) {
        try {
            # Build EXO parameter set
            $exoParams = Convert-ToExoCustomAttributeParams -Attributes $ExtensionAttributes
            if ($exoParams.Count -eq 0) {
                Write-Error "No valid extensionAttribute1..15 keys were provided for '$UserIdOrUpn'."
                return $false
            }
            $exoParams['Identity'] = $UserIdOrUpn

            # Try Set-User first (works for most recipient types)
            try {
                Set-User @exoParams -ErrorAction Stop
                Write-Host "[EXO] Set-User updated CustomAttribute(s) for '$UserIdOrUpn'." -ForegroundColor Green
                return $true
            } catch {
                Write-Warning "[EXO] Set-User failed for '$UserIdOrUpn': $($_.Exception.Message). Trying Set-Mailbox as a secondary fallback..."

                try {
                    Set-Mailbox @exoParams -ErrorAction Stop
                    Write-Host "[EXO] Set-Mailbox updated CustomAttribute(s) for '$UserIdOrUpn'." -ForegroundColor Green
                    return $true
                } catch {
                    Write-Error "[EXO] Fallback failed for '$UserIdOrUpn': $($_.Exception.Message)"
                    return $false
                }
            }
        } catch {
            Write-Error "Failed to process Exchange Online fallback for '$UserIdOrUpn': $($_.Exception.Message)"
            return $false
        }
    }

    return $false
}

# ---- Optional connects ----
if (-not $SkipConnect) {
    # Graph (needs at least User.ReadWrite.All)
    try {
        if (-not (Get-Module Microsoft.Graph -ListAvailable)) { Import-Module Microsoft.Graph -ErrorAction Stop }
        if (-not (Get-MgContext)) { Connect-MgGraph -Scopes 'User.ReadWrite.All' -ErrorAction Stop }
    } catch {
        throw "Failed to ensure Microsoft Graph connection: $($_.Exception.Message)"
    }

    # We'll connect to EXO only if fallback is required
}

# ---- Main execution based on parameter set ----
if ($PSCmdlet.ParameterSetName -eq 'SingleUser') {
    # Single user mode
    Write-Host "Processing single user: $UserIdOrUpn" -ForegroundColor Cyan
    $success = Update-UserExtensionAttributes -UserIdOrUpn $UserIdOrUpn -ExtensionAttributes $ExtensionAttributes
    if (-not $success) {
        throw "Failed to update extension attributes for user '$UserIdOrUpn'"
    }
} 
elseif ($PSCmdlet.ParameterSetName -eq 'BulkCsv') {
    # Bulk CSV mode
    Write-Host "Processing bulk update from CSV: $CsvPath" -ForegroundColor Cyan
    
    # Check if EXO connection might be needed and establish it upfront for bulk operations
    $needsExoConnection = $false
    
    # Read and validate CSV data
    $bulkUsers = Read-BulkUpdateCsv -CsvPath $CsvPath
    
    $successCount = 0
    $failureCount = 0
    $totalUsers = $bulkUsers.Count
    
    foreach ($userUpdate in $bulkUsers) {
        Write-Progress -Activity "Updating Extension Attributes" -Status "Processing $($userUpdate.UserIdOrUpn)" -PercentComplete (($successCount + $failureCount) / $totalUsers * 100)
        
        # If this is the first EXO fallback needed, establish connection
        if (-not $needsExoConnection) {
            # Test if we might need EXO by attempting a quick graph call on first user
            # If it fails with external service error, we'll connect to EXO
        }
        
        $success = Update-UserExtensionAttributes -UserIdOrUpn $userUpdate.UserIdOrUpn -ExtensionAttributes $userUpdate.ExtensionAttributes
        
        if ($success) {
            $successCount++
        } else {
            $failureCount++
            # Connect to EXO if not already connected and we encounter external service errors
            if (-not $needsExoConnection -and -not $SkipConnect) {
                try {
                    if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
                        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -ErrorAction Stop
                    }
                    if (-not (Get-Module ExchangeOnlineManagement)) {
                        Import-Module ExchangeOnlineManagement -ErrorAction Stop
                    }
                    if (-not (Get-PSSession | Where-Object {$_.ConfigurationName -eq 'Microsoft.Exchange'})) {
                        Connect-ExchangeOnline -ShowProgress:$false -ErrorAction Stop
                        $needsExoConnection = $true
                        Write-Host "Established Exchange Online connection for fallback processing." -ForegroundColor Yellow
                    }
                } catch {
                    Write-Warning "Failed to establish Exchange Online session for fallback: $($_.Exception.Message)"
                }
            }
        }
    }
    
    Write-Progress -Activity "Updating Extension Attributes" -Completed
    
    Write-Host "`nBulk update completed:" -ForegroundColor Cyan
    Write-Host "  Successful: $successCount" -ForegroundColor Green
    Write-Host "  Failed: $failureCount" -ForegroundColor Red
    Write-Host "  Total: $totalUsers" -ForegroundColor Cyan
    
    if ($failureCount -gt 0) {
        Write-Warning "Some users failed to update. Check the error messages above for details."
    }
}
