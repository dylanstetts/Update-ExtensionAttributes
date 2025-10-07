# Update-ExtensionAttributes.ps1 - Complete User Guide

## Overview
This PowerShell script provides a robust solution for updating Active Directory extension attributes (extensionAttribute1-15) for Azure AD users. It supports both single user updates and bulk CSV processing, with automatic fallback from Microsoft Graph to Exchange Online when needed.

## Features
- **Dual Update Methods**: Single user or bulk CSV processing
- **Smart Fallback**: Automatically falls back from Graph API to Exchange Online when Graph updates are blocked
- **Comprehensive Error Handling**: Detailed logging and error reporting for troubleshooting
- **Progress Tracking**: Real-time progress updates for bulk operations
- **Flexible Authentication**: Optional connection management with `-SkipConnect` parameter

## Prerequisites

### Required PowerShell Modules
The script will automatically attempt to install and import required modules if they're not available:

1. **Microsoft.Graph** (Primary method)
   - **Required Scope**: `User.ReadWrite.All`
   - **Purpose**: Updates `onPremisesExtensionAttributes` via Graph API
   - **Installation**: `Install-Module Microsoft.Graph -Scope CurrentUser`

2. **ExchangeOnlineManagement** (Fallback method)
   - **Required Permissions**: Exchange Administrator role or equivalent
   - **Purpose**: Updates `CustomAttribute1-15` when Graph API is blocked
   - **Installation**: `Install-Module ExchangeOnlineManagement -Scope CurrentUser`

### Azure AD Permissions
- **For Graph API**: User.ReadWrite.All application permission or delegated permission
- **For Exchange Online**: Exchange Administrator role or User Administrator role with mailbox permissions

## Script Capabilities

### What the Script DOES Handle:
**Extension Attributes**: Updates extensionAttribute1 through extensionAttribute15  
**Multiple User Types**: Works with cloud-only and hybrid users  
**Automatic Fallback**: Graph API → Exchange Online when blocked by external service authority  
**Flexible Input**: User Principal Name (UPN) or Object ID (GUID)  
**Bulk Processing**: CSV file with multiple users  
**Null Values**: Can clear existing attributes by setting them to null  
**Error Recovery**: Continues processing other users if individual updates fail  
**Progress Reporting**: Real-time status updates for bulk operations  

### What the Script Does NOT Handle:
**Other User Properties**: Only extension attributes, not other AD properties  
**Groups or Contacts**: Only works with user objects  
**On-Premises AD**: Does not directly update on-premises Active Directory  
**Schema Extensions**: Only built-in extension attributes 1-15  
**Batch Transactions**: Each user is processed individually (no rollback capability)  
**Custom Attributes**: Does not support custom schema extensions beyond standard extension attributes  

## Usage

### Method 1: Single User Update

Update extension attributes for a single user using a hashtable:

```powershell
# Basic single user update
.\Update-ExtensionAttributes.ps1 -UserIdOrUpn "john.doe@contoso.com" -ExtensionAttributes @{ 
    extensionAttribute1 = 'Department-IT'
    extensionAttribute2 = 'Location-NYC'
    extensionAttribute10 = 'Manager-JaneSmith'
}

# Using Object ID instead of UPN
.\Update-ExtensionAttributes.ps1 -UserIdOrUpn "12345678-1234-1234-1234-123456789012" -ExtensionAttributes @{ 
    extensionAttribute1 = 'Department-HR'
    extensionAttribute5 = $null  # Clears the attribute
}

# Skip authentication (if already connected)
.\Update-ExtensionAttributes.ps1 -UserIdOrUpn "user@contoso.com" -ExtensionAttributes @{ extensionAttribute1 = 'Value' } -SkipConnect
```

### Method 2: Bulk CSV Update

Process multiple users from a CSV file:

```powershell
# Basic bulk update
.\Update-ExtensionAttributes.ps1 -CsvPath ".\users-to-update.csv"

# Bulk update with skip authentication
.\Update-ExtensionAttributes.ps1 -CsvPath "C:\Data\ExtensionUpdates.csv" -SkipConnect
```

## CSV Template Usage

### CSV Format Requirements

#### Required Column
- **UserIdOrUpn**: User identifier - can be either:
  - User Principal Name (UPN): `john.doe@contoso.com`
  - User Object ID (GUID): `12345678-1234-1234-1234-123456789012`

#### Extension Attribute Columns
Include any combination of the following columns for the attributes you want to update:
- `extensionAttribute1` through `extensionAttribute15`

### CSV File Instructions

1. **Copy the Template**: Use `ExtensionAttributes-Template.csv` as your starting point
2. **Rename the File**: Save as your desired filename (e.g., `users-to-update.csv`)
3. **Fill in the Data**:
   - **UserIdOrUpn**: Enter the user's UPN or Object ID
   - **Extension Attributes**: Enter values for the attributes you want to set
   - **Empty Values**: Leave cells empty for attributes you don't want to change
   - **Clear Values**: Use `null` to clear/empty an existing attribute value
4. **Save and Run**: Execute the script with `-CsvPath` parameter

### CSV Examples

#### Setting Values
```csv
UserIdOrUpn,extensionAttribute1,extensionAttribute2,extensionAttribute5
john.doe@contoso.com,Department-IT,Location-NYC,Project-Alpha
jane.smith@contoso.com,Department-HR,Location-Remote,Project-Beta
```

#### Clearing Values
```csv
UserIdOrUpn,extensionAttribute1,extensionAttribute2
jane.smith@contoso.com,null,Location-Remote
user3@contoso.com,Department-Finance,null
```

#### Mixed Operations (Set, Clear, and Skip)
```csv
UserIdOrUpn,extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute10
user1@contoso.com,Department-HR,,Project-Alpha,
user2@contoso.com,null,Location-Remote,,Manager-Updated
user3@contoso.com,Department-Finance,Location-Chicago,null,null
```

## Authentication and Connection Management

### Automatic Connection (Default)
The script will automatically:
1. Connect to Microsoft Graph with `User.ReadWrite.All` scope
2. Connect to Exchange Online only if Graph fallback is needed

### Manual Connection Management
Use `-SkipConnect` if you want to manage connections yourself:

```powershell
# Connect manually before running script
Connect-MgGraph -Scopes 'User.ReadWrite.All'
Connect-ExchangeOnline

# Run script without auto-connection
.\Update-ExtensionAttributes.ps1 -CsvPath "users.csv" -SkipConnect
```

## Error Handling and Troubleshooting

### Common Scenarios

1. **Graph API Blocked**: When users originate from external services (hybrid environments), Graph API may be blocked. The script automatically falls back to Exchange Online.

2. **Permission Issues**: Ensure you have:
   - `User.ReadWrite.All` for Graph API
   - Exchange Administrator role for Exchange Online fallback

3. **User Not Found**: Verify UPN or Object ID is correct and user exists in the tenant.

4. **Module Installation**: The script attempts automatic module installation. Run PowerShell as Administrator if installation fails.

### Progress and Reporting

For bulk operations, the script provides:
- Real-time progress updates
- Success/failure counts
- Detailed error messages for failed updates
- Summary report at completion

### Example Output
```
Processing bulk update from CSV: .\users-to-update.csv
Loaded 150 user(s) from CSV for bulk processing.
[Graph] Updated onPremisesExtensionAttributes for 'user1@contoso.com' successfully.
[Graph] Blocked by external-service authority for 'user2@contoso.com'. Will attempt Exchange Online fallback.
[EXO] Set-User updated CustomAttribute(s) for 'user2@contoso.com'.

Bulk update completed:
  Successful: 148
  Failed: 2
  Total: 150
```

## Parameter Reference

| Parameter | Type | Required | Parameter Set | Description |
|-----------|------|----------|---------------|-------------|
| `UserIdOrUpn` | String | Yes | SingleUser | User's UPN or Object ID for single user updates |
| `ExtensionAttributes` | Hashtable | Yes | SingleUser | Hashtable of extension attributes to update |
| `CsvPath` | String | Yes | BulkCsv | Path to CSV file for bulk processing |
| `SkipConnect` | Switch | No | Both | Skip automatic authentication connections |

## Best Practices

1. **Test First**: Always test with a small subset of users before bulk operations
2. **Backup Data**: Document current extension attribute values before making changes
3. **Use Object IDs**: Object IDs are more reliable than UPNs for API calls
4. **Monitor Progress**: For large bulk operations, monitor the progress output
5. **Check Permissions**: Verify all required permissions before starting bulk operations
6. **Handle Failures**: Review failure messages and retry failed users if needed

## Limitations

- **Rate Limits**: Large bulk operations may hit API rate limits
- **Session Timeouts**: For very large operations, authentication sessions may timeout
- **Concurrent Execution**: Do not run multiple instances simultaneously on the same users
- **Exchange Online**: Some recipient types may not support all CustomAttribute updates

## Support and Troubleshooting

For issues:
1. Check the detailed error messages in the console output
2. Verify user permissions and module installations
3. Test with a single user before bulk operations
4. Review the CSV file format if using bulk mode
5. Ensure network connectivity to Microsoft Graph and Exchange Online endpoints