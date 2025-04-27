<#
.SYNOPSIS
This script automates the creation of shared mailboxes in Exchange Online using a CSV file.

.DESCRIPTION
The script first ensures that the `ExchangeOnlineManagement` module is installed and imported. It then connects to Exchange Online and reads mailbox details from a CSV file (`SharedMailboxes.csv`). 

For each entry in the CSV file:
- Checks if the shared mailbox already exists.
- If it exists, logs it in a separate file (`ExistingSharedMailboxes.csv`).
- If it does not exist, creates a new shared mailbox with:
  - A modified display name prefixed with "SPOT-".
  - A primary SMTP address using the `target.com` domain.
  - Optional settings such as access rights (FullAccess, SendAs) based on the CSV data and user mappings.

All failures encountered during mailbox creation are logged in `SMB_Failures.csv`.

Once the process is complete, the script disconnects from Exchange Online.

.NOTES
- The script requires administrative privileges in Exchange Online.
- The `SharedMailboxes.csv` file must be located in the same directory as the script.
- Output files (`ExistingSharedMailboxes.csv`, `SMB_Failures.csv`, and `UnmappedUsersReport.csv`) will be saved in the same directory.
- Ensure PowerShell execution policies allow the script to run.

.AUTHOR
SubjectData

.EXAMPLE
.\CreateSharedMailboxes.ps1
This will execute the script, processing the 'SharedMailboxes.csv' file, connecting to Exchange Online, and generating reports on created and existing mailboxes.
#>

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force
    Import-Module ExchangeOnlineManagement
} else {
    Import-Module ExchangeOnlineManagement
}

# Connect to Exchange Online
Connect-ExchangeOnline

# Initialize working paths and load input CSVs
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CSVPath = Join-Path $myDir "SharedMailboxes.csv"
$MappingPath = Join-Path $myDir "UserMapping.csv"
$UnmappedUsersReportPath = Join-Path $myDir "UnmappedUsersReport.csv"

$SharedMailboxes = Import-Csv -Path $CSVPath
$Mappings = Import-Csv -Path $MappingPath

$UnmappedUsers = @()
$ExistingMailboxes = @()

# Set Target Domain
$TargetDomain = "target.com"

function Resolve-Users {
    param($usersString, $mailboxDisplayName)

    $resolvedUsers = @()
    $unmapped = @()

    if (![string]::IsNullOrEmpty($usersString)) {
        $users = $usersString -split ";"  # assuming semicolon-separated
        foreach ($user in $users) {
            $user = $user.Trim()
            $mappedUPN = $UserMap[$user.ToLower()]
            if ($mappedUPN) {
                $resolvedUsers += $mappedUPN
            } else {
                $unmapped += $user
            }
        }
    }
    return @($resolvedUsers, $unmapped)
}

# Create a hash table for mappings
$UserMap = @{}
foreach ($entry in $Mappings) {
    $UserMap[$entry.'NetApp Email/UPN'.ToLower()] = $entry.'Flexera UPN'
}

foreach ($mailbox in $SharedMailboxes) {
    $firstPart = $mailbox.alias
    $originalDisplayName = $mailbox.displayname
    $prefixedDisplayName = "SPOT-$originalDisplayName"  # SPOT- prefix Name

    # Generate primary SMTP and alias addresses
    $newPrimarySmtpAddress = "$firstPart@$TargetDomain"
    #$alias1 = ($originalDisplayName + "@$TargetDomain").Replace(" ", "")
    $alias2 = ($originalDisplayName + "@target.onmicrosoft.com").Replace(" ", "")

    <#
    # List of addresses to check for existing mailbox
    $emailsToCheck = @($newPrimarySmtpAddress, $alias1, $alias2) | ForEach-Object { $_.ToLower() }

    # Search all mailboxes if any address already exists
    $existingMailbox = Get-Mailbox -ResultSize Unlimited | Where-Object {
        $mailEmails = $_.EmailAddresses | ForEach-Object { $_.ToString().ToLower() }
        $emailsToCheck | Where-Object { $mailEmails -contains "smtp:$_" -or $mailEmails -contains "SMTP:$_" }
    } | Select-Object -First 1

    if ($existingMailbox) {
        Write-Host "Mailbox already exists for $originalDisplayName, skipping creation." -ForegroundColor Yellow
        # Track the existing mailbox info
        $ExistingMailboxes += [PSCustomObject]@{
            DisplayName         = $existingMailbox.DisplayName
            PrimarySmtpAddress  = $existingMailbox.PrimarySmtpAddress
            Alias               = $existingMailbox.Alias
            ExternalDirectoryObjectId = $existingMailbox.ExternalDirectoryObjectId
        }
        continue
    } 
    else { #>
        # Prepare parameters for the new shared mailbox
        $params = @{
            Name = $prefixedDisplayName  # SPOT- prefix name
            DisplayName = $prefixedDisplayName
            # Alias = $firstPart            # alias without SPOT-
            PrimarySmtpAddress = $newPrimarySmtpAddress
            Shared = $true
        }

        try {
            $newEmail = New-Mailbox @params
            Write-Host "Created Shared Mailbox: $prefixedDisplayName ($newPrimarySmtpAddress)" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to create mailbox for $($prefixedDisplayName): $_"
            continue
        }
    #}

    # Process MbxAccess
    $mbxAccessResults = Resolve-Users $mailbox.MbxAccess $prefixedDisplayName
    $mbxAccessMapped = $mbxAccessResults[0]
    $mbxAccessUnmapped = $mbxAccessResults[1]

    foreach ($user in $mbxAccessMapped) {
        try {
            Add-MailboxPermission -Identity $newPrimarySmtpAddress -User $user -AccessRights FullAccess -AutoMapping $false
            Write-Host "Granted FullAccess to $user on $prefixedDisplayName" -ForegroundColor Cyan
        } catch {
            Write-Warning "Failed to grant FullAccess to $user on $($prefixedDisplayName): $_"
        }
    }

    # Process SendAccess
    $sendAccessResults = Resolve-Users $mailbox.SendAccess $prefixedDisplayName
    $sendAccessMapped = $sendAccessResults[0]
    $sendAccessUnmapped = $sendAccessResults[1]

    foreach ($user in $sendAccessMapped) {
        try {
            Add-RecipientPermission -Identity $newPrimarySmtpAddress -Trustee $user -AccessRights SendAs -Confirm:$false
            Write-Host "Granted SendAs to $user on $prefixedDisplayName" -ForegroundColor Yellow
        } catch {
            Write-Warning "Failed to grant SendAs to $user on $($prefixedDisplayName): $_"
        }
    }

    # Track unmapped users with displayname
    foreach ($unmappedUser in $mbxAccessUnmapped + $sendAccessUnmapped) {
        $UnmappedUsers += [PSCustomObject]@{
            DisplayName = $originalDisplayName
            UnmappedUser = $unmappedUser
        }
    }
}

if ($ExistingMailboxes.Count -gt 0) {
    $ExistingMailboxesReportPath = Join-Path $myDir "ExistingSharedMailboxes.csv"
    $ExistingMailboxes | Export-Csv -Path $ExistingMailboxesReportPath -NoTypeInformation
    Write-Host "`nExisting shared mailboxes report generated at: $ExistingMailboxesReportPath" -ForegroundColor Green
}

if ($UnmappedUsers.Count -gt 0) {
    $UnmappedUsers | Export-Csv -Path $UnmappedUsersReportPath -NoTypeInformation
    Write-Host "`nUnmapped users report generated at: $UnmappedUsersReportPath" -ForegroundColor Green
    Write-Host "`nUnmapped users report generated at: $UnmappedUsersReportPath" -ForegroundColor Green
} else {
    Write-Host "`nNo unmapped users found. No report generated." -ForegroundColor Cyan
}

Write-Host "Total shared mailboxes processed: $($SharedMailboxes.Count)"