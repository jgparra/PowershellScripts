<#
.SYNOPSIS
    Generates a report of Microsoft Entra group members and their managers,
    suitable for Organizational Data uploads into M365 / Viva Insights.

.DESCRIPTION
    ###############Disclaimer#####################################################
    The sample scripts are not supported under any Microsoft standard support 
    program or service. The sample scripts are provided AS IS without warranty  
    of any kind. Microsoft further disclaims all implied warranties including,  
    without limitation, any implied warranties of merchantability or of fitness for 
    a particular purpose. The entire risk arising out of the use or performance of  
    the sample scripts and documentation remains with you. In no event shall 
    Microsoft, its authors, or anyone else involved in the creation, production, or 
    delivery of the scripts be liable for any damages whatsoever (including, 
    without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use 
    of or inability to use the sample scripts or documentation, even if Microsoft 
    has been advised of the possibility of such damages.
    ###############Disclaimer#####################################################

    PURPOSE:
        Queries Microsoft Entra (via Microsoft Graph) to collect user profile details,
        manager information, and group membership context. The resulting CSV output is
        intended as a source file for Organizational Data uploads in Microsoft 365 or
        Viva Advanced Insights.

    The script supports three input modes selected through an interactive menu:
        1. All users from all Entra groups.
        2. Members of a single group searched by display name.
        3. Users listed in a CSV file (requires a column named 'UPN').

    OUTPUT FILE: EntraGroupMembersExport.csv (saved in the current working directory).

    EXPORTED COLUMNS:
        GroupName, DisplayName, JobTitle, Department, OfficeLocation, CompanyName,
        UsageLocation, PreferredLanguage, StreetAddress, City, State, Country,
        PostalCode, UserPrincipalName, ManagerUPN

    REQUIREMENTS:
        - Microsoft Graph PowerShell Module
          https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0
        - Required Graph API scopes: Group.Read.All, User.Read.All

    VERSION HISTORY:
        - 07/26/2024: V1 - Initial release (Alejandro Lopez, Dean Cron)
        - 04/21/2026: V2 - Added inline documentation and comments (J.G. Parra, GitHub Copilot)

    AUTHOR(S):
        - Alejandro Lopez  - Alejanl@Microsoft.com
        - Dean Cron        - DeanCron@Microsoft.com
        - J.G. Parra       - (contributor)
          Assisted by      : GitHub Copilot (AI assistant)

.EXAMPLE
    # Run the script and follow the interactive menu
    .\Get-GroupsMembersManagers.ps1

.EXAMPLE
    # To use CSV mode, prepare a file like:
    #   UPN
    #   user1@contoso.com
    #   user2@contoso.com
    # Then select option 3 and provide the CSV path when prompted.
#>

# Ensure the script can be run by the current user without requiring a system-wide policy change.
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Read-MenuOption
# PURPOSE : Displays the interactive source-selection menu and returns the
#           validated option chosen by the operator (integer 1-4).
#           Loops until a valid selection is entered.
# RETURNS : [int] 1 = All groups | 2 = Single group | 3 = CSV | 4 = Exit
# ─────────────────────────────────────────────────────────────────────────────
function Read-MenuOption {
    while ($true) {
        Write-Host ""
        Write-Host "Select source option:" -ForegroundColor Yellow
        Write-Host "  1) All users from all groups (original logic)"
        Write-Host "  2) Members from one group (search by group display name)"
        Write-Host "  3) Users from CSV (required column: UPN)"
        Write-Host "  4) Exit"

        $selectedOption = Read-Host "Enter your option (1-4)"

        # Accept only digits 1 through 4; re-prompt on any other input.
        if ($selectedOption -match '^[1-4]$') {
            return [int]$selectedOption
        }

        Write-Host "Invalid option. Please select 1, 2, 3, or 4." -ForegroundColor Red
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Test-CtrlCPressed
# PURPOSE : Non-blocking check for a Ctrl+C key combination.
#           Used inside loops to allow graceful interruption without crashing.
#           When [Console]::TreatControlCAsInput is $true, Ctrl+C does NOT
#           terminate the process; it is instead read as a regular key event
#           and evaluated here.
# RETURNS : [bool] $true if Ctrl+C was detected; $false otherwise.
# ─────────────────────────────────────────────────────────────────────────────
function Test-CtrlCPressed {
    try {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq [ConsoleKey]::C -and ($key.Modifiers -band [ConsoleModifiers]::Control)) {
                return $true
            }
        }
    }
    catch {
        # If the console host does not support key reading (e.g. ISE, non-interactive),
        # silently swallow the exception and report no key pressed.
        return $false
    }

    return $false
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Add-GroupMembersToWorkItems
# PURPOSE : Retrieves all members of the given Entra group via Microsoft Graph
#           and appends a work-item entry (GroupName + UserLookup) for each
#           member to the shared $WorkItems collection.
#           UserLookup stores the member's object ID (used later with Get-MgUser).
# PARAMS  :
#   -Group      [MicrosoftGraphGroup]  The group object to enumerate.
#   -WorkItems  [System.Collections.ArrayList]  Shared list accumulating work items.
# ─────────────────────────────────────────────────────────────────────────────
function Add-GroupMembersToWorkItems {
    param(
        [Parameter(Mandatory = $true)]$Group,
        [System.Collections.ArrayList]$WorkItems
    )

    if ($null -eq $Group) {
        throw "Group cannot be null."
    }

    if ($null -eq $WorkItems) {
        throw "WorkItems collection cannot be null."
    }

    # Fetch every member of the group (paging handled automatically by -All).
    $members = Get-MgGroupMember -GroupId $Group.Id -All
    foreach ($member in $members) {
        # Store the member's object ID as UserLookup; the display name is not
        # used here because Get-MgUser accepts the object ID directly.
        [void]$WorkItems.Add(
            [PSCustomObject]@{
                GroupName  = $Group.DisplayName
                UserLookup = $member.Id
            }
        )
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN SCRIPT BODY
# ─────────────────────────────────────────────────────────────────────────────

# Present the interactive menu and capture the operator's choice.
$menuOption = Read-MenuOption

# Exit immediately if the operator chose to quit (option 4).
if ($menuOption -eq 4) {
    Write-Host "Exit selected. No data exported." -ForegroundColor Yellow
    return
}

# ── Runtime counters and result collections ───────────────────────────────────
$interrupted     = $false   # Tracks whether Ctrl+C was detected during processing.
$processedUsers  = 0        # Number of work items visited in the processing loop.
$successfulUsers = 0        # Work items that were resolved successfully via Graph.
$failedUsers     = 0        # Work items that raised an error during Graph lookup.

# ArrayList is used instead of a standard array to avoid costly re-allocation
# on each append (+=), which matters for large data sets.
$userDetails = New-Object System.Collections.ArrayList  # Final rows written to CSV.
$workItems   = New-Object System.Collections.ArrayList  # Intermediate work queue.

# ── Output configuration ──────────────────────────────────────────────────────
# Output CSV is always written to the current working directory.
$outputPath   = Join-Path -Path (Get-Location) -ChildPath "EntraGroupMembersExport.csv"

# Comma-separated list of user properties to request from Graph.
# Limiting the property set reduces payload size and improves performance.
$propertyList = "id,displayName,userPrincipalName,department,jobTitle,officeLocation,companyName,usageLocation,preferredLanguage,streetAddress,city,state,country,postalCode"

# ── Ctrl+C interception setup ─────────────────────────────────────────────────
# Save the original setting so it can be restored in the finally block.
$originalTreatControlCAsInput = $false

try {
    $originalTreatControlCAsInput         = [Console]::TreatControlCAsInput
    # Redirect Ctrl+C from the OS-level process terminator to the key-read buffer
    # so Test-CtrlCPressed can detect it without killing the process mid-export.
    [Console]::TreatControlCAsInput = $true
}
catch {
    # Some PowerShell hosts (e.g. VS Code integrated terminal, ISE) do not expose
    # the Console class in the same way; warn and continue without custom handling.
    Write-Host "Warning: Ctrl+C custom handling is not available in this host." -ForegroundColor Yellow
}

try {
    # ── Step 1: Authenticate to Microsoft Graph ───────────────────────────────
    # Requests the minimum scopes needed: read groups and read user profiles.
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes "Group.Read.All", "User.Read.All"

    # ── Step 2: Build the work-item queue based on the selected input mode ────
    switch ($menuOption) {

        # ── Option 1: All users from every group in the tenant ────────────────
        1 {
            Write-Host "Collecting users from all groups (original behavior). This might take some time..." -ForegroundColor Yellow
            $groups     = Get-MgGroup -All   # Retrieve every Entra group (handles paging).
            $groupIndex = 1

            foreach ($group in $groups) {
                # Check for Ctrl+C between group iterations to allow graceful exit.
                if (Test-CtrlCPressed) {
                    Write-Host "`n`nCapture interrupted by user (Ctrl+C)." -ForegroundColor Yellow
                    Write-Host "Continuing with data collected so far..." -ForegroundColor Yellow
                    $interrupted = $true
                    break
                }

                Write-Host "Processed Group: $groupIndex of $($groups.Count)" -ForegroundColor Yellow
                Add-GroupMembersToWorkItems -Group $group -WorkItems $workItems
                $groupIndex++
            }
        }

        # ── Option 2: Members of a single group (searched by display name) ────
        2 {
            $groupName = Read-Host "Enter group display name"
            if ([string]::IsNullOrWhiteSpace($groupName)) {
                throw "Group display name is required."
            }

            # Escape single quotes to prevent OData filter injection.
            $escapedGroupName  = $groupName.Replace("'", "''")
            $matchingGroups    = @(Get-MgGroup -Filter "DisplayName eq '$escapedGroupName'" -All)

            if ($matchingGroups.Count -eq 0) {
                throw "No group found with display name '$groupName'."
            }

            if ($matchingGroups.Count -gt 1) {
                # Warn the operator; the first match (alphabetical by Graph default) is used.
                Write-Host "More than one group matched '$groupName'. The first match will be used." -ForegroundColor Yellow
            }

            $targetGroup = $matchingGroups[0]
            if ($null -eq $targetGroup -or [string]::IsNullOrWhiteSpace($targetGroup.Id)) {
                throw "Unable to resolve a valid group object for '$groupName'."
            }

            Write-Host "Using group: $($targetGroup.DisplayName) ($($targetGroup.Id))" -ForegroundColor Yellow
            Add-GroupMembersToWorkItems -Group $targetGroup -WorkItems $workItems
        }

        # ── Option 3: Users listed in a CSV file ──────────────────────────────
        3 {
            Write-Host "CSV mode selected." -ForegroundColor Yellow
            Write-Host "Required input format: CSV file with fixed column name 'UPN'." -ForegroundColor Yellow

            <#
            TODO (future enhancement):
            - Add an option to load CSV and display detected columns.
            - Let the operator choose which column contains the user identifier.
            #>

            $csvPath = Read-Host "Enter full path of CSV file"
            if (-not (Test-Path -Path $csvPath)) {
                throw "CSV file does not exist: $csvPath"
            }

            $csvRows = Import-Csv -Path $csvPath
            if (-not $csvRows) {
                throw "CSV file is empty."
            }

            # Validate that the mandatory 'UPN' column is present before processing.
            if (-not ($csvRows[0].PSObject.Properties.Name -contains "UPN")) {
                throw "CSV must contain a column named 'UPN'."
            }

            foreach ($row in $csvRows) {
                # Skip rows where UPN is empty or whitespace-only.
                if ([string]::IsNullOrWhiteSpace($row.UPN)) {
                    continue
                }

                # For CSV mode, GroupName is intentionally left blank because
                # there is no group context associated with these users.
                [void]$workItems.Add(
                    [PSCustomObject]@{
                        GroupName  = ""
                        UserLookup = $row.UPN   # UPN is passed directly to Get-MgUser.
                    }
                )
            }
        }
    }

    # ── Step 3: Validate that at least one work item was collected ─────────────
    if ($workItems.Count -eq 0) {
        Write-Host "No users found to process. Empty file will be exported." -ForegroundColor Yellow
    }

    $totalUsers = $workItems.Count
    Write-Host "Total users to export: $totalUsers" -ForegroundColor Yellow

    # ── Step 4: Fetch detailed user profiles from Microsoft Graph ─────────────
    foreach ($workItem in $workItems) {
        # Allow the operator to abort between user lookups without losing data.
        if (Test-CtrlCPressed) {
            Write-Host "`n`nCapture interrupted by user (Ctrl+C)." -ForegroundColor Yellow
            Write-Host "Continuing with data collected so far..." -ForegroundColor Yellow
            $interrupted = $true
            break
        }

        $processedUsers++

        # Calculate and display percentage progress.
        $percentComplete = 0
        if ($totalUsers -gt 0) {
            $percentComplete = [Math]::Round(($processedUsers / $totalUsers) * 100, 2)
        }

        Write-Progress -Activity "Collecting user profile data" -Status "Processed $processedUsers of $totalUsers" -PercentComplete $percentComplete

        try {
            # Retrieve user profile with the pre-defined property set and expand
            # the Manager navigation property in a single Graph call.
            $user       = Get-MgUser -UserId $workItem.UserLookup -Property $propertyList -ExpandProperty Manager
            $managerUpn = $null

            if ($null -ne $user.Manager) {
                # Manager UPN is stored inside AdditionalProperties because it is
                # not a top-level field on the DirectoryObject returned by -ExpandProperty.
                $managerUpn = $user.Manager.AdditionalProperties['userPrincipalName']
            }

            # Build the output record and add it to the results collection.
            [void]$userDetails.Add(
                [PSCustomObject]@{
                    GroupName         = $workItem.GroupName
                    DisplayName       = $user.DisplayName
                    JobTitle          = $user.JobTitle
                    Department        = $user.Department
                    OfficeLocation    = $user.OfficeLocation
                    CompanyName       = $user.CompanyName
                    UsageLocation     = $user.UsageLocation
                    PreferredLanguage = $user.PreferredLanguage
                    StreetAddress     = $user.StreetAddress
                    City              = $user.City
                    State             = $user.State
                    Country           = $user.Country
                    PostalCode        = $user.PostalCode
                    UserPrincipalName = $user.UserPrincipalName
                    ManagerUPN        = $managerUpn
                }
            )

            $successfulUsers++
        }
        catch {
            # Log failures but continue processing remaining work items.
            $failedUsers++
            Write-Host "Unable to get details for this account: $($workItem.UserLookup)" -ForegroundColor Yellow
        }
    }

    # Clear the progress bar from the console once processing is complete.
    Write-Progress -Activity "Collecting user profile data" -Completed

    # ── Step 5: Export results to CSV ─────────────────────────────────────────
    try {
        # Export-Csv with -NoTypeInformation omits the PowerShell type header row,
        # producing a clean CSV that can be imported directly into M365 / Viva Insights.
        $userDetails | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

        if ($interrupted) {
            Write-Host "Partial export completed due to user interruption." -ForegroundColor Yellow
        }
        else {
            Write-Host "Finished processing." -ForegroundColor Yellow
        }

        Write-Host "Results exported to: $outputPath" -ForegroundColor Yellow
        Write-Host "Summary -> Total work items: $totalUsers | Processed: $processedUsers | Success: $successfulUsers | Failed: $failedUsers | Exported rows: $($userDetails.Count)" -ForegroundColor Yellow
    }
    catch {
        Write-Host "Hit error while exporting results: $($_)" -ForegroundColor Red
    }
}
catch {
    Write-Host "Execution failed: $($_)" -ForegroundColor Red
}
finally {
    # ── Cleanup: Restore Ctrl+C behavior and disconnect from Graph ────────────
    try {
        # Always restore the original Ctrl+C behavior, regardless of how the script ended.
        [Console]::TreatControlCAsInput = $originalTreatControlCAsInput
    }
    catch {
        # Suppress errors from hosts that do not support TreatControlCAsInput.
    }

    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
}