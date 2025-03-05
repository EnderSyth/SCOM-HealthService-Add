<#
.SYNOPSIS
Adds HealthServiceWatcher objects to the UROC group in SCOM 2019.

.DESCRIPTION
This script connects to a SCOM 2019 Management Server, finds a group named UROC,
identifies all Microsoft.Windows.Computer objects in that group, finds their
associated HealthServiceWatcher objects, and adds those HealthServiceWatcher
objects to the UROC group if they aren't already members.

.PARAMETER ManagementServer
The name of the SCOM Management Server to connect to.

.EXAMPLE
.\Add-HealthServiceWatchersToGroup.ps1 -ManagementServer "SCOM01.contoso.com"

.NOTES
Author: System Administrator
Date: $(Get-Date -Format "yyyy-MM-dd")
Version: 1.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$ManagementServer
)

# Function to write log messages
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Output to console with appropriate color
    switch ($Level) {
        'Info'    { Write-Host $logMessage -ForegroundColor Green }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error'   { Write-Host $logMessage -ForegroundColor Red }
    }
}

try {
    # Step 1: Import the OperationsManager module
    Write-Log "Importing OperationsManager module..."
    Import-Module OperationsManager -ErrorAction Stop
    Write-Log "OperationsManager module imported successfully."
    
    # Step 2: Connect to the SCOM Management Server
    Write-Log "Connecting to SCOM Management Server: $ManagementServer..."
    $connection = New-SCOMManagementGroupConnection -ComputerName $ManagementServer -ErrorAction Stop
    Write-Log "Connected to SCOM Management Server successfully."
    
    # Step 3: Find the UROC group
    Write-Log "Finding UROC group..."
    $urocGroup = Get-SCOMGroup -DisplayName "UROC" -ErrorAction Stop
    
    if (-not $urocGroup) {
        throw "UROC group not found. Please check the group name and try again."
    }
    
    Write-Log "UROC group found with ID: $($urocGroup.Id)"
    
    # Step 4: Get all Microsoft.Windows.Computer objects in the UROC group
    Write-Log "Getting Microsoft.Windows.Computer objects from the UROC group..."
    $computerClass = Get-SCOMClass -Name "Microsoft.Windows.Computer" -ErrorAction Stop
    $computersInGroup = $urocGroup | Get-SCOMClassInstance -Class $computerClass -ErrorAction Stop
    
    if (-not $computersInGroup -or $computersInGroup.Count -eq 0) {
        Write-Log "No Microsoft.Windows.Computer objects found in the UROC group." -Level Warning
        return
    }
    
    Write-Log "Found $($computersInGroup.Count) Microsoft.Windows.Computer objects in the UROC group."
    
    # Step 5: Get current members of the UROC group to avoid adding duplicates
    $currentMembers = Get-SCOMGroupInstance -Group $urocGroup -ErrorAction Stop
    Write-Log "Current UROC group has $($currentMembers.Count) members."
    
    # Step 6: Find HealthServiceWatcher for each computer and add to the UROC group
    $healthServiceWatcherClass = Get-SCOMClass -Name "Microsoft.SystemCenter.HealthServiceWatcher" -ErrorAction Stop
    $addedCount = 0
    $alreadyMemberCount = 0
    
    foreach ($computer in $computersInGroup) {
        Write-Log "Processing computer: $($computer.DisplayName)..."
        
        # Find the HealthServiceWatcher for this computer
        $healthServiceWatcher = Get-SCOMClassInstance -Class $healthServiceWatcherClass | 
            Where-Object { $_.DisplayName -like "*$($computer.DisplayName)*" } -ErrorAction Continue
        
        if (-not $healthServiceWatcher) {
            Write-Log "No HealthServiceWatcher found for computer: $($computer.DisplayName)" -Level Warning
            continue
        }
        
        Write-Log "Found HealthServiceWatcher: $($healthServiceWatcher.DisplayName) for computer: $($computer.DisplayName)"
        
        # Check if the HealthServiceWatcher is already a member of the UROC group
        $isAlreadyMember = $false
        foreach ($member in $currentMembers) {
            if ($member.Id -eq $healthServiceWatcher.Id) {
                $isAlreadyMember = $true
                break
            }
        }
        
        if ($isAlreadyMember) {
            Write-Log "HealthServiceWatcher for $($computer.DisplayName) is already a member of the UROC group." -Level Info
            $alreadyMemberCount++
        } else {
            # Add the HealthServiceWatcher to the UROC group
            Write-Log "Adding HealthServiceWatcher for $($computer.DisplayName) to the UROC group..."
            Add-SCOMGroupInstance -Group $urocGroup -Instance $healthServiceWatcher -ErrorAction Continue
            Write-Log "HealthServiceWatcher for $($computer.DisplayName) added to the UROC group successfully."
            $addedCount++
        }
    }
    
    # Step 7: Summary
    Write-Log "Operation completed. Added $addedCount new HealthServiceWatcher objects to the UROC group."
    Write-Log "$alreadyMemberCount HealthServiceWatcher objects were already members of the UROC group."
    
} catch {
    Write-Log "An error occurred: $_" -Level Error
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level Error
} finally {
    # No cleanup needed as the PowerShell session handles SCOM connection cleanup
    Write-Log "Script execution completed."
}

