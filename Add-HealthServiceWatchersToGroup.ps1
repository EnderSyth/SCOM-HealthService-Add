#=================================================================================
#  Add Computer to SCOM Group (Simplified Version)
#
#  This script connects to a SCOM Management Server, finds a specified computer 
#  object (using its DisplayName) and adds it to an existing SCOM Monitoring Group.
#
#  Usage:
#      .\AddComputerToSCOMGroup.ps1 -ManagementServer "SCOMServer" -GroupID "GroupName" -ComputerToAdd "ComputerDisplayName"
#
#  Note:
#      This version only supports adding a computer. It does not support removal or 
#      creating new groups.
#=================================================================================

Param(
    [Parameter(Mandatory = $true)]
    [string]$ManagementServer,
    
    [Parameter(Mandatory = $true)]
    [string]$GroupID,
    
    [Parameter(Mandatory = $true)]
    [string]$ComputerToAdd
)

# Import the SCOM Operations Manager module
Import-Module OperationsManager

# Connect to the SCOM Management Group using the provided management server
$mg = Get-SCOMManagementGroup -ComputerName $ManagementServer

# Retrieve the SCOM class for Windows Computers and then find the instance that matches our computer name
$computerClass = Get-SCOMClass -Name "Microsoft.Windows.Computer"
$computerObject = Get-SCOMClassInstance -Class $computerClass | Where-Object { $_.DisplayName -eq $ComputerToAdd }

if (-not $computerObject) {
    Write-Error "Computer '$ComputerToAdd' not found in SCOM."
    exit
}

# Get the target monitoring group by its name
$group = Get-SCOMMonitoringGroup -Name $GroupID

if (-not $group) {
    Write-Error "Group '$GroupID' not found in SCOM."
    exit
}

# Check if the computer is already a member of the group
$existingMember = $group.Members | Where-Object { $_.Id -eq $computerObject.Id }
if ($existingMember) {
    Write-Host "Computer '$ComputerToAdd' is already a member of group '$GroupID'."
    exit
}

# Add the computer object to the group
$group | Add-SCOMMonitoringGroupMember -Member $computerObject

Write-Host "Computer '$ComputerToAdd' has been added to group '$GroupID'."
