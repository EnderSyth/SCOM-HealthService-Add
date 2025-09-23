#Requires -Version 5.1

<#
.SYNOPSIS
    Connects to a single server, retrieves pending SCCM updates, and initiates their installation
.DESCRIPTION
    This script connects to a specified server using smart card credentials, queries for pending 
    SCCM updates, displays the details, and optionally installs them. Provides detailed feedback
    on the installation process.
.PARAMETER ComputerName
    Name of the computer to connect to and install updates on
.PARAMETER Force
    Skip confirmation prompt and install updates automatically
.PARAMETER TimeoutMinutes
    Timeout for remote connection in minutes (default: 5)
.EXAMPLE
    .\InstallUpdatesOnServer.ps1 -ComputerName "SERVER01"
    Connects to SERVER01, shows pending updates, and prompts for installation confirmation
.EXAMPLE
    .\InstallUpdatesOnServer.ps1 -ComputerName "SERVER01" -Force
    Connects to SERVER01 and installs updates without prompting
.NOTES
    Requires PowerShell remoting to be enabled on target server
    Requires appropriate SCCM permissions for update installation
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ComputerName,
    
    [switch]$Force,
    [int]$TimeoutMinutes = 5
)

# Function to get and install SCCM updates on remote server
function Install-SCCMUpdatesRemotely {
    param(
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential,
        [bool]$InstallUpdates = $false,
        [int]$TimeoutMinutes = 5
    )
    
    try {
        Write-Host "Connecting to $ComputerName..." -ForegroundColor Yellow
        
        # Create remote session with timeout
        $sessionOption = New-PSSessionOption -IdleTimeout (New-TimeSpan -Minutes $TimeoutMinutes).TotalMilliseconds
        $session = New-PSSession -ComputerName $ComputerName -Credential $Credential -SessionOption $sessionOption -ErrorAction Stop
        
        Write-Host "Successfully connected to $ComputerName" -ForegroundColor Green
        
        $result = Invoke-Command -Session $session -ScriptBlock {
            param($InstallFlag)
            
            try {
                Write-Host "Querying SCCM for software updates..." -ForegroundColor Cyan
                
                # Get SCCM software updates
                $updates = Get-WmiObject -Namespace "root\ccm\clientsdk" -Class "ccm_softwareupdate" -ErrorAction Stop
                
                if (-not $updates) {
                    return @{
                        ComputerName = $env:COMPUTERNAME
                        Status = "No_Updates_Found"
                        Message = "No updates found in SCCM client"
                        TotalUpdates = 0
                        PendingUpdates = @()
                        InstallationResults = @{}
                    }
                }
                
                # Categorize updates by evaluation state
                $pendingUpdates = @($updates | Where-Object { $_.EvaluationState -eq 0 -or $_.EvaluationState -eq 1 })
                $downloadingUpdates = @($updates | Where-Object { $_.EvaluationState -eq 2 -or $_.EvaluationState -eq 3 })
                $installingUpdates = @($updates | Where-Object { $_.EvaluationState -eq 6 -or $_.EvaluationState -eq 7 })
                $installedUpdates = @($updates | Where-Object { $_.EvaluationState -eq 13 })
                $failedUpdates = @($updates | Where-Object { $_.EvaluationState -eq 11 -or $_.EvaluationState -eq 12 })
                $rebootRequiredUpdates = @($updates | Where-Object { $_.EvaluationState -eq 8 -or $_.EvaluationState -eq 9 })
                
                # Get detailed info for pending updates
                $pendingUpdateDetails = @()
                foreach ($update in $pendingUpdates) {
                    $pendingUpdateDetails += @{
                        UpdateID = $update.UpdateID
                        PackageID = $update.PackageID
                        Name = $update.Name
                        ArticleID = $update.ArticleID
                        Description = $update.Description
                        Size = $update.MaxExecutionTime
                        Severity = $update.Severity
                    }
                }
                
                # Check for pending reboot
                $pendingReboot = $false
                try {
                    $pendingReboot = (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue) -or
                                   ((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue) -ne $null)
                } catch {
                    $pendingReboot = $false
                }
                
                Write-Host "Found $($updates.Count) total updates:" -ForegroundColor White
                Write-Host "  - Pending: $($pendingUpdates.Count)" -ForegroundColor Yellow
                Write-Host "  - Downloading: $($downloadingUpdates.Count)" -ForegroundColor Cyan
                Write-Host "  - Installing: $($installingUpdates.Count)" -ForegroundColor Magenta
                Write-Host "  - Installed: $($installedUpdates.Count)" -ForegroundColor Green
                Write-Host "  - Failed: $($failedUpdates.Count)" -ForegroundColor Red
                Write-Host "  - Reboot Required: $($rebootRequiredUpdates.Count)" -ForegroundColor Yellow
                
                # Installation logic
                $installationResults = @{}
                
                if ($InstallFlag -and $pendingUpdates.Count -gt 0) {
                    Write-Host "`nInitiating installation of $($pendingUpdates.Count) pending updates..." -ForegroundColor Green
                    
                    try {
                        # Get the CCM_SoftwareUpdatesManager class
                        $updateManager = [wmiclass]"root\ccm\clientsdk:CCM_SoftwareUpdatesManager"
                        
                        # Prepare update list for installation (using UpdateID)
                        # SCCM expects an array of CCM_SoftwareUpdate objects, not hashtables
                        $updateList = @()
                        foreach ($update in $pendingUpdates) {
                            Write-Host "  Queuing: $($update.Name)" -ForegroundColor Gray
                            $updateList += $update
                        }
                        
                        # Call the InstallUpdates method with the actual update objects
                        Write-Host "Calling SCCM InstallUpdates method..." -ForegroundColor Cyan
                        $installResult = $updateManager.InstallUpdates($updateList)
                        
                        # Interpret the result
                        switch ($installResult.ReturnValue) {
                            0 { 
                                $installationResults = @{
                                    Status = "Success"
                                    Message = "Installation initiated successfully"
                                    ReturnCode = $installResult.ReturnValue
                                    UpdatesQueued = $pendingUpdates.Count
                                }
                                Write-Host "✓ Installation successfully initiated!" -ForegroundColor Green
                            }
                            1 { 
                                $installationResults = @{
                                    Status = "Failed"
                                    Message = "Installation failed - Invalid method parameters"
                                    ReturnCode = $installResult.ReturnValue
                                    UpdatesQueued = 0
                                }
                                Write-Host "✗ Installation failed - Invalid parameters" -ForegroundColor Red
                            }
                            2 { 
                                $installationResults = @{
                                    Status = "Failed"
                                    Message = "Installation failed - Invalid update list"
                                    ReturnCode = $installResult.ReturnValue
                                    UpdatesQueued = 0
                                }
                                Write-Host "✗ Installation failed - Invalid update list" -ForegroundColor Red
                            }
                            default { 
                                $installationResults = @{
                                    Status = "Failed"
                                    Message = "Installation failed with return code: $($installResult.ReturnValue)"
                                    ReturnCode = $installResult.ReturnValue
                                    UpdatesQueued = 0
                                }
                                Write-Host "✗ Installation failed with return code: $($installResult.ReturnValue)" -ForegroundColor Red
                            }
                        }
                        
                    } catch {
                        $installationResults = @{
                            Status = "Error"
                            Message = "Installation error: $($_.Exception.Message)"
                            ReturnCode = -1
                            UpdatesQueued = 0
                        }
                        Write-Host "✗ Installation error: $($_.Exception.Message)" -ForegroundColor Red
                    }
                } elseif ($InstallFlag -and $pendingUpdates.Count -eq 0) {
                    $installationResults = @{
                        Status = "No_Action"
                        Message = "No pending updates to install"
                        ReturnCode = 0
                        UpdatesQueued = 0
                    }
                    Write-Host "No pending updates to install" -ForegroundColor Yellow
                }
                
                return @{
                    ComputerName = $env:COMPUTERNAME
                    Status = "Success"
                    Message = "Query completed successfully"
                    TotalUpdates = $updates.Count
                    PendingUpdates = $pendingUpdateDetails
                    DownloadingCount = $downloadingUpdates.Count
                    InstallingCount = $installingUpdates.Count
                    InstalledCount = $installedUpdates.Count
                    FailedCount = $failedUpdates.Count
                    RebootRequiredCount = $rebootRequiredUpdates.Count
                    PendingReboot = $pendingReboot
                    InstallationResults = $installationResults
                }
                
            } catch {
                Write-Host "Error querying SCCM: $($_.Exception.Message)" -ForegroundColor Red
                return @{
                    ComputerName = $env:COMPUTERNAME
                    Status = "SCCM_Error"
                    Message = "Failed to query SCCM: $($_.Exception.Message)"
                    TotalUpdates = 0
                    PendingUpdates = @()
                    InstallationResults = @{}
                }
            }
        } -ArgumentList $InstallUpdates
        
        Remove-PSSession $session
        return $result
        
    } catch {
        Write-Host "Connection failed: $($_.Exception.Message)" -ForegroundColor Red
        return @{
            ComputerName = $ComputerName
            Status = "Connection_Failed" 
            Message = "Failed to connect: $($_.Exception.Message)"
            TotalUpdates = 0
            PendingUpdates = @()
            InstallationResults = @{}
        }
    }
}

# Main script execution
try {
    Write-Host "Single Server SCCM Update Installer" -ForegroundColor Green
    Write-Host "====================================" -ForegroundColor Green
    Write-Host "Target Server: $ComputerName" -ForegroundColor Cyan
    
    # Test if server is reachable
    Write-Host "`nTesting connectivity to $ComputerName..." -ForegroundColor Yellow
    if (-not (Test-Connection -ComputerName $ComputerName -Count 2 -Quiet)) {
        Write-Warning "Server $ComputerName is not responding to ping"
        $continue = Read-Host "Continue anyway? (y/n)"
        if ($continue -ne 'y' -and $continue -ne 'Y') {
            Write-Host "Operation cancelled by user" -ForegroundColor Yellow
            exit 0
        }
    } else {
        Write-Host "✓ Server is reachable" -ForegroundColor Green
    }
    
    # Prompt for smart card credentials
    Write-Host "`nPlease provide your admin smart card credentials:" -ForegroundColor Cyan
    try {
        $credential = Get-Credential -Message "Enter your admin smart card credentials"
        if (-not $credential) {
            Write-Error "Credentials are required to connect to the remote server."
            exit 1
        }
    } catch {
        Write-Error "Failed to get credentials: $($_.Exception.Message)"
        exit 1
    }
    
    # Query for updates first
    Write-Host "`nQuerying $ComputerName for pending updates..." -ForegroundColor Cyan
    $result = Install-SCCMUpdatesRemotely -ComputerName $ComputerName -Credential $credential -InstallUpdates $false -TimeoutMinutes $TimeoutMinutes
    
    if ($result.Status -ne "Success") {
        Write-Host "`nFailed to query updates on $ComputerName" -ForegroundColor Red
        Write-Host "Status: $($result.Status)" -ForegroundColor Red
        Write-Host "Message: $($result.Message)" -ForegroundColor Red
        exit 1
    }
    
    # Display results
    Write-Host "`n" + "="*50 -ForegroundColor Green
    Write-Host "UPDATE SUMMARY FOR $($result.ComputerName.ToUpper())" -ForegroundColor Green
    Write-Host "="*50 -ForegroundColor Green
    
    Write-Host "Total Updates Found: $($result.TotalUpdates)" -ForegroundColor White
    Write-Host "Pending Updates: $($result.PendingUpdates.Count)" -ForegroundColor Yellow
    Write-Host "Currently Downloading: $($result.DownloadingCount)" -ForegroundColor Cyan
    Write-Host "Currently Installing: $($result.InstallingCount)" -ForegroundColor Magenta
    Write-Host "Already Installed: $($result.InstalledCount)" -ForegroundColor Green
    Write-Host "Failed Updates: $($result.FailedCount)" -ForegroundColor Red
    Write-Host "Require Reboot: $($result.RebootRequiredCount)" -ForegroundColor Yellow
    Write-Host "System Pending Reboot: $($result.PendingReboot)" -ForegroundColor $(if ($result.PendingReboot) { "Red" } else { "Green" })
    
    if ($result.PendingUpdates.Count -gt 0) {
        Write-Host "`nPENDING UPDATES DETAILS:" -ForegroundColor Yellow
        Write-Host "-" * 40 -ForegroundColor Yellow
        
        $result.PendingUpdates | ForEach-Object {
            Write-Host "• $($_.Name)" -ForegroundColor White
            Write-Host "  Package ID: $($_.PackageID)" -ForegroundColor Gray
            Write-Host "  Update ID: $($_.UpdateID)" -ForegroundColor Gray
            if ($_.ArticleID) {
                Write-Host "  KB Article: $($_.ArticleID)" -ForegroundColor Gray
            }
            Write-Host ""
        }
        
        # Installation prompt or force installation
        $shouldInstall = $false
        
        if ($Force) {
            Write-Host "Force flag specified - proceeding with installation..." -ForegroundColor Green
            $shouldInstall = $true
        } else {
            Write-Host "INSTALLATION OPTIONS:" -ForegroundColor Cyan
            Write-Host "This will initiate installation of $($result.PendingUpdates.Count) pending update(s)." -ForegroundColor Yellow
            Write-Host "Updates will be processed by SCCM and may require server restart." -ForegroundColor Yellow
            
            $confirmation = Read-Host "`nProceed with installation? (Y/N)"
            $shouldInstall = ($confirmation -eq 'Y' -or $confirmation -eq 'y')
        }
        
        if ($shouldInstall) {
            Write-Host "`nInitiating update installation..." -ForegroundColor Green
            $installResult = Install-SCCMUpdatesRemotely -ComputerName $ComputerName -Credential $credential -InstallUpdates $true -TimeoutMinutes $TimeoutMinutes
            
            Write-Host "`n" + "="*50 -ForegroundColor Green
            Write-Host "INSTALLATION RESULTS" -ForegroundColor Green
            Write-Host "="*50 -ForegroundColor Green
            
            if ($installResult.InstallationResults.Status) {
                $status = $installResult.InstallationResults.Status
                $message = $installResult.InstallationResults.Message
                $returnCode = $installResult.InstallationResults.ReturnCode
                $queued = $installResult.InstallationResults.UpdatesQueued
                
                switch ($status) {
                    "Success" {
                        Write-Host "✓ Installation Status: SUCCESS" -ForegroundColor Green
                        Write-Host "✓ Updates Queued: $queued" -ForegroundColor Green
                        Write-Host "✓ Message: $message" -ForegroundColor Green
                        Write-Host "`nNext Steps:" -ForegroundColor Cyan
                        Write-Host "- Monitor SCCM console for installation progress" -ForegroundColor White
                        Write-Host "- Check server for pending reboots" -ForegroundColor White
                        Write-Host "- Run this script again later to verify completion" -ForegroundColor White
                    }
                    "Failed" {
                        Write-Host "✗ Installation Status: FAILED" -ForegroundColor Red
                        Write-Host "✗ Return Code: $returnCode" -ForegroundColor Red
                        Write-Host "✗ Message: $message" -ForegroundColor Red
                    }
                    "Error" {
                        Write-Host "✗ Installation Status: ERROR" -ForegroundColor Red
                        Write-Host "✗ Message: $message" -ForegroundColor Red
                    }
                    "No_Action" {
                        Write-Host "- Installation Status: NO ACTION NEEDED" -ForegroundColor Yellow
                        Write-Host "- Message: $message" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "No installation results returned" -ForegroundColor Yellow
            }
        } else {
            Write-Host "`nInstallation cancelled by user" -ForegroundColor Yellow
        }
        
    } else {
        Write-Host "`n✓ No pending updates found!" -ForegroundColor Green
        Write-Host "Server is up to date." -ForegroundColor Green
    }
    
    Write-Host "`nScript completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Error $_.ScriptStackTrace
}
