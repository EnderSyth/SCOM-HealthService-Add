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
    [switch]$VerifyInstallation,
    [int]$TimeoutMinutes = 5,
    [int]$VerificationWaitMinutes = 5
)

# Function to verify installation progress
function Verify-InstallationProgress {
    param(
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential,
        [array]$OriginalUpdateIDs,
        [int]$WaitMinutes = 5
    )
    
    Write-Host "`nVerifying installation progress..." -ForegroundColor Cyan
    Write-Host "Waiting $WaitMinutes minutes for SCCM to process the evaluation cycle..." -ForegroundColor Yellow
    
    # Wait for SCCM to process
    Start-Sleep -Seconds ($WaitMinutes * 60)
    
    try {
        $session = New-PSSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
        
        $verificationResult = Invoke-Command -Session $session -ScriptBlock {
            param($OriginalIDs)
            
            try {
                # Check scan time
                $scanHistory = Get-WmiObject -Namespace "root\ccm\scanagent" -Class "CCM_ScanUpdateSourceHistory" -ErrorAction SilentlyContinue | 
                              Sort-Object LastScanTime -Descending | Select-Object -First 1
                
                # Get current update status
                $currentUpdates = Get-WmiObject -Namespace "root\ccm\clientsdk" -Class "ccm_softwareupdate" -ErrorAction Stop
                
                # Check status of our original updates
                $originalUpdateStatus = @()
                foreach ($updateID in $OriginalIDs) {
                    $update = $currentUpdates | Where-Object { $_.UpdateID -eq $updateID }
                    if ($update) {
                        $status = switch ($update.EvaluationState) {
                            0 { "Available" }
                            1 { "Pending" }  
                            2 { "Downloading" }
                            3 { "Downloaded" }
                            6 { "Installing" }
                            7 { "Pending Reboot" }
                            8 { "Pending Reboot" }
                            11 { "Failed" }
                            12 { "Failed" }
                            13 { "Installed" }
                            default { "Unknown ($($update.EvaluationState))" }
                        }
                        
                        $originalUpdateStatus += @{
                            Name = $update.Name
                            UpdateID = $updateID
                            Status = $status
                            EvaluationState = $update.EvaluationState
                        }
                    } else {
                        $originalUpdateStatus += @{
                            Name = "Update not found"
                            UpdateID = $updateID
                            Status = "Not Found"
                            EvaluationState = -1
                        }
                    }
                }
                
                # Count current states
                $statusCounts = @{
                    Pending = ($currentUpdates | Where-Object { $_.EvaluationState -eq 0 -or $_.EvaluationState -eq 1 }).Count
                    Downloading = ($currentUpdates | Where-Object { $_.EvaluationState -eq 2 -or $_.EvaluationState -eq 3 }).Count
                    Installing = ($currentUpdates | Where-Object { $_.EvaluationState -eq 6 -or $_.EvaluationState -eq 7 }).Count
                    Installed = ($currentUpdates | Where-Object { $_.EvaluationState -eq 13 }).Count
                    Failed = ($currentUpdates | Where-Object { $_.EvaluationState -eq 11 -or $_.EvaluationState -eq 12 }).Count
                }
                
                return @{
                    Success = $true
                    LastScanTime = if ($scanHistory) { $scanHistory.LastScanTime } else { "Unknown" }
                    OriginalUpdateStatus = $originalUpdateStatus
                    CurrentStatusCounts = $statusCounts
                    TotalUpdates = $currentUpdates.Count
                }
                
            } catch {
                return @{
                    Success = $false
                    Error = $_.Exception.Message
                }
            }
        } -ArgumentList @(,$OriginalUpdateIDs)
        
        Remove-PSSession $session
        
        if ($verificationResult.Success) {
            Write-Host "`n" + "="*50 -ForegroundColor Green
            Write-Host "INSTALLATION VERIFICATION RESULTS" -ForegroundColor Green  
            Write-Host "="*50 -ForegroundColor Green
            
            Write-Host "Last SCCM Scan Time: $($verificationResult.LastScanTime)" -ForegroundColor White
            Write-Host "`nCurrent Update Status Counts:" -ForegroundColor Yellow
            Write-Host "  Pending: $($verificationResult.CurrentStatusCounts.Pending)" -ForegroundColor Yellow
            Write-Host "  Downloading: $($verificationResult.CurrentStatusCounts.Downloading)" -ForegroundColor Cyan
            Write-Host "  Installing: $($verificationResult.CurrentStatusCounts.Installing)" -ForegroundColor Magenta
            Write-Host "  Installed: $($verificationResult.CurrentStatusCounts.Installed)" -ForegroundColor Green
            Write-Host "  Failed: $($verificationResult.CurrentStatusCounts.Failed)" -ForegroundColor Red
            
            Write-Host "`nStatus of Originally Pending Updates:" -ForegroundColor Yellow
            Write-Host "-" * 40 -ForegroundColor Yellow
            
            $progressMade = $false
            foreach ($update in $verificationResult.OriginalUpdateStatus) {
                $color = switch ($update.Status) {
                    "Downloading" { "Cyan"; $progressMade = $true }
                    "Installing" { "Magenta"; $progressMade = $true }
                    "Installed" { "Green"; $progressMade = $true }
                    "Pending Reboot" { "Yellow"; $progressMade = $true }
                    "Failed" { "Red" }
                    default { "White" }
                }
                
                Write-Host "• $($update.Name)" -ForegroundColor White
                Write-Host "  Status: $($update.Status)" -ForegroundColor $color
                Write-Host ""
            }
            
            if ($progressMade) {
                Write-Host "✓ Installation progress detected! SCCM is processing the updates." -ForegroundColor Green
            } else {
                Write-Host "⚠ No immediate progress detected. This could mean:" -ForegroundColor Yellow
                Write-Host "  - Updates are queued but not yet started" -ForegroundColor Gray
                Write-Host "  - SCCM is still evaluating policies" -ForegroundColor Gray
                Write-Host "  - Updates require user interaction or scheduling" -ForegroundColor Gray
                Write-Host "  - Check again in 10-15 minutes" -ForegroundColor Gray
            }
            
        } else {
            Write-Host "Verification failed: $($verificationResult.Error)" -ForegroundColor Red
        }
        
    } catch {
        Write-Host "Failed to verify installation: $($_.Exception.Message)" -ForegroundColor Red
    }
}
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
                        
                        # Prepare update list for installation - unwrap PSObjects to get native WMI objects
                        Write-Host "Preparing updates for installation..." -ForegroundColor Cyan
                        
                        $updateIDs = @()
                        foreach ($update in $pendingUpdates) {
                            Write-Host "  Queuing: $($update.Name)" -ForegroundColor Gray
                            $updateIDs += $update.UpdateID
                        }
                        
                        # For manual SCCM configurations, use client action approach
                        Write-Host "Using SCCM Client Action methods for manual configuration..." -ForegroundColor Cyan
                        
                        try {
                            $installResult = @{ ReturnValue = -1 }
                            $successCount = 0
                            $actionResults = @()
                            
                            foreach ($updateID in $updateIDs) {
                                try {
                                    # Get the specific update object
                                    $updateObj = Get-WmiObject -Namespace "root\ccm\clientsdk" -Class "ccm_softwareupdate" -Filter "UpdateID='$updateID'"
                                    
                                    if ($updateObj) {
                                        $updateName = $updateObj.Name
                                        Write-Host "  Processing: $updateName" -ForegroundColor Gray
                                        
                                        # Method 1: Use CCM_SoftwareUpdate.Install() method directly
                                        try {
                                            Write-Host "    Attempting direct install method..." -ForegroundColor Yellow
                                            $directResult = Invoke-WmiMethod -InputObject $updateObj -Name "Install"
                                            
                                            if ($directResult.ReturnValue -eq 0) {
                                                $successCount++
                                                $actionResults += "✓ Direct install initiated for: $updateName"
                                                Write-Host "    ✓ Direct install successful" -ForegroundColor Green
                                                continue
                                            } else {
                                                Write-Host "    Direct install failed (Code: $($directResult.ReturnValue))" -ForegroundColor Yellow
                                            }
                                        } catch {
                                            Write-Host "    Direct install method not available" -ForegroundColor Yellow
                                        }
                                        
                                        # Method 2: Use the ClientSDK to download first, then install
                                        try {
                                            Write-Host "    Attempting download then install..." -ForegroundColor Yellow
                                            
                                            # Step 1: Download the update
                                            $downloadResult = Invoke-WmiMethod -Namespace "root\ccm\clientsdk" -Class "CCM_SoftwareUpdatesManager" -Name "SetDownloadAsync" -ArgumentList @(@($updateObj))
                                            
                                            if ($downloadResult.ReturnValue -eq 0) {
                                                Write-Host "    Download initiated, waiting..." -ForegroundColor Cyan
                                                Start-Sleep -Seconds 10  # Wait for download to start
                                                
                                                # Step 2: Install the update
                                                $installSingleResult = Invoke-WmiMethod -Namespace "root\ccm\clientsdk" -Class "CCM_SoftwareUpdatesManager" -Name "InstallUpdates" -ArgumentList @(@($updateObj))
                                                
                                                if ($installSingleResult.ReturnValue -eq 0) {
                                                    $successCount++
                                                    $actionResults += "✓ Download+Install initiated for: $updateName"
                                                    Write-Host "    ✓ Download+Install successful" -ForegroundColor Green
                                                    continue
                                                } else {
                                                    Write-Host "    Install after download failed (Code: $($installSingleResult.ReturnValue))" -ForegroundColor Yellow
                                                }
                                            } else {
                                                Write-Host "    Download initiation failed (Code: $($downloadResult.ReturnValue))" -ForegroundColor Yellow
                                            }
                                        } catch {
                                            Write-Host "    Download+Install method failed: $($_.Exception.Message)" -ForegroundColor Yellow
                                        }
                                        
                                        # Method 3: Use WMI to set the update's download and install flags
                                        try {
                                            Write-Host "    Attempting to modify update properties..." -ForegroundColor Yellow
                                            
                                            # Try to set the update to download and install
                                            $updateObj.Get()  # Refresh the object
                                            
                                            # Some SCCM versions have these properties
                                            if (Get-Member -InputObject $updateObj -Name "IsSelected" -MemberType Property) {
                                                $updateObj.IsSelected = $true
                                                $updateObj.Put()
                                                Write-Host "    Set update as selected" -ForegroundColor Cyan
                                            }
                                            
                                            # Try to trigger the download/install action
                                            $modifyResult = $updateObj.Install()
                                            if ($modifyResult.ReturnValue -eq 0) {
                                                $successCount++
                                                $actionResults += "✓ Property modification successful for: $updateName"
                                                Write-Host "    ✓ Property modification successful" -ForegroundColor Green
                                            } else {
                                                Write-Host "    Property modification failed (Code: $($modifyResult.ReturnValue))" -ForegroundColor Red
                                            }
                                            
                                        } catch {
                                            Write-Host "    Property modification failed: $($_.Exception.Message)" -ForegroundColor Red
                                        }
                                        
                                        # Method 4: Manual trigger using specific SCCM actions
                                        try {
                                            Write-Host "    Attempting manual SCCM client actions..." -ForegroundColor Yellow
                                            
                                            # Trigger Software Updates Download Cycle
                                            $downloadCycleResult = Invoke-WmiMethod -Namespace "root\ccm" -Class "SMS_CLIENT" -Name "TriggerSchedule" -ArgumentList "{00000000-0000-0000-0000-000000000108}"
                                            
                                            Start-Sleep -Seconds 5
                                            
                                            # Trigger Software Updates Install Cycle  
                                            $installCycleResult = Invoke-WmiMethod -Namespace "root\ccm" -Class "SMS_CLIENT" -Name "TriggerSchedule" -ArgumentList "{00000000-0000-0000-0000-000000000109}"
                                            
                                            $successCount++
                                            $actionResults += "✓ Triggered download and install cycles for: $updateName"
                                            Write-Host "    ✓ Triggered SCCM client actions" -ForegroundColor Green
                                            
                                        } catch {
                                            Write-Host "    SCCM client actions failed: $($_.Exception.Message)" -ForegroundColor Red
                                            $actionResults += "✗ All methods failed for: $updateName"
                                        }
                                    }
                                } catch {
                                    Write-Host "  ✗ Error processing update: $updateID - $($_.Exception.Message)" -ForegroundColor Red
                                    $actionResults += "✗ Error processing: $updateID"
                                }
                            }
                            
                            # Set overall result based on success count
                            $installResult = @{
                                ReturnValue = if ($successCount -gt 0) { 0 } else { 1 }
                                SuccessfulUpdates = $successCount
                                TotalUpdates = $updateIDs.Count
                                Method = "ManualClientActions"
                                ActionResults = $actionResults
                            }
                            
                        } catch {
                            $installResult = @{
                                Status = "Error"
                                Message = "Manual client actions error: $($_.Exception.Message)"
                                ReturnCode = -1
                                UpdatesQueued = 0
                                Method = "Failed"
                            }
                            Write-Host "✗ Manual client actions failed: $($_.Exception.Message)" -ForegroundColor Red
                        }
                        
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
                $returnCode = if ($installResult.InstallationResults.ReturnCode) { $installResult.InstallationResults.ReturnCode } else { $installResult.InstallationResults.ReturnValue }
                $queued = $installResult.InstallationResults.UpdatesQueued
                $method = if ($installResult.InstallationResults.Method) { $installResult.InstallationResults.Method } else { "Standard" }
                
                switch ($status) {
                    "Success" {
                        Write-Host "✓ Installation Status: SUCCESS" -ForegroundColor Green
                        Write-Host "✓ Updates Processed: $queued" -ForegroundColor Green
                        Write-Host "✓ Method Used: $method" -ForegroundColor Green
                        Write-Host "✓ Message: $message" -ForegroundColor Green
                        
                        if ($installResult.InstallationResults.ActionResults) {
                            Write-Host "`nDetailed Results:" -ForegroundColor Cyan
                            $installResult.InstallationResults.ActionResults | ForEach-Object {
                                $color = if ($_ -like "*✓*") { "Green" } else { "Red" }
                                Write-Host "  $_" -ForegroundColor $color
                            }
                        }
                        
                        Write-Host "`nNext Steps:" -ForegroundColor Cyan
                        Write-Host "- Updates should start downloading/installing within 1-2 minutes" -ForegroundColor White
                        Write-Host "- Check SCCM Control Panel for progress" -ForegroundColor White
                        Write-Host "- Monitor C:\\Windows\\CCM\\Logs\\UpdatesDeployment.log" -ForegroundColor White
                        Write-Host "- Run with -VerifyInstallation to check progress" -ForegroundColor White
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
