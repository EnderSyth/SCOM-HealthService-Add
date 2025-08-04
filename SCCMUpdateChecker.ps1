#Requires -Version 5.1
#Requires -Modules PSExcel

<#
.SYNOPSIS
    Checks SCCM updates across multiple servers from Excel inventory
.DESCRIPTION
    Reads server list from Excel file, connects via PowerShell remoting,
    checks SCCM update status, and provides comprehensive summary
.PARAMETER ExcelPath
    Path to Excel file containing server inventory
.PARAMETER WorksheetName
    Name of worksheet containing server data (default: first sheet)
.PARAMETER HostnameColumn
    Column name containing hostnames (default: "Hostname")
.PARAMETER MaxConcurrentJobs
    Maximum number of concurrent remote jobs (default: 10)
.PARAMETER TimeoutMinutes
    Timeout for remote connections in minutes (default: 5)
#>

param(
    [string]$ExcelPath = $null,
    [string]$WorksheetName = $null,
    [string]$HostnameColumn = "Hostname",
    [int]$MaxConcurrentJobs = 10,
    [int]$TimeoutMinutes = 5
)

# Function to get SCCM updates from remote server
function Get-SCCMUpdates {
    param(
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Host "Connecting to $ComputerName..." -ForegroundColor Yellow
        
        $session = New-PSSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
        
        $result = Invoke-Command -Session $session -ScriptBlock {
            try {
                # Get SCCM software updates
                $updates = Get-WmiObject -Namespace "root\ccm\clientsdk" -Class "ccm_softwareupdate" -ErrorAction Stop
                
                # Get reboot status
                $rebootStatus = Get-WmiObject -Namespace "root\ccm\clientsdk" -Class "ccm_clientutilities" -ErrorAction SilentlyContinue
                $pendingReboot = $false
                
                # Check for pending reboot using multiple methods
                try {
                    $pendingReboot = (Get-WmiObject -Class "Win32_Registry" -ErrorAction SilentlyContinue) -and
                                   (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue) -or
                                   (Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue) -or
                                   (Get-WmiObject -Query "SELECT * FROM Win32_Registry WHERE Hive=2147483650 AND KeyPath='SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\WindowsUpdate\\Auto Update\\RebootRequired'" -ErrorAction SilentlyContinue)
                    
                    # Alternative reboot check
                    if (-not $pendingReboot) {
                        $pendingReboot = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue) -ne $null
                    }
                } catch {
                    $pendingReboot = $false
                }
                
                # Categorize updates
                $pendingUpdates = @($updates | Where-Object { $_.EvaluationState -eq 0 -or $_.EvaluationState -eq 1 })
                $downloadingUpdates = @($updates | Where-Object { $_.EvaluationState -eq 2 -or $_.EvaluationState -eq 3 })
                $installingUpdates = @($updates | Where-Object { $_.EvaluationState -eq 6 -or $_.EvaluationState -eq 7 })
                $installedUpdates = @($updates | Where-Object { $_.EvaluationState -eq 13 })
                $failedUpdates = @($updates | Where-Object { $_.EvaluationState -eq 11 -or $_.EvaluationState -eq 12 })
                $rebootRequiredUpdates = @($updates | Where-Object { $_.EvaluationState -eq 8 -or $_.EvaluationState -eq 9 })
                
                # Get last scan time
                try {
                    $scanHistory = Get-WmiObject -Namespace "root\ccm" -Class "ccm_scanagent" -ErrorAction SilentlyContinue
                    $lastScanTime = if ($scanHistory) { $scanHistory.LastScanTime } else { "Unknown" }
                } catch {
                    $lastScanTime = "Unknown"
                }
                
                return @{
                    ComputerName = $env:COMPUTERNAME
                    Status = "Success"
                    TotalUpdates = $updates.Count
                    PendingUpdates = $pendingUpdates.Count
                    DownloadingUpdates = $downloadingUpdates.Count
                    InstallingUpdates = $installingUpdates.Count
                    InstalledUpdates = $installedUpdates.Count
                    FailedUpdates = $failedUpdates.Count
                    RebootRequiredUpdates = $rebootRequiredUpdates.Count
                    PendingReboot = $pendingReboot
                    LastScanTime = $lastScanTime
                    PendingUpdatesList = ($pendingUpdates | Select-Object -ExpandProperty Name) -join "; "
                    FailedUpdatesList = ($failedUpdates | Select-Object -ExpandProperty Name) -join "; "
                    Error = $null
                }
                
            } catch {
                return @{
                    ComputerName = $env:COMPUTERNAME
                    Status = "SCCM_Error"
                    Error = $_.Exception.Message
                    TotalUpdates = 0
                    PendingUpdates = 0
                    DownloadingUpdates = 0
                    InstallingUpdates = 0
                    InstalledUpdates = 0
                    FailedUpdates = 0
                    RebootRequiredUpdates = 0
                    PendingReboot = $false
                    LastScanTime = "Unknown"
                    PendingUpdatesList = ""
                    FailedUpdatesList = ""
                }
            }
        }
        
        Remove-PSSession $session
        return $result
        
    } catch {
        return @{
            ComputerName = $ComputerName
            Status = "Connection_Failed"
            Error = $_.Exception.Message
            TotalUpdates = 0
            PendingUpdates = 0
            DownloadingUpdates = 0
            InstallingUpdates = 0
            InstalledUpdates = 0
            FailedUpdates = 0
            RebootRequiredUpdates = 0
            PendingReboot = $false
            LastScanTime = "Unknown"
            PendingUpdatesList = ""
            FailedUpdatesList = ""
        }
    }
}

# Main script execution
try {
    Write-Host "SCCM Update Checker Starting..." -ForegroundColor Green
    Write-Host "=================================" -ForegroundColor Green
    
    # Prompt for Excel file path if not provided
    if (-not $ExcelPath) {
        do {
            $ExcelPath = Read-Host "Enter the path to your Excel file"
            if (-not $ExcelPath) {
                Write-Host "Excel path is required!" -ForegroundColor Red
            } elseif (-not (Test-Path $ExcelPath)) {
                Write-Host "File not found: $ExcelPath" -ForegroundColor Red
                $ExcelPath = $null
            }
        } while (-not $ExcelPath)
    }
    
    # Prompt for smart card credentials
    Write-Host "`nPlease provide your admin smart card credentials:" -ForegroundColor Cyan
    try {
        $credential = Get-Credential -Message "Enter your admin smart card credentials"
        if (-not $credential) {
            Write-Error "Credentials are required to connect to remote servers."
            exit 1
        }
    } catch {
        Write-Error "Failed to get credentials: $($_.Exception.Message)"
        exit 1
    }
    
    # Check if PSExcel module is available
    if (-not (Get-Module -ListAvailable -Name PSExcel)) {
        Write-Error "PSExcel module is required. Install it with: Install-Module PSExcel -Force"
        exit 1
    }
    
    # Import Excel file using PSExcel
    Write-Host "Reading Excel file: $ExcelPath" -ForegroundColor Cyan
    
    if (-not (Test-Path $ExcelPath)) {
        Write-Error "Excel file not found: $ExcelPath"
        exit 1
    }
    
    # Import the Excel data using PSExcel
    try {
        if ($WorksheetName) {
            $excelData = Import-XLSX -Path $ExcelPath -WorksheetName $WorksheetName
        } else {
            $excelData = Import-XLSX -Path $ExcelPath
        }
    } catch {
        Write-Error "Failed to read Excel file: $($_.Exception.Message)"
        Write-Host "Tip: Ensure Excel file is not open in another application" -ForegroundColor Yellow
        exit 1
    }
    
    # Extract hostnames
    $hostnames = $excelData | Where-Object { $_.$HostnameColumn -and $_.$HostnameColumn.Trim() -ne "" } | 
                Select-Object -ExpandProperty $HostnameColumn | 
                ForEach-Object { $_.Trim() }
    
    if (-not $hostnames) {
        Write-Error "No hostnames found in column '$HostnameColumn'"
        exit 1
    }
    
    Write-Host "Found $($hostnames.Count) servers to check" -ForegroundColor Cyan
    
    # Initialize results array
    $results = @()
    $jobQueue = @()
    $activeJobs = @()
    
    # Process servers with job management
    $totalServers = $hostnames.Count
    $processedServers = 0
    
    foreach ($hostname in $hostnames) {
        # Wait if we have too many active jobs
        while ($activeJobs.Count -ge $MaxConcurrentJobs) {
            $completedJobs = @()
            foreach ($job in $activeJobs) {
                if ($job.Job.State -eq 'Completed' -or $job.Job.State -eq 'Failed') {
                    $completedJobs += $job
                }
            }
            
            # Process completed jobs
            foreach ($completedJob in $completedJobs) {
                $result = Receive-Job -Job $completedJob.Job
                $results += $result
                Remove-Job -Job $completedJob.Job
                $activeJobs = $activeJobs | Where-Object { $_.Job.Id -ne $completedJob.Job.Id }
                $processedServers++
                
                Write-Progress -Activity "Checking SCCM Updates" -Status "Processed $processedServers of $totalServers servers" -PercentComplete (($processedServers / $totalServers) * 100)
            }
            
            Start-Sleep -Milliseconds 500
        }
        
        # Start new job
        $job = Start-Job -ScriptBlock ${function:Get-SCCMUpdates} -ArgumentList $hostname, $credential
        $activeJobs += @{
            Job = $job
            Hostname = $hostname
            StartTime = Get-Date
        }
    }
    
    # Wait for remaining jobs to complete
    while ($activeJobs.Count -gt 0) {
        $completedJobs = @()
        foreach ($job in $activeJobs) {
            $elapsed = (Get-Date) - $job.StartTime
            
            if ($job.Job.State -eq 'Completed' -or $job.Job.State -eq 'Failed' -or $elapsed.TotalMinutes -gt $TimeoutMinutes) {
                $completedJobs += $job
            }
        }
        
        foreach ($completedJob in $completedJobs) {
            if ($completedJob.Job.State -eq 'Running') {
                Stop-Job -Job $completedJob.Job
                $result = @{
                    ComputerName = $completedJob.Hostname
                    Status = "Timeout"
                    Error = "Connection timeout after $TimeoutMinutes minutes"
                    TotalUpdates = 0
                    PendingUpdates = 0
                    DownloadingUpdates = 0
                    InstallingUpdates = 0
                    InstalledUpdates = 0
                    FailedUpdates = 0
                    RebootRequiredUpdates = 0
                    PendingReboot = $false
                    LastScanTime = "Unknown"
                    PendingUpdatesList = ""
                    FailedUpdatesList = ""
                }
            } else {
                $result = Receive-Job -Job $completedJob.Job
            }
            
            $results += $result
            Remove-Job -Job $completedJob.Job
            $activeJobs = $activeJobs | Where-Object { $_.Job.Id -ne $completedJob.Job.Id }
            $processedServers++
            
            Write-Progress -Activity "Checking SCCM Updates" -Status "Processed $processedServers of $totalServers servers" -PercentComplete (($processedServers / $totalServers) * 100)
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    Write-Progress -Completed -Activity "Checking SCCM Updates"
    
    # Generate summary report
    Write-Host "`n`nSCCM UPDATE SUMMARY REPORT" -ForegroundColor Green
    Write-Host "===========================" -ForegroundColor Green
    Write-Host "Generated: $(Get-Date)" -ForegroundColor Gray
    Write-Host "Total Servers Checked: $($results.Count)" -ForegroundColor Gray
    
    # Categorize results
    $successfulServers = $results | Where-Object { $_.Status -eq "Success" }
    $connectionFailedServers = $results | Where-Object { $_.Status -eq "Connection_Failed" }
    $sccmErrorServers = $results | Where-Object { $_.Status -eq "SCCM_Error" }
    $timeoutServers = $results | Where-Object { $_.Status -eq "Timeout" }
    
    # Connection Status Summary
    Write-Host "`nCONNECTION STATUS:" -ForegroundColor Yellow
    Write-Host "  Successful Connections: $($successfulServers.Count)" -ForegroundColor Green
    Write-Host "  Connection Failed: $($connectionFailedServers.Count)" -ForegroundColor Red
    Write-Host "  SCCM Errors: $($sccmErrorServers.Count)" -ForegroundColor Red
    Write-Host "  Timeouts: $($timeoutServers.Count)" -ForegroundColor Red
    
    if ($successfulServers.Count -gt 0) {
        # Update Status Summary
        $serversWithPendingUpdates = $successfulServers | Where-Object { $_.PendingUpdates -gt 0 }
        $serversNeedingReboot = $successfulServers | Where-Object { $_.PendingReboot -eq $true -or $_.RebootRequiredUpdates -gt 0 }
        $serversWithFailedUpdates = $successfulServers | Where-Object { $_.FailedUpdates -gt 0 }
        $serversInstalling = $successfulServers | Where-Object { $_.InstallingUpdates -gt 0 }
        $serversDownloading = $successfulServers | Where-Object { $_.DownloadingUpdates -gt 0 }
        
        Write-Host "`nUPDATE STATUS:" -ForegroundColor Yellow
        Write-Host "  Servers with Pending Updates: $($serversWithPendingUpdates.Count)" -ForegroundColor Cyan
        Write-Host "  Servers Needing Reboot: $($serversNeedingReboot.Count)" -ForegroundColor Magenta
        Write-Host "  Servers with Failed Updates: $($serversWithFailedUpdates.Count)" -ForegroundColor Red
        Write-Host "  Servers Currently Installing: $($serversInstalling.Count)" -ForegroundColor Yellow
        Write-Host "  Servers Currently Downloading: $($serversDownloading.Count)" -ForegroundColor Yellow
        
        # Detailed Server Lists
        if ($serversWithPendingUpdates.Count -gt 0) {
            Write-Host "`nSERVERS WITH PENDING UPDATES:" -ForegroundColor Cyan
            $serversWithPendingUpdates | ForEach-Object {
                Write-Host "  $($_.ComputerName): $($_.PendingUpdates) pending" -ForegroundColor White
            }
        }
        
        if ($serversNeedingReboot.Count -gt 0) {
            Write-Host "`nSERVERS NEEDING REBOOT:" -ForegroundColor Magenta
            $serversNeedingReboot | ForEach-Object {
                $rebootReason = if ($_.PendingReboot) { "System reboot required" } else { "$($_.RebootRequiredUpdates) updates require reboot" }
                Write-Host "  $($_.ComputerName): $rebootReason" -ForegroundColor White
            }
        }
        
        if ($serversWithFailedUpdates.Count -gt 0) {
            Write-Host "`nSERVERS WITH FAILED UPDATES:" -ForegroundColor Red
            $serversWithFailedUpdates | ForEach-Object {
                Write-Host "  $($_.ComputerName): $($_.FailedUpdates) failed updates" -ForegroundColor White
                if ($_.FailedUpdatesList) {
                    Write-Host "    Failed: $($_.FailedUpdatesList)" -ForegroundColor Gray
                }
            }
        }
    }
    
    # Connection Failures
    if ($connectionFailedServers.Count -gt 0) {
        Write-Host "`nCONNECTION FAILURES:" -ForegroundColor Red
        $connectionFailedServers | ForEach-Object {
            Write-Host "  $($_.ComputerName): $($_.Error)" -ForegroundColor White
        }
    }
    
    if ($sccmErrorServers.Count -gt 0) {
        Write-Host "`nSCCM ERRORS:" -ForegroundColor Red
        $sccmErrorServers | ForEach-Object {
            Write-Host "  $($_.ComputerName): $($_.Error)" -ForegroundColor White
        }
    }
    
    if ($timeoutServers.Count -gt 0) {
        Write-Host "`nTIMEOUT SERVERS:" -ForegroundColor Red
        $timeoutServers | ForEach-Object {
            Write-Host "  $($_.ComputerName): Connection timeout" -ForegroundColor White
        }
    }
    
    # Export detailed results to CSV
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = "SCCM_Update_Report_$timestamp.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation
    
    Write-Host "`nDetailed report exported to: $csvPath" -ForegroundColor Green
    Write-Host "Script completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Error $_.ScriptStackTrace
}
