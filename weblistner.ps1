# PowerShell Webhook Listener for SCOM Alerts
# Listens on port 8080 for JSON webhook data

# Initialize logging
$LogFile = "C:\Scripts\weblistener.txt"

# Function to write log entries
function Write-Log {
    param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] $Message"
    Add-Content -Path $LogFile -Value $LogEntry
    Write-Host $LogEntry
}

# No collection needed - processing each webhook immediately in background job

# Initialize HTTP Listener
$HttpListener = New-Object System.Net.HttpListener
$HttpListener.Prefixes.Add("http://+:8080/")

try {
    # Start the HTTP Listener
    $HttpListener.Start()
    Write-Log "HTTP Listener started successfully on port 8080"
    
    # Main listening loop
    while ($HttpListener.IsListening) {
        try {
            Write-Log "Waiting for incoming webhook request..."
            
            # Get the incoming request context
            $Context = $HttpListener.GetContext()
            $Request = $Context.Request
            $Response = $Context.Response
            
            Write-Log "Received request from $($Request.RemoteEndPoint)"
            
            # Read the request body
            $StreamReader = New-Object System.IO.StreamReader($Request.InputStream)
            $RequestBody = $StreamReader.ReadToEnd()
            $StreamReader.Close()
            
            Write-Log "Request body received: $RequestBody"
            
            # Parse JSON payload
            if ($RequestBody) {
                try {
                    $JsonData = $RequestBody | ConvertFrom-Json
                    
                    # Start background job to process SCOM alert immediately
                    Write-Log "Processing webhook data - AlertId: $($JsonData.AlertId), Acknowledged: $($JsonData.Acknowledged)"
                    
                    # Start background job to process SCOM alert
                    $JobScript = {
                        param($AlertId, $AcknowledgedState, $LogFile)
                        
                        # Function to write log entries from job
                        function Write-JobLog {
                            param([string]$Message)
                            $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            $LogEntry = "[$Timestamp] [JOB] $Message"
                            Add-Content -Path $LogFile -Value $LogEntry
                        }
                        
                        try {
                            Write-JobLog "Processing alert ID: $AlertId with acknowledged state: $AcknowledgedState"
                            
                            # Get SCOM alert
                            $Alert = Get-SCOMAlert -Id $AlertId
                            
                            if ($Alert) {
                                Write-JobLog "Found SCOM alert: $($Alert.Name)"
                                
                                # Compare acknowledged states
                                $CurrentAckState = $Alert.IsAcknowledged
                                Write-JobLog "Current acknowledged state: $CurrentAckState, Received state: $AcknowledgedState"
                                
                                if ($CurrentAckState -ne $AcknowledgedState) {
                                    Write-JobLog "Acknowledged state mismatch detected for alert $AlertId"
                                    # Add your processing logic here for state mismatch
                                } else {
                                    Write-JobLog "Acknowledged states match for alert $AlertId"
                                }
                            } else {
                                Write-JobLog "SCOM alert not found for ID: $AlertId"
                            }
                        }
                        catch {
                            Write-JobLog "Error processing alert $AlertId : $($_.Exception.Message)"
                        }
                    }
                    
                    # Start the background job
                    Start-Job -ScriptBlock $JobScript -ArgumentList $JsonData.AlertId, $JsonData.Acknowledged, $LogFile | Out-Null
                    
                    # Send success response
                    $Response.StatusCode = 200
                    $ResponseString = "OK"
                    
                } catch {
                    Write-Log "Error parsing JSON: $($_.Exception.Message)"
                    
                    # Send error response
                    $Response.StatusCode = 400
                    $ResponseString = "Bad Request - Invalid JSON"
                }
            } else {
                Write-Log "Empty request body received"
                
                # Send error response
                $Response.StatusCode = 400
                $ResponseString = "Bad Request - Empty Body"
            }
            
            # Write response and close
            $Buffer = [System.Text.Encoding]::UTF8.GetBytes($ResponseString)
            $Response.ContentLength64 = $Buffer.Length
            $Response.OutputStream.Write($Buffer, 0, $Buffer.Length)
            $Response.Close()
            
            Write-Log "Response sent: $($Response.StatusCode) - $ResponseString"
            
        } catch {
            Write-Log "Error processing request: $($_.Exception.Message)"
            
            # Try to send error response if possible
            try {
                if ($Response -and !$Response.OutputStream.CanWrite -eq $false) {
                    $Response.StatusCode = 500
                    $ErrorResponse = "Internal Server Error"
                    $Buffer = [System.Text.Encoding]::UTF8.GetBytes($ErrorResponse)
                    $Response.ContentLength64 = $Buffer.Length
                    $Response.OutputStream.Write($Buffer, 0, $Buffer.Length)
                    $Response.Close()
                }
            } catch {
                Write-Log "Could not send error response: $($_.Exception.Message)"
            }
        }
        
        # Clean up completed jobs periodically
        Get-Job | Where-Object { $_.State -eq 'Completed' } | Remove-Job
    }
}
catch {
    Write-Log "Fatal error in HTTP Listener: $($_.Exception.Message)"
}
finally {
    # Cleanup
    if ($HttpListener.IsListening) {
        $HttpListener.Stop()
        Write-Log "HTTP Listener stopped"
    }
    
    # Clean up any remaining jobs
    Get-Job | Remove-Job -Force
    Write-Log "Webhook listener script ended"
}
