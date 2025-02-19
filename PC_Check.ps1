# PC_Check.ps1

# Function to get system information
function Get-SystemInfo {
    try {
        $sysInfo = Get-CimInstance -ClassName Win32_OperatingSystem
        $processor = Get-CimInstance -ClassName Win32_Processor
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        
        return @{
            OS = $sysInfo
            Processor = $processor
            ComputerSystem = $computerSystem
        }
    } catch {
        Write-Warning "Error getting system information: $_"
        return $null
    }
}

# Function to get CPU usage
function Get-CPUUsage {
    try {
        $cpuUsage = Get-Counter -Counter "\Processor(_Total)\% Processor Time" -SampleInterval 1 -MaxSamples 1
        return [math]::Round($cpuUsage.CounterSamples.CookedValue, 2)
    } catch {
        Write-Warning "Error getting CPU usage: $_"
        return "N/A"
    }
}

# Function to get memory usage
function Get-MemoryUsage {
    try {
        $memInfo = Get-CimInstance -ClassName Win32_OperatingSystem
        $totalMemoryGB = [math]::Round($memInfo.TotalVisibleMemorySize / 1MB, 2)
        $freeMemoryGB = [math]::Round($memInfo.FreePhysicalMemory / 1MB, 2)
        $usedMemoryGB = $totalMemoryGB - $freeMemoryGB
        $memoryUsage = [math]::Round(($usedMemoryGB / $totalMemoryGB) * 100, 2)
        
        return @{
            UsagePercent = $memoryUsage
            TotalGB = $totalMemoryGB
            UsedGB = $usedMemoryGB
            FreeGB = $freeMemoryGB
        }
    } catch {
        Write-Warning "Error getting memory usage: $_"
        return $null
    }
}

# Function to get disk usage
function Get-DiskUsage {
    $diskInfo = Get-PSDrive -PSProvider FileSystem
    $diskUsage = @()
    foreach ($disk in $diskInfo) {
        if ($disk.Used -ne $null -and $disk.Free -ne $null) {
            $total = $disk.Used + $disk.Free
            $usage = [PSCustomObject]@{
                Drive = $disk.Name
                Used = if ($total -gt 0) { [math]::round(($disk.Used / $total * 100), 2) } else { 0 }
            }
            $diskUsage += $usage
        }
    }
    return $diskUsage
}

# Function to check network connectivity
function Test-NetworkConnectivity {
    try {
        $pingGoogle = Test-Connection -ComputerName google.com -Count 1 -Quiet
        $pingMS = Test-Connection -ComputerName microsoft.com -Count 1 -Quiet
        return @{
            Google = $pingGoogle
            Microsoft = $pingMS
        }
    } catch {
        Write-Warning "Error testing network connectivity: $_"
        return $null
    }
}

# Function to check for Windows updates
function Check-WindowsUpdates {
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $searchResult = $updateSearcher.Search("IsInstalled=0")
    return $searchResult.Updates.Count
}

# Function to check audio devices
function Test-AudioDevices {
    try {
        $audioDevices = Get-CimInstance -ClassName Win32_SoundDevice
        $results = @()
        foreach ($device in $audioDevices) {
            # StatusInfo values from CIM_LogicalDevice official documentation
            $statusInfo = switch ($device.StatusInfo) {
                1 { @{Status = "Other"; Message = "Other/Unknown status"} }
                2 { @{Status = "Unknown"; Message = "Status cannot be determined"} }
                3 { @{Status = "OK"; Message = "Device is running at full power"} }
                4 { @{Status = "Warning"; Message = "Device requires attention"} }
                5 { @{Status = "In Test"; Message = "Device is in testing mode"} }
                6 { @{Status = "Not Applicable"; Message = "Status not applicable for this device"} }
                7 { @{Status = "Power Off"; Message = "Device is powered off"} }
                8 { @{Status = "Off Line"; Message = "Device is offline"} }
                9 { @{Status = "Off Duty"; Message = "Device is off duty"} }
                10 { @{Status = "Degraded"; Message = "Device is running in degraded state"} }
                11 { @{Status = "Not Installed"; Message = "Device is not installed"} }
                12 { @{Status = "Install Error"; Message = "Installation error detected"} }
                13 { @{Status = "Power Save"; Message = "Device is in power save mode"} }
                default { @{Status = "Unknown"; Message = "Status code not recognized"} }
            }

            # Additional device state check from ConfigManagerErrorCode
            $state = switch ($device.ConfigManagerErrorCode) {
                0 { "Device is working properly" }
                1 { "Device is not configured correctly" }
                2 { "Windows cannot load the driver" }
                3 { "Driver is corrupted" }
                4 { "Device is not working properly" }
                5 { "Device requires a restart" }
                6 { "Device has a resource conflict" }
                7 { "Driver installation is required" }
                8 { "Resource settings are invalid" }
                9 { "Device is not responding" }
                10 { "Device cannot start" }
                default { $null }
            }

            if ($state) {
                $statusInfo.Message += ". $state"
            }

            $results += [PSCustomObject]@{
                Name = $device.Name
                Status = $statusInfo.Status
                Description = $device.Description
                StatusMessage = $statusInfo.Message
            }
        }
        return $results
    } catch {
        Write-Warning "Error checking audio devices: $_"
        return $null
    }
}

# Function to get detailed network adapter information
function Get-NetworkAdapterInfo {
    try {
        $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
        $results = @()
        foreach ($adapter in $adapters) {
            $ipConfig = Get-NetIPAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4
            # Fix: Convert LinkSpeed string to numeric value
            $speedString = $adapter.LinkSpeed
            $speed = if ($speedString -match '(\d+\.?\d*)\s*\w+') {
                [double]$matches[1]
            } else {
                0
            }
            
            $results += [PSCustomObject]@{
                Name = $adapter.Name
                Description = $adapter.InterfaceDescription
                Status = $adapter.Status
                Speed = $speed  # Already in Mbps
                IPAddress = $ipConfig.IPAddress
                MacAddress = $adapter.MacAddress
            }
        }
        return $results
    } catch {
        Write-Warning "Error getting network adapter information: $_"
        return $null
    }
}

# Function to check system services
function Test-CriticalServices {
    try {
        $criticalServices = @(
            "wuauserv",      # Windows Update
            "AudioSrv",      # Windows Audio
            "BITS",          # Background Intelligent Transfer Service
            "DPS",           # Diagnostic Policy Service
            "Dnscache",      # DNS Client
            "nsi"            # Network Store Interface Service
        )
        
        $results = @()
        foreach ($service in $criticalServices) {
            $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
            if ($svc) {
                $results += [PSCustomObject]@{
                    Name = $svc.DisplayName
                    Status = $svc.Status
                    StartType = $svc.StartType
                }
            }
        }
        return $results
    } catch {
        Write-Warning "Error checking critical services: $_"
        return $null
    }
}

# Function to generate HTML report
function Generate-HTMLReport {
    $sysInfo = Get-SystemInfo
    $cpuUsage = Get-CPUUsage
    $memoryUsage = Get-MemoryUsage
    $diskUsage = Get-DiskUsage
    $networkConnectivity = Test-NetworkConnectivity
    $windowsUpdates = Check-WindowsUpdates
    $audioDevices = Test-AudioDevices
    $networkAdapters = Get-NetworkAdapterInfo
    $criticalServices = Test-CriticalServices

    $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $computerName = $env:COMPUTERNAME

    $html = @"
    <html>
    <head>
        <title>PC Check Report - $computerName</title>
        <style>
            body { 
                font-family: Arial, sans-serif;
                margin: 20px;
                background-color: #f5f5f5;
            }
            .container {
                max-width: 1200px;
                margin: 0 auto;
                background-color: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            table { 
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
            }
            th, td { 
                border: 1px solid #ddd;
                padding: 12px;
                text-align: left;
            }
            th { 
                background-color: #4CAF50;
                color: white;
            }
            tr:nth-child(even) { 
                background-color: #f9f9f9;
            }
            .section {
                margin-bottom: 30px;
            }
            .status-good {
                color: #4CAF50;
            }
            .status-bad {
                color: #f44336;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>PC Check Report - $computerName</h1>
            <p>Report generated on: $date</p>

            <div class="section">
                <h2>System Information</h2>
                <table>
                    <tr><th>Property</th><th>Value</th></tr>
                    <tr><td>OS</td><td>$($sysInfo.OS.Caption)</td></tr>
                    <tr><td>Version</td><td>$($sysInfo.OS.Version)</td></tr>
                    <tr><td>Processor</td><td>$($sysInfo.Processor.Name)</td></tr>
                    <tr><td>Total Memory</td><td>$($memoryUsage.TotalGB) GB</td></tr>
                    <tr><td>Computer Name</td><td>$($sysInfo.ComputerSystem.Name)</td></tr>
                </table>
            </div>

            <div class="section">
                <h2>System Performance</h2>
                <table>
                    <tr><th>Metric</th><th>Value</th></tr>
                    <tr><td>CPU Usage</td><td>$cpuUsage%</td></tr>
                    <tr><td>Memory Usage</td><td>$($memoryUsage.UsagePercent)% ($($memoryUsage.UsedGB)GB used of $($memoryUsage.TotalGB)GB)</td></tr>
                </table>
            </div>

            <div class="section">
                <h2>Storage</h2>
                <table>
                    <tr><th>Drive</th><th>Usage (%)</th></tr>
                    $($diskUsage | ForEach-Object { "<tr><td>$($_.Drive)</td><td>$($_.Used)%</td></tr>" })
                </table>
            </div>

            <div class="section">
                <h2>Network Connectivity</h2>
                <table>
                    <tr><th>Service</th><th>Status</th></tr>
                    <tr>
                        <td>Google.com</td>
                        <td class="$(if($networkConnectivity.Google){'status-good'}else{'status-bad'})">
                            $(if($networkConnectivity.Google){'Connected'}else{'Not Connected'})
                        </td>
                    </tr>
                    <tr>
                        <td>Microsoft.com</td>
                        <td class="$(if($networkConnectivity.Microsoft){'status-good'}else{'status-bad'})">
                            $(if($networkConnectivity.Microsoft){'Connected'}else{'Not Connected'})
                        </td>
                    </tr>
                </table>
            </div>

            <div class="section">
                <h2>Windows Updates</h2>
                <p>Pending Updates: <span class="$(if($windowsUpdates -eq 0){'status-good'}else{'status-bad'})">$windowsUpdates</span></p>
            </div>

            <div class="section">
                <h2>Audio Devices</h2>
                <table>
                    <tr><th>Device Name</th><th>Status</th><th>Description</th><th>Status Details</th></tr>
                    $($audioDevices | ForEach-Object { 
                        "<tr><td>$($_.Name)</td><td class='$(if($_.Status -eq 'OK'){'status-good'}else{'status-bad'})'>$($_.Status)</td><td>$($_.Description)</td><td>$($_.StatusMessage)</td></tr>" 
                    })
                </table>
            </div>

            <div class="section">
                <h2>Network Adapters</h2>
                <table>
                    <tr><th>Name</th><th>Status</th><th>Speed (Mbps)</th><th>IP Address</th><th>MAC Address</th></tr>
                    $($networkAdapters | ForEach-Object { 
                        "<tr><td>$($_.Name)</td><td>$($_.Status)</td><td>$($_.Speed)</td><td>$($_.IPAddress)</td><td>$($_.MacAddress)</td></tr>" 
                    })
                </table>
            </div>

            <div class="section">
                <h2>Critical Services</h2>
                <table>
                    <tr><th>Service Name</th><th>Status</th><th>Start Type</th></tr>
                    $($criticalServices | ForEach-Object { 
                        "<tr><td>$($_.Name)</td><td class='$(if($_.Status -eq 'Running'){'status-good'}else{'status-bad'})'>$($_.Status)</td><td>$($_.StartType)</td></tr>" 
                    })
                </table>
            </div>
        </div>
    </body>
    </html>
"@

    try {
        $html | Out-File -FilePath "PC_Check_Report.html" -Encoding utf8
        Write-Host "Report generated successfully at PC_Check_Report.html"
    } catch {
        Write-Error "Failed to generate report: $_"
    }
}

# Run the report generation
Generate-HTMLReport