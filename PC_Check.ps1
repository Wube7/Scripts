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

# Function to generate HTML report
function Generate-HTMLReport {
    $sysInfo = Get-SystemInfo
    $cpuUsage = Get-CPUUsage
    $memoryUsage = Get-MemoryUsage
    $diskUsage = Get-DiskUsage
    $networkConnectivity = Test-NetworkConnectivity
    $windowsUpdates = Check-WindowsUpdates

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