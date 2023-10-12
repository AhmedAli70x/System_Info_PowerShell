function ConvertTo-DateTime ($DateValue) {
    if ($null -eq $DateValue) {
        Write-Warning "Null date value detected."
        return $null
    }

    try {
        return [DateTime]$DateValue
    }
    catch {
        Write-Warning "Invalid DateTime format detected: $DateValue"
        return $null
    }
}


function SafeGet-CimInstance {
    param ($ClassName)

    try {
        return Get-CimInstance -ClassName $ClassName
    }
    catch {
        Write-Warning "Failed to query $ClassName. Error: $_"
        return $null
    }
}


function ConvertTo-MarkdownTable {
    param ($InputObject)

    $header = $InputObject[0].psobject.properties.name -join " | "
    $separator = ($InputObject[0].psobject.properties.name | ForEach-Object { "---" }) -join " | "

    $body = $InputObject | ForEach-Object {
        $props = $_.psobject.properties
      ($props.value -join " | ")
    }

    return @"
| $header |
| $separator |
$(($body | ForEach-Object { "| $_ |" }) -join "`n")
"@
}


function Get-ComputerMonitorInfo {
    ForEach ($Computer in $ComputerName) {
  
        $Mon_Attached_Computer = $Computer
    
        #Grabs the Monitor objects from WMI
        $Monitors = Get-WmiObject -Namespace "root\WMI" -Class "WMIMonitorID" -ComputerName $Computer -ErrorAction SilentlyContinue
    
        #Creates an empty array to hold the data
        $Monitor_Array = @()
    
    
        #Takes each monitor object found and runs the following code:
        ForEach ($Monitor in $Monitors) {
      
            # Grabs respective data and converts it from ASCII encoding and removes any trailing ASCII null values
            If ($Monitor.UserFriendlyName -and $Monitor.UserFriendlyName.Length -gt 0) {
                $Mon_Model = ([System.Text.Encoding]::ASCII.GetString($Monitor.UserFriendlyName)).Replace("$([char]0x0000)", "")
            }
            else {
                $Mon_Model = $null
            }

            If ($Monitor.SerialNumberID -and $Monitor.SerialNumberID.Length -gt 0) {
                $Mon_Serial_Number = ([System.Text.Encoding]::ASCII.GetString($Monitor.SerialNumberID)).Replace("$([char]0x0000)", "")
            }

            If ($Monitor.ManufacturerName -and $Monitor.ManufacturerName.Length -gt 0) {
                $Mon_Manufacturer = ([System.Text.Encoding]::ASCII.GetString($Monitor.ManufacturerName)).Replace("$([char]0x0000)", "")
            }
      
            #Filters out "non monitors". Place any of your own filters here. These two are all-in-one computers with built in displays. I don't need the info from these.
            If ($Mon_Model -like "*800 AIO*" -or $Mon_Model -like "*8300 AiO*") { Break }
      
            #Sets a friendly name based on the hash table above. If no entry found sets it to the original 3 character code
            $Mon_Manufacturer_Friendly = $ManufacturerHash.$Mon_Manufacturer
            If ($Mon_Manufacturer_Friendly -eq $null) {
                $Mon_Manufacturer_Friendly = $Mon_Manufacturer
            }
      
            #Creates a custom monitor object and fills it with 4 NoteProperty members and the respective data
            $Monitor_Obj = [PSCustomObject]@{
                Manufacturer     = $Mon_Manufacturer_Friendly
                Model            = $Mon_Model
                SerialNumber     = $Mon_Serial_Number
                AttachedComputer = $Mon_Attached_Computer
            }
      
            #Appends the object to the array
            $Monitor_Array += $Monitor_Obj

        } #End ForEach Monitor
  
        #Outputs the Array
        # $Monitor_Array
        #Outputs the Array as Markdown Table
        return ConvertTo-MarkdownTable -InputObject $Monitor_Array
    }
}


function Get-PhysicalDriveDetails {
    # Get all physical drives
    $disks = Get-WmiObject Win32_DiskDrive

    # Create a markdown table header
    $output = "| Disk Model | Disk Size (GB) |`n|-------------|------------------|`n"

    # Loop through each disk to fetch details and format the markdown table
    foreach ($disk in $disks) {
        # Convert bytes to GB and round it
        $sizeInGB = [Math]::Round($disk.Size / 1GB)

        $output += "| $($disk.Model) | $sizeInGB GB |`n"
    }

    return $output
}


function Get-LogicalDriveDetails {
    # Get Total and Free space
    $logicalDrives = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | 
    Select-Object Caption, VolumeName, VolumeSerialNumber, 
    @{Name = "SizeGB"; Expression = { [math]::Round($_.Size / 1GB) } }, 
    @{Name = "FreeSpaceGB"; Expression = { [math]::Round($_.FreeSpace / 1GB) } }, FileSystem

    # Create a markdown table header
    $output = "| Caption | Volume Name | Serial Number | Size (GB) | Free Space (GB) | File System |`n|---------|-------------|--------------|----------|----------------|--------------|`n"

    # Format the markdown table
    foreach ($drive in $logicalDrives) {
        $output += "| $($drive.Caption) | $($drive.VolumeName) | $($drive.VolumeSerialNumber) | $($drive.SizeGB) | $($drive.FreeSpaceGB) | $($drive.FileSystem) |`n"
    }

    return $output
}

# Data Extraction
$Username = $env:USERNAME
$ComputerName = $env:ComputerName
$SystemInfo = SafeGet-CimInstance Win32_ComputerSystem
$board = SafeGet-CimInstance Win32_BaseBoard
$BIOSInfo = SafeGet-CimInstance Win32_BIOS
$Processor = SafeGet-CimInstance Win32_Processor
$Memory = SafeGet-CimInstance Win32_PhysicalMemory
$OSInfo = SafeGet-CimInstance Win32_OperatingSystem
$Disks = SafeGet-CimInstance Win32_LogicalDisk
$GraphicsCards = SafeGet-CimInstance Win32_VideoController | Where-Object { $_.Name -notlike '*DisplayLink*' }
$GraphicsInfo = $GraphicsCards | Select-Object Name, @{Expression={[math]::Round($_.AdapterRAM/1GB, 2)};label="RAM (GB)"}

$Monitors = SafeGet-CimInstance Win32_DesktopMonitor
$Monitors = $Monitors | Where-Object { $_.ScreenHeight -and $_.ScreenWidth }
$NetworkAdapters = SafeGet-CimInstance Win32_NetworkAdapter | Where-Object {
    ($_.NetConnectionID) -and 
    (($_.AdapterType -eq 'Ethernet 802.3') -or ($_.AdapterType -eq 'Wireless')) -and
    ($_.Name -notmatch 'Hyper-V|VMware|VirtualBox|vEthernet')
}


$SoftwareLicensing = SafeGet-CimInstance SoftwareLicensingService

$LastBootUp = ConvertTo-DateTime $OSInfo.LastBootUpTime
$CurrentTime = ConvertTo-DateTime $OSInfo.LocalDateTime
$UpTime = $CurrentTime - $LastBootUp
$MemoryTotal = [math]::Round(($Memory | Measure-Object Capacity -Sum).Sum / 1GB, 2)
$MemoryUsed = [math]::Round(($OSInfo.TotalVisibleMemorySize - $OSInfo.FreePhysicalMemory) / 1MB, 2)
$MemoryFree = [math]::Round($OSInfo.FreePhysicalMemory / 1MB, 2)
$MemoryUsage = [math]::Round((($OSInfo.TotalVisibleMemorySize - $OSInfo.FreePhysicalMemory) / $OSInfo.TotalVisibleMemorySize) * 100, 2)


$DiskData = $Disks | ForEach-Object {
    @{
        'DeviceID'   = $_.DeviceID
        'TotalSpace' = [math]::Round($_.Size / 1GB, 2)
        'UsedSpace'  = [math]::Round(($_.Size - $_.FreeSpace) / 1GB, 2)
    }
}

$CacheSizeKB = $(($Processor.L2CacheSize + $Processor.L3CacheSize))
$CacheSizeMB = [math]::Round($CacheSizeKB / 1024, 2)

$totalVirtualKB = $OSInfo.TotalVirtualMemorySize
$totalPhysicalKB = $OSInfo.TotalVisibleMemorySize
$virtualOnlyKB = $totalVirtualKB - $totalPhysicalKB

$TotalMemory = [math]::Round($totalVirtualKB / 1MB, 2)
$virtualOnlyGB = [math]::Round($virtualOnlyKB / 1MB, 2)


$ManufacturerHash = @{ 
    "AAC" =	"AcerView";
    "ACR" = "Acer";
    "AOC" = "AOC";
    "AIC" = "AG Neovo";
    "APP" = "Apple Computer";
    "AST" = "AST Research";
    "AUO" = "Asus";
    "BNQ" = "BenQ";
    "CMO" = "Acer";
    "CPL" = "Compal";
    "CPQ" = "Compaq";
    "CPT" = "Chunghwa Pciture Tubes, Ltd.";
    "CTX" = "CTX";
    "DEC" = "DEC";
    "DEL" = "Dell";
    "DPC" = "Delta";
    "DWE" = "Daewoo";
    "EIZ" = "EIZO";
    "ELS" = "ELSA";
    "ENC" = "EIZO";
    "EPI" = "Envision";
    "FCM" = "Funai";
    "FUJ" = "Fujitsu";
    "FUS" = "Fujitsu-Siemens";
    "GSM" = "LG Electronics";
    "GWY" = "Gateway 2000";
    "HEI" = "Hyundai";
    "HIT" = "Hyundai";
    "HSL" = "Hansol";
    "HTC" = "Hitachi/Nissei";
    "HWP" = "HP";
    "IBM" = "IBM";
    "ICL" = "Fujitsu ICL";
    "IVM" = "Iiyama";
    "KDS" = "Korea Data Systems";
    "LEN" = "Lenovo";
    "LGD" = "Asus";
    "LPL" = "Fujitsu";
    "MAX" = "Belinea"; 
    "MEI" = "Panasonic";
    "MEL" = "Mitsubishi Electronics";
    "MS_" = "Panasonic";
    "NAN" = "Nanao";
    "NEC" = "NEC";
    "NOK" = "Nokia Data";
    "NVD" = "Fujitsu";
    "OPT" = "Optoma";
    "PHL" = "Philips";
    "REL" = "Relisys";
    "SAN" = "Samsung";
    "SAM" = "Samsung";
    "SBI" = "Smarttech";
    "SGI" = "SGI";
    "SNY" = "Sony";
    "SRC" = "Shamrock";
    "SUN" = "Sun Microsystems";
    "SEC" = "Hewlett-Packard";
    "TAT" = "Tatung";
    "TOS" = "Toshiba";
    "TSB" = "Toshiba";
    "VSC" = "ViewSonic";
    "ZCM" = "Zenith";
    "UNK" = "Unknown";
    "_YV" = "Fujitsu";
}

$domain = $SystemInfo.Domain
$hostname = $SystemInfo.Name
$manufacturer = $SystemInfo.Manufacturer
$model = $SystemInfo.Model
$boardSerialNumber = $board.SerialNumber
$BIOSVersion = $BIOSInfo.Version
$BIOSReleaseDate = if ($BIOSInfo.ReleaseDate -ne $null) { $BIOSInfo.ReleaseDate.ToString("yyyy-MM-dd") } else { "N/A" }

$OSName = $OSInfo.Caption
$OSVersion = $OSInfo.Version
$OSBuild = $OSInfo.BuildNumber
$OS_product_key = $SoftwareLicensing.OA3xOriginalProductKey
$installation_date = if ($OSInfo.InstallDate -ne $null) { $OSInfo.InstallDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
$up_time = "$($UpTime.Days) days, $($UpTime.Hours) hours, $($UpTime.Minutes) minutes"

$processor_name = $Processor.Name
$total_cores = $Processor.NumberOfCores
$total_threads = $Processor.ThreadCount
$clock_speed = "$($Processor.MaxClockSpeed) MHz"


# Create the header
$memory_table = "Manufacturer | SerialNumber | BankLabel | DeviceLocator | Size (GB) | ClockSpeed (MHz)`n"
$memory_table += "| --- | --- | --- | --- | --- | --- |`n"

# Append each row
$Memory | ForEach-Object {
    $size = [math]::Round($_.Capacity/1GB, 2)
    $clockSpeed = $_.ConfiguredClockSpeed
    $memory_table += "| $($_.Manufacturer) | $($_.SerialNumber) | $($_.BankLabel) | $($_.DeviceLocator) | $size | $clockSpeed |`n"
}


# Start markdown table header
$GraphicTable = "| Name | RAM (GB) | `n| --- | --- |`n"

# Add rows to the markdown table
$GraphicsInfo | ForEach-Object {
    $GraphicTable += "| $($_.Name) | $($_.'RAM (GB)') GB |`n"
}


# Markdown Formation
$Output = @"
# System Information Report

## System Details:

| Property       | Value                      |
| -------------- | -------------------------- |
| **User**          | $Username        |
| **Domain**   | $domain        |
| **Hostname**   | $hostname        |
| **Manufacturer** | $manufacturer |
| **Model**       | $model       |
| **Serial Number** | $boardSerialNumber |
| **BIOS Version** | $BIOSVersion      |
| **BIOS Release Date** | $BIOSReleaseDate       |

## Operating System:

| Property         | Value                                                           |
| ---------------- | --------------------------------------------------------------- |
| **OS Name**      | $OSName                                              |
| **OS Version**   | $OSVersion                                              |
| **OS Build**     | $OSBuild                                          |
| **Installation Date** | $installation_date |
| **System Up Time** | $up_time    |
| **OS Serial Number** | $OS_product_key                  |

## Processor:

| Property       | Value                     |
| -------------- | ------------------------- |
| **Processor** | $processor_name |
| **Total Cores** | $total_cores |
| **Total Threads** | $total_threads  |
| **Clock Speed** | $clock_speed |
| **Cache Size**   | $CacheSizeMB MB       |

## Memory:

| Property       | Value                 |
| -------------- | --------------------- |
| **Total Memory** | $TotalMemory GB       |
| **Virtual Memory** | $virtualOnlyGB GB   |
| **Physical Memory** | $MemoryTotal GB    |
| **Used Memory**  | $MemoryUsed GB        |
| **Free Memory**  | $MemoryFree GB        |
| **Memory Usage** | $MemoryUsage%         |

$memory_table

## Disk Storage:
"@


# Call the functions to get markdown tables
$physicalDriveDetails = Get-PhysicalDriveDetails
$logicalDriveDetails = Get-LogicalDriveDetails


# Combine markdown results and save to a file
$Output += "`n## Physical Drive Details`n$physicalDriveDetails`n## Logical Drive Details`n$logicalDriveDetails`n"


$Output += @"
## Graphics Card:
$GraphicTable

## Monitors:
| Monitor | Manufacturer | Height | Width |
| --- | --- | --- | --- |`n
"@

$Output += $Monitors | ForEach-Object { "| $($_.Name) | $($_.MonitorManufacturer) | $($_.ScreenHeight) | $($_.ScreenWidth) |`n" }

$connectedMonitors = Get-ComputerMonitorInfo
$Output += "`n## Connected Monitors`n$connectedMonitors`n"

$Output += @"
`n## Network:
| AdapterType | Manufacturer | ProductName | MACAddress |
| --- | --- | --- | --- |`n
"@

$Output += $NetworkAdapters | ForEach-Object {
    "| $($_.NetConnectionID) | $($_.Manufacturer) | $($_.ProductName) | $($_.MACAddress) |`n"
}


# Save to File (Markdown)
$OutputFileNameMD = "C:\" + $SystemInfo.Name + "_SystemInfo.md"
$Output | Out-File -FilePath $OutputFileNameMD -Encoding UTF8
Write-Output "Markdown Report saved to: $OutputFileNameMD"



# CSV Data Formation
$csvData = @(
    [PSCustomObject]@{
        'User'              = $Username
        'Hostname'              = $hostname
        'Domain'                = $domain
        'Manufacturer'          = $manufacturer
        'Model'                 = $model
        'Serial Number'         = $boardSerialNumber
        'BIOS Version'          = $BIOSVersion
        'OS Name'               = $OSName
        'OS Version'            = $OSVersion
        'OS Build'              = $OSBuild
        'Installation Date'     = $installation_date
        'System Up Time'        = $up_time
        'OS Serial Number'      = $OS_product_key
        'Total Cores'           = $total_cores
        'Total Threads'         = $total_threads
        'Clock Speed'           = $clock_speed
        'Cache Size'            = "$CacheSizeMB MB"
        'Total Memory'          = "$TotalMemory GB"
        'Virtual Memory'        = "$virtualOnlyGB GB"
        'Physical Memory'       = "$MemoryTotal GB"
        'Used Memory'           = "$MemoryUsed GB"
        'Free Memory'           = "$MemoryFree GB"
        'Memory Usage'          = "$MemoryUsage%"
    }
)


# Save to File (CSV)
$OutputFileNameCSV = "C:\" + $SystemInfo.Name + "_SystemInfo.csv"
$csvData | Export-Csv -Path $OutputFileNameCSV -NoTypeInformation

# Output Result File Locations
Write-Output "CSV Report saved to: $OutputFileNameCSV"


