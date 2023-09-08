# Improved System Information Gathering Script

# Functions
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


# Data Extraction
$SystemAll = @{
}

function System_Details {
    $SystemDetail = @{
    }

    $Username = $env:USERNAME
    $ComputerName = $env:ComputerName
    $SystemInfo = SafeGet-CimInstance Win32_ComputerSystem
    $board = SafeGet-CimInstance Win32_BaseBoard
    $BIOSInfo = SafeGet-CimInstance Win32_BIOS
    $board = SafeGet-CimInstance Win32_BaseBoard
    $BIOSVersion = $BIOSInfo.Version
    $BIOSReleaseDate = if ($BIOSInfo.ReleaseDate -ne $null) { $BIOSInfo.ReleaseDate.ToString("yyyy-MM-dd") } else { "N/A" }

    $SystemDetail["User_Name"] = $env:USERNAME
    $SystemDetail["Computer_Name"] = $env:ComputerName
    $SystemDetail["Domain"] = $SystemInfo.Domain
    $SystemDetail["Hostname"] = $SystemInfo.Name 
    $SystemDetail["Manufacturer"] = $SystemInfo.Name
    $SystemDetail["Model"] = $SystemInfo.Model
    $SystemDetail["Serial_Number"] = $board.SerialNumber
    $SystemDetail["BIOS_Version"] = $BIOSInfo.Version
    $SystemDetail["BIOS_Release_Date"] = $BIOSReleaseDate

    return $SystemDetail 
}

#ٍShared Variables
$OSInfo = SafeGet-CimInstance Win32_OperatingSystem
$Processor = SafeGet-CimInstance Win32_Processor
$Memory = SafeGet-CimInstance Win32_PhysicalMemory

function Operating_System{
    $OS = @{}

    $SoftwareLicensing = SafeGet-CimInstance SoftwareLicensingService
    $LastBootUp = ConvertTo-DateTime $OSInfo.LastBootUpTime
    $CurrentTime = ConvertTo-DateTime $OSInfo.LocalDateTime
    $UpTime = $CurrentTime - $LastBootUp
    
    $OS["OS_Name"] = $OSInfo.Caption
    $OS["OS_Version"] = $OSInfo.Version
    $OS["OS_Build"] = $OSInfo.BuildNumber
    $OS["OS_product_key"] = $SoftwareLicensing.OA3xOriginalProductKey
    $OS["Installation_Date"] = if ($OSInfo.InstallDate -ne $null) { $OSInfo.InstallDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
    $OS["Up_Time"] = "$($UpTime.Days) days, $($UpTime.Hours) hours, $($UpTime.Minutes) minutes"

    return $OS
}

function System_Processor {
    $System_Processor = @{
    }

    $CacheSizeKB = $(($Processor.L2CacheSize + $Processor.L3CacheSize))
    $CacheSizeMB = [math]::Round($CacheSizeKB / 1024, 2)

    $System_Processor["Processor_Name"] = $Processor.Name
    $System_Processor["Total_Cores"] = $Processor.NumberOfCores
    $System_Processor["Total_Threads"] = $Processor.ThreadCount
    $System_Processor["Clock_Speed"] = "$($Processor.MaxClockSpeed) MHz"
    $System_Processor["Cache_Size"] = "$CacheSizeMB MB "

    return $System_Processor 
}

function System_Memory {
    $System_Memory = @{
    }

    $totalVirtualKB = $OSInfo.TotalVirtualMemorySize
    $totalPhysicalKB = $OSInfo.TotalVisibleMemorySize
    $virtualOnlyKB = $totalVirtualKB - $totalPhysicalKB

    $TotalMemory = [math]::Round($totalVirtualKB / 1MB, 2)
    $virtualOnlyGB = [math]::Round($virtualOnlyKB / 1MB, 2)

    $MemoryTotal = [math]::Round(($Memory | Measure-Object Capacity -Sum).Sum / 1GB, 2)
    $MemoryUsed = [math]::Round(($OSInfo.TotalVisibleMemorySize - $OSInfo.FreePhysicalMemory) / 1MB, 2)
    $MemoryFree = [math]::Round($OSInfo.FreePhysicalMemory / 1MB, 2)
    $MemoryUsage = [math]::Round((($OSInfo.TotalVisibleMemorySize - $OSInfo.FreePhysicalMemory) / $OSInfo.TotalVisibleMemorySize) * 100, 2)

    $System_Memory["Total_Memory"] = "$TotalMemory GB"
    $System_Memory["Virtual_Memory"] = "$virtualOnlyGB GB" 
    $System_Memory["Physical_Memory"] = "$MemoryTotal GB"
    $System_Memory["Used Memory"] = "$MemoryUsed GB"
    $System_Memory["Free_Memory"] = "$MemoryFree GB"
    $System_Memory["Memory_Usage"] = "$MemoryUsage %"

    return $System_Memory 
}

function Physical_Memory{

    $Memory_Data=@{

    }
        $index = 0
        $Memory | ForEach-Object {
        $Physical_Memory=@{

        }

        $size = [math]::Round($_.Capacity/1GB, 2)
        $clockSpeed = $_.ConfiguredClockSpeed
        $Physical_Memory["Manufacturer"]=$($_.Manufacturer)
        $Physical_Memory["Serial_Number"]=$($_.SerialNumber)
        $Physical_Memory["Bank_Label"]=$($_.BankLabel)
        $Physical_Memory["Device_Locator"]=$($_.DeviceLocator)
        $Physical_Memory["Size_(GB)"]=$size 
        $Physical_Memory["ClockSpeed_(MHz)"]=$clockSpeed
        $Memory_Data["RAM_$index"]=$Physical_Memory
        $index++
    }

    return $Memory_Data
}

$SystemAll["System_Detail"]= System_Details
$SystemAll["Operating_System"]= Operating_System
$SystemAll["Processor"]= System_Processor
$SystemAll["System_Memory"]= System_Memory
$SystemAll["Physical_Memory"]= Physical_Memory


# Convert the dictionary to a JSON string with indentation for readability
$jsonString = $SystemAll | ConvertTo-Json -Depth 10

# Print the formatted JSON string
Write-Host $jsonString
