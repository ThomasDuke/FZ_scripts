# Create the Office_Install directory
$installPath = "C:\Office_Install"
if (-Not (Test-Path -Path $installPath)) {
    New-Item -Path $installPath -ItemType Directory
}

# Download files
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/ThomasDuke/FZ_scripts/main/Configuration.xml" -OutFile "$installPath\Configuration.xml"
Invoke-WebRequest -Uri "https://github.com/ThomasDuke/FZ_scripts/raw/main/officedeploymenttool_17531-20046.exe" -OutFile "$installPath\officedeploymenttool.exe"

# Execute the installer
Start-Process -FilePath "$installPath\officedeploymenttool.exe" -ArgumentList "/quiet" -Wait

# Run setup to install Office
Start-Process -FilePath "$installPath\setup.exe" -ArgumentList "/configure $installPath\Configuration.xml" -NoNewWindow -WindowStyle Hidden -Wait

# Function to check activation status
function Check-Activation {
    $result = & "cscript" "C:\Program Files\Microsoft Office\Office16\ospp.vbs" "/dstatus" 2>&1
    return $result -like "*LICENSE STATUS: ---LICENSED---*"
}

# Function to activate Office
function Activate-Office {
    param (
        [string]$kmsServer
    )
    & "cscript" "C:\Program Files\Microsoft Office\Office16\ospp.vbs" "/sethst:$kmsServer"
    & "cscript" "C:\Program Files\Microsoft Office\Office16\ospp.vbs" "/act"
}

# Attempt to activate Office with different KMS servers
$kmsServers = @("kms8.MSGuides.com", "kms.03k.org", "kms9.MSGuides.com")
foreach ($server in $kmsServers) {
    Activate-Office -kmsServer $server
    Start-Sleep -Seconds 10  # Wait for activation to process
    if (Check-Activation) {
        Write-Output "Office successfully activated using $server."
        exit 0
    }
}

Write-Output "Office activation failed with all KMS servers."
exit 1
