# Create the Office_Install directory
$installPath = "C:\Office_Install"
if (-Not (Test-Path -Path $installPath)) {
    New-Item -Path $installPath -ItemType Directory
}

# Download files
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/ThomasDuke/FZ_scripts/main/Configuration.xml" -OutFile "$installPath\Configuration.xml"
Invoke-WebRequest -Uri "https://github.com/ThomasDuke/FZ_scripts/raw/main/officedeploymenttool_17531-20046.exe" -OutFile "$installPath\officedeploymenttool.exe"

# Execute the installer in background
$installerPath = "$installPath\officedeploymenttool.exe"
if (Test-Path -Path $installerPath) {
    Start-Process -FilePath $installerPath -ArgumentList "/extract:$installPath /quiet" -Wait
} else {
    Write-Output "Installer not found at $installerPath"
    exit 1
}

# Check if setup.exe is extracted
$setupPath = "$installPath\setup.exe"
if (Test-Path -Path $setupPath) {
    # Run setup to install Office
    Start-Process -FilePath $setupPath -ArgumentList "/configure $installPath\Configuration.xml" -WindowStyle Hidden -Wait
} else {
    Write-Output "Setup file not found at $setupPath"
    exit 1
}

# Function to check activation status
function Check-Activation {
    $result = & "cscript" "C:\Program Files\Microsoft Office\Office16\ospp.vbs" "/dstatus" 2>&1
    if ($result -contains "<Product activation successful>"){
        exit 0
    }else{
        exit 1
    }
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
    else{
        Write-Output "Office activation failed with all KMS servers."
    }
}


exit 1
