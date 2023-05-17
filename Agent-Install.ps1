[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$Location,

    [Parameter(Mandatory = $true)]
    [string]$Key
)

# Step 2: Check if services are running
Write-Host "Looking for services to see if they are running..."
$services = Get-Service | Where-Object {$_.Name -eq "ITSPlatform" -or $_.Name -eq "ITSPlatformManager"}

if ($services) {
    Write-Host "Services found:"
    $services | Format-Table Name, Status
    Write-Host
    # Step 3: Start services if they are stopped
    $stoppedServices = $services | Where-Object {$_.Status -eq "Stopped"}
    if ($stoppedServices) {
        Write-Host "The following services are stopped and will be started:"
        $stoppedServices | ForEach-Object {
            Write-Host $_.Name
            try {
                Start-Service $_.Name -ErrorAction Stop
                Write-Host "Service started."
            } catch {
                Write-Host "Failed to start the service."
            }
        }
        Write-Host
    }
    else {
        Write-Host "All services are already running."
        Write-Host "Script execution completed."
        return
    }
} else {
    # Step 4: Create Support directory if no services found
    Write-Host "Service not found."
    Write-Host "Proceeding with the installation."
    $supportDir = "C:\Support"
    if (-not (Test-Path $supportDir)) {
        New-Item -ItemType Directory -Path $supportDir | Out-Null
        Write-Host "Support directory created."
    }
}

# Step 5: Download file
$fileUrl = "https://drive.google.com/uc?id=1D-0qpQYxgFEukqs2fsF2ZybIu1d8X72N&authuser=0&export=download&confirm=t&uuid=6a923679-d2a0-4d66-b8c8-94171796735d&at=AKKF8vyXz8agiJg36W8zUHoYWzHv:1684328105399"

$constantFileName = "Windows_OS_ITSPlatform_"
$fileName = "C:\Support\$Location`_$constantFileName`TKN$Key.msi"

# Check if the file already exists and overwrite it
if (Test-Path $fileName) {
    Write-Host "File already exists. Overwriting..."
    Remove-Item $fileName -Force
}

# Download the file
Write-Host "Downloading file..."
Invoke-WebRequest -Uri $fileUrl -OutFile $fileName

# Step 9: Silently install MSI file
Write-Host "Installing MSI file..."
Start-Process -FilePath msiexec.exe -ArgumentList "/i `"$fileName`" /qn" -Wait
Write-Host "Installation completed."

# Step 10: Print script completion message
Write-Host "Script execution completed."
