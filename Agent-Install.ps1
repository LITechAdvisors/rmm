[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$Location,

    [Parameter(Mandatory = $true)]
    [string]$Key
)


Write-Host "Looking for services to see if they are running..."
$services = Get-Service | Where-Object {$_.Name -eq "ITSPlatform" -or $_.Name -eq "ITSPlatformManager"}

if ($services) {
    Write-Host "Services found:"
    $services | Format-Table Name, Status
    Write-Host

    $stoppedServices = $services | Where-Object {$_.Status -eq "Stopped"}
    if ($stoppedServices) {
        Write-Host "The following services are stopped and will be started:"
        $stoppedServices | ForEach-Object {
            Write-Host $_.Name
            Start-Service $_.Name
            Write-Host "Service started."
        }
        Write-Host
    }
    else {
        Write-Host "All services are already running."
    }
} else {

    Write-Host "Service not found."
    Write-Host "Proceeding with the installation."
    $supportDir = "C:\Support"
    if (-not (Test-Path $supportDir)) {
        New-Item -ItemType Directory -Path $supportDir | Out-Null
        Write-Host "Support directory created."
    }
}


$fileUrl = "https://drive.google.com/uc?id=1D-0qpQYxgFEukqs2fsF2ZybIu1d8X72N&authuser=0&export=download&confirm=t&uuid=6a923679-d2a0-4d66-b8c8-94171796735d&at=AKKF8vyXz8agiJg36W8zUHoYWzHv:1684328105399"

$constantFileName = "Windows_OS_ITSPlatform_"
$fileName = "C:\Support\$Location`_$constantFileName`TKN$Key.msi"


if (Test-Path $fileName) {
    Write-Host "File already exists. Overwriting..."
    Remove-Item $fileName -Force
}


Write-Host "Downloading file..."
Invoke-WebRequest -Uri $fileUrl -OutFile $fileName

Write-Host "Installing MSI file..."
Start-Process -FilePath msiexec.exe -ArgumentList "/i `"$fileName`" /qn" -Wait
Write-Host "Installation completed."

Write-Host "Script execution completed."

