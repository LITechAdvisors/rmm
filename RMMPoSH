# Uninstall CW RMM
Write-Output ""
Write-Output "Beginning CW RMM Uninstall"

# Uninstall ITSPlatform
Write-Output "Triggering ITSPlatform uninstall"
Get-WmiObject -Class Win32_Product -Filter "Name='ITSPlatform'" | ForEach-Object { $_.Uninstall() }

# Run Uninstall.exe for SAAZOD
Write-Output "Triggering SAAZOD uninstall"
if (Test-Path -Path "C:\Program Files (x86)\SAAZOD\Uninstall\Uninstall.exe") {
	Start-Process -FilePath "C:\Program Files (x86)\SAAZOD\Uninstall\Uninstall.exe" -Wait -ArgumentList '/silent /u:"C:\Program Files (x86)\SAAZOD\Uninstall\Uninstall.xml"'
	Start-Sleep -s 10
}
else {
	Write-Output "WARNING: C:\Program Files (x86)\SAAZOD\Uninstall\Uninstall.exe does not exist"
}


# Stop and force close related processes
#Get-Process -Name 'platform-agent*', 'SAAZ*', 'rthlpdk' | Stop-Process -Force
Write-Output "Stopping CW RMM processes"
$CWRMMProcesses = @(
	'platform-agent*',
	'SAAZ*',
	'rthlpdk'
)

$ProcessesToStop = Get-Process -Name $CWRMMProcesses | Select-Object -ExpandProperty Name
forEach ($ProcessToStop in $ProcessesToStop) {
	if (Get-Process -Name "$ProcessToStop") {
		Write-Output "Stopping Process: $ProcessToStop"
		Get-Process -Name "$ProcessToStop" | Stop-Process -Force
	}
}

# Stop and force close related services
#Get-Service -Name 'ITSPlatform*', 'SAAZ*' | Stop-Service -Force
Write-Output "Stopping CW RMM services"
$CWRMMServices = @(
	'ITSPlatform*',
	'SAAZ*'
)

$ServicesToStop = Get-Service -Name $CWRMMServices | Select-Object -ExpandProperty Name
forEach ($ServiceToStop in $ServicesToStop) {
	if (Get-Service -Name "$ServiceToStop") {
		Write-Output "Stopping Service: $ServiceToStop"
		Get-Service -Name "$ServiceToStop" | Stop-Service -Force
	}
}

# Delete specified services
Write-Output "Deleting CW RMM services"
foreach ($service in $ServicesToStop) {
	#sc.exe delete $service
	Write-Output "Deleting service: $service"
	Start-Process -FilePath "sc.exe" -ArgumentList "delete", $service -NoNewWindow -Wait 
}

#Alternative service deletion
Get-CimInstance -ClassName win32_service | Where-Object { ($_.PathName -like 'ITSPlatform*') -or ($_.PathName -like 'SAAZ*') } | ForEach-Object { Invoke-CimMethod $_ -Name StopService; Remove-CimInstance $_ -Verbose -Confirm:$false }

# Delete specified folders
Write-Output "Deleting CW RMM related folders."
$FoldersToDelete = @(
	'C:\Program Files (x86)\ITSPlatformSetupLogs',
	'C:\Program Files (x86)\ITSPlatform',
	'C:\Program Files (x86)\SAAZOD',
	'C:\Program Files (x86)\SAAZODBKP',
	'C:\ProgramData\SAAZOD'
)

foreach ($folder in $FoldersToDelete) {
	if (Test-Path $folder) {
		Write-Output "Deleting folder: $folder"
		Get-ChildItem $folder -Recurse -Force | Remove-Item -Recurse -Force -Confirm:$false -Verbose
		Start-Sleep -Seconds 1
		Remove-Item -Path $folder -Recurse -Force -Confirm:$false -Verbose
	}
}


# Remove "C:\Program" file
Write-Output "Testing for rogue 'C:\Program' file"
if (Test-Path -LiteralPath "C:\Program" -PathType leaf) {
	Write-Output "Deleting rogue 'C:\Program' file"
	Remove-Item -LiteralPath "C:\Program" -Force -Confirm:$false -Verbose
}


# Remove registry keys
Write-Output "Deleting CW RMM registry keys"
$RegistryKeysToRemove = @(
	'HKLM:\SOFTWARE\WOW6432Node\SAAZOD',
	'HKLM:\SOFTWARE\WOW6432Node\ITSPlatform'    
)

foreach ($RegistryKey in $RegistryKeysToRemove) {
	if (Test-Path -LiteralPath $RegistryKey) {
		Write-Output "Deleting registry key: $RegistryKey"
		Remove-Item -Path $RegistryKey -Recurse -Force -Confirm:$false -Verbose
	}
}

# Remove the registry key with spaces
$BlankRegistryKey = "HKLM:\SOFTWARE\WOW6432Node\  \ITSPlatform"

# Test if the specific registry key exists
if (Test-Path -LiteralPath $BlankRegistryKey) {
	Write-Output "Deleting registry key: $BlankRegistryKey"

	# Get parent key's path
	$parentKey = Split-Path $BlankRegistryKey -Parent

	try {
		# Delete the parent key recursively
		Remove-Item -Path $parentKey -Recurse -Force -Confirm:$false -Verbose -ErrorAction Stop
		Write-Output "Parent registry key deleted successfully."
	}
	catch {
		Write-Output "Failed to delete parent registry key: $_"
	}
}
else {
	Write-Output "Blank registry key not found. Nothing to delete."
}



### Clean up installer registry keys

# Define the path to start the search
$startPath = 'HKLM:\SOFTWARE\Classes\Installer\Products\'

# Define the target product name
$targetProductName = 'ITSPlatform'

# Function to search for keys with the specified product name
function Find-RegistryKeys {
	param (
		[string]$Path
	)

	# Get all subkeys
	$subKeys = Get-ChildItem -Path $Path -ErrorAction SilentlyContinue

	# Loop through each subkey
	foreach ($subKey in $subKeys) {
		# Get the value of the productName property
		$productName = (Get-ItemProperty -Path $subKey.PSPath -Name 'productName' -ErrorAction SilentlyContinue).productName

		# Check if the productName matches the target product name
		if ($productName -eq $targetProductName) {
			# Output the path of the matching key
			#Write-Output $subKey.PSPath
			[PSCustomObject]@{
				Path = $subKey.PSPath
				ProductName = $productName
			}
		}

		# Search for matching keys in the current subkey
		Find-RegistryKeys -Path $subKey.PSPath
	}
}

# Start the search
$ProductRegistryKeysToRemove = Find-RegistryKeys -Path $startPath

# Delete the found keys
if ($ProductRegistryKeysToRemove) {
    foreach ($ProductRegistryKeyToRemove in $ProductRegistryKeysToRemove) {
        if (Test-Path -LiteralPath $ProductRegistryKeyToRemove.Path) {
            Write-Output "Deleting registry key: $($ProductRegistryKeyToRemove.Path)"
            Remove-Item -Path $ProductRegistryKeyToRemove.Path -Recurse -Force -Confirm:$false -Verbose
        }
    }
}

Write-Output "Done!"

Start-Sleep -Seconds 60

# Uninstall CW RMM
Write-Output ""
Write-Output "Beginning CW RMM Uninstall"

# Uninstall ITSPlatform
Write-Output "Triggering ITSPlatform uninstall"
Get-WmiObject -Class Win32_Product -Filter "Name='ITSPlatform'" | ForEach-Object { $_.Uninstall() }

# Run Uninstall.exe for SAAZOD
Write-Output "Triggering SAAZOD uninstall"
if (Test-Path -Path "C:\Program Files (x86)\SAAZOD\Uninstall\Uninstall.exe") {
	Start-Process -FilePath "C:\Program Files (x86)\SAAZOD\Uninstall\Uninstall.exe" -Wait -ArgumentList '/silent /u:"C:\Program Files (x86)\SAAZOD\Uninstall\Uninstall.xml"'
	Start-Sleep -s 10
}
else {
	Write-Output "WARNING: C:\Program Files (x86)\SAAZOD\Uninstall\Uninstall.exe does not exist"
}


# Stop and force close related processes
#Get-Process -Name 'platform-agent*', 'SAAZ*', 'rthlpdk' | Stop-Process -Force
Write-Output "Stopping CW RMM processes"
$CWRMMProcesses = @(
	'platform-agent*',
	'SAAZ*',
	'rthlpdk'
)

$ProcessesToStop = Get-Process -Name $CWRMMProcesses | Select-Object -ExpandProperty Name
forEach ($ProcessToStop in $ProcessesToStop) {
	if (Get-Process -Name "$ProcessToStop") {
		Write-Output "Stopping Process: $ProcessToStop"
		Get-Process -Name "$ProcessToStop" | Stop-Process -Force
	}
}

# Stop and force close related services
#Get-Service -Name 'ITSPlatform*', 'SAAZ*' | Stop-Service -Force
Write-Output "Stopping CW RMM services"
$CWRMMServices = @(
	'ITSPlatform*',
	'SAAZ*'
)

$ServicesToStop = Get-Service -Name $CWRMMServices | Select-Object -ExpandProperty Name
forEach ($ServiceToStop in $ServicesToStop) {
	if (Get-Service -Name "$ServiceToStop") {
		Write-Output "Stopping Service: $ServiceToStop"
		Get-Service -Name "$ServiceToStop" | Stop-Service -Force
	}
}

# Delete specified services
Write-Output "Deleting CW RMM services"
foreach ($service in $ServicesToStop) {
	#sc.exe delete $service
	Write-Output "Deleting service: $service"
	Start-Process -FilePath "sc.exe" -ArgumentList "delete", $service -NoNewWindow -Wait 
}

#Alternative service deletion
Get-CimInstance -ClassName win32_service | Where-Object { ($_.PathName -like 'ITSPlatform*') -or ($_.PathName -like 'SAAZ*') } | ForEach-Object { Invoke-CimMethod $_ -Name StopService; Remove-CimInstance $_ -Verbose -Confirm:$false }

# Delete specified folders
Write-Output "Deleting CW RMM related folders."
$FoldersToDelete = @(
	'C:\Program Files (x86)\ITSPlatformSetupLogs',
	'C:\Program Files (x86)\ITSPlatform',
	'C:\Program Files (x86)\SAAZOD',
	'C:\Program Files (x86)\SAAZODBKP',
	'C:\ProgramData\SAAZOD'
)

foreach ($folder in $FoldersToDelete) {
	if (Test-Path $folder) {
		Write-Output "Deleting folder: $folder"
		Get-ChildItem $folder -Recurse -Force | Remove-Item -Recurse -Force -Confirm:$false -Verbose
		Start-Sleep -Seconds 1
		Remove-Item -Path $folder -Recurse -Force -Confirm:$false -Verbose
	}
}


# Remove "C:\Program" file
Write-Output "Testing for rogue 'C:\Program' file"
if (Test-Path -LiteralPath "C:\Program" -PathType leaf) {
	Write-Output "Deleting rogue 'C:\Program' file"
	Remove-Item -LiteralPath "C:\Program" -Force -Confirm:$false -Verbose
}


# Remove registry keys
Write-Output "Deleting CW RMM registry keys"
$RegistryKeysToRemove = @(
	'HKLM:\SOFTWARE\WOW6432Node\SAAZOD',
	'HKLM:\SOFTWARE\WOW6432Node\ITSPlatform'    
)

foreach ($RegistryKey in $RegistryKeysToRemove) {
	if (Test-Path -LiteralPath $RegistryKey) {
		Write-Output "Deleting registry key: $RegistryKey"
		Remove-Item -Path $RegistryKey -Recurse -Force -Confirm:$false -Verbose
	}
}

# Remove the registry key with spaces
$BlankRegistryKey = "HKLM:\SOFTWARE\WOW6432Node\  \ITSPlatform"

# Test if the specific registry key exists
if (Test-Path -LiteralPath $BlankRegistryKey) {
	Write-Output "Deleting registry key: $BlankRegistryKey"

	# Get parent key's path
	$parentKey = Split-Path $BlankRegistryKey -Parent

	try {
		# Delete the parent key recursively
		Remove-Item -Path $parentKey -Recurse -Force -Confirm:$false -Verbose -ErrorAction Stop
		Write-Output "Parent registry key deleted successfully."
	}
	catch {
		Write-Output "Failed to delete parent registry key: $_"
	}
}
else {
	Write-Output "Blank registry key not found. Nothing to delete."
}



### Clean up installer registry keys

# Define the path to start the search
$startPath = 'HKLM:\SOFTWARE\Classes\Installer\Products\'

# Define the target product name
$targetProductName = 'ITSPlatform'

# Function to search for keys with the specified product name
function Find-RegistryKeys {
	param (
		[string]$Path
	)

	# Get all subkeys
	$subKeys = Get-ChildItem -Path $Path -ErrorAction SilentlyContinue

	# Loop through each subkey
	foreach ($subKey in $subKeys) {
		# Get the value of the productName property
		$productName = (Get-ItemProperty -Path $subKey.PSPath -Name 'productName' -ErrorAction SilentlyContinue).productName

		# Check if the productName matches the target product name
		if ($productName -eq $targetProductName) {
			# Output the path of the matching key
			#Write-Output $subKey.PSPath
			[PSCustomObject]@{
				Path = $subKey.PSPath
				ProductName = $productName
			}
		}

		# Search for matching keys in the current subkey
		Find-RegistryKeys -Path $subKey.PSPath
	}
}

# Start the search
$ProductRegistryKeysToRemove = Find-RegistryKeys -Path $startPath

# Delete the found keys
if ($ProductRegistryKeysToRemove) {
    foreach ($ProductRegistryKeyToRemove in $ProductRegistryKeysToRemove) {
        if (Test-Path -LiteralPath $ProductRegistryKeyToRemove.Path) {
            Write-Output "Deleting registry key: $($ProductRegistryKeyToRemove.Path)"
            Remove-Item -Path $ProductRegistryKeyToRemove.Path -Recurse -Force -Confirm:$false -Verbose
        }
    }
}

Write-Output "Done!"

## This has run two times now, ready for reinstall

Start-Sleep 60


#=======================================================================
# Run PowerShell script based on OS architecture
#=======================================================================
if ( $env:PROCESSOR_ARCHITEW6432 -eq "AMD64" ) {
    if ( $myInvocation.Line ) {
        &"$env:systemroot\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive `
            -NoProfile $myInvocation.Line
    }
    else {
        &"$env:systemroot\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive `
            -NoProfile -file "$( $myInvocation.InvocationName )" $args
    }
    EXIT $lastEXITcode
}

#=======================================================================
# Variable Initializations
#=======================================================================
$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'

$MSMAURL = "https://prod.setup.itsupport247.net/windows/MSMA/32/LITechAdvisors-NewComputers_MSMA_ITSPlatform_TKN225b35ae-1794-4b69-87c1-d212db94bf85/MSI/setup"
$DPMAURL = "https://prod.setup.itsupport247.net/windows/DPMA/32/LITechAdvisors-NewComputers_DPMA_ITSPlatform_TKN225b35ae-1794-4b69-87c1-d212db94bf85/MSI/setup"
$TargetFileServer = "${env:TEMP}\LITechAdvisors-NewComputers_MSMA_ITSPlatform_TKN225b35ae-1794-4b69-87c1-d212db94bf85.msi"
$TargetFileDesktop = "${env:TEMP}\LITechAdvisors-NewComputers_DPMA_ITSPlatform_TKN225b35ae-1794-4b69-87c1-d212db94bf85.msi"

#=======================================================================
# Functions Definitions
#=======================================================================
function fcn-downloadFile {
    Param (
        [ Parameter( Mandatory = $true ) ]
        $url,
        [ Parameter( Mandatory = $true ) ]
        $targetFile
    )
    Process {
        try {
            $uri = New-Object "System.Uri" "$url"
            #$targetFile = "$TargetFileServer"
            $request = [System.Net.HttpWebRequest]::Create($uri)
            $request.set_Timeout(1200000)
            $response = $request.GetResponse()
            $responseStream = $response.GetResponseStream()
            $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile, Create
            $buffer = new-object byte[] 10KB
            $count = $responseStream.Read($buffer, 0, $buffer.length)
            $downloadedBytes = $count
            while ($count -gt 0) {
                $targetStream.Write($buffer, 0, $count)
                $count = $responseStream.Read($buffer, 0, $buffer.length)
                $downloadedBytes = $downloadedBytes + $count
            }
            $targetStream.Flush()
            $targetStream.Close()
            $targetStream.Dispose()
            $responseStream.Dispose()
    
            if ( Test-Path $targetFile ) {
                return 0
            }
            else {
                return 1
            }
        }
        catch {
            return -1
        }
    }
}

function fcn-run-exe {
    Param (
        [ Parameter( Mandatory = $true ) ]
        $FPATH
    )
    Process {
        try {
            $ErrorActionPreference = 'Stop'
            $arg = @("/qn"; "/i"; "$FPATH") 
            $returncode = (Start-Process -FilePath msiexec.exe -ArgumentList $arg -Wait -Passthru).ExitCode
            if ( $returncode -eq 0 ) {
                return 0
            }
            else {
                return $returncode
            }
        }
        catch {
            return -1
        }
    }
}

#=======================================================================
# Code to Check Pre-Requisite to run script
#=======================================================================
try {
    #   OS Version & PS Version Check
    [int]$PSVersion = $PSVersionTable.PSVersion.Major
    [double]$OSVersion = [Environment]::OSVersion.Version.ToString( 2 ) 
    if ( ( $OSVersion -lt 6.1 ) -or ( $PSVersion -lt 2 ) ) {
        Write-Output "[MSG]: System is not compatible with the requirement. Either machine is below Server 2008R2/WIndows 7 or Powershell version is lower than 2.0"
        EXIT
    }
}
catch {
    Write-Output "Error Message: $($Error[0].Exception.Message)"
    EXIT
}

"$("-"*40)`n"
#=======================================================================
# Check for existence of services of RMM Agent
#=======================================================================
try {
    "Attempting to fetch status of RMM Agent services"
    $services = Get-Service -Name 'ITSPlatformManager', 'ITSPlatform' -ErrorAction Stop
    $services | ForEach-Object {
        "`nService Name   : $($_.DisplayName)"
        "Service Status : $($_.Status)"     
    }
    "RMM Agent found on machine. Hence, cannot proceed further for agent installation."
    EXIT
}
catch {
    "`n[ERROR] : Unable to get details of RMM Agent services due to below error"
    "$($Error[0].Tostring())"
    "`nSince, RMM Agent is not found on machine. Hence, proceeding further for agent installation."
}

#=======================================================================
# Check if current machine is DPMA or MSMA
#=======================================================================
$OSType = (Get-WmiObject -Class win32_ComputerSystem -ErrorAction Stop).DomainRole
if (($OSType -eq 1) -or ($OSType -eq 0)) {
    $Machine_type = "DPMA"
    "[INFO] : Current machine is DPMA"
}
else {
    $Machine_type = "MSMA"
    "[INFO] : Current machine is MSMA"
}

try {
    "`n$("="*40)`n"
    "Attempting to GET Agent Installer.`n"
    if ($Machine_type -eq "DPMA") {
        "[INFO] : Downloading DPMA RMM Agent Installer"
        $n_dl_result = fcn-downloadFile -url "$DPMAURL" -targetFile "$TargetFileDesktop"
        if ($n_dl_result -eq 0) {
            "[INFO] : Successfully downloaded DPMA RMM Agent Installer"
        }
        else {
            "[ERROR] : Failed to download DPMA RMM Agent Installer"
            EXIT
        }
    }
    else {
        "[INFO] : Downloading MSMA RMM Agent Installer"
        $n_dl_result = fcn-downloadFile -url "$MSMAURL" -targetFile "$TargetFileServer"
        if ($n_dl_result -eq 0) {
            "[INFO] : Successfully downloaded MSMA RMM Agent Installer"
        }
        else {
            "[ERROR] : Failed to download MSMA RMM Agent Installer"
            EXIT
        }
    }
}
catch {
    "Error Message: $($Error[0].Exception.Message)"
    EXIT
}

#=======================================================================
# Installing Agent Installer
#=======================================================================
try {
    "`n$("="*40)`n"
    "Attempting to install RMM Agent Installer.`n"
    if ($Machine_type -eq "DPMA") {
        "[INFO] : Installing DPMA RMM Agent Installer"
        $n_exe_result = fcn-run-exe -FPATH "$TargetFileDesktop"
        if ($n_exe_result -eq 0) {
            "[INFO] : Successfully installed DPMA RMM Agent Installer"
        }
        else {
            "[ERROR] : Failed to install DPMA RMM Agent Installer"
            EXIT
        }
    }
    else {
        "[INFO] : Installing MSMA RMM Agent Installer"
        $n_exe_result = fcn-run-exe -FPATH "$TargetFileServer"
        if ($n_exe_result -eq 0) {
            "[INFO] : Successfully installed MSMA RMM Agent Installer"
        }
        else {
            "[ERROR] : Failed to install MSMA RMM Agent Installer"
            EXIT
        }
    }
}
catch {
    "Error Message: $($Error[0].Exception.Message)"
    EXIT
}

## We'll put SAAZ manual install in here if necessary but it does not install


#=======================================================================
# Cleanup
#=======================================================================
finally {
    # Delete downloaded Installer.
    if ($TargetFileDesktop) {
        Remove-Item -Path $TargetFileDesktop -Force -ErrorAction SilentlyContinue
    }
    if ($TargetFileServer) {
        Remove-Item -Path $TargetFileServer -Force -ErrorAction SilentlyContinue
    }
}

