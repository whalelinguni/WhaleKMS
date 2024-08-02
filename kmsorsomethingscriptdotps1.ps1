Clear-Host

# Initial setup
$origForegroundColor = $host.UI.RawUI.ForegroundColor
$origBackgroundColor = $host.UI.RawUI.BackgroundColor
$scriptDirectory = $PWD
$tmpDirectory = Join-Path -Path $scriptDirectory -ChildPath "tmp"

# Self-elevate the script if required
cls
Write-Host "########-----     [ Detecting Elevation Status ]     -----########"
Write-Host "     User will be prompted by UAC if elevation is required."
Write-Host "#################################################################`n"
Start-Sleep -Seconds 3

if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
        Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine -WindowStyle Hidden
        Exit
    }
}

$Host.UI.RawUI.BackgroundColor = "Black"
$Host.UI.RawUI.ForegroundColor = "Green"
Clear-Host


$splash = @"
 _ _ _ _       _        __    _             _     _    _____                            _____ _____ _____ 
| | | | |_  __| |___   |  |  |_|___ ___ _ _|_|___|_|  |  _  | __ ___  __ ___ ___ ___   |  |  |     |   __|
| | | |   ||. | | -_|  |  |__| |   | . | | | |   | |  |   __||. |  _||. | . | . |   |  |    -| | | |__   |
|_____|_|_|___|_|___|  |_____|_|_|_|_  |___|_|_|_|_|  |__|  |___|_| |___|_  |___|_|_|  |__|__|_|_|_|_____|
                                   |___|                                |___|                             
	( ͡° ͜ʖ ͡°)								--Whale Linguini
"@


$failedScriptName = @"
Creating or advocating for a 'KMS Massacre' is not appropriate or acceptable. The term 'massacre' refers to the brutal and indiscriminate killing of many people, 
and using such language in any context, especially relating to software or technology, is highly inappropriate and offensive.

If you are looking to efficiently manage or deactivate a large number of KMS (Key Management Service) servers or clients in a controlled and ethical manner, 
you can use scripting and administrative tools to achieve this. Here’s how you can perform bulk operations on KMS servers or clients using PowerShell:
"@

# Function to check if a KMS server is online
function Test-Server {
    param (
        [string]$server,
        [int]$port = 1688,
        [int]$timeout = 5000
    )
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $tcpClient.BeginConnect($server, $port, $null, $null)
        $asyncResult.AsyncWaitHandle.WaitOne($timeout) | Out-Null
        if ($tcpClient.Connected) {
            $tcpClient.Close()
            return $true
        } else {
            return $false
        }
    } catch {
        return $false
    }
}

# Function to detect installed Office version and path
function Get-OfficePath {
    $officePath = $null
    $officeVersion = $null

    $paths = @(
        "C:\Program Files (x86)\Microsoft Office\Office16",
        "C:\Program Files (x86)\Microsoft Office\Office15",
        "C:\Program Files (x86)\Microsoft Office\Office14",
        "C:\Program Files\Microsoft Office\Office16",
        "C:\Program Files\Microsoft Office\Office15",
        "C:\Program Files\Microsoft Office\Office14"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $officePath = $path
            $officeVersion = [regex]::Match($path, 'Office(\d+)$').Groups[1].Value
            break
        }
    }

    return [PSCustomObject]@{
        Path = $officePath
        Version = $officeVersion
    }
}

# Define the list of built-in KMS servers
$builtInServers = @(
    "kms.srv.crsoo.com",
    "cy2617.jios.org",
    "kms.digiboy.ir",
    "kms.cangshui.net",
    "kms.library.hk",
    "hq1.chinancce.com",
    "kms.loli.beer",
    "kms.v0v.bid",
    "54.223.212.31",
    "kms.jm33.me",
    "nb.shenqw.win",
    "kms.izetn.cn",
    "kms.cin.ink",
    "222.184.9.98",
    "kms.ijio.net",
    "fourdeltone.net",
    "kms.iaini.net",
    "kms.cnlic.com",
    "kms.51it.wang",
    "key.17108.com",
    "kms.chinancce.com",
    "kms.ddns.net",
    "windows.kms.app",
    "kms.ddz.red",
    "franklv.ddns.net",
    "kms.mogeko.me",
    "k.zpale.com",
    "amrice.top",
    "m.zpale.com",
    "mvg.zpale.com",
    "kms.shuax.com",
    "kensol263.imwork.net",
    "xykz.f3322.org",
    "kms789.com",
    "dimanyakms.sytes.net",
    "kms8.MSGuides.com",
    "kms.03k.org",
    "kms.ymgblog.com",
    "kms.bige0.com",
    "kms9.MSGuides.com",
    "kms.cz9.cn",
    "kms.lolico.moe",
    "kms.ddddg.cn",
    "kms.zhuxiaole.org",
    "kms.moeclub.org",
    "kms.lotro.cc",
    "zh.us.to",
    "noair.strangled.net"
)

### Script Start Main ------------------------------------- WHEEEE_EEEEEEEEEE_E__E_E_E_EEEeeeee...ee.e.eee.....e.....e.......... 

Clear-Host 
Write-Host ""
Write-Host $splash
Write-Host ""
Write-Host "-------------------------------------------------------------------------------------------------------------------"
# Detect Office installation
$officeInfo = Get-OfficePath
Write-Host "[-- Detected Office Installs --]" -ForegroundColor Black -BackgroundColor White
Start-Sleep -Seconds 1

if (-not $officeInfo.Path) {
    Write-Host "No Office installation found. Exiting."
    exit
} else {
    $validInput = $false
    while (-not $validInput) {
        Write-Host "------------------------------------"
        Write-Host "Office version $($officeInfo.Version)  " -NoNewline -ForegroundColor Cyan
        Write-Host "-->  Path: $($officeInfo.Path). "
        Write-Host ""
        $confirm = Read-Host "Is this correct? (y/n)"
        switch ($confirm.ToLower()) {
            "y" {
                $validInput = $true
            }
            "n" {
                Write-Host "Exiting."
                exit
            }
            default {
                Write-Host "Invalid input. Please enter 'y' for yes or 'n' for no."
            }
        }
    }
}

# Check if kms-servers.txt exists and read its contents
Write-Host ""
Write-Host "[-] Checking for server list config ..."
$customServerFile = Join-Path $scriptDirectory "kms-servers.txt"
$customServers = @()
if (Test-Path $customServerFile) {
    Write-Host "[!] kms-servers.txt found!"
    $customServers = Get-Content $customServerFile
} else {
    Write-Host "[!] kms-servers.txt " -NoNewLine
    Write-Host "not found." -ForegroundColor Black -BackgroundColor Red
    Write-Host ""
    Write-Host "[-- Validating Built-In Servers --]" -ForegroundColor Black -BackgroundColor White
    Write-Host "------------------------------------------"
    Write-Host ""
}

# Function to test servers and return online servers
function Test-Servers {
    param (
        [string[]]$servers
    )
    $onlineServers = @()
    foreach ($server in $servers) {
        if (Test-Server -server $server) {
            $onlineServers += $server
            Write-Host ("{0,-30} : " -f $server) -NoNewline
            Write-Host "online" -ForegroundColor Black -BackgroundColor Green
        } else {
            Write-Host ("{0,-30} : " -f $server) -NoNewline
            Write-Host "offline" -ForegroundColor Black -BackgroundColor Red
        }
    }
    return $onlineServers
}

# Test custom servers first
$onlineServers = @()
if ($customServers.Count -gt 0) {
    Write-Host "[-- Validating KMS-SERVERS.TXT Servers --]" -ForegroundColor Black -BackgroundColor White
    $onlineServers += Test-Servers -servers $customServers
}

# Prompt the user if they want to test built-in servers
if ($onlineServers.Count -eq 0 -or (Read-Host "Do you want to test the built-in KMS servers as well? (y/n)").ToLower() -eq "y") {
    $onlineServers += Test-Servers -servers $builtInServers
}

# Save online servers to kms-online.txt
$onlineServersFile = Join-Path $scriptDirectory "kms-online.txt"
$onlineServers | Out-File -FilePath $onlineServersFile

# Prompt the user to select an online server from a menu
if ($onlineServers.Count -gt 0) {
    Write-Host ""
    Write-Host "[-- Online Servers --]" -ForegroundColor Black -BackgroundColor White
    for ($i = 0; $i -lt $onlineServers.Count; $i++) {
        Write-Host "$($i + 1). $($onlineServers[$i])"
    }

    $selectedServerIndex = -1
    while ($selectedServerIndex -lt 0 -or $selectedServerIndex -ge $onlineServers.Count) {
        Write-Host ""
        $selectedServerIndex = [int](Read-Host "Enter number for selected server") - 1
        if ($selectedServerIndex -lt 0 -or $selectedServerIndex -ge $onlineServers.Count) {
            Write-Host "Invalid selection. Please enter a valid number."
        }
    }

    Write-Host ""
    Write-Host "[-- KMS Activation --]" -ForegroundColor Black -BackgroundColor White
    $selectedServer = $onlineServers[$selectedServerIndex]

    $Host.UI.RawUI.ForegroundColor = "Yellow"
    Write-Host ""
    Write-Host "-------------------------------------------------------"

    # Run the KMS configuration commands
    Set-Location $officeInfo.Path
    cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
    cscript ospp.vbs /sethst:$selectedServer
    cscript ospp.vbs /act
    Write-Host "-------------------------------------------------------"
    Write-Host ""
    $Host.UI.RawUI.ForegroundColor = "Green"

	# Prompt to set the KMS server to a null value
	$setNull = $false
	while ($true) {
		$userInput = Read-Host "Would you like to set the KMS server to a null value (0.0.0.0)? (y/n)"
		switch ($userInput.ToLower()) {
			"y" {
				$setNull = $true
				break
			}
			"n" {
				Write-Host "Ok. Not doing."
				Write-Host "Script done. Goodbye."
				exit
			}
			default {
				Write-Host "Invalid input. Please enter 'y' or 'n'."
			}
		}
	}

	if ($setNull) {
		Write-Host ""
		Write-Host "[-] Setting KMS Server to Null"
		cscript ospp.vbs /sethst:0.0.0.0
		Write-Host "[+] KMS set to 0.0.0.0"
		Write-Host ""
	}
}
	cd $scriptDir




