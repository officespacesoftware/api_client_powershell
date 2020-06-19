########################
# User config
########################
$ossHostname = "subdomain.officespacesoftware.com"                                    # OfficeSpace instance hostname
$ossApiKey = ""               # OfficeSpace API key
$useLogFile = $true                                            # Log output to file ($true=log, $false=do not log)
$logFile = "$PSScriptRoot\asset_posting_log.txt"                   # Path to log file

########################

$ossProtocol = "https://"
$ossHeaders = @{Authorization = "Token token=" + $ossApiKey }
$ossSeatsUrl = "/api/1/seats/"
$ossAssetsUrl = $ossProtocol + $ossHostname + $ossSeatsUrl
$pathToAssets = "assets.txt"
$splitDelimiter = ","   # Each row should be seatId{splitDelimiter}assetName. E.g.: 1234,Ergo Chair
$version = 1

########################

########################
# Functions
########################
# Cleanup and exit the script
function Exit-Script {
    param(
        [Parameter(Mandatory = $false)] [Int]$Code = 0
    )
    # Stop logging
    if ($useLogFile) {
        Stop-Transcript
    }
    Exit $Code
}

########################
# Script execution start
########################
# Capture the start time of script
$startTime = Get-Date

# Start logging stdout and stderr to file
if ($useLogFile) {
    Start-Transcript -Path "$logFile"
}
Write-Host "Script version $version start"

# Force the use of TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

Write-Host "Attempting to read assets file."

# Get list of seat ids and assets: seatId,assetName  
$assets = Get-Content -Path $pathToAssets

Write-Host "Attempting to add $($assets.Count) assets."

$assetsAttempted = 0
$assetsAdded = 0

# For each seat
ForEach ($asset in $assets) {
    # Split the row into two substrings
    $splitup = $asset.Split($splitDelimiter, 2)
    $seatId = $splitup[0]
    $assetName = $splitup[1]

    Write-Host "Attempting to add $assetName to seat id $seatId"

    # Create the seat asset URL
    $ossPostSeatAssetUrl = $ossAssetsUrl + $seatId + "/assets"

    # Create body for the request
    $ossAssetPostBody = "name=" + $assetName

    # Add the asset
    try {
        $resp = Invoke-WebRequest -UseBasicParsing -Uri $ossPostSeatAssetUrl -Method Post -Body $ossAssetPostBody -Headers $ossHeaders
        if ($resp.StatusCode -ne 201) {
            Write-Host "        Non-201 status code [$($resp.StatusCode)/$($resp.StatusDescription)] seen when attempting to assign asset" -ForegroundColor Red
            Exit-Script 1
        }
        else {
            $assetsAdded++
        }
        
    }
    catch [System.Net.WebException] {
        $e = $_
        Write-Host "        Exception caught while attempting to assign asset: $($e.Exception)" -ForegroundColor Red
        Exit-Script 1
    }
    catch {
        Write-Host "        General Exception caught while attempting to assign asset: $_" -ForegroundColor Red
        Exit-Script 1
    }
    $assetsAttempted++
}

# Stats
Write-Host "$assetsAttempted assets attempted"
Write-Host "$assetsAdded assets added"

# Timing
$endTime = Get-Date
$elapsedTime = $endTime - $startTime
Write-Host "Completed in $($elapsedTime.TotalSeconds) seconds"

Exit-Script 0
