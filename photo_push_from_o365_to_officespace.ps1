#
# photo_push_from_o365_to_officespace.ps1 - upload photos from Exchange Online, Azure AD to OfficeSpace
#
###################
# Script parameters
###################
param (
    [Parameter(Mandatory = $false)]$creds = $false         # (optional) -creds path_to_creds_file
)                                                          #   If supplied, will override User config for $promptForCreds and $credsFile.

#############
# User config
#############
$promptForCreds  = $true                                   # $true: will prompt user for credentials
                                                           # $false: will use creds stored in $credsFile
                                                           # If -creds command line option is provided, this variable is ignored.
$credsFile       = "C:\ps\serviceacct.creds"               # encrypted user credentials file
                                                           # If -creds command line option is provided, this variable is ignored.
$azureUsername   = "serviceacct@customer.com"              # user with Azure AD access
$useLogFile      = $true                                   # Log output to file ($true=log, $false=do not log)
$logFile         = "$PSScriptRoot\photo_import.txt"        # Path to log file
$photoSource     = "exchange-azuread"                      # source for photos: azuread, exchange, exchange-azuread, none
$photosDir       = "C:\ps\photos"                          # directory to download thumbnail photos
                                                           #    (only used if $photoSource = 'azuread' or 'exchange-azuread')
$ossToken        = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"      # OfficeSpace API key
$ossHostname     = "customer.officespacesoftware.com"      # OfficeSpace instance hostname
#############

$supportedPhotoSources      = @( 'azuread', 'exchange', 'exchange-azuread', 'none' )
$ossProtocol                = "https://"
$ossHeaders                 = @{Authorization = "Token token=" + $ossToken}
$ossEmployeesUrl            = "/api/1/employees"
$ossGetEmployeesUrl         = $ossProtocol + $ossHostname + $ossEmployeesUrl
$version                    = 2

###########
# Functions
###########
# Cleanup and exit the script
function Exit-Script {
    param(
        [Parameter(Mandatory=$false)] [Int]$Code = 0 
    )   
    # End Exchange session
    if (Get-PSSession) {
        Remove-PSSession $exchangeSession -Verbose
    }
    # End AzureAD session
    Try {
        Disconnect-AzureAD -Verbose
    } catch {
    }
    # Stop logging
    if ($useLogFile) {
        Stop-Transcript
    }   
    Exit $Code
}


# Test JSON data set capability
function Test-Json {
    Write-Host -NoNewLine "Test JSON parsing capability: "
    Try {
        "a" * 2mb | ConvertTo-Json | ConvertFrom-Json | Out-Null
        Write-Host "PASS"
        return $true
    } catch [System.ArgumentException] {
        Write-Host "FAIL"
        Write-Host "NOTICE: Your PowerShell environment may not support parsing larger JSON data sets."
        return $false
    } catch {
        Write-Host "Test-Json() General exception caught: $_.Exception.Message"
        Stop-Transcript
        Exit-Script 2
    }   
}


# Get JSON data from URL
function Get-WebJson {
    param(
        [Parameter(Mandatory=$true)]  [String]$Url,
        [Parameter(Mandatory=$false)] [HashTable]$Headers = @{}
    )   

    if ($jsonOk) {
        # Invoke-RestMethod
        $resp = Invoke-RestMethod -Uri $Url -Method Get -Headers $Headers
    } else {
        # Invoke-WebRequest
        $w = Invoke-WebRequest -Uri $Url -UseBasicParsing -Method Get -Headers $Headers
        $rawContentLength = $w.RawContentLength
        Write-Host "debug: raw content length: $rawContentLength"
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
        $jsonSerial = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer
        # If MaxJsonLength < $rawContentLength, you'll hit an error.
        $jsonSerial.MaxJsonLength  = $rawContentLength
        $resp = $jsonSerial.DeserializeObject($w.Content)
    }   
    return $resp
}


# Get the Azure users
function Get-UserObjects {
    param ($azureObjects)
    $users = @()
    foreach ($o in $azureObjects) {
        #Write-Host "-> $($o.DisplayName)  [$($o.ObjectType)]"
        if ($o.ObjectType -eq 'User') {
            # Collect the user object
            $users += $o
        } elseif ($o.ObjectType -eq 'Group') {
            #Write-Host "---> will parse this group"
            $gm = Get-AzureADGroupMember -All $true -ObjectId $o.ObjectId
            $users += Get-UserObjects -azureObjects $gm
        } else {
            #Write-Host "---> skip this object"
        }
    }
    return $users
}


# Upload user photo to OSS
function Upload-Photo {
    param(
        [Parameter(Mandatory=$true)]  [String]$ossUrl,
        [Parameter(Mandatory=$true)]  [HashTable]$Headers = @{},
        [Parameter(Mandatory=$true)]  [Array]$imageDataRaw,
        [Parameter(Mandatory=$true)]  [PSCustomObject]$ossUser 
    )
    $imageData = [System.Convert]::ToBase64String($imageDataRaw)
    $request = @{
        record = @{
            imageData = $imageData
        }
    }
    $JSONrequest = $request | ConvertTo-Json
    Write-Host "   >> Updating photo for user: $($ossUser.client_employee_id)"
    $putUrl = "$ossUrl/$($employee.id)"

    # Check response; log and continue on error.
    try {
        $resp = Invoke-WebRequest -UseBasicParsing -Uri $putUrl -ContentType 'application/json; charset=utf-8' -Method PUT -Body $JSONrequest -Headers $Headers  
        Write-Host "      [response: $($resp.StatusCode)/$($resp.StatusDescription)]"
        if ($resp.StatusCode -ne 200) {
            Write-Host "Non-200 status code [$($resp.StatusCode)/$($resp.StatusDescription)] seen when attempting photo upload to $putUrl" -ForegroundColor Red 
            $script:uploadErrors++
        } else {
            $script:photosUploaded++
        }   
    } catch [System.Net.WebException] {
        $e = $_
        Write-Host "Exception caught while attempting photo upload to $putUrl : $($e.Exception)" -ForegroundColor Red 
        $script:uploadErrors++
    } catch {
        Write-Host "General Exception caught while attempting photo upload to $putUrl : $_" -ForegroundColor Red 
        $script:uploadErrors++
    }
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

# Sort out creds source. If user supplied -creds command line option, don't popup a prompt for creds.
if ($creds) {
    $promptForCreds = $false
    $credsFile = $creds
}
# Check that photoSource is valid
if ($supportedPhotoSources -contains $photoSource) {
    Write-Host "photoSource: $photoSource"
} else {
    Write-Host "$photoSource is not a supported photoSource. Available values: $supportedPhotoSources"
    Stop-Transcript
    Exit-Script 2
}

# Force the use of TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
# Test JSON capability
$jsonOk = Test-Json

# Connect to OfficeSpace to get count of existing records.
Write-Host "Communicating with OfficeSpace to get count of existing employee records..."
$r = (Get-WebJson -Url $ossGetEmployeesUrl -Headers $ossHeaders)
$ossCount = $r.count
$arrayCount = $r.response.Count
Write-Host "debug: ossCount: $ossCount, arrayCount: $arrayCount"

if ($ossCount -ne $arrayCount) {
    Write-Host "Count mismatch when querying OfficeSpace! Exiting."
    Stop-Transcript
    Exit-Script 2
}

Write-Host "$ossCount total existing OfficeSpace records"

# Get all the active employees from OfficeSpace
$employees = $r.response
# employeesMap = hash where user email address is the key and OSS user record is value
$employeesMap = @{}
ForEach($employee in $employees) {
    if ( -not ([string]::IsNullOrEmpty($employee.email)) ) { #in case this person from OfficeSpace doesn't have an email
        if (-not $employeesMap.ContainsKey($employee.email)) { #in case there is more than one person with the same email
            $employeesMap.add($employee.email, $employee);
        }
    }
}
Write-Output "$($employees.Count) OfficeSpace records added to employeesMap."


# AzureAD
if (Get-Module -ListAvailable -Name AzureAD) {
    Import-Module AzureAD
} else {
    Write-Host "The 'AzureAD' module may not be installed. Check and run as an Administrator 'Install-Module AzureAD'"
    Exit-Script 2
}
$azureadModuleInfo = Get-Module AzureAD
Write-Host "AzureAD module version: $($azureadModuleInfo.Version)"
if ($promptForCreds) {
    Write-Host "(will prompt for Azure AD user credentials)"
    $azureCredential = Get-Credential
} else {
    $azurePassword = gc $credsFile
    $azurePassword = ConvertTo-SecureString $azurePassword -Force
    $azureCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $azureUsername,$azurePassword
}

# Connect to AzureAD using supplied credentials
Write-Host "Connecting to AzureAD..."
Try {
    Connect-AzureAD -Credential $azureCredential | Out-Null
} Catch {
    Write-Host "Problem connecting to Azure AD. Exiting..."
    Stop-Transcript
    Exit-Script 2
}

Write-Host "Getting Azure AD users..."
# Get an array of Azure AD User objects (filtering out disabled accounts)
$azureObjects = Get-AzureADUser -All $true -Filter "AccountEnabled eq true"
$azureObjectsCount = $azureObjects.Count
Write-Host "$azureObjectsCount AzureAD objects returned"
$result = Get-UserObjects -azureObjects $azureObjects
Write-Host "Ensuring uniqueness..."
$azureUsers = $result | Sort-Object UserPrincipalName -Unique
$azureUserCount = $azureUsers.Count
Write-Host "$azureUserCount unique Azure AD users found (from $($result.count) users)"
# Exit script if no user objects were returned
if ($azureUsers.Count -eq 0) {
    Stop-Transcript
    Exit-Script 1
}


# Connect to Exchange if we are sourcing photos from there
if ($photoSource.Contains('exchange')) {
    Write-Host "Connecting to Exchange..."
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $azureCredential -Authentication Basic -AllowRedirection
    Import-PSSession $exchangeSession -DisableNameChecking | Out-Null
    Write-Host "Getting user mailboxes..."
    # Get an array of Exchange user mailboxes
    $userMailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | select UserPrincipalName,HasPicture
    Write-Host "$($userMailboxes.Count) Exchange user mailboxes returned"
}


# Create photos directory if needed.
if ($photoSource.Contains('azuread') -and !(Test-Path $photosDir)) {
    New-Item -Path $photosDir -ItemType Directory | Out-Null
    Write-Host "Created photosDir: $photosDir"
}

# Initialize variables.
$photosUploaded = 0
$uploadErrors = 0
$photosMatchOss = 0
$photosFound = 0
$photosAzureAD = 0
$photosExchange = 0
$exchangePhotoNoImage = 0
$noPictureSetOnMailbox = 0
$usersNotFoundInEmployeesMap = @()
$noThumbnailPhotoInAzureAD = 0
$exceptionsCheckingAzureThumbnailPhoto = 0
$userCounter = 0
# for each record from AzureAD
#  compare AzureAD.md5 with OSS.md5 image_source_fingerprint
#  if they are different then update OSS
foreach ($azureUser in $azureUsers) {
    $userCounter++
    $checkAzureForPhoto = $false
    $azureUPN = $azureUser.UserPrincipalName
    write-Host "-> Checking $azureUPN [$userCounter/$azureUserCount]"
    # Check if this employee is in OfficeSpace.
    if ($employeesMap.ContainsKey($azureUPN))
    {
        write-host "   found in employeesMap"
        # Get the OfficeSpace employee object from the map.
        $employee = $employeesMap.get_item($azureUPN)
        # If the photoSource includes 'exchange', first check if they have an Exchange mailbox.
        if ($photoSource.Contains('exchange')) {
            if ($userMailboxes.UserPrincipalname -contains $azureUPN) {
                Write-Host "   has an Exchange Mailbox"
                if ($userMailboxes | where UserPrincipalName -eq $azureUPN | % { $_.HasPicture }) {
                    $photo = Get-UserPhoto -Identity $azureUPN
                    $imageDataRaw = ""
                    $imageDataRaw = $photo.PictureData
                    if ($imageDataRaw -ne "") {
                        $photosFound++
                        $photosExchange++
                        Write-Host "   has photo [exch]"
                        $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
                        $photoMd5 = [System.BitConverter]::ToString($md5.ComputeHash($imageDataRaw))
                        $photoMd5 = $photoMd5 -replace '-',''
                        $photoMd5 = $photoMd5.ToLower()
                        $imageDataBase64 = [System.Convert]::ToBase64String($imageDataRaw)
                        #Write-Host "debug: [exch] photo=$photoMd5 image_source_fingerprint=$($employee.image_source_fingerprint)"
                        if ($photoMd5 -ine $employee.image_source_fingerprint) {
                            Upload-Photo -imageDataRaw $imageDataRaw -ossUser $employee -ossUrl $ossGetEmployeesUrl -Headers $ossHeaders
                        } else {
                            $photosMatchOss++
                        }
                        # We found a valid photo image for an Exchange mailbox, so now we can continue to check the next user.
                        continue
                    } else {
                        $exchangePhotoNoImage++
                        Write-Host "    user HasPicture, but no image data found!"
                        if ($photoSource -eq "exchange-azuread") {
                            $checkAzureForPhoto = $true
                        }
                    }
                } else {
                    $noPictureSetOnMailbox++
                    Write-Host "   has Exchange mailbox but does not have picture set"
                    if ($photoSource -eq "exchange-azuread") {
                        $checkAzureForPhoto = $true
                    }
                }
            } else {
                Write-Host "   user does not have an Exchange mailbox"
            }
        } 
        if ($photoSource.Contains("azuread")) {
            $checkAzureForPhoto = $true
        }

        if ($photoSource -eq "azuread" -or $checkAzureForPhoto) {
            # Get the image data from AzureAD.
            $imageDataRaw = ""
            Try {
                $photoFile = Join-Path -Path $photosDir -ChildPath $azureUPN
                Get-AzureADUserThumbnailPhoto -ObjectId $azureUser.ObjectId -FileName $photoFile
                $photosFound++
                $photosAzureAD++
                Write-Host "   has photo [aad]"
                # Get-AzureADUserThumbnailPhoto download appends '.jpeg' to the filename.
                $photoFile += '.jpeg'
                $photoMd5 = Get-FileHash -Path $($photoFile) -Algorithm MD5
                $photoMd5.Hash = $photoMd5.Hash.ToLower()
                $imageDataRaw = Get-Content -Raw $photoFile -Encoding byte
                #Write-Host "debug: [aad] photo=$($photoMd5.Hash) image_source_fingerprint=$($employee.image_source_fingerprint)"
                if ($photoMd5.Hash -ine $employee.image_source_fingerprint) {
                    Upload-Photo -imageDataRaw $imageDataRaw -ossUser $employee -ossUrl $ossGetEmployeesUrl -Headers $ossHeaders
                } else {
                    $photosMatchOss++
                }
            } Catch [Microsoft.Open.AzureAD16.Client.ApiException] {
                # No photo in AzureAD.
                Write-Host "   no ThumbnailPhoto in AzureAD"
                $noThumbnailPhotoInAzureAD++
                continue
            } Catch {
                Write-Host "**** General exception caught: $_.Exception.Message ****"
                $exceptionsCheckingAzureThumbnailPhoto++
                continue
            }
        }
    } #end of if employee is in OfficeSpace
    else {
        Write-Host "   user not found in employeesMap"
        $usersNotFoundInEmployeesMap += $azureUPN
    }
}

# Stats.
Write-Host
Write-Host "User records found in OSS             : $($employees.Count)"
Write-Host "Photos found                          : $photosFound"
Write-Host "Photo matched OSS version             : $photosMatchOss"
Write-Host "Photos uploaded                       : $photosUploaded"
Write-Host "Upload errors                         : $uploadErrors"
Write-Host "Azure AD users checked                : $azureUserCount"
Write-Host "usersNotFoundInEmployeesMap           : $($usersNotFoundInEmployeesMap.Count)"
if ($photoSource.Contains('exchange')) {
    Write-Host "photosExchange                        : $photosExchange"
    Write-Host "Exchange Mailboxes checked            : $($userMailboxes.Count)"
    Write-Host "noPictureSetOnMailbox                 : $noPictureSetOnMailbox"
    Write-Host "exchangePhotoNoImage                  : $exchangePhotoNoImage"
}
if ($photoSource.Contains('azuread')) {
    Write-Host "photosAzureAD                         : $photosAzureAD"
    Write-Host "noThumbnailPhotoInAzureAD             : $noThumbnailPhotoInAzureAD"
    Write-Host "exceptionsCheckingAzureThumbnailPhoto : $exceptionsCheckingAzureThumbnailPhoto"
}

# Timing.
$endTime = Get-Date
$elapsedTime = $endTime - $startTime
Write-Host "Completed in $($elapsedTime.TotalSeconds) seconds"

Exit-Script 0

