#
# push_from_o365_to_officespace.ps1 - import users from Exchange Online, Azure AD to OfficeSpace
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
                                                           # If -creds command line option is provided, $promptForCreds is ignored.
$azureUsername   = "oss_service@mycompany.com"             # user with Azure AD access
$credsFile       = "C:\ps\oss_service.creds"               # encrypted user credentials file

$useLogFile      = $true                                   # Log output to file ($true=log, $false=do not log)
$logFile         = "$PSScriptRoot\oss_import.txt"          # Path to log file
$photoSource     = "exchange-azuread"                      # source for photos: azuread, exchange, exchange-azuread, none
$photosDir       = "C:\ps\photos"                          # directory to download thumbnail photos
                                                           #    (only used if $photoSource = 'azuread' or 'exchange-azuread')
$ossToken        = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"      # OfficeSpace API key
$ossHostname     = "mycompany.officespacesoftware.com"     # OfficeSpace instance hostname
$tryNicknames    = $false                                  # look at displayName for possible preferred name
$importThreshold = 60                                      # minimum percentage of import count compared to existing OSS record count
                                                           #   If the user count to import is less than this threshold %, don't import.
###########

$source                     = "AzureAD"
$batchSize                  = 300
$supportedPhotoSources      = @( 'azuread', 'exchange', 'exchange-azuread', 'none' )
$ossProtocol                = "https://"
$ossHeaders                 = @{Authorization = "Token token=" + $ossToken}
$ossBatchUrl                = "/api/1/employee_batch_imports"
$ossImportUrl               = "/api/1/employee_directory"
$ossEmployeesUrl            = "/api/1/employees"
$ossGetEmployeesUrl         = $ossProtocol + $ossHostname + $ossEmployeesUrl
$ossEmployeeBatchUrl        = $ossProtocol + $ossHostname + $ossBatchUrl
$ossEmployeeBatchStagingUrl = $ossProtocol + $ossHostname + $ossImportUrl + "/" + $source
$ossEmployeeImportUrl       = $ossProtocol + $ossHostname + $ossImportUrl
$version                    = 4

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


# Get nickname from display name
function Get-Nickname {
   param ($givenName, $displayName, $objType)
   if ($displayName -eq $null -Or $displayName -eq "") {
       # (no displayName)
       return $givenName
   }
   if ($displayName.Split().Count -ge 2) {
       if ($givenName -eq $null -Or $givenName -eq "") {
           if ($objType -eq 'user') {
               $script:nicknamesAssigned++
               Write-Host "    (no givenName, so took from displayName)"
           }
           return $displayName.Split()[0]
       } elseif ($displayName.IndexOf($givenName) -eq -1) {
           if ($objType -eq 'user') {
               $script:nicknamesAssigned++
               Write-Host "    (givenName not found in displayName, so took from displayName)"
           }
           return $displayName.Split()[0]
       } else {
           # (givenName found in displayName, so no nickname)
           return $givenName
       }
    } else {
        Write-Host "    (single word displayName)"
        return $givenName
    }
}


# Get the Azure users
function Get-UserObjects {
    param ($azureObjects)
    $users = @()
    foreach ($o in $azureObjects) {
        Write-Host "-> $($o.DisplayName)  [$($o.ObjectType)]"
        if ($o.ObjectType -eq 'User') {
            # Collect the user object
            $users += $o
        } elseif ($o.ObjectType -eq 'Group') {
            Write-Host "---> will parse this group"
            $gm = Get-AzureADGroupMember -All $true -ObjectId $o.ObjectId
            $users += Get-UserObjects -azureObjects $gm
        } else {
            Write-Host "---> skip this object"
        }
    }
    return $users
}


########################
# Script execution start
########################
# Capture the start time of script
$startTime = Get-Date

# Azure AD User object fields to collect
$azureUserFields = @(
    'Department',
    'DisplayName',
    'GivenName',
    'JobTitle',
    'Mobile',
    'Surname',
    'TelephoneNumber',
    'UserPrincipalName'
)

# Azure AD Manager fields to collect
$azureManagerFields = @(
    'DisplayName',
    'GivenName',
    'Surname',
    'UserPrincipalName'
)

######################################################
# OfficeSpace <-> Azure AD attribute attribute mapping
######################################################
$EmployeeId = "UserPrincipalName"  # UPN is unique and easy to recongize. There is also a unique objectId.
$FirstName = "PrefName"            # PrefName = GivenName. PrefName can update if $tryNicknames.
$LastName = "Surname"
$Title = "JobTitle"
$WorkPhone = "TelephoneNumber"
$Extension = ""
$ImageData = "ImageData"
$Department = "Department"
$Bio = ""
$Email = "UserPrincipalName"
$StartDate = ""
$EndDate = ""
$ShowInVd = ""
$Udf0 = "Mobile"                   # User's mobile number
$Udf1 = "ManagerUserPrincipalName" # Manager's email
$Udf2 = "ManagerPrefName"          # Manager's Given/first name. ManagerPrefName can update if $tryNicknames.
$Udf3 = "ManagerSurname"           # Manager's surname/last name
$Udf4 = ""
$Udf5 = ""
$Udf6 = ""
$Udf7 = ""
$Udf8 = ""
$Udf9 = ""
$Udf10 = ""
$Udf11 = ""
$Udf12 = ""
$Udf13 = ""
$Udf14 = ""
$Udf15 = ""
$Udf16 = ""
$Udf17 = ""
$Udf18 = ""
$Udf19 = ""
$Udf20 = ""
$Udf21 = ""
$Udf22 = ""
$Udf23 = ""
$Udf24 = ""
######################

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
    Exit-Script 2
}
Write-Host "tryNicknames: $tryNicknames"


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
    Exit-Script 2
}
$ossCurrentUserCount = $ossCount
Write-Host "$ossCurrentUserCount existing OfficeSpace records"

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
    Exit-Script 2
}

Write-Host "Getting Azure AD users..."
# Get an array of Azure AD User objects (filtering out disabled accounts)
$azureObjects = Get-AzureADUser -All $true -Filter "AccountEnabled eq true"
$azureObjectsCount = $azureObjects.count
Write-Host "$azureObjectsCount AzureAD objects returned"
$result = Get-UserObjects -azureObjects $azureObjects
Write-Host "Ensuring uniqueness..."
$azureUsers = $result | Sort-Object UserPrincipalName -Unique
$azureUserCount = $azureUsers.Count
Write-Host "$azureUserCount unique Azure AD users found (from $($result.count) users)"
# Exit script if no user objects were returned
if ($azureUsers.Count -eq 0) {
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
    $exchangePhoto = @{}
    # Store whether there is a picture associated with the mailbox
    foreach ($userMailbox in $userMailboxes) {
        $exchangePhoto.Add($userMailbox.UserPrincipalName, $userMailbox.HasPicture)
    }
}

# Create photos directory if needed
if ($photoSource.Contains('azuread') -and !(Test-Path $photosDir)) {
    New-Item -Path $photosDir -ItemType Directory | Out-Null
    Write-Host "Created photosDir: $photosDir"
}

# Build an array from the Azure AD Users that we can modify
$ossUsersArray = New-Object System.Collections.ArrayList

$mgrsFound         = 0
$photosFound       = 0
$photosAzureAD     = 0
$photosExchange    = 0
$nicknamesAssigned = 0
$usersSkipped      = @()
Write-Host "Inspecting users..."
for ($counter = 1; $counter -le $azureUserCount; $counter++) {
    # OSS user hash table
    $ossUser = @{}
    $u = $azureUsers[$counter - 1]
    $userId = $u.UserPrincipalName
    Write-Host "-> $userId  [$counter/$azureUserCount]"

    # Skip if user has no last name
    if ($u.Surname -eq $null) {
        Write-Host "    (no last name, skipping)"
        $usersSkipped += $userId
        continue
    }

    # Try to look for nickname
    if ($tryNicknames) {
        $ossUser.PrefName = Get-Nickname -givenName $u.GivenName -displayName $u.DisplayName -objType "user"
    } else {
        $ossUser.Add("PrefName", $u.GivenName)
    }
    # Skip if user has no PrefName
    if ($ossUser.PrefName -eq $null -Or $ossUser.PrefName -eq "") {
        Write-Host "    (no first/pref name, skipping)"
        $userSkipped += $userId
        continue
    }

    foreach ($i in $azureUserFields) {
        $ossUser.Add($i, $u.$i)
    }

    # Fetch manager info and add it to the user
    $mgr = Get-AzureADUserManager -ObjectId $userId
    if ($mgr -ne $null) {
        $mgrsFound++
        Write-Host "    (has manager)"
        # Check if we should use manager nickname
        if ($tryNicknames) {
            $ossUser.ManagerPrefName = Get-Nickname -givenName $mgr.GivenName -displayName $mgr.DisplayName -objType "manager"
        } else {
            $ossUser.Add("ManagerPrefName", $mgr.GivenName)
        }
    }
    foreach ($i in $azureManagerFields) {
        $ossUser.Add("Manager" + $i, $mgr.$i)
    }

    # Fetch thumbnail photo and add it to the OSS user
    $ossUser.Add("PhotoMd5", $null)
    $ossUser.Add("ImageData", $null)
    $checkAzureForPhoto = $false
    if ($photoSource.Contains('exchange')) {
        if ($exchangePhoto.$userId -eq $true) {
            Try {
                 $photo = $null
                 $photo = Get-UserPhoto -Identity $userId
                 $imageDataRaw = $null
                 $imageDataRaw = $photo.PictureData
                 if ($imageDataRaw -ne "") {
                     $photosFound++
                     $photosExchange++
                     Write-Host "    (has photo [exch])"
                     $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
                     $md5Hash = $null
                     $md5Hash = [System.BitConverter]::ToString($md5.ComputeHash($imageDataRaw))
                     $md5Hash = $md5Hash -replace '-',''
                     $imageDataBase64 = $null
                     $imageDataBase64 = [System.Convert]::ToBase64String($imageDataRaw)
                     $ossUser.Set_Item("PhotoMd5", $md5Hash)
                     $ossUser.Set_Item("ImageData", $imageDataBase64)
                 }
             } Catch {
                Write-Host "    (couldn't retrieve photo [exch])"
            }
        }
        if ($ossUser.ImageData -eq $null -and $photoSource -eq "exchange-azuread") {
            $checkAzureForPhoto = $true
        }
    }
    if ($photoSource -eq "azuread" -or $checkAzureForPhoto) {
        Try {
            $photoFile = Join-Path -Path $photosDir -ChildPath $userId
            Get-AzureADUserThumbnailPhoto -ObjectId $userId -FileName $photoFile
            $photosFound++
            $photosAzureAD++
            Write-Host "    (has photo [aad])"
            # Get-AzureADUserThumbnailPhoto download appends '.jpeg' to the filename
            $photoFile += '.jpeg'
            $photoMd5 = Get-FileHash -Path $($photoFile) -Algorithm MD5
            $imageDataRaw = Get-Content -Raw $photoFile -Encoding byte
            $imageDataBase64 = [System.Convert]::ToBase64String($imageDataRaw)
            $ossUser.Set_Item("PhotoMd5", $photoMd5.Hash)
            $ossUser.Set_Item("ImageData", $imageDataBase64)
        } Catch {
            # No photo in AzureAD
        }
    }

    # Add the OSS user hash table to the OSS users array
    $ossUsersArray.Add($ossUser) | Out-Null
}

# Stats
Write-Host "$azureUserCount users inspected"
Write-Host "$mgrsFound manager assignments found"
if ($photoSource -ne "none") {
    Write-Host -NoNewLine "$photosFound user photos found"
    if ($photoSource -eq "exchange-azuread") {
        Write-Host " ($photosExchange exch, $photosAzureAD aad)"
    } else {
        Write-Host
    }
}
if ($tryNicknames) {
    Write-Host "$nicknamesAssigned nicknames assigned to users"
}
$numUsersSkipped = $usersSkipped.Count
Write-Host "$numUsersSkipped users skipped"
if ($numUsersSkipped -gt 0) {
    Write-Host "Users skipped: $usersSkipped"
}
$userCount = $ossUsersArray.Count
Write-Host "$userCount users to import to OfficeSpace"
# Import only if we meet the import threshold %
if ($ossCurrentUserCount -gt 0) {
    $importPercentage = [math]::round(($userCount / $ossCurrentUserCount) * 100, 2)
    Write-Host "importPercentage=$importPercentage, importThreshold=$importThreshold"
    if ($importPercentage -lt $importThreshold) {
        Write-Host "The number of users to import is too low to perform the import. Exiting."
        Exit-Script 1
    }
}


# Setup batching
$totalBatches = [math]::Ceiling($userCount / $batchSize)
$startIdx     = 0
$endIdx       = $batchSize - 1
$currentBatch = 1

# Start communicating with OfficeSpace
Write-Host "Preparing OfficeSpace for data push (staging records) ..." -NoNewline
# Check response; exit on error.
try {
    $resp = Invoke-WebRequest -UseBasicParsing -Uri $ossEmployeeBatchStagingUrl -ContentType application/json -Method Delete -Headers $ossHeaders
    Write-Host "  [response: $($resp.StatusCode)/$($resp.StatusDescription)]"
    if ($resp.StatusCode -ne 200) {
        Write-Host "Exiting due to non-200 status code [$($resp.StatusCode)/$($resp.StatusDescription)] seen when attempting staging records to $ossEmployeeBatchStagingUrl" -ForegroundColor Red
        Exit-Script 2
    }
} catch [System.Net.WebException] {
    $e = $_
    Write-Host "Exception caught while attempting staging records to $ossEmployeeBatchStagingUrl : $($e.Exception)" -ForegroundColor Red
    Exit-Script 2
} catch {
    Write-Host "General Exception caught while attempting staging records to $ossEmployeeBatchStagingUrl : $_" -ForegroundColor Red
    Exit-Script 2
}

Write-Host "Sending records to $ossEmployeeBatchUrl"
do {
    if ($endIdx -ge $userCount) {
        $endIdx = $userCount - 1
    }
    $ossImportBatch = New-Object System.Collections.ArrayList
    foreach ($user in $ossUsersArray[$startIdx..$endIdx]) {
        $ossImportBatch.Add([PSCustomObject]@{
            "EmployeeId" = $user.$EmployeeId
            "Source"     = $source
            "FirstName"  = $user.$FirstName
            "LastName"   = $user.$LastName
            "Title"      = $user.$Title
            "WorkPhone"  = $user.$WorkPhone
            "Extension"  = $user.$Extension
            "ImageData"  = $user.$ImageData
            "Department" = $user.$Department
            "Bio"        = $user.$Bio
            "Email"      = $user.$Email
            "StartDate"  = $user.$StartDate
            "EndDate"    = $user.$EndDate
            "ShowInVd"   = $user.$ShowInVd
            "Udf0"       = $user.$Udf0
            "Udf1"       = $user.$Udf1
            "Udf2"       = $user.$Udf2
            "Udf3"       = $user.$Udf3
            "Udf4"       = $user.$Udf4
            "Udf5"       = $user.$Udf5
            "Udf6"       = $user.$Udf6
            "Udf7"       = $user.$Udf7
            "Udf8"       = $user.$Udf8
            "Udf9"       = $user.$Udf9
            "Udf10"      = $user.$Udf10
            "Udf11"      = $user.$Udf11
            "Udf12"      = $user.$Udf12
            "Udf13"      = $user.$Udf13
            "Udf14"      = $user.$Udf14
            "Udf15"      = $user.$Udf15
            "Udf16"      = $user.$Udf16
            "Udf17"      = $user.$Udf17
            "Udf18"      = $user.$Udf18
            "Udf19"      = $user.$Udf19
            "Udf20"      = $user.$Udf20
            "Udf21"      = $user.$Udf21
            "Udf22"      = $user.$Udf22
            "Udf23"      = $user.$Udf23
            "Udf24"      = $user.$Udf24
        }) | Out-Null
    }

    $JSONArray = ConvertTo-Json -InputObject $ossImportBatch
    $JSONArrayUTF8 = [System.Text.Encoding]::UTF8.GetBytes($JSONArray)

    try {
        # Check response; exit on error
        $resp = Invoke-WebRequest -UseBasicParsing -Uri $ossEmployeeBatchUrl -ContentType 'application/json; charset=utf-8' -Method Post -Body $JSONArrayUTF8 -Headers $ossHeaders -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        if ($resp.StatusCode -ne 200) {
            Write-Host "Exiting due to non-200 status code [$($resp.StatusCode)/$($resp.StatusDescription)] seen when attempting sending records to $ossEmployeeBatchUrl" -ForegroundColor Red
            Exit-Script 2
        }
        $startNum = $startIdx + 1
        $endNum = $endIdx + 1
        Write-Host "$startNum-$endNum " -NoNewline
        Write-Host "Done" -ForegroundColor Green -NoNewline
        Write-Host "  [response: $($resp.StatusCode)/$($resp.StatusDescription)]"
    } catch [System.Net.WebException] {
        $e = $_
        Write-Host "Exception caught while attempting sending records to $ossEmployeeBatchUrl : $($e.Exception)" -ForegroundColor Red
        Exit-Script 2
    } catch {
        Write-Host "General Exception caught while attempting sending records to $ossEmployeeBatchUrl : $_" -ForegroundColor Red
        Exit-Script 2
    }

    $startIdx += $batchSize
    $endIdx += $batchSize
    $currentBatch++
} while ($currentBatch -le $totalBatches)

Write-Host "Triggering migration ..." -NoNewline
$ossImportUrlPostBody = "Source=" + $source
# Check response.
try {
    $resp = Invoke-WebRequest -UseBasicParsing -Uri $ossEmployeeImportUrl -Method Post -Body $ossImportUrlPostBody -Headers $ossHeaders
    Write-Host "  [response: $($resp.StatusCode)/$($resp.StatusDescription)]"
    if ($resp.StatusCode -ne 200) {
        Write-Host "Exiting due to non-200 status code [$($resp.StatusCode)/$($resp.StatusDescription)] seen when attempting migration to $ossEmployeeImportUrl" -ForegroundColor Red
        Exit-Script 2
    }
} catch [System.Net.WebException] {
    $e = $_
    Write-Host "Exception caught while attempting migration to $ossEmployeeImportUrl : $($e.Exception)" -ForegroundColor Red
    Exit-Script 2
} catch {
    Write-Host "General Exception caught while attempting migration to $ossEmployeeImportUrl : $_" -ForegroundColor Red
    Exit-Script 2
}


# Timing
$endTime = Get-Date
$elapsedTime = $endTime - $startTime
Write-Host "Completed in $($elapsedTime.TotalSeconds) seconds"

Exit-Script 0

