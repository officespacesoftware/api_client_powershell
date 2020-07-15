###############
# User config
###############
# According to Zenefits (https://developers.zenefits.com/docs/pagination), their service will return 20 records at a time, by default.
# You can control this by adding the parameter "limit=xx". The maximum value allowed is 100.
# Example: Return 50 people records at a time: https://api.zenefits.com/core/people?limit=50

$onlyActive                 = $true                                                                         # When true, only request employees with the "active" status.
                                                                                                            #   If false, we validate results against list of desired statuses.
if($onlyActive){
    $zenefitsPeopleUrl      = "https://api.zenefits.com/core/people?includes=department&status=active"      # Zenefits URL to access people data and department info
} else {
    $zenefitsPeopleUrl      = "https://api.zenefits.com/core/people?includes=department"                    # Zenefits URL to access people data and department info
}

$zenefitsPeopleUrl      = "https://api.zenefits.com/core/people?includes=department"    # Zenefits URL to access people data and department info
$zenefitsApiKey         = "zenefitsApiKey123456"                                        # Zenefits API key
$logFile                = "C:\ps\oss_import.txt"                                        # log file
$photoSource            = "zenefits"                                                    # source for photos: zenefits, none  [none=don't use zenefits for user photos]
$photosDir              = "C:\ps\photos2"                                               # directory to download thumbnail photos
$ossApiKey              = "0123456789abcdef1011121314151617"                            # OfficeSpace API key
$ossHostname            = "mycompany.officespacesoftware.com"                           # OfficeSpace instance hostname
$tryNicknames           = $false                                                        # when $true, use preferred name over first name
$employmentStatus       = @{active = "active"}                                          # Hash table with the accepted employment statuses. The name should match a valid Zenefits status,
                                                                                        #   per their API documentation: active, terminated, leave_of_absence, requested, setup, deleted.
                                                                                        #   Example: $employmentStatus     = @{active = "active"; terminated = "terminated"}
$importThreshold        = 60                                                            # minimum percentage of import count compared to existing OSS record count
                                                                                        #   If the user count to import is less than this threshold %, don't import.

######################
$source                     = "Zenefits"                                                # Source of user data for OSS import
$batchSize                  = 300
$supportedPhotoSources      = @( 'none', 'zenefits' )
$ossProtocol                = "https://"
$ossHeaders                 = @{Authorization = "Token token=" + $ossApiKey}
$ossBatchUrl                = "/api/1/employee_batch_imports"
$ossImportUrl               = "/api/1/employee_directory"
$ossEmployeesUrl            = "/api/1/employees"
$ossGetEmployeesUrl         = $ossProtocol + $ossHostname + $ossEmployeesUrl
$ossEmployeeBatchUrl        = $ossProtocol + $ossHostname + $ossBatchUrl
$ossEmployeeBatchStagingUrl = $ossProtocol + $ossHostname + $ossImportUrl + "/" + $source
$ossEmployeeImportUrl       = $ossProtocol + $ossHostname + $ossImportUrl
$zenefitsHeaders            = @{Authorization = "Bearer " + $zenefitsApiKey}
$version                    = 2
######################

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
               Write-Output "    (no givenName, so took from displayName)"
           }
           return $displayName.Split()[0]
       } elseif ($displayName.IndexOf($givenName) -eq -1) {
           if ($objType -eq 'user') {
               $script:nicknamesAssigned++
               Write-Output "    (givenName not found in displayName, so took from displayName)"
           }
           return $displayName.Split()[0]
       } else {
           # (givenName found in displayName, so no nickname)
           return $givenName
       }
    } else {
        Write-Output "    (single word displayName)"
        return $givenName
    }
}

# Test JSON data set capability
function Test-Json {
    Write-Host -NoNewLine "Test JSON parsing performance: "
    Try {
        "a" * 2mb | ConvertTo-Json | ConvertFrom-Json | Out-Null
        Write-Output "PASS"
        return $true
    } catch [System.ArgumentException] {
        Write-Output "FAIL"
        Write-Output "NOTICE: Your PowerShell environment may not support parsing larger JSON data sets."
        return $false
    } catch {
        Write-Output "Test-Json() General exception caught: $_.Exception.Message"
        Stop-Transcript
        Exit 2
    }
}

######################

# Capture the start time of script
$startTime = Get-Date
Write-Output "Script version $version start"

# Force the use of TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# Zenefits User object fields to collect
$zenefitsUserFields = @(
    'id',
    'first_name',
    'last_name',
    'personal_phone',
    'title',
    'work_email',
    'work_phone'
)

# Zenefits Manager fields to collect
$zenefitsManagerFields = @(
    'last_name',
    'work_email'
)

########################################################
# OfficeSpace <-> Zenefits attribute attribute mapping
########################################################
$EmployeeId = "id"                   # Employee ID number
$FirstName = "pref_name"
$LastName = "last_name"
$Title = "title"
$WorkPhone = "work_phone"
$Extension = ""
$ImageData = "ImageData"
$Department = "department"
$Bio = ""
$Email = "work_email"
$StartDate = ""
$EndDate = ""
$ShowInVd = "" 
$Udf0 = "personal_phone"             # User's personal phone number
$Udf1 = "manager_work_email"         # Manager's email
$Udf2 = "manager_pref_name"          # Manager's Given/first name. ManagerPrefName can update if $tryNicknames.
$Udf3 = "manager_last_name"          # Manager's surname/last name
$Udf4 = "PhotoMd5"
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
Start-Transcript -Path "$logFile"
if ($supportedPhotoSources -contains $photoSource) {
    Write-Output "photoSource: $photoSource"
} else {
    Write-Output "$photoSource is not a supported photoSource. Available values: $supportedPhotoSources"
    Stop-Transcript
    Exit 2
}
Write-Output "tryNicknames: $tryNicknames"
$jsonOk = Test-Json

# Connect to OfficSpace to get count of existing records.
Write-Output "Communicating with OfficeSpace to get count of existing employee records..."
if ($jsonOk) {
    $r = (Invoke-RestMethod -Uri $ossGetEmployeesUrl -Method Get -Headers $ossHeaders)
    $ossCount = $r.Count
    $arrayCount = $r.Response.Count
} else {
    Write-Output "Using Invoke-WebRequest..."
    $w = (Invoke-WebRequest -Uri $ossGetEmployeesUrl -UseBasicParsing -Method Get -Headers $ossHeaders)
    $rawContentLength = $w.RawContentLength
    Write-Output "debug: raw content length: $rawContentLength"
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
    $jsonSerial = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer
    # If you set MaxJsonLength to $rawContentLength - 1, you'll hit the error.
    $jsonSerial.MaxJsonLength  = $rawContentLength 
    $r = $jsonSerial.DeserializeObject($w.Content)
    $ossCount = $r.count
    $arrayCount = $r.response.count
}
Write-Output "debug: ossCount: $ossCount  arrayCount: $arrayCount"

if ($ossCount -ne $arrayCount) {
    Write-Output "Count mismatch when querying OfficeSpace! Exiting."
    Stop-Transcript
    Exit 2
}
$ossCurrentUserCount = $ossCount
Write-Output "$ossCurrentUserCount existing OfficeSpace records"

# Initialize arrays to store the Zenefits users and departments
$zenefitsUsers = @()
$zenefitsDepts = @()
# Communicate with Zenefits
Write-Output "Connecting to Zenefits..."
# Get Users
$getZenefitsUsers = $true
$zenefitsUrl = $zenefitsPeopleUrl
while ($getZenefitsUsers) {
    Try {
        # To see the URL we are fetching: Write-Host "debug: zenefitsUrl = $zenefitsUrl"
        $z = (Invoke-RestMethod -Uri $zenefitsUrl -Method Get -Headers $zenefitsHeaders)
        $zenefitsUsers += $z.data.data
        # if 'next_url' contains a value, there's more data available.
        if ($z.data.next_url) {
            $zenefitsUrl = $z.data.next_url
        } else {
            $getZenefitsUsers = $false
        }
    } Catch {
        Write-Output "Problem connecting to Zenefits. Exiting..."
        Stop-Transcript
        Exit 2
    }
}
$zenefitsUserCount = $zenefitsUsers.Count
Write-Output "$zenefitsUserCount Zenefits users returned"
# Exit script if no users were returned
if ($zenefitsUsers.Count -eq 0) {
    Stop-Transcript
    Exit
}

# Create photos directory if needed
if ($photoSource.Contains('zenefits') -and !(Test-Path $photosDir)) {
    New-Item -Path $photosDir -ItemType Directory | Out-Null
    Write-Output "Created photosDir: $photosDir"
}

# Build an array from the Zenefits Users that we can modify
$ossUsersArray = New-Object System.Collections.ArrayList

$mgrsFound         = 0
$photosFound       = 0
$nicknamesAssigned = 0
$usersSkipped      = @()
Write-Output "Inspecting users..."

for ($counter = 1; $counter -le $zenefitsUserCount; $counter++) {
    # OSS user hash table
    $ossUser = @{}
    $u = $zenefitsUsers[$counter - 1]
    $userId = $u.id
    Write-Output "-> ${userId}:$($u.first_name) $($u.last_name) [$counter/$zenefitsUserCount]"
    
    # Skip if user doesn't have a desired employment status and onlyActive is false
    if (-not $onlyActive) {
        if (-not $employmentStatus[$u.status]){
            Write-Output "    (filtered employment status: $($u.status), skipping)"
            $usersSkipped += $userId
            continue
        }
    }
    
    # Skip if user has no last name
    if ($u.last_name -eq $null) {
        Write-Output "    (no last name, skipping)"
        $usersSkipped += $userId
        continue
    }

    # Try to look for nickname
    if ($tryNicknames) {
        $ossUser.pref_name = Get-Nickname -givenName $u.first_name -displayName ($u.preferred_name + " " + $u.last_name) -objType "user"
    } else {
        $ossUser.Add("pref_name", $u.first_name)
    }
    # Skip if user has no pref name
    if ($ossUser.pref_name -eq $null -Or $ossUser.pref_name -eq "") {
        Write-Output "    (no first/pref name, skipping)"
        $usersSkipped += $userId
        continue
    }

    foreach ($i in $zenefitsUserFields) {
        $ossUser.Add($i, $u.$i)
    }
    
    # Fetch department info and add it to the user
    $ossUser.Add("department", $u.department.name)

    # Fetch manager info and add it to the user
    $mgr = $zenefitsUsers | Where-Object url -eq $u.manager.url
    if ($mgr -ne $null) {
        $mgrsFound++
        Write-Output "    (has manager)"
        # Check if we should use manager nickname
        if ($tryNicknames) {
            $ossUser.manager_pref_name = Get-Nickname -givenName $mgr.first_name -displayName ($mgr.preferred_name + " " + $mgr.last_name) -objType "manager"
        } else {
            $ossUser.Add("manager_pref_name", $mgr.first_name)
        }
    } 
    foreach ($i in $zenefitsManagerFields) {
        $ossUser.Add("manager_" + $i, $mgr.$i)
    }

    # Fetch thumbnail photo and add it to the OSS user
    $ossUser.Add("PhotoMd5", $null)
    $ossUser.Add("ImageData", $null)
    $imageDataRaw = $null
    $imageDataBase64 = $null    
    if ($photoSource.Contains('zenefits')) {
        if ($u.photo_thumbnail_url) {
             $photosFound++
             Write-Host "    (has photo [$($u.photo_thumbnail_url)])"
             $photoFile = $photosDir + '\' + $userId + '.png'
             Invoke-WebRequest $u.photo_thumbnail_url -OutFile $photoFile
             $photoMd5 = Get-FileHash -Path $($photoFile) -Algorithm MD5
             # Subtle difference if using PSCore
             if ($PSEdition -eq 'Core') {
                 $imageDataRaw = Get-Content -Raw $photoFile -AsByteStream
             } else {
                 $imageDataRaw = Get-Content -Raw $photoFile -Encoding Byte
             }
             $imageDataBase64 = [System.Convert]::ToBase64String($imageDataRaw)
             $ossUser.Set_Item("PhotoMd5", $photoMd5.Hash)
             $ossUser.Set_Item("ImageData", $imageDataBase64)
        }
    }

    # Add the OSS user hash table to the OSS users array
    $ossUsersArray.Add($ossUser) | Out-Null
}


# Stats
Write-Output "$zenefitsUserCount users inspected"
Write-Output "$mgrsFound manager assignments found"
if ($photoSource -ne "none") {
    Write-Output "$photosFound user photos found"
}
if ($tryNicknames) {
    Write-Output "$nicknamesAssigned nicknames assigned to users"
}
$numUsersSkipped = $usersSkipped.Count
Write-Output "$numUsersSkipped users skipped"
if ($numUsersSkipped -gt 0) {
    Write-Output "User ids skipped: $usersSkipped"
}
$userCount = $ossUsersArray.Count
Write-Output "$userCount users to import to OfficeSpace"
# Import only if we meet the import threshold %
if ($ossCurrentUserCount -gt 0) {
    $importPercentage = [math]::round(($userCount / $ossCurrentUserCount) * 100, 2)
    Write-Output "importPercentage=$importPercentage, importThreshold=$importThreshold"
    if ($importPercentage -lt $importThreshold) {
        Write-Output "The number of users to import is too low to perform the import. Exiting."
        Stop-Transcript
        Exit 1
    }
}

# Setup batching
$totalBatches = [math]::Floor($userCount / $batchSize) + 1
$startIdx     = 0
$endIdx       = $batchSize - 1
$currentBatch = 1

# Start communicating with OfficeSpace
Write-Output "Preparing OfficeSpace for data push..."
Invoke-WebRequest -UseBasicParsing -Uri $ossEmployeeBatchStagingUrl -ContentType application/json -Method Delete -Headers $ossHeaders | Out-Null

Write-Output "Sending records to $ossEmployeeBatchUrl"
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
    Try {
        Invoke-WebRequest -UseBasicParsing -Uri $ossEmployeeBatchUrl -ContentType 'application/json; charset=utf-8' -Method Post -Body $JSONArrayUTF8 -Headers $ossHeaders -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Out-Null
        $startNum = $startIdx + 1
        $endNum = $endIdx + 1
        Write-Output "$startNum-$endNum " -NoNewline
        Write-Output "Done" -ForegroundColor Green
    } Catch {
        $startNum = $startIdx + 1
        $endNum = $endIdx + 1
        Write-Output "$startNum-$endNum " -NoNewline
        Write-Output $_.Exception.Message -ForegroundColor Red
    }
    $startIdx += $batchSize
    $endIdx += $batchSize
    $currentBatch++
} while ($currentBatch -le $totalBatches)

Write-Output "Triggering migration..."
$ossImportUrlPostBody = "Source=" + $source
Invoke-WebRequest -UseBasicParsing -Uri $ossEmployeeImportUrl -Method Post -Body $ossImportUrlPostBody -Headers $ossHeaders | Out-Null

# Timing
$endTime = Get-Date
$elapsedTime = $endTime - $startTime
Write-Output "Completed in $($elapsedTime.TotalSeconds) seconds"
# Stop logging
Stop-Transcript

