###############
# User config
###############
$ossHostname     = "mycompany.officespacesoftware.com"              # OfficeSpace instance hostname
$ossApiKey       = "0123456789abcdef1011121314151617"               # OfficeSpace API key
$dateTime        = Get-Date -UFormat "%Y%m%d_%H%m%S"                # Date/Time format that can be used in the output filename
$outputPath      = "$($env:userprofile)\desktop\"                   # Directory/path to write output file
$outputFile      = "OfficeSpace_Employee_Location_$($dateTime).csv" # Output filename
$outputDelimiter = ","                                              # Output field delimiter to use

######################
$ossProtocol                = "https://"
$ossHeaders                 = @{Authorization = "Token token=" + $ossApiKey}
$ossDirectoriesUrl          = "/api/1/directories"
$ossEmployeesUrl            = "/api/1/employees"
$ossFloorsUrl               = "/api/1/floors"
$ossSeatsUrl                = "/api/1/seats"
$ossSitesUrl                = "/api/1/sites"
$ossGetDirectoriesUrl       = $ossProtocol + $ossHostname + $ossDirectoriesUrl
$ossGetEmployeesUrl         = $ossProtocol + $ossHostname + $ossEmployeesUrl
$ossGetFloorsUrl            = $ossProtocol + $ossHostname + $ossFloorsUrl
$ossGetSeatsUrl             = $ossProtocol + $ossHostname + $ossSeatsUrl
$ossGetSitesUrl             = $ossProtocol + $ossHostname + $ossSitesUrl
######################

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
        Exit 2
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
        # If you set MaxJsonLength to $rawContentLength - 1, you'll hit the error.
        $jsonSerial.MaxJsonLength  = $rawContentLength
        $resp = $jsonSerial.DeserializeObject($w.Content)
    }
    return $resp
}

# Force the use of TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# Test JSON capability
$jsonOk = Test-Json

################################################################
# SYS: Make Public and SYS: Make Private are two special system
# directories that always exists and that describe the online and
# offline floor respectively.
$directoryOnlineName = "SYS: Make Public"
$directoryOfflineName = "SYS: Make Private"

# A string based set of directory names to filter by
$directoriesToInclude = New-Object System.Collections.Generic.HashSet[string]
$directoriesToInclude.Add($directoryOnlineName) | Out-Null
#$directoriesToInclude.Add($directoryOfflineName)
#$directoriesToInclude.Add("SOME OTHER DIRECTORY NAME")
#$directoriesToInclude.Add(...)
###############################################################

# A string based set of directory URLs to filter by
$directoryFilter = New-Object System.Collections.Generic.HashSet[string]


# Get all the directories and put them in a hashmap
$directories = (Get-WebJson -Url $ossGetDirectoriesUrl -Headers $ossHeaders).response
$directoriesMap = @{}
ForEach($directory in $directories) {
    $directoriesMap.add($ossDirectoriesUrl + "/" + $directory.id, $directory)
    # While we are looping through these we may as well create the set of directory URLs that we care about
    if ($directoriesToInclude.Contains($directory.name)) {
        $directoryFilter.add($ossDirectoriesUrl + "/" + $directory.id) | Out-Null
    }
}
Write-Output "Found $($directories.Count) DIRECTORIES"

# Get all the sites and put them in a hashmap
$sites = (Get-WebJson -Url $ossGetSitesUrl -Headers $ossHeaders).response
$sitesMap = @{}
ForEach($site in $sites) {
    $sitesMap.add($site.id, $site)
}

Write-Output "Found $($sites.Count) SITES"

# Get all the floors and put them in a hashmap where key is the api URL to that floor
$floors = (Get-WebJson -Url $ossGetFloorsUrl -Headers $ossHeaders).response
$floorsMap = @{}
ForEach($floor in $floors) {
    $floorsMap.add($ossFloorsUrl + "/" + $floor.id, $floor)
}

Write-Output "Found $($floors.Count) FLOORS"

# Get all the seats and put them in a hashmap
$seats = (Get-WebJson -Url $ossGetSeatsUrl -Headers $ossHeaders).response
$seatsMap = @{}
ForEach($seat in $seats) {
    $seatsMap.add($ossSeatsUrl + "/" + $seat.id, $seat)
}

Write-Output "Found $($seats.Count) SEATS"

# Get all the active employees from OfficeSpace
$employees = (Get-WebJson -Url $ossGetEmployeesUrl -Headers $ossHeaders).response
Write-Output "Found $($employees.Count) EMPLOYEES"

# For each employee
ForEach($employee in $employees){

    $employee | Add-Member -Name seat -Value "" -MemberType NoteProperty
    $employee | Add-Member -Name floor -Value "" -MemberType NoteProperty
    $employee | Add-Member -Name site -Value "" -MemberType NoteProperty
    $employee | Add-Member -Name directory -Value "" -MemberType NoteProperty

    # If the employee is seated
    if ( ($employee.seating.seated -ne "not seated") -and
         ($employee.seating.seat_urls.length -gt 0)
        ) {

        $employee_seat = $null

        # And if at least one seat is in a directory that we listed in $directoriesToInclude variable
        ForEach($seat_url in $employee.seating.seat_urls) {
            if ($directoryFilter.Contains($floorsMap.get_item(($seatsMap.get_item($seat_url).floor_url)).directories[0])) {

                $employee_seat = $seatsMap.get_item($seat_url)

                # Get the employee's seat's label
                $employee.seat = $employee_seat.label
                $employee.floor = $floorsMap.get_item($employee_seat.floor_url).label
                if ($floorsMap.get_item($employee_seat.floor_url).site_id) {
                    $employee.site = $sitesMap.get_item($floorsMap.get_item($employee_seat.floor_url).site_id).name
                }
                else {
                    $employee.site = ""
                }
                $employee.directory = $directoriesMap.get_item($floorsMap.get_item(($employee_seat.floor_url)).directories[0]).name
            }
        }
    }
}

$employees | Export-Csv -Delimiter $outputDelimiter -Path "$($outputPath)\$($outputFile)" -NoTypeInformation
