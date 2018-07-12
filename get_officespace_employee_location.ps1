$baseURL = "https://xxxxxxxxxxx.officespacesoftware.com";
$apiToken = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";

$apiGetEmployees = "/api/1/employees";
$apiGetSeats = "/api/1/seats";
$apiGetDirectories = "/api/1/directories";
$apiGetSites = "/api/1/sites";
$apiGetFloors = "/api/1/floors";
$authorizationHeader = @{"Authorization"="Token token="+$apiToken};

$dateTime = Get-Date -UFormat "%Y%m%d_%H%m%S"

$outputPath = "$($env:userprofile)\desktop\"
$outputFile = "OfficeSpace_Seat_Occupancy_$($dateTime).csv"
$outputDelimiter = ","

#Force TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

################################################################
# SYS: Make Public and SYS: Make Private are two special system
# directories that always exists and that describe the online and
# offiline floor respectively.
$directoryOnlineName = "SYS: Make Public";
$directoryOfflineName = "SYS: Make Private";

# A string based set of directory names to filter by
$directoriesToInclude = New-Object System.Collections.Generic.HashSet[string];
$directoriesToInclude.Add($directoryOnlineName) | Out-Null
#$directoriesToInclude.Add($directoryOfflineName)
#$directoriesToInclude.Add("SOME OTHER DIRECTORY NAME")
#$directoriesToInclude.Add(...)
###############################################################

# A string based set of directory URLs to filter by
$directoryFilter = New-Object System.Collections.Generic.HashSet[string];


#Get all the directories and put them in a hashmap
$directories = (Invoke-RestMethod -Uri $baseURL$apiGetDirectories -Method get -Headers $authorizationHeader).response;
$directoriesMap = @{};
ForEach($directory in $directories){
    $directoriesMap.add($apiGetDirectories + "/" + $directory.id, $directory);
    #while we are loping through these we may as well create the set of directory URLs that we care about
    if ($directoriesToInclude.Contains($directory.name)) {
        $directoryFilter.add($apiGetDirectories + "/" + $directory.id) | Out-Null
    }
}
Write-Output "Found $($directories.Count) DIRECTORIES"

#Get all the sites and put them in a hashmap
$sites = (Invoke-RestMethod -Uri $baseURL$apiGetSites -Method get -Headers $authorizationHeader).response;
$sitesMap = @{};
ForEach($site in $sites){
    $sitesMap.add($site.id, $site);
}

Write-Output "Found $($sites.Count) SITES"

#Get all the floors and put them in a hashmap where key is the api URL to that floor
$floors = (Invoke-RestMethod -Uri $baseURL$apiGetFloors -Method get -Headers $authorizationHeader).response;
$floorsMap = @{};
ForEach($floor in $floors){
    $floorsMap.add($apiGetFloors + "/" + $floor.id, $floor);
}

Write-Output "Found $($floors.Count) FLOORS"

#Get all the seats and put them in a hashmap 
$seats = (Invoke-RestMethod -Uri $baseURL$apiGetSeats -Method get -Headers $authorizationHeader).response;
$seatsMap = @{};
ForEach($seat in $seats){
    $seatsMap.add($apiGetSeats + "/" + $seat.id, $seat);
}

Write-Output "Found $($seats.Count) SEATS"

#Get all the active employees from OfficeSpace
$employees = (Invoke-RestMethod -Uri $baseURL$apiGetEmployees -Method get -Headers $authorizationHeader).response;

Write-Output "Found $($employees.Count) EMPLOYEES"

#For each employee
ForEach($employee in $employees){

    $employee | Add-Member -Name seat -Value "" -MemberType NoteProperty
    $employee | Add-Member -Name floor -Value "" -MemberType NoteProperty
    $employee | Add-Member -Name site -Value "" -MemberType NoteProperty
    $employee | Add-Member -Name directory -Value "" -MemberType NoteProperty


    #If the employee is seated
    if ( ($employee.seating.seated -ne "not seated") -and 
         ($employee.seating.seat_urls.length -gt 0)
        ) {

        $employee_seat = $null

        #And if at least one seat is in a directory that we listed in $directoriesToInclude variable
        ForEach($seat_url in $employee.seating.seat_urls){
            if ($directoryFilter.Contains($floorsMap.get_item(($seatsMap.get_item($seat_url).floor_url)).directories[0])){

                $employee_seat = $seatsMap.get_item($seat_url)

                #Get the employee's seat's label
                $employee.seat = $employee_seat.label;
                $employee.floor = $floorsMap.get_item($employee_seat.floor_url).label;
                if ($floorsMap.get_item($employee_seat.floor_url).site_id) {
                    $employee.site = $sitesMap.get_item($floorsMap.get_item($employee_seat.floor_url).site_id).name;
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
