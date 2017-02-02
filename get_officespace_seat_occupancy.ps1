$baseURL = "https://xxxxxxxxxxx.officespacesoftware.com";
$apiGetEmployees = "/api/1/employees";
$apiGetSeats = "/api/1/seats";
$apiGetDirectories = "/api/1/directories";
$apiGetSites = "/api/1/sites";
$apiGetFloors = "/api/1/floors";
$apiToken = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";
$authorizationHeader = @{"Authorization"="Token token="+$apiToken};
$dateTime = Get-Date -UFormat "%Y%m%d_%H%m%S"

$outputPath = "C:\users\ashkon\desktop\"
$outputFile = "OfficeSpace_Seat_Occupancy_$($dateTime).csv"
$outputDelimiter = ","

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
$employeesMap = @{};
ForEach($employee in $employees) {
    $employeesMap.add($apiGetEmployees + "/" + $employee.id, $employee);
}

Write-Output "Found $($employees.Count) EMPLOYEES"

$seatsToReturn = @()

#For each seat
ForEach($seat in $seats){

    $seatToReturn = @{}
    $seatToReturn | Add-Member -Name id -Value "" -MemberType NoteProperty
    $seatToReturn | Add-Member -Name seat -Value "" -MemberType NoteProperty
    $seatToReturn | Add-Member -Name floor -Value "" -MemberType NoteProperty
    $seatToReturn | Add-Member -Name site -Value "" -MemberType NoteProperty
    $seatToReturn | Add-Member -Name directory -Value "" -MemberType NoteProperty
    $seatToReturn | Add-Member -Name occupant -Value "" -MemberType NoteProperty

    if ( $directoryFilter.Contains($floorsMap.get_item($seat.floor_url).directories[0]) -and
         $seat.utility -eq "-1"
       ) {
        $seatToReturn.id = $seat.id
        $seatToReturn.seat = $seat.label
        $seatToReturn.floor = $floorsMap.get_item($seat.floor_url).label
        $seatToReturn.site = $sitesMap.get_item($floorsMap.get_item($seat.floor_url).site_id).name
        $seatToReturn.directory = $directoriesMap.get_item($floorsMap.get_item($seat.floor_url).directories[0]).name
        if ($seat.occupancy.occupied -eq "occupied") {
            $seatToReturn.occupant = ($employeesMap.get_item($seat.occupancy.employee_url)).client_employee_id
        }

        $seatsToReturn += $seatToReturn
    }
    		
}

$seatsToReturn | Export-Csv -Delimiter $outputDelimiter -Path "$($outputPath)\$($outputFile)"
