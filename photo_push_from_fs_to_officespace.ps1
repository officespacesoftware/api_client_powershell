$token         = "xxxxxxxxxxxxxxxxxxxxxx"
$protocol      = "https://"
$headers       = @{Authorization = "Token token="+$token}
$hostname      = "xxxxxxxxxxxxx.officespacesoftware.com"
$baseURL       = $protocol+$hostname;
$apiGetEmployees = "/api/1/employees";
$apiPutEmployees = "/api/1/employees/"

$PhotoPath = "$($env:userprofile)\desktop\photos\"

function Test-Image { # This function was copied from https://blogs.technet.microsoft.com/heyscriptingguy/2015/03/19/psimaging-part-1-test-image/
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param(
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias('PSPath')]
        [string] $Path
    )

    PROCESS {
        $knownHeaders = @{
            jpg = @( "FF", "D8" );
            #bmp = @( "42", "4D" );
            gif = @( "47", "49", "46" );
            tif = @( "49", "49", "2A" );
            png = @( "89", "50", "4E", "47", "0D", "0A", "1A", "0A" );
            #pdf = @( "25", "50", "44", "46" );
        }

        # coerce relative paths from the pipeline into full paths
        if($_ -ne $null) {
            $Path = $_.FullName

        }

         # read in the first 8 bits

        $bytes = Get-Content -LiteralPath $Path -Encoding Byte -ReadCount 1 -TotalCount 8 -ErrorAction Ignore
         $retval = $false
        foreach($key in $knownHeaders.Keys) {
             # make the file header data the same length and format as the known header
            $fileHeader = $bytes |
                Select-Object -First $knownHeaders[$key].Length |
                ForEach-Object { $_.ToString("X2") }
            if($fileHeader.Length -eq 0) {
                continue
            }

             # compare the two headers
            $diff = Compare-Object -ReferenceObject $knownHeaders[$key] -DifferenceObject $fileHeader
            if(($diff | Measure-Object).Count -eq 0) {
                $retval = $true
            }
        }

        return $retval
    }
}

#Get all the active employees from OfficeSpace and put them in a hash map
$employees = (Invoke-RestMethod -Uri $baseURL$apiGetEmployees -Method get -Headers $headers).response;
$employeeMap = @{};
ForEach($employee in $employees){
    $employeeMap.add($employee.client_employee_id, $employee);
}

$batchSize = 300

$StartTime = Get-Date


$Photos = Get-ChildItem -Path $PhotoPath -File
Write-Host "Found $($Photos.count) files in $($PhotoPath)"

$batches = [math]::floor($Photos.count/$batchSize) + 1
$start = 0
$end = $batchSize - 1
$batch = 1

Write-Host "Updating photos... "
do
{
    #for each user in this batch
    foreach ($Photo in $Photos[$start..$end])
    {

        $id_from_photo = [io.path]::GetFileNameWithoutExtension($Photo);
        Write-Host "Processing id: $($id_from_photo)"

        #if this employee is in OfficeSpace
        if ($employeeMap.get_item($id_from_photo) -and (Test-Image $PhotoPath$Photo))
        {
            #get the OfficeSpace employee object from the map
            $employee = $employeeMap.get_item($id_from_photo);
            $empId = ($employee).id;

            #get the image data from file system
            $imageDataRaw = "";
            if ($employee) {
                $imageDataRaw = Get-Content -Raw $PhotoPath$Photo -Encoding byte;
            }

            #if we found image data
            if ($imageDataRaw -ne ""){
                
                #calculate md5 of photo in file system
                $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider;
                $md5hash = [System.BitConverter]::ToString($md5.ComputeHash($imageDataRaw))
                $md5hash = $md5hash -replace '-',''

                #if md5 of photo in file system does not match md5 stored in OfficeSpace
                if ($md5hash -ine $employee.image_source_fingerprint){
                    $imageData = [System.Convert]::ToBase64String($imageDataRaw);

                    $request = @{
                        record = @{
                            imageData = $imageData
                            }
                        }

                    $JSONrequest = $request | ConvertTo-Json

                    Write-Host "Updating photo for user: " $employee.client_employee_id;
                    Invoke-WebRequest -Uri $baseURL$apiPutEmployees$empId -ContentType 'application/json; charset=utf-8' -Method PUT -Body $JSONrequest -Headers $headers -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Out-Null

                } #end of if the md5 hashes did not match

            } #end of if we found image data

        } #end of if this employee is in OfficeSpace

        else {
            Write-Host "Could not find $($id_from_photo) in OfficeSpace or it was not a proper image file";
        }

    } #end of for each user in this batch

    $start+=$batchSize
    $end+=$batchSize
    $batch++
    $CurrentTime = Get-Date
    $CurrentRunTime = $CurrentTime - $StartTime
    Write-Host "Current run time $($CurrentRunTime.TotalHours) hours. Started: $StartTime. Current: $CurrentTime."

}
while ($batch -le $batches)

$EndTime = Get-Date
$RunTime = $EndTime - $StartTime
Write-Host "Completed in $($RunTime.TotalHours) hours. Started: $StartTime. Ended: $EndTime."
