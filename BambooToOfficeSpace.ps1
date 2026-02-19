#####################################################################################
# The following is needed to allow PowerShell to query BambooHR's API without error #
#####################################################################################
$AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
[System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
#####################################################################################

$bamboo_subd = "YOUR_BAMBOO_SUBDOMAIN"
$bamboo_user = "YOUR_BAMBOO_API_KEY_GOES_HERE"
$bamboo_pass = "x"
$bamboo_pair = "${bamboo_user}:${bamboo_pass}"

$bamboo_api_url = "https://api.bamboohr.com/api/gateway.php/$($bamboo_subd)/v1/employees/directory"

$bamboo_pair_base64 = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($bamboo_pair))

$bamboo_headers = @{}
$bamboo_headers.Add("Authorization","Basic $($bamboo_pair_base64)")
$bamboo_headers.Add("Accept","application/json")

$bamboo_response = Invoke-RestMethod -UseBasicParsing -Uri $bamboo_api_url -Headers $bamboo_headers -Method Get

# Uncomment the following line to export the BambooHR data to a CSV file
#$bamboo_response.employees | Export-Csv "$($env:userprofile)\desktop\bamboo_export_$($dateTime).csv"

Write-Host "Got $($bamboo_response.employees.Count) employees from Bamboo"

########################
# OfficeSpace API Push #
########################

$hostname      = "YOUR_OFFICESPACE_SUBDOMAIN.officespacesoftware.com"
$token         = "YOUR_OFFICESPACE_API_KEY_GOES_HERE"
$batch_url     = "/api/1/employee_batch_imports"
$import_url    = "/api/1/employee_directory"
$protocol      = "https://"
$headers       = @{Authorization = "Token token="+$token}
$source        = "BambooHR"
$employee_batch_url  = $protocol + $hostname + $batch_url
$employee_batch_staging_url  = $protocol + $hostname + $import_url + "/" + $source
$employee_import_url  = $protocol + $hostname + $import_url

$batchSize = 1000

#  OfficeSpace to BambooHR mapping  (additional possible attributes for UDFs can be found at: https://documentation.bamboohr.com/docs/list-of-field-names - they need testing)
$EmployeeId = "id"
$FirstName = "firstName"
$LastName = "lastName"
$Title = "jobTitle"
$WorkPhone = "workPhone"
$Extension = "workPhoneExtension"
$Photo = "photoUrl"
$Department = "department"
$Bio = ""
$Email = "workEmail"
$StartDate = ""
$EndDate = ""
$ShowInVd = ""
$Udf0 = "location"
$Udf1 = "division"
$Udf2 = "mobilePhone"
$Udf3 = "preferredName"
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

$bamboo_employees = $bamboo_response.employees

$batches = [math]::floor($bamboo_employees.count/$batchSize) + 1
$start = 0
$end = $batchSize - 1
$batch = 1


Write-Host "Staging records in OfficeSpace..."
Invoke-WebRequest -UseBasicParsing -Uri $employee_batch_staging_url -ContentType application/json -Method Delete -Headers $headers | Out-Null

Write-Host "Sending records to $employee_batch_url"
do
{
$Array = new-object system.collections.arraylist
foreach ($bamboo_employee in $bamboo_employees[$start..$end])
    {

$Array.Add([PSCustomObject]@{
    "EmployeeId"= "$(($bamboo_employee).$EmployeeId)"
    "Source"= $source
    "FirstName"= "$(($bamboo_employee).$FirstName)"
    "LastName"= "$(($bamboo_employee).$LastName)"
    "Title"= "$(($bamboo_employee).$Title)"
    "WorkPhone"= "$(($bamboo_employee).$WorkPhone)"
    "Extension"= "$(($bamboo_employee).$Extension)"
    "Photo"= "$(($bamboo_employee).$Photo)"
    "Department"= "$(($bamboo_employee).$Department)"
    "Bio"= "$(($bamboo_employee).$Bio)"
    "Email"= "$(($bamboo_employee).$Email)"
    "StartDate"= "$(($bamboo_employee).$StartDate)"
    "EndDate"= "$(($bamboo_employee).$EndDate)"
    "ShowInVd"= "$(($bamboo_employee).$ShowInVd)"
    "Udf0"= "$(($bamboo_employee).$Udf0)"
    "Udf1"= "$(($bamboo_employee).$Udf1)"
    "Udf2"= "$(($bamboo_employee).$Udf2)"
    "Udf3"= "$(($bamboo_employee).$Udf3)"
    "Udf4"= "$(($bamboo_employee).$Udf4)"
    "Udf5"= "$(($bamboo_employee).$Udf5)"
    "Udf6"= "$(($bamboo_employee).$Udf6)"
    "Udf7"= "$(($bamboo_employee).$Udf7)"
    "Udf8"= "$(($bamboo_employee).$Udf8)"
    "Udf9"= "$(($bamboo_employee).$Udf9)"
    "Udf10"= "$(($bamboo_employee).$Udf10)"
    "Udf11"= "$(($bamboo_employee).$Udf11)"
    "Udf12"= "$(($bamboo_employee).$Udf12)"
    "Udf13"= "$(($bamboo_employee).$Udf13)"
    "Udf14"= "$(($bamboo_employee).$Udf14)"
    "Udf15"= "$(($bamboo_employee).$Udf15)"
    "Udf16"= "$(($bamboo_employee).$Udf16)"
    "Udf17"= "$(($bamboo_employee).$Udf17)"
    "Udf18"= "$(($bamboo_employee).$Udf18)"
    "Udf19"= "$(($bamboo_employee).$Udf19)"
    "Udf20"= "$(($bamboo_employee).$Udf20)"
    "Udf21"= "$(($bamboo_employee).$Udf21)"
    "Udf22"= "$(($bamboo_employee).$Udf22)"
    "Udf23"= "$(($bamboo_employee).$Udf23)"
    "Udf24"= "$(($bamboo_employee).$Udf24)"}) | Out-Null
    }

  $JSONArray = $array | ConvertTo-Json
  $JSONArrayUTF8 = [System.Text.Encoding]::UTF8.GetBytes($JSONArray)

  try {
      Invoke-WebRequest -UseBasicParsing -Uri $employee_batch_url -ContentType 'application/json; charset=utf-8' -Method Post -Body $JSONArrayUTF8 -Headers $headers -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Out-Null
      $startn = $start+1;$endn = $end+1; Write-Host "$startn-$endn " -NoNewline; Write-Host "Done" -ForegroundColor Green
      }
  catch
      {
      $startn = $start+1;$endn = $end+1; Write-Host "$startn-$endn " -NoNewline; Write-Host $_.Exception.Message -ForegroundColor Red
      }

  $start+=$batchSize
  $end+=$batchSize
  $batch++
}
while ($batch -le $batches)

Write-Host "Triggering migration"
$import_url_post_body = "Source=" + $source
Invoke-WebRequest -UseBasicParsing -Uri $employee_import_url -Method Post -Body $import_url_post_body -Headers $headers | Out-Null
Write-Host "Completed"
