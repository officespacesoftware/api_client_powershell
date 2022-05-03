# Force the use of TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

##################
# UltiPro config 
##################
$ultipro_subd = "service4"
$ultipro_key  = "1RYOW"
$ultipro_user = "OfficeSpace"
$ultipro_pass = "XXXXXXXXXXXXXXX"
##################

$ultipro_pair = "${ultipro_user}:${ultipro_pass}"
$ultipro_pair_base64 = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($ultipro_pair))
$ultipro_person_api = "https://$($ultipro_subd).ultipro.com/personnel/v1/person-details"
$ultipro_employment_api = "https://$($ultipro_subd).ultipro.com/personnel/v1/employment-details"
$ultipro_org_api = "https://$($ultipro_subd).ultipro.com/configuration/v1/org-levels"

$ultipro_headers = @{}
$ultipro_headers.Add("Authorization","Basic $($ultipro_pair_base64)")
$ultipro_headers.Add("Accept","application/json")
$ultipro_headers.Add("US-Customer-Api-Key", $ultipro_key)


######################
# OfficeSpace config 
######################
$oss_token         = "YYYYYYYYYYYYYYY" # Your API key can be found here: https://<client_domain>.officespacesoftware.com/api_keys
$oss_hostname      = "<client_domain>.officespacesoftware.com"
######################

$oss_protocol      = "https://"
$oss_headers       = @{Authorization = "Token token="+$oss_token}

$oss_batch_url     = "/api/1/employee_batch_imports"
$oss_import_url    = "/api/1/employee_directory"
$source            = "UltiPro"
$oss_import_url_post_body = "Source=" + $source
$oss_employee_batch_url  = $oss_protocol + $oss_hostname + $oss_batch_url
$oss_employee_batch_staging_url  = $oss_protocol + $oss_hostname + $oss_import_url + "/" + $source
$oss_employee_import_url  = $oss_protocol + $oss_hostname + $oss_import_url

############################################
# OfficeSpace to UltiPro attribute mapping 
############################################
$EmployeeId = "employeeID"
$FirstName = "firstName"
$LastName = "lastName"
$Title = "jobDescription"
$WorkPhone = "workPhoneNumber"
$Extension = ""
$Photo = ""
$Department = "orgLevel1"
$Bio = ""
$Email = "emailAddress"
$StartDate = ""
$EndDate = ""
$ShowInVd = ""
$Udf0 = "orgLevel2"
$Udf1 = "orgLevel3"
$Udf2 = "supervisorID"
$Udf3 = "supervisorFirstName"
$Udf4 = "supervisorLastName"
$Udf5 = "userName"
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
##########################################


$page_size = 100

# get all the "org-codes" data from UltiPro and store them in a hashmap"
$current_page = 1
$current_result_count = 0
$ultipro_all_org_codes_map = @{}

do {

    $paged_orgs = $null
    $current_result_count = 0
    Write-Host "Getting Page: $($current_page) from UltiPro Org API"
    $paged_orgs = Invoke-RestMethod -UseBasicParsing -Uri  "$($ultipro_org_api)?per_page=$($page_size)&page=$($current_page)" -Headers $ultipro_headers -Method Get
    Write-Host "Found $($paged_orgs.count) orgs"
    if ($paged_orgs) {

        foreach ($org in $paged_orgs) {
            $ultipro_all_org_codes_map.Set_Item($org.code, $org.description)
        }

    }
    $current_result_count = $paged_orgs.count
    $current_page++

} while ($current_result_count -gt 0)

Write-Host "Got $($ultipro_all_org_codes_map.Count) unique records from the UltiPro Org Endpoint"


# get all the "person-details" data from UltiPro
$current_page = 1
$current_result_count = 0
$ultipro_all_person_map = @{}

do {

    $paged_employees = $null
    $current_result_count = 0
    Write-Host "Getting Page: $($current_page) from UltiPro Person API"
    $paged_employees = Invoke-RestMethod -UseBasicParsing -Uri  "$($ultipro_person_api)?per_page=$($page_size)&page=$($current_page)" -Headers $ultipro_headers -Method Get
    Write-Host "Found $($paged_employees.count) users"
    if ($paged_employees) {

        foreach ($employee in $paged_employees) {
            $ultipro_all_person_map.Set_Item($employee.employeeID, $employee)
        }

    }
    $current_result_count = $paged_employees.count
    $current_page++

} while ($current_result_count -gt 0)

Write-Host "Got $($ultipro_all_person_map.Count) unique records from the UltiPro Person Endpoint"


# get all the "employment-details" data from UltiPro and add some of it to the data we got from "person-details"
$current_page = 1
$current_result_count = 0

do {

    $paged_employees = $null
    $current_result_count = 0
    Write-Host "Getting Page: $($current_page) from UltiPro Employment API"
    $paged_employees = Invoke-RestMethod -UseBasicParsing -Uri  "$($ultipro_employment_api)?per_page=$($page_size)&page=$($current_page)" -Headers $ultipro_headers -Method Get
    Write-Host "Found $($paged_employees.count) users"
    if ($paged_employees) {

        foreach ($employee in $paged_employees) {
            if ($ultipro_all_person_map."$($employee.employeeID)") {
                #remove any inactive records
                if ($null -ne $employee.termReason){
                    $ultipro_all_person_map.Remove($employee.employeeID)
                }
                else{
                    $ultipro_all_person_map."$($employee.employeeID)" | Add-Member -NotePropertyName jobDescription -NotePropertyValue $employee.jobDescription
                    $ultipro_all_person_map."$($employee.employeeID)" | Add-Member -NotePropertyName workPhoneNumber -NotePropertyValue $employee.workPhoneNumber
                    $ultipro_all_person_map."$($employee.employeeID)" | Add-Member -NotePropertyName supervisorID -NotePropertyValue $employee.supervisorID
                    $ultipro_all_person_map."$($employee.employeeID)" | Add-Member -NotePropertyName supervisorFirstName -NotePropertyValue $employee.supervisorFirstName
                    $ultipro_all_person_map."$($employee.employeeID)" | Add-Member -NotePropertyName supervisorLastName -NotePropertyValue $employee.supervisorLastName
                    $ultipro_all_person_map."$($employee.employeeID)" | Add-Member -NotePropertyName orgLevel1 -NotePropertyValue $(if($employee.orgLevel1Code -and $ultipro_all_org_codes_map.($employee.orgLevel1Code)){$ultipro_all_org_codes_map.($employee.orgLevel1Code)} Else {""})
                    $ultipro_all_person_map."$($employee.employeeID)" | Add-Member -NotePropertyName orgLevel2 -NotePropertyValue $(if($employee.orgLevel2Code -and $ultipro_all_org_codes_map.($employee.orgLevel2Code)){$ultipro_all_org_codes_map.($employee.orgLevel2Code)} Else {""})
                    $ultipro_all_person_map."$($employee.employeeID)" | Add-Member -NotePropertyName orgLevel3 -NotePropertyValue $(if($employee.orgLevel3Code -and $ultipro_all_org_codes_map.($employee.orgLevel3Code)){$ultipro_all_org_codes_map.($employee.orgLevel3Code)} Else {""})
                }
            }            
        }

    }
    $current_result_count = $paged_employees.count
    $current_page++

} while ($current_result_count -gt 0)


# Start communicating with OfficeSpace

Write-Host "Preparing OfficeSpace for new data push"
Invoke-WebRequest -UseBasicParsing -Uri $oss_employee_batch_staging_url -ContentType application/json -Method Delete -Headers $oss_headers | Out-Null

$Array = new-object system.collections.arraylist

foreach ($employee in $ultipro_all_person_map.GetEnumerator()) {

    $Array.Add([PSCustomObject]@{
        "EmployeeId" = "$(($employee.Value).$EmployeeId)"
        "Source"     = $source
        "FirstName"  = "$(($employee.Value).$FirstName)"
        "LastName"   = "$(($employee.Value).$LastName)"
        "Title"      = "$(($employee.Value).$Title)"
        "WorkPhone"  = "$(($employee.Value).$WorkPhone)"
        "Extension"  = "$(($employee.Value).$Extension)"
        "ImageData"  = $imageData
        "Department" = "$(($employee.Value).$Department)"
        "Bio"        = "$(($employee.Value).$Bio)"
        "Email"      = "$(($employee.Value).$Email)"
        "StartDate"  = "$(($employee.Value).$StartDate)"
        "EndDate"    = "$(($employee.Value).$EndDate)"
        "ShowInVd"   = "$(($employee.Value).$ShowInVd)"
        "Udf0"       = "$(($employee.Value).$Udf0)"
        "Udf1"       = "$(($employee.Value).$Udf1)"
        "Udf2"       = "$(($employee.Value).$Udf2)"
        "Udf3"       = "$(($employee.Value).$Udf3)"
        "Udf4"       = "$(($employee.Value).$Udf4)"
        "Udf5"       = "$(($employee.Value).$Udf5)"
        "Udf6"       = "$(($employee.Value).$Udf6)"
        "Udf7"       = "$(($employee.Value).$Udf7)"
        "Udf8"       = "$(($employee.Value).$Udf8)"
        "Udf9"       = "$(($employee.Value).$Udf9)"
        "Udf10"      = "$(($employee.Value).$Udf10)"
        "Udf11"      = "$(($employee.Value).$Udf11)"
        "Udf12"      = "$(($employee.Value).$Udf12)"
        "Udf13"      = "$(($employee.Value).$Udf13)"
        "Udf14"      = "$(($employee.Value).$Udf14)"
        "Udf15"      = "$(($employee.Value).$Udf15)"
        "Udf16"      = "$(($employee.Value).$Udf16)"
        "Udf17"      = "$(($employee.Value).$Udf17)"
        "Udf18"      = "$(($employee.Value).$Udf18)"
        "Udf19"      = "$(($employee.Value).$Udf19)"
        "Udf20"      = "$(($employee.Value).$Udf20)"
        "Udf21"      = "$(($employee.Value).$Udf21)"
        "Udf22"      = "$(($employee.Value).$Udf22)"
        "Udf23"      = "$(($employee.Value).$Udf23)"
        "Udf24"      = "$(($employee.Value).$Udf24)"
    }) | Out-Null
}

$JSONArray = $array | ConvertTo-Json
$JSONArrayUTF8 = [System.Text.Encoding]::UTF8.GetBytes($JSONArray)

Write-Host "Sending records to $($oss_employee_batch_url)"
Invoke-WebRequest -UseBasicParsing -Uri $oss_employee_batch_url -ContentType 'application/json; charset=utf-8' -Method Post -Body $JSONArrayUTF8 -Headers $oss_headers -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Out-Null

Write-Host "Triggering migration"
$oss_import_url_post_body = "Source=" + $source
Invoke-WebRequest -UseBasicParsing -Uri $oss_employee_import_url -Method Post -Body $oss_import_url_post_body -Headers $oss_headers | Out-Null
Write-Host "Completed"