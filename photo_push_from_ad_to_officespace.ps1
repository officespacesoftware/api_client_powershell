# Force the use of TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

$token         = ""
$hostname      = "xxxxxx.officespacesoftware.com"
$protocol      = "https://"
$headers       = @{Authorization = "Token token="+$token}

$OUDN          = "OU=USA,DC=officespacesoftware,DC=com;"
$filter        = "(objectClass=user)"

$ad_Photo_Attribute = "thumbnailPhoto";
$ad_Id_Attribute = "mail";

$ad_attributes = "sAMAccountName",$ad_Id_Attribute,$ad_Photo_Attribute

$baseURL = "$protocol$hostname"
$apiGetEmployees = "/api/1/employees"
$apiPutEmployees = "/api/1/employees/"

#importing module for get_adusers
Import-Module ActiveDirectory

#prepare a MD5 calculator
$md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider;

Write-Host "Getting records from OfficeSpace..."

#Get all the active employees from OfficeSpace
$employees = (Invoke-RestMethod -Uri $baseURL$apiGetEmployees -Method get -Headers $headers).response;
$employeesMap = @{};
ForEach($employee in $employees) {
    if ( -not ([string]::IsNullOrEmpty($employee.email)) ) { #in case this person from OfficeSpace doesn't have an email
        if (-not $employeesMap.ContainsKey($employee.email)) { #in case there is more than one person with the same email
            $employeesMap.add($employee.email, $employee);
        }
    }
}

Write-Output "Found $($employees.Count) EMPLOYEES"

Write-Host "Getting records from AD..."

$OUDNsSplit = $OUDN.Split(";")

$ad_results = $OUDNsSplit[0..($OUDNsSplit.count-2)] | foreach {Get-ADUser -LDAPFilter $filter -SearchScope SubTree -SearchBase $_ -Properties $ad_attributes}

[System.Object[]]$ADusersinOU = @()
$ADusersinOU  = $ADusersinOU + $ad_results #we do it like this in case the get_aduser call returned only one record... we still need it to be an array

Write-Host "Found $($ADusersinOU.Count) RECORDS"

Write-Host "Updating OfficeSpace..."
# for each record from AD
#  compare AD.md5 with OSS.md5 image_source_fingerprint
#  if they are different then update OSS 
foreach ($ADuser in $ADusersinOU) {

    #if this employee is in OfficeSpace
    if ($employeesMap.get_item($ADuser.$ad_Id_Attribute))
    {
        #get the OfficeSpace employee object from the map
        $employee = $employeesMap.get_item($ADuser.$ad_Id_Attribute);
        $empId = $employee.id;
        
        #get the image data from AD
        $imageDataRaw = "";
        if ($ADuser.$ad_Photo_Attribute.value) {
            $imageDataRaw = $ADuser.$ad_Photo_Attribute.value;
        }
        elseif ($ADuser.$ad_Photo_Attribute) {
            $imageDataRaw = $ADuser.$ad_Photo_Attribute;
        }

        #if we found image data
        if ($imageDataRaw -ne ""){

            #calculate md5 of photo in AD
            $md5hash = [System.BitConverter]::ToString($md5.ComputeHash($imageDataRaw))
            $md5hash = $md5hash -replace '-',''

            #if md5 of photo in AD does not match md5 stored in OfficeSpace
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

    } #end of if employee is in OfficeSpace
}

Write-Host "Process completed."
