Param (
    $upn,
    $credsFile
)

try {
    $o365password = Get-Credential $upn -ErrorAction Stop
    $o365password.password | ConvertFrom-SecureString | Set-Content $credsFile -ErrorAction Stop
    write-host "Password updated in $credsFile"
} catch {
    Write-Host "Creds file not updated"
}
