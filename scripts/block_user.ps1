$OutputEncoding=[Text.Encoding]::UTF8
Param([string]$Login,[string]$FirstName,[string]$LastName)
Import-Module ActiveDirectory
try {
  if ($Login) {
    $usr=Get-ADUser -Filter {SamAccountName -eq $Login}
  } else {
    $usr=Get-ADUser -Filter {givenName -eq $FirstName -and sn -eq $LastName}
  }
  Disable-ADAccount -Identity $usr.SamAccountName
  Write-Output "Заблокирован."
  exit 0
} catch { Write-Output "Ошибка: $_"; exit 1 }
