$OutputEncoding=[Text.Encoding]::UTF8
Param([string]$FirstName,[string]$LastName)
Import-Module ActiveDirectory
try {
  $usr=Get-ADUser -Filter {givenName -eq $FirstName -and sn -eq $LastName}
  $pw="Aa0000"|ConvertTo-SecureString -AsPlainText -Force
  Set-ADAccountPassword -Identity $usr.SamAccountName -NewPassword $pw -Reset
  Write-Output "Пароль сброшен."
  exit 0
} catch { Write-Output "Ошибка: $_"; exit 1 }
