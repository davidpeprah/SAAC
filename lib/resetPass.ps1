import-Module ActiveDirectory
$Time= (Get-Date)

function lastFourSSN($lastFourDigitSSN) {
   
   if ($lastFourDigitSSN.length -lt 3) { 
     
       return (get-date).year 
   } else { 
   
       return $lastFourDigitSSN 
   } 
}

Try {

$workEmail = $args[0].ToString()
$lastFourDigitSSN = lastFourSSN $args[1].ToString()

$user = get-aduser -Filter {mail -eq $workEmail} -Properties SamAccountName | select -ExpandProperty SamAccountName


Set-ADAccountPassword -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "milford$lastFourDigitSSN" -Force)
"$Time Password Reset Successful for $workEmail" | out-file .\event_log.log -append
  return 0

} catch {
   $ErrorMessage = $_.Exception.Message

  "$Time Password Reset failed $workEmail. Error Message: $ErrorMessage" | out-file event_log.log -append
  
  
  # Return this information to python
  return 1
  
  Break
}
