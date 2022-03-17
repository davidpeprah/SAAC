<#

   Author: David Peprah
 
   NOTE: Most of the modification should be done here in this script. Different institutions have different
         Active Directory structure and policies in place when it comes to creating staff account.
         You can add new functions and logic and remove any objects not needed for your environment
#>

import-Module ActiveDirectory


$Time= (Get-Date)


<#
 This return the username to be used after checking the AD to make sure
 its not being used already by another staff
 NOTE: This should be modified to reflect the institution's policy for account creation
#>
function SamAccountNm($lastName, $firstName) {
  
  $proposeAccName = $firstName.Tolower() + "." + $lastName.ToLower()

  if ((chrCount $proposeAccName) -lt 21) {
       
       for ($i=1; $i -lt 6; $i++) {
    
            if (-Not(get-aduser -Filter {SamAccountName -eq $proposeAccName})) {return $proposeAccName}
            if ($i -gt 1) {$proposeAccName + ($i - 1)}
       }
   } else {
      #
      #
      $proposeAccName = $firstName[0] + $lastName.substring(0, 19)

   }

   return $proposeAccName
 }


<#
This function returns the fullname and the abbreviation for a building
to be used for description and department
#>
function buildingNm($building) {
  

}

<#
This function returns OU path to be used to create a user.
It uses the building and Department of the user to determine
the OU the user should created in
NOTE: This should be modified to reflect the institution's policy for account creation
#>
function OUPath($Department, $building) {
 
 
}

<#
Returns security and distribution groups base on the user's building assignment and department
NOTE: This should be modified to reflect the institution's policy for account creation
#>
function defaultStaffGroups($Department, $Building) {
   
}


Try {

<# 
Assign the various values from python to their descriptive variable
Remove all white spaces in lastname and firstname and social security
Remove leading and trailing spaces from building, position, and stafftype
#>
$currentEmpEmail = ($args[0].ToString()).replace(' ', '')
$lastName = ($args[1].ToString()).replace(' ', '')
$firstName = ($args[2].ToString()).replace(' ', '')
$personalEmail = ($args[3].ToString()).replace(' ', '')
$position = $args[4].ToString()
$building_raw = $args[5].ToString()
$department = $args[6].ToString()


$profileDir = profile_dir $department $building_raw
$SamAccountName = SamAccountNm $lastName $firstName
$EmailAddName = EmailAddNm $lastName $firstName
$fullName = fullNm $lastName $firstName
$building = buildingNm $building_raw



$emailAddress = ("$EmailAddName@example.org").ToString()
$password = ("P@ssw0rd@!!").ToString() # This password will change once the account is confirmed in Google console
$path = (dirPath $department $building_raw) + ", OU=Users, DC=example, DC=local"
$description = ($position).ToString()
$homeDirectory = ("\\example.local\$profileDir\$SamAccountName").ToString()
$homeDrive = "H:"
$userPrincipalName = "$SamAccountName@hexample.local"
$Build =  $building[1]

# Groups
$defaultBuildADGroup = defaultBuildingGroups $building_raw
$defaultdepartADGroup = defaultStaffGroups $department $building_raw



# Create User Account
 New-ADUser -Name $fullName -GivenName $firstName -Surname $lastName -DisplayName $fullName `
 -AccountPassword (ConvertTo-SecureString -AsPlainText $password -Force) `
 -SamAccountName $SamAccountName -UserPrincipalName $userPrincipalName `
 -Path $path -EmailAddress $emailAddress -Description $description -Company "Example" `
 -Department $Build -Title $position -HomeDirectory  $homeDirectory -HomeDrive $homeDrive `
 -PasswordNeverExpires $True -Enabled $True




# Add user account to default Group
  if ($defaultBuildADGroup) {
 
     forEach ($grp in $defaultBuildADGroup) {
     Add-ADGroupMember $grp -members $SamAccountName 
   }
  }
  
  
  forEach ($grp in $defaultdepartADGroup) {
    Add-ADGroupMember $grp -members $SamAccountName
  }
  
  #>

  "$Time Account was successfully created for $fullName" | out-file logs\event_log.log -append

  # Return this information to python
  return (1, $emailAddress, "Account Successfully Created in Active Directory")

} catch {

  $ErrorMessage = $_.Exception.Message
  $FailedItem = $_.Exception.ItemName

  "$Time An error occured when trying to create an account for $firstname $lastname. Error Message: $ErrorMessage" | out-file logs\event_log.log -append
  $FailedItem | out-file logs\event_log.log -append
  
  # Return this information to python
  return (2, '', "Could not create account, Check event log for details")
  
  Break

}<# Finally {
  
  
  "This script made a read attempt at $Time" | out-file .\event_log.log -append

}#>






