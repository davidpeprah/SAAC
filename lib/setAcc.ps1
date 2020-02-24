import-Module ActiveDirectory

$Time= (Get-Date)

# Default groups 
$defaultStaffADGroup = "Milford", "Milford-All", "All Staff"

# Default groups per building
$defaultBOEADGroup = "COST"

# Default groups per function
#Custodian
$defaultADGroupCus = ""

#Food Service
$defaultADGroupFS = ""

#Extended Day
$defaultADGroupEXD = ""

#Teacher/Health/Media Aides
$defaultADGroupAIDES = ""

#Certified
$defaultADGroupCRT = ""

#Building Secretary
$defaultADGroupBLDGSEC = ""

#Exempt
$defaultADGroupEXEMPT = ""

#Administrative
$defaultADGroupADM = ""


# This is used to determine the directory for the user's profile
function profile_dir($staff_type) {
  

  switch -Wildcard ($staff_type) {
     
     "*teacher*" {return "TEACHERS"}
     "*admin*" {return "ADMIN"}
     "*principal*" {return "ADMIN"}
     default {return "classified"}
  }

}



<#
 This return the username to be used after checking the AD to make sure
 its not being used already by another staff
#>
function SamAccountNm($lastName, $firstName) {
  
  for ($i=1; $i -lt 4; $i++) {

    $proposeAccName = $lastName.ToLower() + "_" + $firstName.Tolower().Substring(0,$i)
    if (-Not(get-aduser -Filter {SamAccountName -eq $proposeAccName})) {return $proposeAccName}

  }

   $proposeAccName = $lastName.ToLower() + "_" + $firstName.Tolower()
   if (-Not(get-aduser -Filter {SamAccountName -eq $proposeAccName})) {return $proposeAccName} else {break}

}

function fullNm($lastName, $firstName) {

    $proposefullName = "$firstName $lastName"
    if (-Not(get-aduser -Filter {Name -eq $proposefullName})) {return $proposefullName} 

    $proposefullName = "$lastName $firstName"
    if (-Not(get-aduser -Filter {Name -eq $proposefullName})) {return $proposefullName}

    return "$firstName L $lastName"
   
}

# Returns character count
function chrCount($str) {
  
  $measureObject = $str | Measure-Object -Character;
  $count = $measureObject.Characters;
  return $count;

}

# Return the current year if last four digit of social security is not provided
function lastFourSSN($lastFourDigitSSN) {
   
   if (!$lastFourDigitSSN) {

       return (get-date).year 

   } elseif (chrCount($lastFourDigitSSN) -lt 3)  { 
   
       return (get-date).year
        
   } else {
       
       return $lastFourDigitSSN 
   }
   
   
}

<#
This function returns the fullname and the abbreviation for a building
to be used for description and department
#>

function buildingNm($building) {
   
   switch -Wildcard ($building) {

   "*High School*" {"MHS","Milford High School"; break}
   "*Junior School*" {"MJH", "Milford Junior School"; break}
   "*Boyd * Smith*" {"BES", "Milford Boyd E Smith ES"; break}
   "*Charles * Seipelt*" {"CLS", "Milford Charles L Seipelt ES"; break}
   "*Mulberry*" {"MLB", "Milford Mulberry ES"; break}
   "*Meadowview*" {"MDV", "Milford Meadowview ES"; break}
   "*McCormick*" {"MCM", "Milford McCormick ES"; break}
   "*Pattison*" {"MJH", "Milford Pattison ES"; break}
   "*Preschool*" {"MPS", "Milford Preschool"; break}
   "*Extended Day*" {"EXT", "Milford Preschool"; break}
   "*Administrative Offices*" {"BOE", "Milford Board of Education"; break}
   "*District*" {"District", "Milford School District"; break}
   "*Wyoming*" {"Wyoming", "Wyoming"; break}
   "*Williamsburg*" {"MJH", "Milford Junior School"; break}
   "*Success Academy*" {"SSA", "Milford Student Success Academy"; break}
   "*St* Columb*" {"St. Columbian", "St. Columbian"; break}
   "*St* Andrew*" {"St. Andrew", "St. Andrew"; break}
   "*Seton*" {"Seton", "Seton"; break}
   "*SEASSA*" {"SEASSA", "SEASSA"; break}
   "*John Paul II*" {"John Paul II", "John Paul II"; break}
   "*Finneytown*" {"Finneytown", "Finneytown"; break}
   "*Maderia*" {"Maderia", "Maderia"; break}
   default {"MEVSD","Milford Exempted Village School District"}
   }

}

<#
This function returns OU path to be used to create a user.
by default ti will return "OU=Classified" its not able to 
determine the proper OU for staff
#>
function dirPath($staffType, $building) {
 

  
 if ($staffType -like "TEACHE*") {
   
       switch -Wildcard ($building) {
       
       "*High School*" {return "OU=High School, OU=Teachers"; break}
       "*Junior School*" {return "OU=Junior High, OU=Teachers"; break}
       "*Boyd * Smith*" {return "OU=Boyd E Smith, OU=Teachers"; break}
       "*Charles * Seipelt*" {return "OU=Charles L Seipelt, OU=Teachers"; break}
       "*Mulberry*" {return "OU=Mulberry, OU=Teachers"; break}
       "*Meadowview*" {return "OU=Meadowview, OU=Teachers"; break}
       "*McCormick*" {return "OU=McCormick, OU=Teachers"; break}
       "*Pattison*" {return "OU=Pattison, OU=Teachers"; break}
       "*Preschool*" {return "OU=Miami PreSchool Extended Day, OU=Teachers"; break}
       "*Extended Day*" {return "OU=Miami PreSchool Extended Day, OU=Teachers"; break}
       "*Success Academy*" {return "OU=Success Academy, OU=Teachers"; break}
       "*Long*Term*Sub*" {return "OU=Long Term Substitutes, OU=Teachers"; break} 
       default {return "OU=District, OU=Teachers"}
       }  
  } elseif ($staffType -like "*adm*") {

      return "OU=Administrative"

  } else {

      switch -Wildcard ($staffType) {

      "*custodian*" {return "OU=Custodial Facilities, OU=Classified"; break}
      "*mainten*" {return "OU=Classified"; break}
      "*food*" {return "OU=Nutrition Services POS, OU=Classified"; break}
       "*kitch*" {return "OU=Nutrition Services POS, OU=Classified"; break}
       "*transport*" {return "OU=Transportation, OU=Classified"; break}
       "*principal*" {return "OU=Administrative"; break}
       "*board*Member*" {return "OU=Board Members, OU=Administrative"; break}
      }

  }

   if ($building -like "District") {return "OU=District, OU=Teachers"} 

  return "OU=Classified"
 
}

Try {

<# 
Assign the various values from python to their descriptive variable
Remove all white spaces in lastname and firstname and social security
Remove leading and trailing spaces from building, position, and stafftype
#>
$lastName = ($args[0].ToString()).replace(' ', '')
$firstName = ($args[1].ToString()).replace(' ', '')
$lastFourDigitSSN = ($args[2].ToString()).replace(' ', '')
$position = $args[3].ToString()
$building_raw = $args[4].ToString()
$staffType = $args[5].ToString()


$profileDir = profile_dir $staffType
$SamAccountName = SamAccountNm $lastName $firstName
$lastFourDigitSSN = lastFourSSN $lastFourDigitSSN
$fullName = fullNm $lastName $firstName
$building = buildingNm $building_raw


$emailAddress = ("$SamAccountName@milfordschools.org").ToString()
$password = ("milford$lastFourDigitSSN").ToString()
$path = (dirPath $staffType $building_raw) + ", OU=Managed Users, DC=MEVSD, DC=NET"
$description = ($building[0] + "  - $position").ToString()
$homeDirectory = ("\\eagle-nest\$profileDir\$SamAccountName").ToString()
$homeDrive = "Z:"
$userPrincipalName = "$SamAccountName@MEVSD.NET"




# Create User Account
 New-ADUser -Name $fullName -GivenName $firstName -Surname $lastName -DisplayName $fullName `
 -AccountPassword (ConvertTo-SecureString -AsPlainText $password -Force) `
 -SamAccountName $SamAccountName -UserPrincipalName $userPrincipalName `
 -Path $path -EmailAddress $emailAddress -Description $description -Company "Milford" `
 -Department $building[1] -Title $position -HomeDirectory  $homeDirectory -HomeDrive $homeDrive `
 -PasswordNeverExpires $True -Enabled $True




# Add user account to default Group
  forEach ($grp in $defaultStaffADGroup) {
    Add-ADGroupMember $grp -members $SamAccountName
    
  }#>

  "$Time Account was successfully created for $fullName" | out-file .\event_log.log -append

  # Return this information to python
  return (1, $emailAddress, "Account Successfully Created")

} catch {

  $ErrorMessage = $_.Exception.Message
  $FailedItem = $_.Exception.ItemName

  "$Time An error occured when trying to create an account for $firstname $lastname. Error Message: $ErrorMessage" | out-file event_log.log -append
  $FailedItem | out-file .\event_log.log -append
  
  # Return this information to python
  return (2, '', "Could not create account, Check event log for details")
  
  Break

}<# Finally {
  
  
  "This script made a read attempt at $Time" | out-file .\event_log.log -append

}#>






