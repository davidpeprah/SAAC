<#

   Author: David Peprah

#>

import-Module ActiveDirectory


$Time= (Get-Date)


# This is used to determine the directory for the user's profile
function profile_dir($Department, $building) {  

  if ($building -eq "BOE") {

      return "BOE"
  }

  if ($Department -eq "Teachers/Subs/Aides") {

      return "Teachers"
  }

  if ($Department -eq "Administrators") {

      return "Administrators"
  }


  return "Operations"
}



<#
 This return the username to be used after checking the AD to make sure
 its not being used already by another staff
#>
function SamAccountNm($lastName, $firstName) {
  
  $proposeAccName = $firstName.Tolower() + "." + $lastName.ToLower()

  if ((chrCount $proposeAccName) -lt 21) {
       
       for ($i=1; $i -lt 6; $i++) {
    
            if (-Not(get-aduser -Filter {SamAccountName -eq $proposeAccName})) {return $proposeAccName}
            if ($i -gt 1) {$proposeAccName + ($i - 1)}
       }

   
   } else {
       if (-Not($firstName.Length -eq 1)) {
             $proposeAccName = SamAccountNm $lastName $firstName.Substring(0, $firstName.Length-1)
             return $proposeAccName 

       } else {

             $proposeAccName = SamAccountNm $lastName.Substring(0, $lastName.Length-1) $firstName
             return $proposeAccName

       }

   }
 }


 <#
 Check to make sure the email address is not already inuse
 if the email address is already inuse, the second character of
 the firstname will be added to the first initial until the third character
 If the unique email address has still not been found yet, the script will
 use the full firstname.
 #>
function EmailAddNm($lastName, $firstName) {
   
  for ($i=1; $i -lt 4; $i++) {

    $proposeAccName =  $firstName.Tolower().Substring(0,$i) + $lastName.ToLower()
    $fmail = "$proposeAccName@hcsdoh.org"
    if (-Not(get-aduser -Filter {mail -eq $fmail})) {return $proposeAccName}

  }

   $proposeAccName = $firstName.Tolower() + $lastName.ToLower()
   $fmail = "$proposeAccName@hcsdoh.org"
   if (-Not(get-aduser -Filter {mail -eq $fmail})) {return $proposeAccName} else {break}



}


# Returns Fullname
function fullNm($lastName, $firstName) {

    $proposefullName = "$firstName $lastName"
    if (-Not(get-aduser -Filter {Name -eq $proposefullName})) {return $proposefullName} 

    $proposefullName = "$lastName $firstName"
    if (-Not(get-aduser -Filter {Name -eq $proposefullName})) {return $proposefullName}

    return "$firstName $lastName"
   
}

# Returns character count
function chrCount($str) {
  
  $measureObject = $str | Measure-Object -Character;
  $count = $measureObject.Characters;
  return $count;

}



<#
This function returns the fullname and the abbreviation for a building
to be used for description and department
#>

function buildingNm($building) {
   
   switch -Wildcard ($building) {

   "BOE" {"HBOE","Board of Education"; break}
   "Bridgeport" {"HABP", "Bridgeport ES"; break}
   "Brookwood" {"HABW", "Brookwood ES"; break}
   "Crawford Woods" {"HACW", "Crawford Woods ES"; break}
   "Fairwood" {"HAFW", "Fairwood ES"; break}
   "Freshman" {"HAFS", "Freshman School"; break}
   "Garfield" {"HAGF", "Garfield MS"; break}
   "Highland" {"HAHL", "Highland ES"; break}
   "High School" {"HAHS", "High School/CTE"; break}
   "Linden" {"HALN", "Linden ES"; break}
   "Miami School" {"HAMI", "Miami School"; break}
   "Ridgeway" {"HARW", "Ridgeway ES"; break}
   "Riverview" {"HARV", "Riverview ES"; break}
   "Wilson" {"HAWS", "Wilson MS"; break}
   default {"HAM","School District"}
   }

}

<#
This function returns OU path to be used to create a user.
by default ti will return "OU=Classified" its not able to 
determine the proper OU for staff
#>
function dirPath($Department, $building) {
 
   $Operations = ("Food Services", "Maintenance", "Transportation")

   if (($Department -eq "Teachers/Subs/Aides") -and ($building -eq "High School")) {

        return "OU=Teachers, OU=Users, OU=HS & JDC"

   }
  
  if ( $Department -eq "Teachers/Subs/Aides") {

        return "OU=Teachers, OU=Users, OU=$building"

     }


  if ( $Department -eq "Treasure") {

     return "OU=Treasurer NEW, OU=Users, OU=$building"
  }


  if (($Department -in $Operations) -and ($building -eq "District")) {
      
     return "OU=Users, OU=$Department"

  }

  if (($Department -eq "Teachers/Subs/Aides") -and ($building -eq "District")){
     return "OU=Sub Aides, OU=Subs & Misc Employees"
  }

  if ($Department -eq "Student Services"){
     return "OU=Student Services,OU=Users,OU=BOE"
  }

  if ($Department -in $Operations) {
     return "OU=Operations, OU=Users, OU=$building"
  }
   return "OU=$Department, OU=Users, OU=$building"
 
}


function defaultStaffGroups($Department, $Building) {

    $defaultStaffgroups = ("DL-HCSDSTAFF", "HCS-OfflineFiles", "password-reset", "TechGuard")

    if ($Department -eq "Administrators") {
       
       $defaultStaffgroups += "HCS-Administrators", "DL-ADMINISTRATORS"

       switch -Exact ($Building) {
        
           "Bridgeport" {$defaultStaffgroups += "BRGPT-Administrators" ; break}
           "Brookwood" {$defaultStaffgroups += "BRKWD-Administrators"; break}
           "Crawford Woods" {$defaultStaffgroups += "CW-Administrators" ; break}
           "Fairwood" {$defaultStaffgroups += "FRWD-Administrators" ; break}
           "Freshman" {$defaultStaffgroups += "FHS-Administrators"; break}
           "Garfield" {$defaultStaffgroups += "GMS-Administrators"; break}
           "Highland" {$defaultStaffgroups += "HGHLND-Administrators"; break}
           "High School" {$defaultStaffgroups += "HS-Administrators"; break}
           "Linden" {$defaultStaffgroups += "LNDN-Administrators"; break}
           "Miami School" {$defaultStaffgroups += "MIA-Administrators"; break}
           "Ridgeway" {$defaultStaffgroups += "RDGWY-Administrators"; break}
           "Riverview" {$defaultStaffgroups += "RVRVW-Administrators"; break}
           "Wilson" {$defaultStaffgroups += "WMS-Administrators"; break}

       }


    } elseif ($Department -eq "Operations") {
      
       $defaultStaffgroups += "HCS-Operations"

       switch -Exact ($Building) {
        
           "BOE" {$defaultStaffgroups += "BOE-Operations"; break}
           "Bridgeport" {$defaultStaffgroups += "BRGPT-Operations" ; break}
           "Brookwood" {$defaultStaffgroups +="BRKWD-Operations"; break}
           "Crawford Woods" {$defaultStaffgroups +="CW-Operations"; break}
           "Fairwood" {$defaultStaffgroups += "FRWD-Operations"; break}
           "Freshman" {$defaultStaffgroups += "FHS-Operations"; break}
           "Garfield" {$defaultStaffgroups += "GMS-Operations"; break}
           "Highland" {$defaultStaffgroups += "HGHLND-Operations"; break}
           "High School" {$defaultStaffgroups += "HS-Operations"; break}
           "Linden" {$defaultStaffgroups += "LNDN-Operations"; break}
           "Miami School" {$defaultStaffgroups += "MIA-Operations"; break}
           "Ridgeway" {$defaultStaffgroups += "RDGWY-Operations"; break}
           "Riverview" {$defaultStaffgroups += "RVRVW-Operations"; break}
           "Wilson" {$defaultStaffgroups += "WMS-Operations"; break}

       }

    } elseif ($Department -eq "Teachers/Subs/Aides") {

       $defaultStaffgroups += "HCS-Teachers"

       switch -Exact ($Building) {
        
           "Bridgeport" {$defaultStaffgroups += "BRGPT-Teachers" ; break}
           "Brookwood" {$defaultStaffgroups += "BRKWD-Teachers"; break}
           "Crawford Woods" {$defaultStaffgroups +="CW-Teachers"; break}
           "Fairwood" {$defaultStaffgroups += "FRWD-Teachers"; break}
           "Freshman" {$defaultStaffgroups += "FHS-Teachers"; break}
           "Garfield" {$defaultStaffgroups += "GMS-Teachers"; break}
           "Highland" {$defaultStaffgroups += "HGHLND-Teachers"; break}
           "High School" {$defaultStaffgroups += "HS-Teachers", "DL-HHSTeachers"; break}
           "Linden" {$defaultStaffgroups += "LNDN-Teachers"; break}
           "Miami School" {$defaultStaffgroups += "MIA-Teachers"; break}
           "Ridgeway" {$defaultStaffgroups += "RDGWY-Teachers"; break}
           "Riverview" {$defaultStaffgroups += "RVRVW-Teachers"; break}
           "Wilson" {$defaultStaffgroups += "WMS-Teachers"; break}

       }

    }


    if ($Department -eq "Transportation") {

       $defaultStaffgroups += "HCS-TRANS", "TRANS-USERS"

    } elseif ($Department -eq "Maintenance") {

       $defaultStaffgroups += "MAINT-Users"

    } elseif ($Department -eq "Food Service") {

       $defaultStaffgroups += "FOOD-Users", "HCS-FOOD"

    }

    if ($Department -eq "Board Members") {

        $defaultStaffgroups += "boemembers", "HCS-VPN"
     }

     if ($Department -eq "Business and Planning") {

        $defaultStaffgroups += "BOE-Business and Planning"
     }

     if ($Department -eq "Human Resources") {

        $defaultStaffgroups += "BOE-Human Resources"
     }

     if ($Department -eq "Instructional Services") {

        $defaultStaffgroups += "BOE-Instructional Services", "DL-INSTRUCTIONS"
     }


     if ($Department -eq "Pupil Personnel") {

        $defaultStaffgroups += "BOE-Pupil Personnel"
     }
   

     if ($Department -eq "Treasure") {

        $defaultStaffgroups += "BOE-Treasury", "HCS-VPN"
     }
   return $defaultStaffgroups

}

function defaultBuildingGroups($Building) {

    switch -Exact ($Building) {
       
       "BOE" {"DL-BOESTAFF", "HCS-BOE", "HCS-BOE-Staff"; break}
       "Bridgeport" {"HCS-BRGPT","DL-BRIDGEPORTSTAFF"; break}
       "Brookwood" {"HCS-BRKWD", "DL-BROOKWOODSTAFF"; break}
       "Crawford Woods" {"HCS-CW","DL-CRAWFORDWOODSSTAFF"; break}
       "Fairwood" {"HCS-FRWD","DL-FAIRWOODSTAFF"; break}
       "Freshman" {"HCS-FHS","DL-FRESHMANSTAFF"; break}
       "Garfield" {"HCS-GMS","DL-GARFIELDSTAFF"; break}
       "Highland" {"HCS-HGHLND","DL-HIGHLANDSTAFF"; break}
       "High School" {"HCS-HS","DL-HIGHSCHOOLSTAFF"; break}
       "Linden" {"HCS-LNDN","DL-LINDENSTAFF"; break}
       "Miami School" {"HCS-MIA","DL-MiamiSchoolStaff"; break}
       "Ridgeway" {"HCS-RDGWY","DL-RIDGEWAYSTAFF"; break}
       "Riverview" {"HCS-RVRVW","DL-RIVERVIEWSTAFF"; break}
       "Wilson" {"HCS-WMS","DL-WILSONSTAFF"; break}
       "District" {""; break}

    }

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



$emailAddress = ("$EmailAddName@hcsdoh.org").ToString()
$password = ("P@ssw0rd@!!").ToString() # This password will change once the account is confirmed in Google console
$path = (dirPath $department $building_raw) + ", OU=HCS, DC=hcs, DC=local"
$description = ($position).ToString()
$homeDirectory = ("\\hcs.local\$profileDir\$SamAccountName").ToString()
$homeDrive = "H:"
$userPrincipalName = "$SamAccountName@hcs.local"
$Build = "Hamilton " + $building[1]

# Groups
$defaultBuildADGroup = defaultBuildingGroups $building_raw
$defaultdepartADGroup = defaultStaffGroups $department $building_raw



# Create User Account
 New-ADUser -Name $fullName -GivenName $firstName -Surname $lastName -DisplayName $fullName `
 -AccountPassword (ConvertTo-SecureString -AsPlainText $password -Force) `
 -SamAccountName $SamAccountName -UserPrincipalName $userPrincipalName `
 -Path $path -EmailAddress $emailAddress -Description $description -Company "Hamilton" `
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






