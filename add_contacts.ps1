# Add Contacts to Active Directory
# Author: John Patrick McCarthy
# <http://www.midactstech.blogspot.com> <https://www.github.com/Midacts>
# Date: 24th February, 2014
# Version 1.0
#
# To God only wise, be glory through Jesus Christ forever. Amen.
#
# Romans 16:27 ; I Corinthians 15:1-4
#----------------------------------------------------------------
#---------------------#
#		      #
#  Creating Contacts  #
# =================== #
#		      #
#---------------------#
# This first section must be done on the exchange management console, 
# or remotely if you have the Microsoft.Exchange.Management.PowerShell.Admin module
# Variables
$name=@()
$first=@()
$last=@()
$email=@()
# Text file containing the contact's first name
$first = get-content C:\users\USERNAME\desktop\contact_first.txt
# Text file containing the contact's last name
$last = get-content C:\users\USERNAME\desktop\contact_last.txt
# Text file containing the contact's email
$email = get-content C:\users\USERNAME\desktop\contact_email.txt

# Combines the first and last names to get the user's full name
$i=0
$fullname = @()
foreach ($name in $first){
$result = $first[$i] + " " + $last[$i]
$fullname += $result
$i++
}

# Creates the contact- with their email, first and lastname
$i=0
foreach ($member in $fullname){
new-mailcontact -name $member[$i] -externalemailaddress $email[$i] -organizationalunit "Contacts" -firstname $first[$i] -lastname $last[$i]
$i ++
}

#--------------------------------#
#				 #
#  Adding Contact's Description  #
# ============================== #
#				 #
#--------------------------------#
# Variables
$name=@()
$first=@()
$last=@()
$email=@()
# Text file containing the contact's first name
$first = get-content C:\users\USERNAME\desktop\contact_first.txt
# Text file containing the contact's last name
$last = get-content C:\users\USERNAME\desktop\contact_last.txt
# Text file containing the contact's email
$email = get-content C:\users\USERNAME\desktop\contact_email.txt

# Combines the first and last names to get the user's full name
$i=0
$fullname = @()
foreach ($name in $first){
$result = $first[$i] + " " + $last[$i]
$fullname += $result
$i++
}

# This goes through and gets the company (or domain) of each of the contact's emails
$i = 0
$company_lc = @()
foreach ($addr in $email){
$result = $email[$i].split('@')[1].split('.')[0]
$company_lc += $result
$i ++
}

# Capitalizes the first letter of each company
$company = @()
foreach ($name in $company_lc){
$result = $name.substring(0,1).toupper()+$name.substring(1).tolower()
$company += $result
}

# Adds a Description to the contact based on their email address domain
$i=0
foreach ($name in $fullname){
set-adobject -identity "CN=$name,OU=Contacts,DC=DOMAIN,DC=LOCAL" -description $company[$i]
$i++
}

#----------------------------------------#
#					 #
#  Adding contact to distribution group  #
# ====================================== #
#					 #
#----------------------------------------#
# Adds all the users in the $fullname array to the $dist_grp Distribution Group
# Distribution Group Name
$dist_grp="NAME OF DISTRIBUTION GROUP HERE"

foreach ($contact in $fullname){
add-qadgroupmember -identity $dist_grp -member "CN=$contact,OU=Contacts,dc=DOMAIN,dc=LOCAL"
}

#-------------#
#	      #
#  Resources  #
# =========== #
#	      #
#-------------#
# Getting the company name (between '@' and '.' in an email)
#	http://stackoverflow.com/questions/19168475/powershell-to-remove-text-from-a-string
# Capitalizing the first letter in a string in Powershell
#	http://rafdelgado.blogspot.com/2012/05/powershell-converting-first-character.html
# Interpolating arrays inside double quotes
#	http://stackoverflow.com/questions/9196525/poweshell-outputting-array-items-when-interpolating-within-double-quotes
