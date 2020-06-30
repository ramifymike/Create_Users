#requires -Module ActiveDirectory
#requires -RunAsAdministrator

$Creds = Get-Credential -Message "PLEASE ENTER YOUR OFFICE 365 EMAIL ACCOUNT INFORMATION!!"

### UPDATE THESE ###
$ExchangeServer = 'exchange.ramify.local'
$homedrive_location = "\\homeserver\homedrive"
$domain = 'ramify.local'
$homedrive = 'U:'
$default_ou = 'OU=Users,DC=ramify,DC=local'
$sync_server = 'dsc01.ramify.local' #OFFICE 365 SYNC SERVER
$domain_prefix = 'ramify'
### UPDATE ###
  
  
Write-Output "Importing Active Directory Module"
Import-Module ActiveDirectory -ErrorAction Stop
Write-Host "Done..."
Write-Host
Write-Host
 
 
Write-Output "Importing OnPrem Exchange Module"
$OnPrem = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/powershell #-Credential $Creds
Import-PSSession $OnPrem | Out-Null
Write-Host "Done..."
Write-Host
Write-Host
  
  
Start-Sleep 2
Clear-Host
Write-Host "Before we create the account..."
$CopyUser = Read-Host "Would you like to copy from another user? (y/n)"
Write-Host
  
  Do {
  if ($CopyUser -ieq 'y') {
  $CUser = Read-Host "Enter in the USERNAME that you would like to copy FROM"
  Write-Host
  
  
  Write-Host "Checking if $CUser is a valid user..." -ForegroundColor:Green
  If ($(Get-ADUser -Filter {SamAccountName -eq $CUser})) {
  Write-Host "Copying from user account" (Get-ADUser $CUser | Select-Object -ExpandProperty DistinguishedName)
  Write-Host
  
  $Proceed = Read-Host "Continue? (y/n)"
  Write-Host
  
  if ($Proceed -ieq 'y') {
  $CUser = Get-ADUser $CUser -Properties *
  $Exit = $true
  }
  
  } else {
  Write-Host "$CUser was not a valid user" -ForegroundColor:Red
  Start-Sleep 4
  $Exit = $false
  Clear-Host
  }
  
  } else {
  $Exit = $true
  }
  
  } until ($Exit -eq $true)
  
  
Clear-Host
Write-Host "Gathering information for new account creation."
Write-Host
$firstname = Read-Host "Enter in the First Name"
Write-Host
$lastname = Read-Host "Enter in the Last Name"
Write-Host
$middlename = Read-Host "Enter in the Middle Initial"
Write-Host
 
#If middle name is set, put into correct attribute and add period if not applied. 
if($middlename){
  
  if($middlename.Contains('.')){
  continue
  } else { $middlename = "$middlename.".ToUpper()}
 
}
 
$pattern = '[^a-zA-Z]'
$fullname = "$firstname $middlename $lastname"
$fullname = $fullname.Replace('  ', ' ') #fix double whitespace 
 
#Write-Host
$i = 1
$logonname = $firstname.substring(0,$i) + $middlename + $lastname
$logonname = $logonname -replace $pattern,''
 
#clear extra white space when no middle name was created
 
#Write-Host
#$EmployeeID = Read-Host "Enter in the Employee ID"
#Write-Host
$password = ConvertTo-SecureString "defaultpassword" -AsPlainText -Force
  
$server = Get-ADDomain | Select-Object -ExpandProperty PDCEmulator
  
  if ($CUser)
  {
  #Getting OU from the copied User.
  $Object = $CUser | Select-Object -ExpandProperty DistinguishedName
  $pos = $Object.IndexOf(",OU")
  $OU = $Object.Substring($pos+1)
  
  
  #Getting Description from the copied User.
  $Description = $CUser.description
  
  #Getting Office from the copied User.
  $Office = $CUser.Office
  
  #Getting Street Address from the copied User.
  $StreetAddress = $CUser.StreetAddress
  
  #Getting City from copied user.
  $City = $CUser.City
  
  #Getting State from copied user.
  $State = $CUser.State
  
  #Getting PostalCode from copied user.
  $PostalCode = $CUser.PostalCode
  
  #Getting Country from copied user.
  $Country = $CUser.Country
  
  #Getting Title from copied user.
  $Title = $CUser.Title
  
  #Getting Department from copied user.
  $Department = $CUser.Department
  
  #Getting Company from copied user.
  $Company = $CUser.Company
  
  #Getting Manager from copied user.
  $Manager = $CUser.Manager
 
  #Getting Script Path from copied user.
  $scriptpath = $CUser.ScriptPath
  
  #Set Homepath
  $homepath = $homedrive_location + $logonname
 
  #Getting Membership groups from copied user.
  $MemberOf = Get-ADPrincipalGroupMembership $CUser | Where-Object {$_.Name -ine "Domain Users"}
  
  
  } else {
  #Getting the default Users OU for the domain.
  $OU = Get-ADOrganizationalUnit -Identity $default_ou | Select-Object -ExpandProperty Name
  
  }
  
Clear-Host
Write-Host "======================================="
Write-Host
Write-Host "Firstname:  $firstname"
Write-Host "Lastname:  $lastname"
Write-Host "Display name:  $fullname"
Write-Host "Logon name:  $logonname"
Write-Host "Email Address:  $logonname@$domain"
Write-Host "OU:  $OU"
Write-Host "======================================="
  
 
DO
{
If ($(Get-ADUser -Filter {SamAccountName -eq $logonname})) {
  Write-Host "WARNING: Logon name" $logonname.toUpper() "already exists!!" -ForegroundColor:Green
  $i++
  $logonname = $firstname.substring(0,$i) + $lastname
  Write-Host
  Write-Host
  Write-Host "Changing Logon name to" $logonname.toUpper() -ForegroundColor:Green
  Write-Host
  $taken = $true
  Start-Sleep 4
  } else {
  $taken = $false
  }
} Until ($taken -eq $false)
$logonname = $logonname.toLower()
Start-Sleep 3
  
Clear-Host
Write-Host "======================================="
Write-Host
Write-Host "Firstname:  $firstname"
Write-Host "Lastname:  $lastname"
Write-Host "Display name:  $fullname"
Write-Host "Logon name:  $logonname"
Write-Host "Email Address:  $logonname@$domain"
Write-Host "OU:  $OU"
Write-Host "======================================="
Write-Host
  
Write-Host "Continuing will create the AD account and O365 Email." -ForegroundColor:Green
Write-Host
$Proceed = $null
$Proceed = Read-Host "Continue? (y/n)"
  
  if ($Proceed -ieq 'y') {
  
  Write-Host "Creating the O365 mailbox and AD Account."
 
  New-RemoteMailbox -Name $fullname -FirstName $firstname -LastName $lastname -DisplayName $fullname `
  -SamAccountName $logonname -UserPrincipalName $logonname@$domain -PrimarySmtpAddress $logonname@$domain `
  -Password $password -OnPremisesOrganizationalUnit $OU  -DomainController $Server
 
  Write-Host "Done..."
  Write-Host
  Write-Host
  Start-Sleep 15
  
  
  Write-Host "Adding Properties to the new user account."
  Start-Sleep 5
 
  Get-ADUser $logonname -Server $Server | Set-ADUser -Server $Server -Description $Description `
  -Office $Office -StreetAddress $StreetAddress -City $City -State $State -PostalCode $PostalCode `
  -Country $Country -Title $Title -Department $Department -Company $Company -Manager $Manager `
  -EmployeeID $EmployeeID -ScriptPath $scriptpath -HomeDirectory $homepath `
  -HomeDrive $homedrive -ChangePasswordAtLogon:$True
 
  if($middlename){
  Get-ADUser $logonname -Server $server | Set-ADUser -Initials $middlename
  }
 
 
  Write-Host "Done..."
  Write-Host
  Write-Host
  
  if ($MemberOf) {
  Write-Host "Adding Membership Groups to the new user account."
  Get-ADUser $logonname | Add-ADPrincipalGroupMembership -Server $Server -MemberOf $MemberOf
  Write-Host "Done..."
  Write-Host
  Write-Host
  }  
 
  Write-Host "Syncing Office 365... "
  Invoke-Command -ComputerName $sync_server -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }
  Write-Host "Done..."
  Write-Host
  Write-Host
 
  Write-Host "Creating Home Drive... "
 
  #Create Folder
  New-Item "$homedrive_location\$logonname" -ItemType Directory
 
  #Assign Permissions + Inheretence Settings
  $acl = Get-Acl $homedrive_location\$logonname
 
  #set up object
   $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("$domain_prefix\$logonname","FullControl","Allow")
 
  #apply access rule to object
  $acl.SetAccessRule($AccessRule)
 
  #set permissions
  $acl | Set-Acl $homedrive_location\$logonname
 
  #display permissions
  foreach($id in $acl){$id}
 
  Write-Host "Assigning Licenses... "
  Start-Sleep 75
  
  #Import module
  Import-Module MSOnline
 
  #Connect to Microsoft Online Services
  Connect-MsolService -Credential $Creds
 
  #Assign Licenses
  Set-MsolUser -UserPrincipalName $logonname@$domain -UsageLocation US
  Write-Host "Location set to US." -ForegroundColor Green
  Start-Sleep 2
  Set-MsolUserLicense -UserPrincipalName $logonname@$domain -AddLicenses "$domain_prefix:ENTERPRISEPACK"
  Write-Host "E3 Enterprise has been assigned." -ForegroundColor Green
  Start-Sleep 2
  Set-MsolUserLicense -UserPrincipalName $logonname@$domain -AddLicenses "$domain_prefix:ATP_ENTERPRISE"
  Write-Host "Advanced Threat Preventation has been assigned." -ForegroundColor Green
  Start-Sleep 2
  try {Set-MsolUserLicense -UserPrincipalName $logonname@$domain -AddLicenses "$domain_prefix:EMS"
  Write-Host "EMS License has been assigned." -ForegroundColor Green }
  Catch { Set-MsolUserLicense -UserPrincipalName $logonname@$domain -AddLicenses "$domain_prefix:AAD_PREMIUM"
  Write-Host "AAD Premium License has been asigned." -ForegroundColor Green }
 
  Write-Host
  Write-Host "Account has been created!"
 
  }
  
  
Get-PSSession | Remove-PSSession
