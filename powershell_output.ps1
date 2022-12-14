Import-Module ActiveDirectory
$user ='TEST.CLONE'
Get-ADUser -Filter 'samAccountName -like $user' | ForEach-Object{ $DN=$_.distinguishedname -split',' 
$clone_location =$DN[1..($DN.count -1)] -join ','} 
$ou_path = $clone_location 
$New_Starter = New-ADUser -Name "TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME"  -ChangePasswordAtLogon $true  -GivenName TEST_USER_FIRST_NAME  -Surname TEST_USER_LAST_NAME  -SamAccountName TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME  -UserPrincipalName TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME@testcompany.com  -Path $ou_path  -AccountPassword(ConvertTo-SecureString -AsPlainText "ValidPassword1234CZ!" -Force)  -PassThru | Enable-ADAccount 
$new_starter_sam_account = "TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME"
$new_starter_name = "TEST_USER_FIRST_NAME TEST_USER_LAST_NAME"

$SourceUsersGroup = "TEST.CLONE" 
$DestinationUser = $new_starter_sam_account 
$sourceUserMemberOf = Get-ADUser $SourceUsersGroup -Properties MemberOf | Select-Object -ExpandProperty MemberOf 

foreach($group in $SourceUserMemberOf){Get-ADGroup -Identity $group | Add-ADGroupMember -Members $DestinationUser}
$SourceUsersMemberOf = Get-ADUser $DestinationUser -Properties MemberOf | Select-Object -ExpandProperty memberof 
foreach($group in $SourceUsersMemberOf){Get-ADGroup -Identity $group | Select-Object -ExpandProperty samAccountName}

Set-ADUser TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME -description "28-7-2022" 
Set-ADUser TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME -EmployeeNumber Unknown 
Set-ADUSer TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME -Title "MANAGER_TEST"
Set-ADUser TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME -Manager West North
Set-ADUser TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME -StreetAddress "QQ Street nr 515" 
Set-AdUser TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME -Office "Bahamas"
Set-ADUser TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME -Displayname "TEST_USER_FIRST_NAME TEST_USER_LAST_NAME"
Set-ADUser TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME -Department Sales
Set-ADUser TEST_USER_FIRST_NAME.TEST_USER_LAST_NAME -EmailAddress TEST_USER_FIRST_NAME.TEST_USER_LAST_NAM@testcompany.com
