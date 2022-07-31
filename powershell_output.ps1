Import-Module ActiveDirectory
$user ='Jack.Sparrow'
Get-ADUser -Filter 'samAccountName -like $user' | ForEach-Object{ $DN=$_.distinguishedname -split',' 
$clone_location =$DN[1..($DN.count -1)] -join ','} 
$ou_path = $clone_location 
$New_Starter = New-ADUser -Name "Salma.Hayek"  -ChangePasswordAtLogon $true  -GivenName Salma  -Surname Hayek  -SamAccountName Salma.Hayek  -UserPrincipalName Salma.Hayek@testcompany.com  -Path $ou_path  -AccountPassword(ConvertTo-SecureString -AsPlainText "ValidPassword1234CZ!" -Force)  -PassThru | Enable-ADAccount 
$new_starter_sam_account = "Salma.Hayek"
$new_starter_name = "Salma Hayek"

$SourceUsersGroup = "Jack.Sparrow" 
$DestinationUser = $new_starter_sam_account 
$sourceUserMemberOf =Get-ADUser $SourceUsersGroup -Properties MemberOf | Select-Object -ExpandProperty MemberOf 

foreach($group in $SourceUserMemberOf){Get-ADGroup -Identity $group | Add-ADGroupMember -Members $DestinationUser}
$SourceUsersMemberOf = Get-ADUser $DestinationUser -Properties MemberOf | Select-Object -ExpandProperty memberof 
foreach($group in $SourceUsersMemberOf){Get-ADGroup -Identity $group | Select-Object -ExpandProperty samAccountName}

Set-ADUser Salma.Hayek -description "28-7-2022" 
Set-ADUser Salma.Hayek -EmployeeNumber Unknown 
Set-ADUSer Salma.Hayek -Title "Beauty Advisor"
Set-ADUser Salma.Hayek -Manager West North
Set-ADUser Salma.Hayek -StreetAddress "QQ Street nr 515" 
Set-AdUser Salma.Hayek -Office "Bahamas"
Set-ADUser Salma.Hayek -Displayname "Salma Hayek"
Set-ADUser Salma.Hayek -Department Sales
Set-ADUser Salma.Hayek -EmailAddress Salma.Hayek@testcompany.com
