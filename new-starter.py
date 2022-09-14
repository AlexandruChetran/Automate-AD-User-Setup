from tkinter import filedialog, ttk

import tkinter as tk
import docx, os, time

#Active Directory library with pre-configured Powershell script in which the relevant variables will be replaced
def active_directory():

    remote_mailbox_address = '@onecwmail.onmicrosoft.com'
    # Not
    new_remote_mailbox = f'New-RemoteMailbox {new_user.first_name}.{new_user.last_name} -RemoteRoutingAddress {new_user.first_name}.{new_user.last_name}{remote_mailbox_address} ' \
     f'-UserPrincipalName {new_user.first_name}.{new_user.last_name} -PrimarySmtpAddress @testcompany.com \n'
    import_AD_module = 'Import-Module ActiveDirectory\n'
    # CHANGED
    user_path_defined = path = "$user =" +"'"+new_user.clone+"'"+'\n'\
                        "Get-ADUser -Filter 'samAccountName -like $user' | ForEach-Object{ $DN=$_.distinguishedname -split',' \n" \
                        "$clone_location =$DN[1..($DN.count -1)] -join ','} \n"
    # Distinguished Name Path
    ou_path = f'$ou_path = $clone_location '
    create_new_ad_user = f'$New_Starter = New-ADUser -Name "{new_user.first_name}.{new_user.last_name}" ' \
                         f' -ChangePasswordAtLogon $true ' \
                         f' -GivenName {new_user.first_name} ' \
                         f' -Surname {new_user.last_name} ' \
                         f' -SamAccountName {new_user.first_name}.{new_user.last_name} ' \
                         f' -UserPrincipalName {new_user.email_address } ' \
                         f' -Path $ou_path ' \
                         f' -AccountPassword(ConvertTo-SecureString -AsPlainText "ValidPassword1234CZ!" -Force) ' \
                         f' -PassThru | Enable-ADAccount \n'
    new_starter_sam_account = f'$new_starter_sam_account = "{new_user.first_name}.{new_user.last_name}"\n'
    new_starter_name = f'$new_starter_name = "{new_user.first_name} {new_user.last_name}"\n'

    source_user_groups = f'$SourceUsersGroup = "{new_user.clone}" \n'
    destination_user = f'$DestinationUser = $new_starter_sam_account \n'
    get_ad_source_user_member_of = f'Get-ADUser $SourceUsersGroup -Properties MemberOf | Select-Object -ExpandProperty MemberOf \n'
    copy_memberof_from_user = f'$sourceUserMemberOf ={get_ad_source_user_member_of}\n'

    # Loop through the member groups in ad (  activedir + loop member of = same command splitted in two chunks)#
    active_diectory_group = "{Get-ADGroup -Identity $group | Add-ADGroupMember -Members $DestinationUser}"
    loop_memberof = f'foreach($group in $SourceUserMemberOf){active_diectory_group}\n'

    set_AD_employee_number = f'Set-ADUser {new_user.first_name}.{new_user.last_name} -EmployeeNumber Unknown \n'
    ad_user_description = f'Set-ADUser {new_user.first_name}.{new_user.last_name} -description "{set_ad_user_description}" \n'
    ad_user_street = f'Set-ADUser {new_user.first_name}.{new_user.last_name} -StreetAddress "{new_user.street_address}" \n'
    ad_user_manager = f'Set-ADUser {new_user.first_name}.{new_user.last_name} -Manager {new_user.line_manager}\n'
    ad_user_job_title = f'Set-ADUSer {new_user.first_name}.{new_user.last_name} -Title "{new_user.job_title}"\n'
    ad_user_office = f'Set-AdUser {new_user.first_name}.{new_user.last_name} -Office "{new_user.office_address}"\n'
    ad_user_department = f'Set-ADUser {new_user.first_name}.{new_user.last_name} -Department {new_user.department}\n'
    ad_user_display_name = f'Set-ADUser {new_user.first_name}.{new_user.last_name} -Displayname "{new_user.first_name} {new_user.last_name}"\n'
    ad_user_email_address = f'Set-ADUser {new_user.first_name}.{new_user.last_name} -EmailAddress {new_user.email_address}\n'
    ad_source_user_member_of = f'$SourceUserMemberof = $Get-AdUser $sourceUserGroup -Properties MemberOf ' \
                               f'| Select-Object -ExpandProperty MemberOf \n'
    ad_groups_loop = 'foreach($group in $sourceUserMemberof)' \
                     '{Get-AdGroup -Identity $Group | Add-ADGroupMember -Members' \
                     '$DestinationUser}\n'
    destination_user_member_of = f'$SourceUsersMemberOf = Get-ADUser $DestinationUser -Properties MemberOf ' \
                                 f'| Select-Object -ExpandProperty memberof \n'
    ad_user_groups_copied = "foreach($group in $SourceUsersMemberOf){Get-ADGroup -Identity $group | Select-Object -ExpandProperty samAccountName}\n"
    space_between_rows = "\n"
    
    
    with open('powershell_output.ps1', 'w') as module:
        # module.write(new_remote_mailbox)
        module.write(import_AD_module)
        module.write(user_path_defined)
        module.write(ou_path)
        module.write(space_between_rows)

        module.write(create_new_ad_user)
        module.write(new_starter_sam_account)
        module.write(new_starter_name)
        module.write(space_between_rows)

        module.write(source_user_groups)
        module.write(destination_user)
        module.write(copy_memberof_from_user)
        module.write(loop_memberof)
        # module.write(destination_user)#
        module.write(destination_user_member_of)
        module.write(ad_user_groups_copied)
        module.write(space_between_rows)

        module.write(ad_user_description)
        module.write(set_AD_employee_number)
        module.write(ad_user_job_title)
        module.write(ad_user_manager)
        module.write(ad_user_street)
        module.write(ad_user_office)
        module.write(ad_user_display_name)
        module.write(ad_user_department)
        module.write(ad_user_email_address)

# New User main attributes that have to be extracted from the docx file
class new_starter:
    def __init__(self, first_name,last_name,job_title,email_address,
                 line_manager,street_address,office_address,department,clone):

        self.first_name = first_name
        self.last_name = last_name
        self.job_title = job_title
        self.email_address = email_address
        self.line_manager = line_manager
        self.street_address = street_address
        self.office_address = office_address
        self.department = department
        self.clone = clone


    def __repr__(self):
        return f"Employee Name : {self.first_name} \n" \
               f"Employee Last Name: {self.last_name} \n" \
               f"Employee Job Title : {self.job_title} \n" \
               f"Employee Email: {self.email_address} \n " \
               f"Employee Line Manager: {self.line_manager} \n" \
               f"Employee Street_address: {self.street_address} \n" \
               f"Employee Office Address: {self.office_address} \n" \
               f"Employee Department: {self.department} \n" \
               f"Employee Clone: {self.clone}"


# File dialog for opening the docx files.
def openFile():
    root = tk.Tk()
    frm = ttk.Frame(root,padding=10)
    frm.grid()
    ttk.Label(frm, text ="New Starter").grid(column=0,row=0)
    ttk.Button(frm, text="Start", command=root.destroy).grid(column=1, row=0)
    ttk.Button(frm, text="Quit", command=root.destroy).grid(column=1, row=1)
    canvas = tk.Canvas(root, width=300, height=100)
    canvas.grid(columnspan=3, rowspan=3)
    filepath = filedialog.askopenfilename(parent=root,title="Open the Word Document")
    root.mainloop()
    return filepath
# Splits the Word  document into a string

# Takes the input a a file path from the openFile function above
# Iterates through each paragraph then returns a string with the relevant data
def word_document_input(docx_file):
    doc = docx.Document(docx_file)
    completedText = []
    for paragraph in doc.paragraphs:
        completedText.append(paragraph.text)
    text_output =''.join(completedText)
    divide_text = text_output.split(':')
    result = [x.strip(' ') for x in divide_text]
    return result

# Extracts the relevant user attributes from the docx document
def extract_user_info():
   text = word_document_input(openFile())
   extract_user_attributes = []
   text_lenght = len(text)
   if text_lenght % 2 != 0:
       text_lenght -=1
   for i in range(text_lenght):
       if i % 2  == 0:
           extract_user_attributes.append(text[i+1])
   return extract_user_attributes


def _time_():
    now = time.gmtime()
    time_map = f'{now[2]}-{now[1]}-{now[0]}'
    return time_map


if __name__ == '__main__':

        set_ad_user_description = _time_()
        user_attributes_list = extract_user_info()
        new_user = new_starter(user_attributes_list[0],user_attributes_list[1],user_attributes_list[2],
                               user_attributes_list[3],user_attributes_list[4],user_attributes_list[5],
                               user_attributes_list[6],user_attributes_list[7],user_attributes_list[8])
        test = new_user.email_address
        print(new_user)
        active_directory()
        os.system('powershell_output.ps1')
