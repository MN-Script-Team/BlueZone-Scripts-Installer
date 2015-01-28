Imports System
Imports System.IO
Imports System.Text
Imports Microsoft.Win32

Public Class scripts_config_form

    Const read_only = 1
    Const read_write = 2



    Private Shared Sub Get45or451FromRegistry(dot_net_from_registry)

    End Sub

    Private Shared Function CheckFor45DotVersion(releaseKey As Integer) As String
        If (releaseKey >= 379893) Then
            Return "4.5.2 or later"
        End If
        If (releaseKey >= 378675) Then
            Return "4.5.1 or later"
        End If
        If (releaseKey >= 378389) Then
            Return "4.5 or later"
        End If
        ' This line should never execute. A non-null release key should mean 
        ' that 4.5 or later is installed. 
        Return "No 4.5 or later version detected"
    End Function

    Private Property list_of_addresses As String

    Private Property fso As Object

    Private Property FSO_new_file_path As Object

    Private Property ObjFSO As Object

    Private Property objFile As Object

    Private Property warning As MsgBoxResult

    Private Property list_of_files_array As String()

    Private Property text_file As String

    Private Property new_text_file As Object

    Private Property GitHub_current_zip_archive As String

    Private Property objXMLHTTP As Object

    Private Property objADOStream As Object

    Private Property objShell As Object

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Dim frmAbout As New AboutBox
        frmAbout.ShowDialog(Me)
    End Sub

    Private Sub EDMS_choice_TextClear(ByVal sender As Object, ByVal e As System.EventArgs) Handles EDMS_choice.Enter
        EDMS_choice.Text = ""
    End Sub

    Private Sub run_configuration_button_Click(sender As Object, e As EventArgs) Handles run_configuration_button.Click



        'FIRST IT CONFIGURES THE ADDRESS----------------------------------------------------------------------------------------
        'Dims the address array as nothing
        Dim list_of_addresses = Nothing

        'Takes each line of the address array and splits it with a "|", and each address with a "~". FUNCTIONS FILE will take it from here and split it.
        If address_entry_form.addr_01_line_01.Text <> "" And address_entry_form.addr_01_line_02.Text <> "" Then list_of_addresses = list_of_addresses & address_entry_form.addr_01_line_01.Text & "|" & address_entry_form.addr_01_line_02.Text & "~"
        If address_entry_form.addr_02_line_01.Text <> "" And address_entry_form.addr_02_line_02.Text <> "" Then list_of_addresses = list_of_addresses & address_entry_form.addr_02_line_01.Text & "|" & address_entry_form.addr_02_line_02.Text & "~"
        If address_entry_form.addr_03_line_01.Text <> "" And address_entry_form.addr_03_line_02.Text <> "" Then list_of_addresses = list_of_addresses & address_entry_form.addr_03_line_01.Text & "|" & address_entry_form.addr_03_line_02.Text & "~"
        If address_entry_form.addr_04_line_01.Text <> "" And address_entry_form.addr_04_line_02.Text <> "" Then list_of_addresses = list_of_addresses & address_entry_form.addr_04_line_01.Text & "|" & address_entry_form.addr_04_line_02.Text & "~"
        If address_entry_form.addr_05_line_01.Text <> "" And address_entry_form.addr_05_line_02.Text <> "" Then list_of_addresses = list_of_addresses & address_entry_form.addr_05_line_01.Text & "|" & address_entry_form.addr_05_line_02.Text & "~"
        If address_entry_form.addr_06_line_01.Text <> "" And address_entry_form.addr_06_line_02.Text <> "" Then list_of_addresses = list_of_addresses & address_entry_form.addr_06_line_01.Text & "|" & address_entry_form.addr_06_line_02.Text & "~"
        If address_entry_form.addr_07_line_01.Text <> "" And address_entry_form.addr_07_line_02.Text <> "" Then list_of_addresses = list_of_addresses & address_entry_form.addr_07_line_01.Text & "|" & address_entry_form.addr_07_line_02.Text & "~"

        'Warning if a county or file path is not selected
        If county_selection.Text = "" Or list_of_addresses = Nothing Or location_to_save_script_files.Text = "" Then
            MsgBox("You must select a county, and enter a complete county address, as well as enter a file location.")
            Exit Sub
        End If

        'Removing the last tilde as it is stupid and not required.
        list_of_addresses = list_of_addresses.Remove(list_of_addresses.Length - 1)




        'ASKS IF YOU ARE A BETA COUNTY, AND WANT TO INSTALL BETA SCRIPTS--------------------------------------------------------
        '   Only some agencies get the option to install beta scripts, based mainly on their status as contributing agencies (they write scripts).
        '   As a result, we have an if/then list. It will need to be updated frequently as agencies join/leave the beta program.

        'Here's where the beta list is kept. Just looks at master branch (shouldn't be different between branches).
        Dim beta_county_list_URL = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script%20Files/LIST%20OF%20BETA%20COUNTIES.txt"

        'We need this for the list of counties
        Dim beta_county_list

        'We need this for the MsgBox
        Dim agency_is_beta

        'Loads the list from GitHub
        Dim beta_county_req = CreateObject("Msxml2.XMLHttp.6.0")            'Creates an object to get a URL
        beta_county_req.open("GET", beta_county_list_URL, False)            'Attempts to open the URL
        beta_county_req.send()                                              'Sends request
        If beta_county_req.Status = 200 Then                                '200 means great success
            fso = CreateObject("Scripting.FileSystemObject")                'Creates an FSO
            beta_county_list = beta_county_req.responseText                 'Loads the list of scripts
        End If

        'Asks user if we're installing beta scripts
        If InStr(beta_county_list, county_selection.Text) Then
            Dim beta_msgbox As String
            beta_msgbox = MsgBox("Your agency is listed as a beta agency. Install beta versions of scripts?", MsgBoxStyle.YesNoCancel)
            If beta_msgbox = vbCancel Then Exit Sub
            If beta_msgbox = vbYes Then
                agency_is_beta = True
            Else
                agency_is_beta = False
            End If
        Else
            agency_is_beta = False
        End If




        'WARNS USER ABOUT WHAT'S GOING TO HAPPEN--------------------------------------------------------------------------------

        'Gives warning about downloads for instances where we're downloading, and gives different warning about manual extraction for instances where we have the folder.
        warning = MsgBox("The following utility will download all of the scripts to the selected directory, and replace the DHS file path with " & _
        "the specified file path. If you choose to move your script directory, you'll have to use this tool again. Are you sure you want to do this?", vbYesNo)
        If warning = vbNo Then Exit Sub

        'Disabling features/enabling others (more of a pretty GUI setup, might rewrite to be even prettier)
        Update_Files_Label.Visible = True
        Tab_Control_Main_Form.Enabled = False
        run_configuration_button.Enabled = False

        'Setting EDMS_choice as DHS eDocs if there is not a local EDMS.
        If EDMS_check.Checked = False Then EDMS_choice.Text = "DHS eDocs"




        'LOADING TEXT FILE OF SCRIPTS TO MAKE FROM GITHUB-----------------------------------------------------------------------

        Dim redirect_list   'We'll need this in a bit
        Dim list_of_scripts_URL 'We'll also need this in a bit
        If agency_is_beta = True Then
            list_of_scripts_URL = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/BETA/Script%20Files/LIST%20OF%20SCRIPTS.txt"
        Else
            list_of_scripts_URL = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/RELEASE/Script%20Files/LIST%20OF%20SCRIPTS.txt"
        End If

        Dim req = CreateObject("Msxml2.XMLHttp.6.0")            'Creates an object to get a URL
        req.open("GET", list_of_scripts_URL, False)             'Attempts to open the URL
        req.send()                                              'Sends request
        If req.Status = 200 Then                                '200 means great success
            fso = CreateObject("Scripting.FileSystemObject")    'Creates an FSO
            redirect_list = req.responseText                    'Loads the list of scripts
        End If
        redirect_list = Split(redirect_list, vbLf)              'Splits redirect list into an array to be used by the next section





        'CREATING ALL OF THE REDIRECTS AND THE AGENCY CUSTOMIZED FOLDER---------------------------------------------------------

        'Declaring variables needed for creating all of these new text files.
        Dim create_VBS_fso
        Dim create_VBS_command
        Dim script_directory

        'Creating each redirect from the script list text file
        For Each redirect_to_make In redirect_list

            'Determines which folder and title the redirect requires
            Dim redirect_path As String
            If InStr(redirect_to_make, "REDIRECT - ACTIONS") Then
                redirect_path = "ACTIONS/ACTIONS - MAIN MENU.vbs"
            ElseIf InStr(redirect_to_make, "REDIRECT - BULK") Then
                redirect_path = "BULK/BULK - MAIN MENU.vbs"
            ElseIf InStr(redirect_to_make, "REDIRECT - DAIL SCRUBBER") Then
                redirect_path = "DAIL/DAIL - DAIL SCRUBBER.vbs"
            ElseIf InStr(redirect_to_make, "REDIRECT - MEMOS") Then
                redirect_path = "MEMOS/MEMOS - MAIN MENU.vbs"
            ElseIf InStr(redirect_to_make, "REDIRECT - NOTES - 0-G") Then
                redirect_path = "NOTES/NOTES - MAIN MENU - 0-G.vbs"
            ElseIf InStr(redirect_to_make, "REDIRECT - NOTES - H-Z") Then
                redirect_path = "NOTES/NOTES - MAIN MENU - H-Z.vbs"
            ElseIf InStr(redirect_to_make, "REDIRECT - NAV") Then
                redirect_path = "NAV/" & Replace(redirect_to_make, "REDIRECT - ", "")
            ElseIf InStr(redirect_to_make, "REDIRECT - OTHER NAV") Then
                redirect_path = "NAV/OTHER NAV/NAV - OTHER NAV MAIN MENU.vbs"
            ElseIf InStr(redirect_to_make, "REDIRECT - AGENCY CUSTOMIZED") Then
                redirect_path = "AGENCY CUSTOMIZED/SOURCE/AGENCY CUSTOMIZED.vbs"
            End If

            'Determines script directory, to be used by the rest of this function (and later).
            If custom_file_path.Text = Nothing Then
                script_directory = location_to_save_script_files.Text & "\Script Files\"
            Else
                'If the path is an unmapped network drive, those slashes are usually reversed. This accounts for that. (Hennepin!!!)
                If InStr(custom_file_path.Text, "/") Then
                    script_directory = custom_file_path.Text & "/Script Files/"
                Else
                    script_directory = custom_file_path.Text & "\Script Files\"
                End If
            End If

            'Here's the file contents for each one
            Dim redirect_file_contents = "'LOADING GLOBAL VARIABLES--------------------------------------------------------------------" & vbCr & _
                "Set run_another_script_fso = CreateObject(""Scripting.FileSystemObject"")" & vbCr & _
                "Set fso_command = run_another_script_fso.OpenTextFile(""" & script_directory & "SETTINGS - GLOBAL VARIABLES.vbs"")" & vbCr & _
                "text_from_the_other_script = fso_command.ReadAll" & vbCr & _
                "fso_command.Close" & vbCr & _
                "Execute text_from_the_other_script" & vbCr & _
                vbCr & _
                "'LOADING SCRIPT" & vbCr & _
                "url = script_repository & ""/" & redirect_path & """" & vbCr & _
                "SET req = CreateObject(""Msxml2.XMLHttp.6.0"")				'Creates an object to get a URL" & vbCr & _
                "req.open ""GET"", url, FALSE									'Attempts to open the URL" & vbCr & _
                "req.send													'Sends request" & vbCr & _
                "IF req.Status = 200 THEN									'200 means great success" & vbCr & _
                "   Set fso = CreateObject(""Scripting.FileSystemObject"")	'Creates an FSO" & vbCr & _
                "	Execute req.responseText								'Executes the script code" & vbCr & _
                "ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script)." & vbCr & _
                "	MsgBox 	""Something has gone wrong. The code stored on GitHub was not able to be reached."" & vbCr &_ " & vbCr & _
                "	vbCr & _" & vbCr & _
                "		""Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com."" & vbCr &_" & vbCr & _
                "		vbCr & _" & vbCr & _
                "		""If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:"" & vbCr &_" & vbCr & _
                "		vbTab & ""- The name of the script you are running."" & vbCr &_" & vbCr & _
                "		vbTab & ""- Whether or not the script is """"erroring out"""" for any other users."" & vbCr &_" & vbCr & _
                "		vbTab & ""- The name and email for an employee from your IT department,"" & vbCr & _" & vbCr & _
                "		vbTab & vbTab & ""responsible for network issues."" & vbCr &_" & vbCr & _
                "		vbTab & ""- The URL indicated below (a screenshot should suffice)."" & vbCr &_" & vbCr & _
                "		vbCr & _" & vbCr & _
                "		""Veronica will work with your IT department to try and solve this issue, if needed."" & vbCr &_ " & vbCr & _
                "		vbCr &_" & vbCr & _
                "		""URL: "" & url" & vbCr & _
                "		StopScript" & vbCr & _
                "END IF"

            'Creates each script file one at a time. If the "Script Files" folder doesn't exist, it'll create it.
            create_VBS_fso = CreateObject("Scripting.FileSystemObject")
            If create_VBS_fso.FolderExists(script_directory) = False Then create_VBS_fso.CreateFolder(script_directory)
            create_VBS_command = create_VBS_fso.CreateTextFile(script_directory & Trim(redirect_to_make), True)
            create_VBS_command.Write(redirect_file_contents)
            create_VBS_command.Close()
            create_VBS_fso = Nothing

            'If the agency customized script is the one we're doing, it'll create the agency customized folder if it doesn't exist yet.
            If InStr(redirect_to_make, "REDIRECT - AGENCY CUSTOMIZED.vbs") Then
                create_VBS_fso = CreateObject("Scripting.FileSystemObject")
                If create_VBS_fso.FolderExists(script_directory & "AGENCY CUSTOMIZED") = False Then create_VBS_fso.CreateFolder(script_directory & "AGENCY CUSTOMIZED")
                create_VBS_command = create_VBS_fso.CreateTextFile(script_directory & "AGENCY CUSTOMIZED\How to use this folder and script.vbs", True)
                create_VBS_command.Write("MsgBox (""This script (and folder) is designed to store any scripts that your county/agency made that are """"customized"""" for your agency (meaning they aren't available statewide for some reason)."" & vbCr & _" & vbCr & _
                                            vbTab & "vbCr & _" & vbCr & _
                                            vbTab & """If your agency has made customized scripts, simply insert them into the folder located at "" & objStartFolder & ""."" & vbCr & _" & vbCr & _
                                            vbTab & "vbCr & _" & vbCr & _
                                            vbTab & """Once you place a script there, this script will """"find"""" it, and run it when selected."" & vbCr & _" & vbCr & _
                                            vbTab & "vbCr & _" & vbCr & _
                                            vbTab & """If you have any questions about how to use this button or folder, please have your alpha user contact Veronica Cary. Thank you!"")")
                create_VBS_command.Close()
                create_VBS_fso = Nothing

                redirect_path = "AGENCY CUSTOMIZED/SOURCE/AGENCY CUSTOMIZED.vbs"
            End If

        Next




        'DOWNLOADING THE POWER PAD----------------------------------------------------------------------------------------------

        'Declares Power Pad variable
        Dim pad_URL

        'URL for Power Pad (depends on if user is beta or not)
        If agency_is_beta = True Then
            pad_URL = "https://github.com/MN-Script-Team/DHS-MAXIS-Scripts/blob/BETA/Script%20Files/PAD%20-%20DHS%20SUPPORTED.pad?raw=true"
        Else
            pad_URL = "https://github.com/MN-Script-Team/DHS-MAXIS-Scripts/blob/RELEASE/Script%20Files/PAD%20-%20DHS%20SUPPORTED.pad?raw=true"
        End If

        'Downloads file
        'Now it downloads the pad file from Github. This code was copied from https://gist.github.com/udawtr/2053179 on 09/13/2014, and modified for our purposes.

        'Creating a server object
        objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")

        'Opening the file
        objXMLHTTP.open("GET", pad_URL, False)
        objXMLHTTP.send()
        If objXMLHTTP.Status = 200 Then
            objADOStream = CreateObject("ADODB.Stream")
            objADOStream.Open()
            objADOStream.Type = 1 'adTypeBinary
            objADOStream.Write(objXMLHTTP.ResponseBody)
            objADOStream.Position = 0    'Set the stream position to the start

            'Writing the file to the hard disk
            ObjFSO = CreateObject("Scripting.FileSystemObject")
            If ObjFSO.Fileexists(script_directory & "PAD - DHS SUPPORTED.pad") Then ObjFSO.DeleteFile(script_directory & "PAD - DHS SUPPORTED.pad")
            ObjFSO = Nothing
            objADOStream.SaveToFile(script_directory & "PAD - DHS SUPPORTED.pad")
            objADOStream.Close()
            objADOStream = Nothing
        End If

        objXMLHTTP = Nothing






        'DOWNLOADING THE GLOBAL VARIABLES FILE AND CONFIGURING IT---------------------------------------------------------------

        'Creating the variable for global variables URL
        Dim global_variables_URL

        'URL for Global Variables (eventually this should be dynamic dependent on the beta/master/other categories, but I need to get this out soon).
        If agency_is_beta = True Then
            global_variables_URL = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/BETA/Script%20Files/SETTINGS%20-%20GLOBAL%20VARIABLES.vbs"
        Else
            global_variables_URL = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/RELEASE/Script%20Files/SETTINGS%20-%20GLOBAL%20VARIABLES.vbs"
        End If


        'Checks if the file exists. If it does, it deletes it.
        If File.Exists(script_directory & "SETTINGS - GLOBAL VARIABLES.vbs") Then File.Delete(script_directory & "SETTINGS - GLOBAL VARIABLES.vbs")

        'Downloads file
        My.Computer.Network.DownloadFile(global_variables_URL, script_directory & "SETTINGS - GLOBAL VARIABLES.vbs")

        'Opening file read-only, to know what path to overwrite from the GitHub repo.
        ObjFSO = CreateObject("Scripting.FileSystemObject")
        objFile = ObjFSO.OpenTextFile(script_directory & "SETTINGS - GLOBAL VARIABLES.vbs", read_only)

        'Reading all lines from it into a string
        Dim text_file() As String = System.IO.File.ReadAllLines(location_to_save_script_files.Text & "\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
        Dim text_line As String

        'Replaces file path and county-specific variables
        For Each text_line In text_file

            'EDMS choice
            If InStr(text_line, "EDMS_choice = ") Then text_line = "EDMS_choice = " & Chr(34) & EDMS_choice.Text & Chr(34)

            'Splits the list_of_addresses into an array, using a tilde as the separating character.
            If InStr(text_line, "county_office_array = split(") Then text_line = "county_office_array = split(" & Chr(34) & list_of_addresses & Chr(34) & ", " & Chr(34) & "~" & Chr(34) & ")"

            'County name (replaces parts of the code_from_installer that aren't needed)
            If InStr(text_line, "county_name = ") Then text_line = "county_name = " & Chr(34) & Strings.Replace(county_selection.Text, Strings.Left(county_selection.Text, 5), "") & Chr(34)

            'Puts the worker county code in, using left two digits of the code_from_installer only. Sets to "MULTICOUNTY" if it's a multicounty site.
            If InStr(text_line, "worker_county_code = ") Then
                If IsNumeric(Strings.Left(county_selection.Text, 2)) = True Then
                    text_line = "worker_county_code = " & Chr(34) & "x1" & Strings.Left(county_selection.Text, 2) & Chr(34)
                Else
                    text_line = "worker_county_code = " & Chr(34) & "MULTICOUNTY" & Chr(34)
                End If
            End If

            'Some counties don't want to case note the intake/reapp dates. This sets that variable.
            If InStr(text_line, "case_noting_intake_dates = ") Then
                If intake_dates_check.Checked = True Then
                    text_line = "case_noting_intake_dates = True"
                Else
                    text_line = "case_noting_intake_dates = False"
                End If
            End If

            'Some counties want "verifs needed" at the top of the CAF case note. This sets that variable.
            If InStr(text_line, "move_verifs_needed = ") Then
                If move_verifs_needed_check.Checked = True Then
                    text_line = "move_verifs_needed = True"
                Else
                    text_line = "move_verifs_needed = False"
                End If
            End If

            'Storing the code from installer (used only for autofill purposes when running the installer, and is not used by any scripts).
            If InStr(text_line, "code_from_installer = ") Then text_line = "code_from_installer = " & Chr(34) & county_selection.Text & Chr(34)

            'County BNDX threshold (defaults to 0). Used by BNDX scrubber to decide the "range" of acceptable change.
            If InStr(text_line, "county_bndx_variance_threshold = ") Then text_line = "county_bndx_variance_threshold = " & Chr(34) & bndx_threshold.Text & Chr(34)

            'Emergency percent rule amount is the amount of income which needs to have been spent in certain categories. Varies county-to-county. This sets it.
            If InStr(text_line, "emer_percent_rule_amt = ") Then text_line = "emer_percent_rule_amt = " & Chr(34) & emer_percent_rule_number.Text & Chr(34)

            'Most counties use 30 days of income when determining emergency program eligibility. Some use more or less. This sets that amount.
            If InStr(text_line, "emer_number_of_income_days = ") Then text_line = "emer_number_of_income_days = " & Chr(34) & emer_number_of_income_days.Text & Chr(34)

            'Some counties use a "CLS" account to send closed cases. This sets that account ID.
            If InStr(text_line, "CLS_x1_number = ") Then text_line = "CLS_x1_number = " & Chr(34) & X1_for_CLS.Text & Chr(34)

            'Modifes the directory ONLY IF it's the default_directory variable in GLOBAL VARIABLES.
            If InStr(text_line, "default_directory = ""C:\DHS-MAXIS-Scripts\Script Files\""") Then text_line = "default_directory = """ & script_directory & """"

            'Sets the all_users_select_a_worker option.
            If InStr(text_line, "all_users_select_a_worker") Then
                If all_users_select_a_worker_CheckBox.Checked = True Then
                    text_line = "all_users_select_a_worker = True"
                Else
                    text_line = "all_users_select_a_worker = False"
                End If
            End If

            'Sets the users who get to select a user manually. If the field was blank, it returns a blank, otherwise it converts the string commas to something working for an array.
            If InStr(text_line, "users_using_select_a_user = array(") Then
                If users_using_select_a_worker.Text = "" Then
                    text_line = "users_using_select_a_user = array()"
                Else
                    text_line = "users_using_select_a_user = array(""" & UCase(Replace(Replace(users_using_select_a_worker.Text, " ", ""), ",", """, """) & """)")
                End If
            End If

            'INSERT COLLECTING STATS FIXES HERE WHEN ACCESS GOES LIVE

            'Sets the agency as beta or otherwise
            If InStr(text_line, "beta_agency =") Then
                If agency_is_beta = True Then
                    text_line = "beta_agency = True"
                Else
                    text_line = "beta_agency = False"
                End If
            End If

            'Writes the file data to the variable for new_text_file, which will get written to the file a bit later.
            new_text_file = new_text_file & text_line & Chr(10)
        Next

        'Splits the array created by the For...next and writes each line into the file
        new_text_file = Split(new_text_file, Chr(10))
        System.IO.File.WriteAllLines(location_to_save_script_files.Text & "\Script Files\SETTINGS - GLOBAL VARIABLES.vbs", new_text_file)

        'Switches controls back to default
        Update_Files_Label.Visible = False
        Tab_Control_Main_Form.Enabled = True
        run_configuration_button.Enabled = True

        'Success!
        Me.Hide()
        MsgBox("Success! BlueZone Scripts installed in this directory.")
        Application.Exit()

    End Sub

    Private Sub FileOpen(Optional p1 As Object = Nothing, Optional file As Object = Nothing, Optional openMode As OpenMode = Nothing, Optional openAccess As OpenAccess = Nothing, Optional p5 As Object = Nothing, Optional p6 As Object = Nothing)
        Throw New NotImplementedException
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        address_entry_form.ShowDialog(Me)

    End Sub


    'This sub finds the current directory, and then offers to copy the existing values from FUNCTIONS FILE. It triggers when the user
    'selects "browse" from the "where are we saving scripts" part of the main tab.
    Private Sub browse_to_folder_button_Click_1(sender As Object, e As EventArgs) Handles browse_to_folder_button.Click
        FolderBrowserDialog1.ShowDialog()

        'It does the next piece in case the user accidentally sets the path as "Script Files". 
        'This will eliminate the recursive "Script Files\Script Files\Script Files" which is bound to happen with less savvy folks.
        If Strings.Right(FolderBrowserDialog1.SelectedPath, 13) = "\Script Files" Then
            FolderBrowserDialog1.SelectedPath = Strings.Left(FolderBrowserDialog1.SelectedPath, FolderBrowserDialog1.SelectedPath.Length - 13)
        End If

        'If there's script files in the folder, this will find them and ask if we want to try and autofill the dialog.
        If File.Exists(FolderBrowserDialog1.SelectedPath & "/Script Files/SETTINGS - GLOBAL VARIABLES.vbs") = True Then
            Dim files_found_messagebox = MsgBox("This folder contains script files! Would you like the installer to try and read the address(es) and agency for you?", MsgBoxStyle.YesNo)
            If files_found_messagebox = MsgBoxResult.Yes Then
                Dim global_variables_text() As String = File.ReadAllLines(FolderBrowserDialog1.SelectedPath & "/Script Files/SETTINGS - GLOBAL VARIABLES.vbs")
                For Each global_variables_line In global_variables_text
                    If InStr(global_variables_line, "code_from_installer") Then
                        global_variables_line = Replace(global_variables_line, "code_from_installer = ", "")
                        global_variables_line = Replace(global_variables_line, Chr(34), "")
                        county_selection.Text = global_variables_line
                    End If
                    If InStr(global_variables_line, "county_bndx_variance_threshold") Then
                        global_variables_line = Replace(global_variables_line, "county_bndx_variance_threshold = ", "")
                        global_variables_line = Replace(global_variables_line, Chr(34), "")
                        bndx_threshold.Text = global_variables_line
                    End If
                    If InStr(global_variables_line, "emer_percent_rule_amt") Then
                        global_variables_line = Replace(global_variables_line, "emer_percent_rule_amt = ", "")
                        global_variables_line = Replace(global_variables_line, Chr(34), "")
                        emer_percent_rule_number.Text = global_variables_line
                    End If
                    If InStr(global_variables_line, "emer_number_of_income_days") Then
                        global_variables_line = Replace(global_variables_line, "emer_number_of_income_days = ", "")
                        global_variables_line = Replace(global_variables_line, Chr(34), "")
                        emer_number_of_income_days.Text = global_variables_line
                    End If
                    If InStr(global_variables_line, "CLS_x1_number") Then
                        global_variables_line = Replace(global_variables_line, "CLS_x1_number = ", "")
                        global_variables_line = Replace(global_variables_line, Chr(34), "")
                        X1_for_CLS.Text = global_variables_line
                    End If
                    If InStr(global_variables_line, "EDMS_choice = ") Then
                        global_variables_line = Replace(global_variables_line, "EDMS_choice = ", "")
                        global_variables_line = Replace(global_variables_line, Chr(34), "")
                        EDMS_choice.Text = global_variables_line
                        If EDMS_choice.Text <> "DHS eDocs" Then EDMS_check.Checked = True
                    End If
                    If InStr(global_variables_line, "case_noting_intake_dates = False") Then
                        intake_dates_check.Checked = False
                    ElseIf InStr(global_variables_line, "case_noting_intake_dates = True") Then
                        intake_dates_check.Checked = True
                    End If
                    If InStr(global_variables_line, "move_verifs_needed = True") Then
                        move_verifs_needed_check.Checked = True
                    ElseIf InStr(global_variables_line, "move_verifs_needed = False") Then
                        move_verifs_needed_check.Checked = False
                    End If
                    If InStr(global_variables_line, "'Set fso_command = run_another_script_fso.OpenTextFile(") Then
                        global_variables_line = Replace(global_variables_line, "'Set fso_command = run_another_script_fso.OpenTextFile(", "")
                        global_variables_line = Replace(global_variables_line, "\Script Files\SETTINGS - GLOBAL VARIABLES.vbs", "")
                        global_variables_line = Replace(global_variables_line, ")", "")
                        global_variables_line = Replace(global_variables_line, Chr(34), "")
                        If global_variables_line <> FolderBrowserDialog1.SelectedPath Then
                            custom_file_path.Text = global_variables_line
                        End If
                    End If
                    If InStr(global_variables_line, "county_office_array = split(") Then
                        Dim reading_address_info As Array
                        global_variables_line = Replace(global_variables_line, "county_office_array = split(" & Chr(34), "")
                        global_variables_line = Replace(global_variables_line, Chr(34) & ", " & Chr(34) & "~" & Chr(34) & ")", "")
                        reading_address_info = Split(global_variables_line, "~")
                        For array_number = 0 To UBound(reading_address_info)
                            If array_number = 0 Then
                                Dim line_to_add = Split(reading_address_info(0), "|")
                                address_entry_form.addr_01_line_01.Text = line_to_add(0)
                                address_entry_form.addr_01_line_02.Text = line_to_add(1)
                            End If
                            If array_number = 1 Then
                                Dim line_to_add = Split(reading_address_info(1), "|")
                                address_entry_form.addr_02_line_01.Text = line_to_add(0)
                                address_entry_form.addr_02_line_02.Text = line_to_add(1)
                            End If
                            If array_number = 2 Then
                                Dim line_to_add = Split(reading_address_info(2), "|")
                                address_entry_form.addr_03_line_01.Text = line_to_add(0)
                                address_entry_form.addr_03_line_02.Text = line_to_add(1)
                            End If
                            If array_number = 3 Then
                                Dim line_to_add = Split(reading_address_info(3), "|")
                                address_entry_form.addr_04_line_01.Text = line_to_add(0)
                                address_entry_form.addr_04_line_02.Text = line_to_add(1)
                            End If
                            If array_number = 4 Then
                                Dim line_to_add = Split(reading_address_info(4), "|")
                                address_entry_form.addr_05_line_01.Text = line_to_add(0)
                                address_entry_form.addr_05_line_02.Text = line_to_add(1)
                            End If
                            If array_number = 5 Then
                                Dim line_to_add = Split(reading_address_info(5), "|")
                                address_entry_form.addr_06_line_01.Text = line_to_add(0)
                                address_entry_form.addr_06_line_02.Text = line_to_add(1)
                            End If
                            If array_number = 6 Then
                                Dim line_to_add = Split(reading_address_info(6), "|")
                                address_entry_form.addr_07_line_01.Text = line_to_add(0)
                                address_entry_form.addr_07_line_02.Text = line_to_add(1)
                            End If
                        Next
                    End If
                    If InStr(global_variables_line, "all_users_select_a_worker") Then
                        'MsgBox(global_variables_line)
                        If global_variables_line = "all_users_select_a_worker = False" Then
                            all_users_select_a_worker_CheckBox.Checked = False
                        Else
                            all_users_select_a_worker_CheckBox.Checked = True
                        End If
                    End If
                    If InStr(global_variables_line, "users_using_select_a_user") Then
                        'Gets the users by taking the array from the line and replacing the function name (and the word array) and the parenthesis with blanks, and the "", "" string with a simple comma-space.
                        users_using_select_a_worker.Text = Replace(Replace(Replace(global_variables_line, "users_using_select_a_user = array(""", ""), """)", ""), """", "")
                    End If
                Next
            End If
        End If

        'Now it updates the text for the file save location.
        location_to_save_script_files.Text = FolderBrowserDialog1.SelectedPath
    End Sub

    Private Sub custom_file_path_TextChanged(sender As Object, e As EventArgs) Handles custom_file_path.Leave
        'Removing the entry if the user puts a slash as the last character.
        If InStrRev(custom_file_path.Text, "\") = Len(custom_file_path.Text) Then
            MsgBox("You cannot enter a slash as the last character. Your entry will now be cleared.")
            custom_file_path.Text = ""
        End If
    End Sub



    Private Sub bndx_threshold_Leave(sender As Object, e As EventArgs) Handles bndx_threshold.Leave
        bndx_threshold.Text = Strings.Replace(bndx_threshold.Text, "$", "")
        If IsNumeric(bndx_threshold.Text) = False Then
            MsgBox("You must enter a numeric amount for the BNDX threshold. It will now revert to 0.")
            bndx_threshold.Text = 0
        End If
    End Sub

    Private Sub emer_percent_rule_number_Leave(sender As Object, e As EventArgs) Handles emer_percent_rule_number.Leave
        emer_percent_rule_number.Text = Strings.Replace(emer_percent_rule_number.Text, "$", "")
        emer_percent_rule_number.Text = Strings.Replace(emer_percent_rule_number.Text, "%", "")
        If IsNumeric(emer_percent_rule_number.Text) = False Then
            MsgBox("You must enter a numeric amount for the EMER percent rule. It will now revert to 30.")
            emer_percent_rule_number.Text = 30
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click


        Dim max_dot_net_version As String
        Using ndpKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                    RegistryView.Registry32).OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\")

            If ndpKey IsNot Nothing AndAlso ndpKey.GetValue("Release") IsNot Nothing Then
                max_dot_net_version = ("Version: " & CheckFor45DotVersion(CInt(ndpKey.GetValue("Release"))))
            Else
                max_dot_net_version = ("Version 4.5 or later is not detected.")
            End If
        End Using

        MsgBox(".NET version currently used: " & Environment.Version.ToString() & vbCr & _
               ".NET 4.5 version installed: " & max_dot_net_version)

    End Sub

    Private Sub emer_number_of_income_days_Leave(sender As Object, e As EventArgs) Handles emer_number_of_income_days.Leave
        If IsNumeric(emer_number_of_income_days.Text) = False Then
            MsgBox("You must enter a numeric amount for the EMER number of intake days. It will now revert to 30.")
            emer_number_of_income_days.Text = 30
        End If
    End Sub

    Private Sub X1_for_CLS_Leave(sender As Object, e As EventArgs) Handles X1_for_CLS.Leave
        If Len(X1_for_CLS.Text) <> 7 And X1_for_CLS.Text <> "" Then
            MsgBox("You must enter a seven digit worker ID. It will now clear what you've typed.")
            X1_for_CLS.Text = ""
        End If
    End Sub
End Class
