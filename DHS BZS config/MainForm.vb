Imports System
Imports System.IO
Imports System.Collections
Imports System.Text



Public Class scripts_config_form

    Const read_only = 1
    Const read_write = 2

    Private Property list_of_addresses As String

    Private Property fso As Object

    Private Property FSO_new_file_path As Object

    Private Property ObjFSO As Object

    Private Property objFile As Object

    Private Property function_file_lines As Object

    Private Property line_to_look_for_in_functions_file As String

    Private Property warning As MsgBoxResult

    Private Property list_of_files_array As String()

    Private Property text_file As String

    Private Property new_text_file As Object

    Private Property GitHub_current_zip_archive As String

    Private Property objXMLHTTP As Object

    Private Property objADOStream As Object

    Private Property objShell As Object

    Private Property old_file_path As Object

    'This is the function that actually modifies the files
    Function update_files(file_name, address_array)

        Dim text_file() As String = System.IO.File.ReadAllLines(file_name)
        Dim text_line As String

        new_text_file = Nothing 'clears the variable

        For Each text_line In text_file
            If custom_file_path.Text = Nothing Then
                text_line = Replace(text_line, old_file_path, location_to_save_script_files.Text & "\")
            Else
                text_line = Replace(text_line, old_file_path, custom_file_path.Text & "\")
            End If
            If InStr(file_name, "FUNCTIONS FILE") <> 0 Then   'Shouldn't do this part for any scripts other than the functions file.
                If InStr(text_line, "worker_county_code = ") Then
                    If IsNumeric(Strings.Left(county_selection.Text, 2)) = True Then
                        text_line = "worker_county_code = " & Chr(34) & "x1" & Strings.Left(county_selection.Text, 2) & Chr(34)
                    Else
                        text_line = "worker_county_code = " & Chr(34) & "MULTICOUNTY" & Chr(34)
                    End If
                End If
                If InStr(text_line, "EDMS_choice = ") Then text_line = "EDMS_choice = " & Chr(34) & EDMS_choice.Text & Chr(34)
                If InStr(text_line, "county_name = ") Then text_line = "county_name = " & Chr(34) & Strings.Replace(county_selection.Text, Strings.Left(county_selection.Text, 5), "") & Chr(34)
                If InStr(text_line, "county_office_array = split(") Then text_line = "county_office_array = split(" & Chr(34) & address_array & Chr(34) & ", " & Chr(34) & "~" & Chr(34) & ")"
                If InStr(text_line, "case_noting_intake_dates = ") Then
                    If intake_dates_check.Checked = True Then
                        text_line = "case_noting_intake_dates = True"
                    Else
                        text_line = "case_noting_intake_dates = False"
                    End If
                End If
                If InStr(text_line, "move_verifs_needed = ") Then
                    If move_verifs_needed_check.Checked = True Then
                        text_line = "move_verifs_needed = True"
                    Else
                        text_line = "move_verifs_needed = False"
                    End If
                End If
                If InStr(text_line, "code_from_installer = ") Then text_line = "code_from_installer = " & Chr(34) & county_selection.Text & Chr(34)
            End If
            'INSERT COLLECTING STATS FIXES HERE WHEN ACCESS GOES LIVE
            new_text_file = new_text_file & text_line & Chr(10)
        Next

        new_text_file = Split(new_text_file, Chr(10))
        System.IO.File.WriteAllLines(file_name, new_text_file)
        Return file_name
    End Function

    


    Sub downloading_files_from_GitHub()
        Dim agency_is_beta As Boolean

        'Only some agencies get the option to install beta scripts, based mainly on their status as contributing agencies (they write scripts).
        'As a result, we have an if/then list. It will need to be updated frequently as agencies join/leave the beta program.
        If county_selection.Text = "02 - Anoka County" Or _
            county_selection.Text = "05 - Benton County" Or _
            county_selection.Text = "19 - Dakota County" Or _
            county_selection.Text = "20 - Dodge County" Or _ 
            county_selection.Text = "55 - Olmsted County" Or _
            county_selection.Text = "57 - Pennington County" Or _
            county_selection.Text = "69 - St. Louis County" Or _
            county_selection.Text = "73 - Stearns County" Or _
            county_selection.Text = "74 - Steele County" Or _
            county_selection.Text = "79 - Wabasha County" Then

            Dim beta_msgbox As String
            beta_msgbox = MsgBox("Your agency is listed as a beta agency. Install beta versions of scripts?", MsgBoxStyle.YesNoCancel)
            If beta_msgbox = 2 Then Exit Sub
            If beta_msgbox = 6 Then agency_is_beta = True
            If beta_msgbox = 7 Then agency_is_beta = False

        End If

        'Variables for local_copy_of_zip_file and temp_folder are used by this sub. These variables are temporary.
        Dim local_copy_of_zip_file = location_to_save_script_files.Text & "\temp\master.zip"
        Dim temp_folder = location_to_save_script_files.Text & "\temp"


        'If the agency is a beta agency, they'll have the beta version of the scripts instead of the main one.
        If agency_is_beta = True Then
            GitHub_current_zip_archive = "https://github.com/MN-DHS-BZS-County-Programmers/MAXIS-BZ-Scripts-County-Beta/archive/beta.zip"
        Else
            GitHub_current_zip_archive = "https://github.com/MN-DHS-BZS-County-Programmers/MAXIS-BZ-Scripts-County-Beta/archive/Release.zip"
        End If

        'First, we create a temp directory for all this madness.
        fso = CreateObject("Scripting.FileSystemObject")
        If fso.folderexists(temp_folder) = True Then fso.deletefolder(temp_folder) 'Clears the temp folder if it exists.
        fso.CreateFolder(temp_folder)

        '---------------------------------------------------------------------------------------
        'Now it downloads the zip file from Github. This code was copied from https://gist.github.com/udawtr/2053179 on 09/13/2014, and modified for our purposes.

        'Creating a server object
        objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")

        'Opening the file
        objXMLHTTP.open("GET", GitHub_current_zip_archive, False)
        objXMLHTTP.send()
        If objXMLHTTP.Status = 200 Then     'Guessing this means "found" but admittedly I'm not sure. -VKC, 09/13/2014
            objADOStream = CreateObject("ADODB.Stream")
            objADOStream.Open()
            objADOStream.Type = 1 'adTypeBinary
            objADOStream.Write(objXMLHTTP.ResponseBody)
            objADOStream.Position = 0    'Set the stream position to the start

            'Writing the file to the hard disk
            ObjFSO = CreateObject("Scripting.FileSystemObject")
            If ObjFSO.Fileexists(local_copy_of_zip_file) Then ObjFSO.DeleteFile(local_copy_of_zip_file)
            ObjFSO = Nothing
            objADOStream.SaveToFile(local_copy_of_zip_file)
            objADOStream.Close()
            objADOStream = Nothing
        End If

        objXMLHTTP = Nothing
        '------------------------------------------------------------------------------------
        'Now, because the ZipFile.ExtractToDirectory method appears to not work with folks running .net 4.0, the
        'app will create a VBS version of the file extractor, and save it to disk.

        'Creating the file and writing each line
        Dim new_zip_file As StreamWriter = File.CreateText(temp_folder & "\unzip.vbs")
        new_zip_file.WriteLine("current_directory = " & Chr(34) & temp_folder & Chr(34))                                                        'Sets current directory to match temp folder
        new_zip_file.WriteLine("set objShell = CreateObject(" & Chr(34) & "Shell.Application" & Chr(34) & ") ")                                 'Creates a shell.application object to execute extraction
        new_zip_file.WriteLine("set FilesInZip = objShell.NameSpace(current_directory & " & Chr(34) & "\master.zip" & Chr(34) & ").items")      'Points to the master.zip file
        new_zip_file.WriteLine("objShell.NameSpace(current_directory).CopyHere FilesInZip, 16 ")                                                'Copies the files out to the temp directory
        new_zip_file.WriteLine("set objShell = Nothing")                                                                                        'Nulls variable like a good little code
        new_zip_file.Flush()
        new_zip_file.Close()

        new_zip_file = Nothing  'clears variable


        'Running the zip file. Calls wscript, as a surprising number of folks have somehow modified their default program for scripts to be notepad.
        'Apparently people want to hack scripts to do their own bidding. Who'd have thought?
        Process.Start("wscript", Chr(34) & temp_folder & "\unzip.vbs" & Chr(34))

        'Now, because we have an independent script running, we need to check to see if wscript.exe is running. So, we set a variable to be our boolean "are we there yet"?
        Dim moving_on As Boolean

        'Now it loops checking to see if wscript.exe is running. If it STOPS running, moving_on is set to true, and we continue on our merry way.
        Do
            Dim processes() As Process
            processes = Process.GetProcessesByName("wscript")
            If processes.Count > 0 Then
                moving_on = False
            Else
                moving_on = True
            End If
        Loop Until moving_on = True

        'There should only be one folder in the download, but the directory checker creates an array. We need to join the array before we can mess with it.
        Dim dirs As List(Of String) = New List(Of String)(Directory.EnumerateDirectories(temp_folder))
        Dim github_folder As String = String.Join(",", dirs)

        'Now we move the folder to its new home.
        fso.CopyFolder(github_folder & "\Script Files", location_to_save_script_files.Text & "\Script Files", True)

        'Now we delete the temp directory.
        fso.DeleteFolder(temp_folder)
    End Sub

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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles run_configuration_button.Click
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

        'Warns user that they can back out
        warning = MsgBox("The following utility will download all of the scripts to the selected directory, and replace the DHS file path with " & _
        "the current file path. If you choose to move your script directory, you'll have to use this tool again. Are you sure you want to do this?", 1)
        If warning = 2 Then Exit Sub

        'Uses a custom sub to download files from GitHub
        downloading_files_from_GitHub()

        'Disabling features/enabling others (more of a pretty GUI setup, might rewrite to be even prettier)
        Update_Files_Label.Visible = True
        Tab_Control_Main_Form.Enabled = False
        run_configuration_button.Enabled = False

        'Setting EDMS_choice as DHS eDocs if there is not a local EDMS.
        If EDMS_check.Checked = False Then EDMS_choice.Text = "DHS eDocs"

        'Opening FUNCTIONS FILE file read-only, to know what path to overwrite from the GitHub repo.
        ObjFSO = CreateObject("Scripting.FileSystemObject")
        objFile = ObjFSO.OpenTextFile(location_to_save_script_files.Text & "\Script Files\FUNCTIONS FILE.vbs", read_only)

        'Reading each line, seeking the old path name from the GitHub repo.
        Do Until objFile.AtEndOfStream
            function_file_lines = objFile.ReadLine
            line_to_look_for_in_functions_file = "'Set fso_command = run_another_script_fso.OpenTextFile("
            If InStr(function_file_lines, line_to_look_for_in_functions_file) Then
                old_file_path = Replace(Replace(Replace(function_file_lines, line_to_look_for_in_functions_file, ""), Chr(34), ""), "Script Files\FUNCTIONS FILE.vbs)", "")
            End If
        Loop

        'Closing the read-only version
        objFile.Close()

        'Grabbing each file
        list_of_files_array = Directory.GetFiles(location_to_save_script_files.Text & "\Script Files")


        'Running the update_files sub on each VBS file
        For Each file_in_array In list_of_files_array
            If UCase(Strings.Right(file_in_array, 4)) = ".VBS" Then update_files(file_in_array, list_of_addresses)
        Next

        Update_Files_Label.Visible = False
        Tab_Control_Main_Form.Enabled = True
        run_configuration_button.Enabled = True

        'Success!
        Me.Hide()
        MsgBox("Success! All scripts modified to work in this directory.")
        Application.Exit()
    End Sub

    Private Sub FileOpen(Optional p1 As Object = Nothing, Optional file As Object = Nothing, Optional openMode As OpenMode = Nothing, Optional openAccess As OpenAccess = Nothing, Optional p5 As Object = Nothing, Optional p6 As Object = Nothing)
        Throw New NotImplementedException
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        address_entry_form.ShowDialog(Me)

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        FolderBrowserDialog1.ShowDialog()

        'It does the next piece in case the user accidentally sets the path as "Script Files". 
        'This will eliminate the recursive "Script Files\Script Files\Script Files" which is bound to happen with less savvy folks.
        If Strings.Right(FolderBrowserDialog1.SelectedPath, 13) = "\Script Files" Then
            FolderBrowserDialog1.SelectedPath = Strings.Left(FolderBrowserDialog1.SelectedPath, FolderBrowserDialog1.SelectedPath.Length - 13)
        End If

        'If there's script files in the folder, this will find them and ask if we want to try and autofill the dialog.
        If File.Exists(FolderBrowserDialog1.SelectedPath & "/Script Files/FUNCTIONS FILE.vbs") = True Then
            Dim files_found_messagebox = MsgBox("This folder contains script files! Would you like the installer to try and read the address(es) and agency for you?", MsgBoxStyle.YesNo)
            If files_found_messagebox = MsgBoxResult.Yes Then
                Dim FUNCTIONS_FILE_text() As String = File.ReadAllLines(FolderBrowserDialog1.SelectedPath & "/Script Files/FUNCTIONS FILE.vbs")
                For Each FUNCTIONS_FILE_line In FUNCTIONS_FILE_text
                    If InStr(FUNCTIONS_FILE_line, "code_from_installer") Then
                        FUNCTIONS_FILE_line = Replace(FUNCTIONS_FILE_line, "code_from_installer = ", "")
                        FUNCTIONS_FILE_line = Replace(FUNCTIONS_FILE_line, Chr(34), "")
                        county_selection.Text = FUNCTIONS_FILE_line

                    End If
                    If InStr(FUNCTIONS_FILE_line, "EDMS_choice = ") Then
                        FUNCTIONS_FILE_line = Replace(FUNCTIONS_FILE_line, "EDMS_choice = ", "")
                        FUNCTIONS_FILE_line = Replace(FUNCTIONS_FILE_line, Chr(34), "")
                        EDMS_choice.Text = FUNCTIONS_FILE_line
                        If EDMS_choice.Text <> "DHS eDocs" Then EDMS_check.Checked = True
                    End If
                    If InStr(FUNCTIONS_FILE_line, "case_noting_intake_dates = False") Then
                        intake_dates_check.Checked = False
                    ElseIf InStr(FUNCTIONS_FILE_line, "case_noting_intake_dates = True") Then
                        intake_dates_check.Checked = True
                    End If
                    If InStr(FUNCTIONS_FILE_line, "move_verifs_needed = True") Then
                        move_verifs_needed_check.Checked = True
                    ElseIf InStr(FUNCTIONS_FILE_line, "move_verifs_needed = False") Then
                        move_verifs_needed_check.Checked = False
                    End If
                    If InStr(FUNCTIONS_FILE_line, "'Set fso_command = run_another_script_fso.OpenTextFile(") Then
                        FUNCTIONS_FILE_line = Replace(FUNCTIONS_FILE_line, "'Set fso_command = run_another_script_fso.OpenTextFile(", "")
                        FUNCTIONS_FILE_line = Replace(FUNCTIONS_FILE_line, "\Script Files\FUNCTIONS FILE.vbs", "")
                        FUNCTIONS_FILE_line = Replace(FUNCTIONS_FILE_line, ")", "")
                        FUNCTIONS_FILE_line = Replace(FUNCTIONS_FILE_line, Chr(34), "")
                        If FUNCTIONS_FILE_line <> FolderBrowserDialog1.SelectedPath Then
                            custom_file_path.Text = FUNCTIONS_FILE_line
                        End If
                    End If
                    If InStr(FUNCTIONS_FILE_line, "county_office_array = split(") Then
                        Dim reading_address_info As Array
                        FUNCTIONS_FILE_line = Replace(FUNCTIONS_FILE_line, "county_office_array = split(" & Chr(34), "")
                        FUNCTIONS_FILE_line = Replace(FUNCTIONS_FILE_line, Chr(34) & ", " & Chr(34) & "~" & Chr(34) & ")", "")
                        reading_address_info = Split(FUNCTIONS_FILE_line, "~")
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
                Next
            End If
        End If

        'Now it updates the text for the file save location.
        location_to_save_script_files.Text = FolderBrowserDialog1.SelectedPath
    End Sub

End Class
