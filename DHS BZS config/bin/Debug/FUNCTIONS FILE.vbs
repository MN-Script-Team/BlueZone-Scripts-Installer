'>>>>THIS IS A DUMMY VERSION, ONLY TO BE USED IN TESTING THE CONFIG PROGRAM<<<<
'
'GATHERING STATS----------------------------------------------------------------------------------------------------
'name_of_script = ""
'start_time = timer
'
''LOADING ROUTINE FUNCTIONS
'Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
'Set fso_command = run_another_script_fso.OpenTextFile("C:\BlueZone-Scripts-Installer\DHS BZS config\bin\Debug\FUNCTIONS FILE.vbs")
'text_from_the_other_script = fso_command.ReadAll
'fso_command.Close
'Execute text_from_the_other_script

'----------------------------------------------------------------------------------------------------

'COUNTY CUSTOM VARIABLES----------------------------------------------------------------------------------------------------

worker_county_code = "x102"
collecting_statistics = False
EDMS_choice = "DHS eDocs"
county_name = "Anoka"
county_address_line_01 = "1234"
county_address_line_02 = "9876"
case_noting_intake_dates = True
move_verifs_needed = False

is_county_collecting_stats = collecting_statistics	'IT DOES THIS BECAUSE THE SETUP SCRIPT WILL OVERWRITE LINES BELOW WHICH DEPEND ON THIS, BY SEPARATING THE VARIABLES WE PREVENT ISSUES

'SHARED VARIABLES----------------------------------------------------------------------------------------------------
checked = 1		'Value for checked boxes
unchecked = 0	'Value for unchecked boxes
cancel = 0		'Value for cancel button in dialogs
OK = -1		'Value for OK button in dialogs

'Some screens require the two digit county code, and this determines what that code is
two_digit_county_code = right(worker_county_code, 2)
If two_digit_county_code = "PW" then two_digit_county_code = "91"	'For DHS purposes












