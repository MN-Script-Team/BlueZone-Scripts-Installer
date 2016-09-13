//This function creates strings containing the required content of redirect files
function redirectFile(scriptToRedirectTo) {

    var filepath = document.getElementById("filepath").value;

    return "'LOADING GLOBAL VARIABLES--------------------------------------------------------------------"                                                                                  + "\n" +
        "Set run_another_script_fso = CreateObject(\"Scripting.FileSystemObject\")"                                                                                                         + "\n" +
        "Set fso_command = run_another_script_fso.OpenTextFile(\"" + filepath + "\\SETTINGS - GLOBAL VARIABLES.vbs\")"                                                                      + "\n" +
        "text_from_the_other_script = fso_command.ReadAll"                                                                                                                                  + "\n" +
        "fso_command.Close"                                                                                                                                                                 + "\n" +
        "Execute text_from_the_other_script"                                                                                                                                                + "\n" +
        ""                                                                                                                                                                                  + "\n" +
        "'LOADING SCRIPT"                                                                                                                                                                   + "\n" +
        "script_URL = script_repository & \"" + scriptToRedirectTo +  "\""                                                                                                                  + "\n" +
        "IF run_locally = False THEN"                                                                                                                                                       + "\n" +
        "   SET req = CreateObject(\"Msxml2.XMLHttp.6.0\")				'Creates an object to get a URL"                                                                                    + "\n" +
        "   req.open \"GET\", script_URL, FALSE							'Attempts to open the URL"                                                                                          + "\n" +
        "   req.send													'Sends request"                                                                                                     + "\n" +
        "   IF req.Status = 200 THEN									'200 means great success"                                                                                           + "\n" +
        "       Set fso = CreateObject(\"Scripting.FileSystemObject\")	'Creates an FSO"                                                                                                    + "\n" +
        "       Execute req.responseText								'Executes the script code"                                                                                          + "\n" +
        "   ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact someone/stops script"         + "\n" +
        "       critical_error_msgbox = MsgBox (\"Something has gone wrong. The code stored on GitHub was not able to be reached.\" & vbNewLine & vbNewLine &_"                             + "\n" +
        "                                \"Script URL: \" & script_URL & vbNewLine & vbNewLine &_"                                                                                          + "\n" +
        "                                \"The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.\", _"                         + "\n" +
        "                                vbOKonly + vbCritical, \"BlueZone Scripts Critical Error\")"                                                                                       + "\n" +
        "                                StopScript"                                                                                                                                        + "\n" +
        "   END IF"                                                                                                                                                                         + "\n" +
        "ELSE"                                                                                                                                                                              + "\n" +
        "   Set run_another_script_fso = CreateObject(\"Scripting.FileSystemObject\")"                                                                                                      + "\n" +
        "   Set fso_command = run_another_script_fso.OpenTextFile(script_URL)"                                                                                                              + "\n" +
        "   text_from_the_other_script = fso_command.ReadAll"                                                                                                                               + "\n" +
        "   fso_command.Close"                                                                                                                                                              + "\n" +
        "   Execute text_from_the_other_script"                                                                                                                                             + "\n" +
        "END IF"                                                                                                                                                                            + "\n";

}

//This actually sets up and creates the zip file, complete with redirects
function scriptsSetupPRISM() {


    var filepath = document.getElementById("filepath").value;
    var countyCALIcode = document.getElementById("countyCALIcode").value;
    var EDMSchoice = document.getElementById("EDMSchoice").value;
    var supportemail = document.getElementById("supportemail").value;


    JSZipUtils.getBinaryContent('https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/Script%20Files/PAD%20-%20CS.pad', function(err, data) {
        if(err) {
            alert(err); //Provide a warning if there's an error
        } else {

            var zip = new JSZip();                                                      //Create a new JSZip object

            zip.file("Script Files/PAD - CS.pad",                                       data);                                                                      //Add the Power Pad which was grabbed using the .getBinaryContent function
            zip.file("Script Files/REDIRECT - ACTIONS.vbs",                             redirectFile("/ACTIONS/ACTIONS - MAIN MENU.vbs"));                          //Add each redirect needed for PRISM scripts
            zip.file("Script Files/REDIRECT - AGENCY CUSTOMIZED.vbs",                   redirectFile("/AGENCY CUSTOMIZED/SOURCE/AGENCY CUSTOMIZED.vbs"));
            zip.file("Script Files/REDIRECT - BULK.vbs",                                redirectFile("/BULK/BULK - MAIN MENU.vbs"));
            zip.file("Script Files/REDIRECT - NAV - CAAD.vbs",                          redirectFile("/NAV/NAV - CAAD.vbs"));
            zip.file("Script Files/REDIRECT - NAV - CAFS.vbs",                          redirectFile("/NAV/NAV - CAFS.vbs"));
            zip.file("Script Files/REDIRECT - NAV - CAPS.vbs",                          redirectFile("/NAV/NAV - CAPS.vbs"));
            zip.file("Script Files/REDIRECT - NAV - CAST.vbs",                          redirectFile("/NAV/NAV - CAST.vbs"));
            zip.file("Script Files/REDIRECT - NAV - CAWT.vbs",                          redirectFile("/NAV/NAV - CAWT.vbs"));
            zip.file("Script Files/REDIRECT - NAV - CPDD.vbs",                          redirectFile("/NAV/NAV - CPDD.vbs"));
            zip.file("Script Files/REDIRECT - NAV - ENFL.vbs",                          redirectFile("/NAV/NAV - ENFL.vbs"));
            zip.file("Script Files/REDIRECT - NAV - FIND THAT SCREEN IN PRISM.vbs",     redirectFile("/NAV/NAV - FIND THAT SCREEN IN PRISM.vbs"));
            zip.file("Script Files/REDIRECT - NAV - MAXIS SCREEN FINDER.vbs",           redirectFile("/NAV/NAV - MAXIS SCREEN FINDER.vbs"));
            zip.file("Script Files/REDIRECT - NAV - NCDD.vbs",                          redirectFile("/NAV/NAV - NCDD.vbs"));
            zip.file("Script Files/REDIRECT - NAV - NCID.vbs",                          redirectFile("/NAV/NAV - NCID.vbs"));
            zip.file("Script Files/REDIRECT - NAV - PALC.vbs",                          redirectFile("/NAV/NAV - PALC.vbs"));
            zip.file("Script Files/REDIRECT - NAV - PAPL.vbs",                          redirectFile("/NAV/NAV - PAPL.vbs"));
            zip.file("Script Files/REDIRECT - NAV - PESE.vbs",                          redirectFile("/NAV/NAV - PESE.vbs"));
            zip.file("Script Files/REDIRECT - NAV - SUCW.vbs",                          redirectFile("/NAV/NAV - SUCW.vbs"));
            zip.file("Script Files/REDIRECT - NAV - USWD.vbs",                          redirectFile("/NAV/NAV - USWD.vbs"));
            zip.file("Script Files/REDIRECT - NOTES.vbs",                               redirectFile("/NOTES/NOTES - MAIN MENU.vbs"));
            zip.file("Script Files/REDIRECT - NOTES - QUICK CAAD.vbs",                  redirectFile("/NOTES/NOTES - QUICK CAAD.vbs"));
            zip.file("Script Files/REDIRECT - UTILITIES.vbs",                           redirectFile("/UTILITIES/UTILITIES - MAIN MENU.vbs"));
            zip.file("Script Files/REDIRECT - UTILITIES - EMAIL SCRIPTS SUPPORT.vbs",   redirectFile("/UTILITIES/UTILITIES - EMAIL SCRIPTS SUPPORT.vbs"));
            zip.file("README.txt",                                                      "This is your BlueZone Scripts directory\n");        //README

            //NOW WE NEED TO CREATE A GLOABL VARIABLES FILE!
            var globVarContent =
                "'COUNTY CUSTOM VARIABLES----------------------------------------------------------------------------------------------------" +
                "\n" + "'The following variables are dynamically added via the installer. They can be modified manually to make changes without re-running the installer, but doing so should not be undertaken lightly." +
                "\n" +
                "\n" + "' DETAILS ABOUT HOW YOUR SCRIPTS WILL RUN ----------------------------------------------------------------------------------" +
                "\n" +
                "\n" + "'Run locally: if this is set to \"True\", the scripts will run locally and bypass GitHub entirely. This is great for debugging or developing scripts. Only scriptwriters should do it. An agency should always be set to \"false\"." +
                "\n" + "run_locally = false" +
                "\n" +
                "\n" + "'This is a variable which signifies the agency uses the master branch or the RELEASE branch. Set to true if you're a scriptwriter agency and all users are going to be on the master branch. Otherwise, set to false." +
                "\n" + "use_master_branch = false" +
                "\n" +
                "\n" + "'This allows a \"beta user\" group to have access to master branch scripts, while everyone else uses release. This is helpful for counties that want to maintain a small test group." +
                "\n" + "'Here is the list of agency super users. These users will have access to the test scripts. Enter the list of users' log-in IDs in the quotes below, comma separated" +
                "\n" + "beta_users = \"\"" +
                "\n" +
                "\n" + "'This is used by the AGENCY CUSTOMIZED process, and can be used elsewhere if needed, but for now it's mostly informational" +
                "\n" + "'	This is modified by the installer, which will determine if this is a scriptwriter or a production user." +
                "\n" + "default_directory = \""+ filepath + "\"" +
                "\n" +
                "\n" + "'DETAILS ABOUT STATISTICS AND GATHERING THEM ------------------------------------------------------------------------------------------" +
                "\n" +
                "\n" + "'This is used for determining whether script_end_procedure will also log usage info in an Access table." +
                "\n" + "collecting_statistics = False" +
                "\n" +
                "\n" + "'This is the file path for the statistics Access database." +
                "\n" + "stats_database_path = \"C:\\DHS-PRISM-Scripts\\Databases for script usage\\usage statistics.accdb\"" +
                "\n" +
                "\n" + "'DETAILS ABOUT WHERE TO FIND DOCS AND WHICH TO USE ------------------------------------------------------------------------------------------" +
                "\n" +
                "\n" + "'This is the folder path for county-specific Word documents. Modify this with your shared-drive location for Word documents." +
                "\n" + "word_documents_folder_path = \"C:\\DHS-PRISM-Scripts\\Word files for script usage\\\"" +
                "\n" +
                "\n" + "'DETAILS ABOUT THE COUNTY ITSELF -------------------------------------------------------------------------------------------------------------" +
                "\n" +
                "\n" + "'This is the county code on the CALI screen." +
                "\n" + "county_cali_code = \"" + countyCALIcode + "\"" +
                "\n" +
                "\n" + "'An array of county attorneys. \"Select one:\" should ALWAYS be in there, and ALWAYS be first. Replace \"County Attorney #\" with your agency's county attorney names." +
                "\n" + "county_attorney_array = array(\"County Attorney 1\", \"County Attorney 2\", \"County Attorney 3\", \"County Attorney 4\", \"County Attorney 5\")" +
                "\n" +
                "\n" + "'An array of child support magistrates. \"Select one:\" should ALWAYS be in there, and ALWAYS be first.  Replace \"Magistrate # with your agency's child support magistrate names." +
                "\n" + "child_support_magistrates_array = array(\"Magistrate 1\", \"Magistrate 2\", \"Magistrate 3\", \"Magistrate 4\", \"Magistrate 5\")" +
                "\n" +
                "\n" + "'An array of judges. \"Select one:\" should ALWAYS be in there, and ALWAYS be first.  Replace \"Judge #\" with your agency's judges names." +
                "\n" + "county_judge_array = array(\"Judge 1\", \"Judge 2\", \"Judge 3\", \"Judge 4\", \"Judge 5\")" +
                "\n" +
                "\n" + "'This is used by scripts which tell the worker where to find a doc to send to a client (ie \"Send form using Compass Pilot\")" +
                "\n" + "EDMS_choice = \"" + EDMSchoice + "\"" +
                "\n" +
                "\n" + "'This is the county's email support address. It can be a distribution list or an individual." +
                "\n" + "support_email_address = \"" + supportemail + "\"" +
                "\n" +
                "\n" + "'ACTIONS TAKEN BASED ON COUNTY CUSTOM VARIABLES------------------------------------------------------------------------------" +
                "\n" + "'**DO NOT EDIT BELOW THIS LINE UNLESS YOU ARE ABSOLUTELY SURE OF WHAT YOU ARE DOING**" +
                "\n" +
                "\n" + "is_county_collecting_stats = collecting_statistics	'IT DOES THIS BECAUSE THE SETUP SCRIPT WILL OVERWRITE LINES BELOW WHICH DEPEND ON THIS, BY SEPARATING THE VARIABLES WE PREVENT ISSUES" +
                "\n" +
                "\n" + "'This loads the user ID for use in determining beta users. May also be used elsewhere in scripts." +
                "\n" + "Set objNet = CreateObject(\"WScript.NetWork\")" +
                "\n" + "windows_user_ID = objNet.UserName" +
                "\n" +
                "\n" + "'This will assign beta users to the master branch." +
                "\n" + "If InStr(beta_users, UCASE(windows_user_ID)) <> 0 then use_master_branch = true" +
                "\n" +
                "\n" + "'This is the URL of our script repository, and should only change if the agency is a scriptwriting agency. Scriptwriters can elect to use the master branch, allowing them to test new tools, etc." +
                "\n" + "IF use_master_branch = TRUE THEN		'scriptwriters typically use the master branch" +
                "\n" + "    script_repository = \"https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/Script Files/\"" +
                "\n" + "ELSE							'Everyone else (who isn't a scriptwriter) typically uses the release branch" +
                "\n" + "    script_repository = \"https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/release/Script Files/\"" +
                "\n" + "END IF" +
                "\n" +
                "\n" + "'If run locally is set to \"True\", the scripts will totally bypass GitHub and run locally." +
                "\n" + "IF run_locally = TRUE THEN script_repository = \"C:\\DHS-PRISM-Scripts\\Script Files\\\"";

            zip.file("Script Files/SETTINGS - GLOBAL VARIABLES.vbs",                    globVarContent);

            //Save the zip file
            zip.generateAsync({type:"blob"})
            .then(function(content) {
                saveAs(content, "PRISM scripts.zip");           // using FileSaver.js
            });
        }
    });
}

function addAttorneys(){

    var addattorneyfield = document.getElementById("addattorneyfield").value;


    var ul = document.getElementById("countyattorneylist");
    var li = document.createElement("li");
    li.appendChild(document.createTextNode(addattorneyfield));
    ul.appendChild(li);

    document.getElementById("addattorneyfield").value = ""
}
