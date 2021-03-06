//This function creates strings containing the required content of redirect files
function redirectFile(scriptToRedirectTo) {

    var filepath = document.getElementById("filepath").value;
    var lastcharfilepath = filepath.substr(filepath.length - 1);
    if (lastcharfilepath != "\\") {
        filepath += "\\";
    }

    return     "'LOADING GLOBAL VARIABLES--------------------------------------------------------------------"
        + "\n" + "GlobVar_path = \"" + filepath + "locally-installed-files\\~globvar.vbs\"													'Setting a default path, which is modified by the installer"
        + "\n" + "Set run_another_script_fso = CreateObject(\"Scripting.FileSystemObject\")														'Creating an FSO for the work"
        + "\n" + "If run_another_script_fso.FileExists(GlobVar_path) then																		'If a Global Variables file is found in above directory..."
        + "\n" + "    Set fso_command = run_another_script_fso.OpenTextFile(GlobVar_path)														'...run it!"
        + "\n" + "Else																														'If a Global Variables file is not found in the above directory..."
        + "\n" + "    Set fso_command = run_another_script_fso.OpenTextFile(\"Scripts\\~globvar-local.vbs\")										'...use the default BlueZone Scripts directory, and insert a custom \"local flavor\" Global Variables file, which can override the default selections."
        + "\n" + "End if"
        + "\n" + "text_from_the_other_script = fso_command.ReadAll																			'Once we have the text from the other script, read it all!"
        + "\n" + "fso_command.Close																											'Close the other script file, and..."
        + "\n" + "Execute text_from_the_other_script																							'...execute the code from the Global Variables file!"
        + "\n" + ""
        + "\n" + "'LOADING SCRIPT"
        + "\n" + "script_URL = script_repository & \"" + scriptToRedirectTo +  "\""
        + "\n" + "IF run_locally = False THEN"
        + "\n" + "   SET req = CreateObject(\"Msxml2.XMLHttp.6.0\")				'Creates an object to get a URL"
        + "\n" + "   req.open \"GET\", script_URL, FALSE							'Attempts to open the URL"
        + "\n" + "   req.send													'Sends request"
        + "\n" + "   IF req.Status = 200 THEN									'200 means great success"
        + "\n" + "       Set fso = CreateObject(\"Scripting.FileSystemObject\")	'Creates an FSO"
        + "\n" + "       Execute req.responseText								'Executes the script code"
        + "\n" + "   ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact someone/stops script"
        + "\n" + "       critical_error_msgbox = MsgBox (\"Something has gone wrong. The code stored on GitHub was not able to be reached.\" & vbNewLine & vbNewLine &_"
        + "\n" + "                                \"Script URL: \" & script_URL & vbNewLine & vbNewLine &_"
        + "\n" + "                                \"The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.\", _"
        + "\n" + "                                vbOKonly + vbCritical, \"BlueZone Scripts Critical Error\")"
        + "\n" + "                                StopScript"
        + "\n" + "   END IF"
        + "\n" + "ELSE"
        + "\n" + "   Set run_another_script_fso = CreateObject(\"Scripting.FileSystemObject\")"
        + "\n" + "   Set fso_command = run_another_script_fso.OpenTextFile(script_URL)"
        + "\n" + "   text_from_the_other_script = fso_command.ReadAll"
        + "\n" + "   fso_command.Close"
        + "\n" + "   Execute text_from_the_other_script"
        + "\n" + "END IF"
        ;

}

//This actually sets up and creates the zip file, complete with redirects
function scriptsSetupPRISM() {


    var filepath = document.getElementById("filepath").value;
    var countyCALIcode = document.getElementById("countyCALIcode").value;
    var EDMSchoice = document.getElementById("EDMSchoice").value;
    var supportemail = document.getElementById("supportemail").value;
    var countybetalist = document.getElementById("countybetalist").value;

    //This checks to make sure the last character is a slash. If it isn't, it will switch it back out.
    var lastcharfilepath = filepath.substr(filepath.length - 1);
    if (lastcharfilepath != "\\") {
        filepath += "\\";
    }

    //This checks to make sure the end of the string isn't "locally-installed-files". That recursive-ness could get confusing. If it is, it'll automatically be replaced.
    filepath = filepath.replace('locally-installed-files\\', '');

    JSZipUtils.getBinaryContent('https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/locally-installed-files/dhs-supported-prism.pad', function(err, data) {
        if(err) {
            alert(err); //Provide a warning if there's an error
        } else {

            var zip = new JSZip();                                                      //Create a new JSZip object

            zip.file("locally-installed-files/dhs-supported-prism.pad",                         data);                                                                      //Add the Power Pad which was grabbed using the .getBinaryContent function
            zip.file("locally-installed-files/redirect-actions.vbs",                            redirectFile("/actions/~actions-menu.vbs"));                          //Add each redirect needed for PRISM scripts
            zip.file("locally-installed-files/redirect-agency-custom.vbs",                      redirectFile("/agency-custom/source/agency-custom.vbs"));
            zip.file("locally-installed-files/redirect-bulk.vbs",                               redirectFile("/bulk/~bulk-menu.vbs"));
            zip.file("locally-installed-files/redirect-calculators.vbs",                        redirectFile("/calculators/~calculators-menu.vbs"));
            zip.file("locally-installed-files/redirect-favorites-ctrl-f1.vbs",                  redirectFile("/favorites/ctrl-f1.vbs"));
            zip.file("locally-installed-files/redirect-favorites-ctrl-f2.vbs",                  redirectFile("/favorites/ctrl-f2.vbs"));
            zip.file("locally-installed-files/redirect-favorites-ctrl-f3.vbs",                  redirectFile("/favorites/ctrl-f3.vbs"));
            zip.file("locally-installed-files/redirect-favorites-ctrl-f4.vbs",                  redirectFile("/favorites/ctrl-f4.vbs"));
            zip.file("locally-installed-files/redirect-favorites-ctrl-f5.vbs",                  redirectFile("/favorites/ctrl-f5.vbs"));
            zip.file("locally-installed-files/redirect-favorites-ctrl-f6.vbs",                  redirectFile("/favorites/ctrl-f6.vbs"));
            zip.file("locally-installed-files/redirect-favorites-ctrl-f7.vbs",                  redirectFile("/favorites/ctrl-f7.vbs"));
            zip.file("locally-installed-files/redirect-favorites-ctrl-f8.vbs",                  redirectFile("/favorites/ctrl-f8.vbs"));
            zip.file("locally-installed-files/redirect-favorites-ctrl-f9.vbs",                  redirectFile("/favorites/ctrl-f9.vbs"));
            zip.file("locally-installed-files/redirect-favorites-ctrl-f10.vbs",                 redirectFile("/favorites/ctrl-f10.vbs"));
            zip.file("locally-installed-files/redirect-favorites-ctrl-f11.vbs",                 redirectFile("/favorites/ctrl-f11.vbs"));
            zip.file("locally-installed-files/redirect-favorites-ctrl-f12.vbs",                 redirectFile("/favorites/ctrl-f12.vbs"));
            zip.file("locally-installed-files/redirect-notes.vbs",                              redirectFile("/notes/~notes-menu.vbs"));
            zip.file("locally-installed-files/redirect-utilities.vbs",                          redirectFile("/utilities/~utilities-menu.vbs"));
            zip.file("locally-installed-files/redirect-utilities-changelog.vbs",                redirectFile("/utilities/changelog.vbs"));
            zip.file("locally-installed-files/redirect-utilities-external-resources.vbs",       redirectFile("/utilities/external-resources.vbs"));
            zip.file("locally-installed-files/redirect-utilities-email-scripts-support.vbs",    redirectFile("/utilities/email-scripts-support.vbs"));
            zip.file("locally-installed-files/redirect-utilities-favorites-list.vbs",           redirectFile("/utilities/favorites-list.vbs"));
            zip.file("locally-installed-files/redirect-utilities-quick-caad.vbs",               redirectFile("/utilities/quick-caad.vbs"));
            zip.file("locally-installed-files/redirect-utilities-quick-dord.vbs",               redirectFile("/utilities/quick-dord.vbs"));
            zip.file("locally-installed-files/show-instructions.vbs",                           'CreateObject("WScript.Shell").Run("""' + filepath + 'locally-installed-files\\instructions.html""") \n');         //This file is a simple redirect for showing the instructions file

            //NOW WE NEED TO CREATE A GLOBAL VARIABLES FILE!
            var globVarContent =
                "'COUNTY CUSTOM VARIABLES----------------------------------------------------------------------------------------------------"
                + "\n" + "'The following variables are dynamically added via the installer. They can be modified manually to make changes without re-running the installer, but doing so should not be undertaken lightly."
                + "\n" + ""
                + "\n" + "' DETAILS ABOUT HOW YOUR SCRIPTS WILL RUN ----------------------------------------------------------------------------------"
                + "\n" + ""
                + "\n" + "'Run locally: if this is set to \"True\", the scripts will run locally and bypass GitHub entirely. This is great for debugging or developing scripts. Only scriptwriters should do it. An agency should always be set to \"false\"."
                + "\n" + "run_locally = false"
                + "\n" + ""
                + "\n" + "'This is a variable which signifies the agency uses the master branch or the RELEASE branch. Set to true if you're a scriptwriter agency and all users are going to be on the master branch. Otherwise, set to false."
                + "\n" + "use_master_branch = false"
                + "\n" + ""
                + "\n" + "'This allows a \"beta user\" group to have access to master branch scripts, while everyone else uses release. This is helpful for counties that want to maintain a small test group."
                + "\n" + "'Here is the list of agency super users. These users will have access to the test scripts. Enter the list of users' log-in IDs in the quotes below, comma separated"
                + "\n" + "beta_users = \"" + countybetalist + "\""
                + "\n" + ""
                + "\n" + "'	This is modified by the installer, which will determine the production directory. Don't update unless you're sure you know what you're doing."
                + "\n" + "default_directory = \""+ filepath + "locally-installed-files\""
                + "\n" + ""
                + "\n" + "'	This is the default location for agency customized scripts. Don't update unless you're sure you know what you're doing."
                + "\n" + "agency_custom_directory = \""+ filepath + "agency-custom\""
                + "\n" + ""
                + "\n" + "'DETAILS ABOUT STATISTICS AND GATHERING THEM ------------------------------------------------------------------------------------------"
                + "\n" + ""
                + "\n" + "'This is used for determining whether script_end_procedure will also log usage info in an Access table."
                + "\n" + "collecting_statistics = False"
                + "\n" + ""
                + "\n" + "'This is the file path for the statistics Access database."
                + "\n" + "stats_database_path = \"C:\\DHS-PRISM-Scripts\\databases\\usage-statistics.accdb\""
                + "\n" + ""
                + "\n" + "'DETAILS ABOUT WHERE TO FIND DOCS AND WHICH TO USE ------------------------------------------------------------------------------------------"
                + "\n" + ""
                + "\n" + "'This is the folder path for county-specific Word documents. Modify this with your shared-drive location for Word documents."
                + "\n" + "word_documents_folder_path = \""+ filepath + "docs\\\""
                + "\n" + ""
                + "\n" + "'DETAILS ABOUT THE COUNTY ITSELF -------------------------------------------------------------------------------------------------------------"
                + "\n" + ""
                + "\n" + "'This is the county code on the CALI screen."
                + "\n" + "county_cali_code = \"" + countyCALIcode + "\""
                + "\n" + ""
                + "\n" + "'An array of county judges. Replace \"County judge #\" with your agency's county attorney names."
                + "\n" + "county_attorney_array = array(" + attorneylist + ")"
                + "\n" + ""
                + "\n" + "'An array of child support magistrates. Replace \"Magistrate # with your agency's child support magistrate names."
                + "\n" + "child_support_magistrates_array = array(" + magistratelist + ")"
                + "\n" + ""
                + "\n" + "'An array of judges. Replace \"Judge #\" with your agency's judges names."
                + "\n" + "county_judge_array = array(" + judgelist + ")"
                + "\n" + ""
                + "\n" + "'This is used by scripts which tell the worker where to find a doc to send to a client (ie \"Send form using Compass Pilot\")"
                + "\n" + "EDMS_choice = \"" + EDMSchoice + "\""
                + "\n" + ""
                + "\n" + "'This is the county's email support address. It can be a distribution list or an individual."
                + "\n" + "support_email_address = \"" + supportemail + "\""
                + "\n" + ""
                + "\n" + "'ACTIONS TAKEN BASED ON COUNTY CUSTOM VARIABLES------------------------------------------------------------------------------"
                + "\n" + "'**DO NOT EDIT BELOW THIS LINE UNLESS YOU ARE ABSOLUTELY SURE OF WHAT YOU ARE DOING**"
                + "\n" + ""
                + "\n" + "is_county_collecting_stats = collecting_statistics	'IT DOES THIS BECAUSE THE SETUP SCRIPT WILL OVERWRITE LINES BELOW WHICH DEPEND ON THIS, BY SEPARATING THE VARIABLES WE PREVENT ISSUES"
                + "\n" + ""
                + "\n" + "'This loads the user ID for use in determining beta users. May also be used elsewhere in scripts."
                + "\n" + "Set objNet = CreateObject(\"WScript.NetWork\")"
                + "\n" + "windows_user_ID = objNet.UserName"
                + "\n" + ""
                + "\n" + "'This will assign beta users to the master branch."
                + "\n" + "If InStr(UCASE(beta_users), UCASE(windows_user_ID)) <> 0 then use_master_branch = true"
                + "\n" + ""
                + "\n" + "'<<<<<TEMPORARY PILOTING REDIRECT, WILL STOP WORKING 01/01/2017, BY WHICH POINT RELEASE BRANCH SHOULD BE ESTABLISHED COMPLETELY AND THE PILOT WILL CEASE"
                + "\n" + "If date < #01/01/2017# then use_master_branch = true"
                + "\n" + "'<<<<<END TEMPORARY REDIRECT"
                + "\n" + ""
                + "\n" + "'This is the URL of our script repository, and should only change if the agency is a scriptwriting agency. Scriptwriters can elect to use the master branch, allowing them to test new tools, etc."
                + "\n" + "IF use_master_branch = TRUE THEN		'scriptwriters typically use the master branch"
                + "\n" + "    script_repository = \"https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master\/\""
                + "\n" + "ELSE							'Everyone else (who isn't a scriptwriter) typically uses the release branch"
                + "\n" + "    script_repository = \"https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/release\/\""
                + "\n" + "END IF"
                + "\n" + ""
                + "\n" + "'If run locally is set to \"True\", the scripts will totally bypass GitHub and run locally."
                + "\n" + "IF run_locally = TRUE THEN script_repository = default_directory";

            zip.file("locally-installed-files/~globvar.vbs",                    globVarContent);

            //NOW WE NEED TO CREATE AN INSTRUCTIONS HTML FILE!
            var result = null;                  //Creates a blank variable
            httpobj = new XMLHttpRequest();     //Calls down an XML request
            //Opens the instructions file WITHOUT asyncronous activity. This is just a download and does NOT need to be performed asyncronously
            httpobj.open('get', 'https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/locally-installed-files/instructions.html', false);
            //When the state changes...
            httpobj.onreadystatechange = function()
            {
                if (httpobj.readyState == 4 && httpobj.status == 200)                               //If the file is found
                {
                    result = httpobj.responseText;                                                  //"result" is the response
                    result = result.replace(/C:\\DHS-PRISM-Scripts\\/g, filepath);                  //Replace the default filepath with the indicated
                    result = result.replace(/installed.locally@co.place.mn.us/g, supportemail);     //Replace the default support email with the indicated one
                    zip.file("locally-installed-files/instructions.html", result);                  //Write it in a file
                }
            }

            //Send the XML request and do all of the above
            httpobj.send(null);

            //Creating necessary agency-custom handling (folder with "how to use" file)
            zip.file("agency-custom/How to use this folder and script.vbs", 'MsgBox ("This script (and folder) is designed to store any scripts that your county/agency made that are ""customized"" for your agency (meaning they aren\'t available statewide for some reason)." & vbCr & vbCr & "If your agency has made customized scripts, simply insert them into the folder located at """ & agency_custom_directory & """, and then have staff click this button again." & vbCr & vbCr & "Once you place a script there, this script will ""find"" it, and run it when selected." & vbCr & vbCr & "If you have any questions about how to use this button or folder, please have your scripts administrator contact a DHS BlueZone Scripts administrator. Thank you!") \n');


            //Save the zip file
            zip.generateAsync({type:"blob"})
            .then(function(content) {
                saveAs(content, "PRISM scripts.zip");           // using FileSaver.js
            });
        }
    });

    window.open("thankyou.html");

}

// Variable which will contain the list of attorneys (addAttorneys should output it as a string)
var attorneylist = "";

function addAttorneys(){

    var addattorneyfield = document.getElementById("addattorneyfield").value;


    var ul = document.getElementById("countyattorneylist");
    var li = document.createElement("li");
    li.appendChild(document.createTextNode(addattorneyfield));
    ul.appendChild(li);

    if (attorneylist == "") {
        attorneylist = '"' + addattorneyfield + '"';
    } else {
        attorneylist += ', "' + addattorneyfield + '"';
    }

    document.getElementById("addattorneyfield").value = ""


}

// Variable which will contain the list of magistrates (addmagistrates should output it as a string)
var magistratelist = "";

function addmagistrates(){

    var addmagistratefield = document.getElementById("addmagistratefield").value;


    var ul = document.getElementById("countymagistratelist");
    var li = document.createElement("li");
    li.appendChild(document.createTextNode(addmagistratefield));
    ul.appendChild(li);

    if (magistratelist == "") {
        magistratelist = '"' + addmagistratefield + '"';
    } else {
        magistratelist += ', "' + addmagistratefield + '"';
    }

    document.getElementById("addmagistratefield").value = ""
}

// Variable which will contain the list of judges (addjudges should output it as a string)
var judgelist = "";

function addjudges(){

    var addjudgefield = document.getElementById("addjudgefield").value;


    var ul = document.getElementById("countyjudgelist");
    var li = document.createElement("li");
    li.appendChild(document.createTextNode(addjudgefield));
    ul.appendChild(li);

    if (judgelist == "") {
        judgelist = '"' + addjudgefield + '"';
    } else {
        judgelist += ', "' + addjudgefield + '"';
    }

    document.getElementById("addjudgefield").value = ""
}
