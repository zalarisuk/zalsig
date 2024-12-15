// attempt to avoid global variables

const ZPH = (function () {

    var g_jsonData = {};
    //list of export templates
    var g_listTemplatesText = new Array();
    //data value associated with an export template
    var g_listTemplatesValue = new Array();
    //Default cookie duration (days) - no one likes a stale cookie.
    const intCookieLife = 9999;

    initialiseTinyMCE()

    //initialise the logo so we don't try and load it before it even exists - gotta love asynchronous processing.
    g_jsonData.logo = '';

    fetchJSON();
    
    //Default the out of office month to this month
    let astrMonths = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    document.getElementById('INPUT-OOOED-MMM').value = astrMonths[new Date().getMonth()];
    //Default the out of office years to this year
    document.getElementById('INPUT-OOOED-YYYY').value = new Date().getFullYear();
    document.getElementById('INPUT-OOOSD-YYYY').value = document.getElementById('INPUT-OOOED-YYYY').value;

       //Initialise drop down lists from JSON
    setupCountries();
    setupTemplates();
    setupLeaveOptions();

    //Initialise the page
    setupPanels();
    templateSelectionUpdate();
    countrySelectionUpdate();

    //If autoload settings has been enabled and saved, load the user settings at page load.
    if (getCookie('ZALARIS-AUTOLOAD') == 'true') LoadSettings();

    //Base initialisation of the country message do this after any auto loading of settings that may have set the country
    document.getElementById('INPUT-MESSAGE').value = g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].msg_corporate;

    //There are several implicit preview updates through the previous functions

    //Preload the hidden images for the download option
    initialiseImagesForDownload();

    Preview();



    //Populate the list of countries
    //This is achieved by enumerating the countries keys in the JSON
    function setupCountries() {
        let astrCountriesHTML = [];
        Object.keys(g_jsonData.countries).forEach(function(key)
        {
            astrCountriesHTML.push('<option value="' + key + '">' + key + '</option>');
        });
        document.getElementById("INPUT-COUNTRY").innerHTML = astrCountriesHTML.join('\n');
    }

    async function fetchJSON() {
        let response = await fetch('./main.json');
        g_jsonData = await response.json();
    }

    //Populate the export templates drop down list
    //This is achieved by enumerating the templates keys in the JSON
    function setupTemplates() {
        let strTemplatesHTML = [];
        Object.keys(g_jsonData.templates).forEach(function(key)
        {
            strTemplatesHTML.push('<option value="' + key + '">' + key + '</option>');
        });
        document.getElementById("INPUT-TEMPLATE").innerHTML = strTemplatesHTML.join('\n');
    }

    //Populate the list of leave options
    //This is achieved by enumerating the leave keys and name sub-values in the JSON
    function setupLeaveOptions() {
        let astrLeaveHTML = [];
        Object.keys(g_jsonData.leave).forEach(function(key)
        {
            astrLeaveHTML.push('<option value="' + key + '">' + g_jsonData.leave[key].name + '</option>');
        });
        document.getElementById("INPUT-TYPEOFLEAVE").innerHTML = astrLeaveHTML.join('\n');
    }

    //Set which panels are expanded and collapsed based on settings in JSON file
    function setupPanels() {
        setupPanelsHelper("panel_signature_mandatory_open", "panel_signaturemandatory");
        setupPanelsHelper("panel_signature_optional_open", "panel_signatureoptional");
        setupPanelsHelper("panel_ooo_mandatory_open", "panel_ooomandatory");
        setupPanelsHelper("panel_ooo_optional_open", "panel_ooooptional");
        setupPanelsHelper("panel_additional_optional_open", "panel_additional_optional");
        setupPanelsHelper("panel_settings_open", "panel_settings");
        setupPanelsHelper("panel_rich_template_open", "panel_richoutput");
        setupPanelsHelper("panel_html_template_open", "panel_htmloutput");
    }

    //Helper to collapse/expand panels
    function setupPanelsHelper(p_strKey, p_strID) {
        if (g_jsonData[p_strKey]) document.getElementById(p_strID).className = "panel-collapse in";
        else document.getElementById(p_strID).className = "panel-collapse collapse";
    }

    //Sets up the preview sections that have optional visibility based on the template selection
    function templateSelectionUpdate() {
        //Get the selected template
        //We'll read the settings from the templates section of the JSON using this value
        let strSelectedTemplate = document.getElementById("INPUT-TEMPLATE").value;

        //Enumerate the template keys and use them to set the section visibility
        //All keys must be specified in the template, but the code will automatically process all of them.
        for(let strKey in g_jsonData.templates[strSelectedTemplate])
        {
               setObjectVisibility(strKey,g_jsonData.templates[strSelectedTemplate][strKey]);
        }

        //Force a refresh of the output preview
        Preview();
    }

    //Sets up the areas affacted by a change in country
    function countrySelectionUpdate() {
        //Initialise the Corporate message
        updateTinyMCEContent();
    
        //Set the dialing code information
        document.getElementById('OUTPUT-COUNTRY').innerHTML = document.getElementById("INPUT-COUNTRY").value;
        document.getElementById('OUTPUT-DIALCODE').innerHTML = g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].dialcode;
        //Because some regions can have multiple dial codes, we're going to pick the first space separated entry from the sentence to get the default.
        //This saves having to have duplicates for the majority of entries in the JSON. The wording supports this approach in any case.
        document.getElementById('INPUT-DIALCODE').value = g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].dialcode.split(' ')[0];

        //Populate the name of the support message info in the out of office details (optional) section
        document.getElementById('OUTPUT-SUPPORT-MSG-COUNTRY').innerHTML = document.getElementById("INPUT-COUNTRY").value;
        if (g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].msg_support.length > 0) document.getElementById('OUTPUT-SUPPORT-MSG').innerHTML = g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].msg_support;
        else document.getElementById('OUTPUT-SUPPORT-MSG').innerHTML = "<em>No support message has been specified for this country.</em>"

        //Update the preview
        Preview();
    }

    //Update the selection for TinMCE from the JSON settings file
    function updateTinyMCEContent() {
        tinymce.get("INPUT-MESSAGE").setContent(g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].msg_corporate);
    }

    //Sets an objects visibility
    function setObjectVisibility(p_strObjectID, p_bMakeVisible) {
        if (p_bMakeVisible) document.getElementById(p_strObjectID).style.display = 'inline';
        else document.getElementById(p_strObjectID).style.display = 'none';
    }

    //Save input fields to cookies and feedback to user by outputting a status area update.
    function SaveSettings() {
        //Signature details
        setCookie('ZALARIS-FIRSTNAME', document.getElementById('INPUT-FIRSTNAME').value, intCookieLife);
        setCookie('ZALARIS-SURNAME', document.getElementById('INPUT-SURNAME').value, intCookieLife);
        setCookie('ZALARIS-JOBTITLE', document.getElementById('INPUT-JOBTITLE').value, intCookieLife);
        setCookie('ZALARIS-COUNTRY', document.getElementById('INPUT-COUNTRY').value, intCookieLife);
    
        //Signature Details (Optional)
        setCookie('ZALARIS-DIALCODE', document.getElementById('INPUT-DIALCODE').value, intCookieLife);
        setCookie('ZALARIS-MOBILE', document.getElementById('INPUT-MOBILE').value, intCookieLife);
        setCookie('ZALARIS-EMAIL', document.getElementById('INPUT-EMAIL').value, intCookieLife);
        setCookie('ZALARIS-LINKEDIN', document.getElementById('INPUT-LINKEDIN').value, intCookieLife);
        setCookie('ZALARIS-XING', document.getElementById('INPUT-XING').value, intCookieLife);
        setCookie('ZALARIS-TWITTER', document.getElementById('INPUT-TWITTER').value, intCookieLife);
        setCookie('ZALARIS-SCN', document.getElementById('INPUT-SCN').value, intCookieLife);
        setCookie('ZALARIS-HOURS', document.getElementById('INPUT-HOURS').value, intCookieLife);
        setCookie('ZALARIS-MEETING-TEXT', document.getElementById('INPUT-MEETING-TEXT').value, intCookieLife);
        setCookie('ZALARIS-MEETING-LINK', document.getElementById('INPUT-MEETING-LINK').value, intCookieLife);
        setCookie('ZALARIS-SIGNOFF', document.getElementById('INPUT-SIGNOFF').value, intCookieLife);
        setCookie('ZALARIS-PREFNAME', document.getElementById('INPUT-PREFNAME').value, intCookieLife);
        setCookie('ZALARIS-CORP-WEB-SOCIAL-COMBINE', document.getElementById('INPUT-CORP-WEB-SOCIAL-COMBINE').checked, intCookieLife);
    
        //Out of office details
        setCookie('ZALARIS-LINEMGR', document.getElementById('INPUT-LINEMGR').value, intCookieLife);
        setCookie('ZALARIS-OOOED-DD', document.getElementById('INPUT-OOOED-DD').value, intCookieLife);
        setCookie('ZALARIS-OOOED-MMM', document.getElementById('INPUT-OOOED-MMM').value, intCookieLife);
        setCookie('ZALARIS-OOOED-YYYY', document.getElementById('INPUT-OOOED-YYYY').value, intCookieLife);
        setCookie('ZALARIS-OOOSD-DD', document.getElementById('INPUT-OOOSD-DD').value, intCookieLife);
        setCookie('ZALARIS-OOOSD-MMM', document.getElementById('INPUT-OOOSD-MMM').value, intCookieLife);
        setCookie('ZALARIS-OOOSD-YYYY', document.getElementById('INPUT-OOOSD-YYYY').value, intCookieLife);
        
        //Out of office details (Optional)
        //Note: Leave type doesn't get saved (on purpose)
        setCookie('ZALARIS-URGENCYMSG', document.getElementById('URGENCYMSG').checked, intCookieLife);
        setCookie('ZALARIS-SUPPORTMSG', document.getElementById('SUPPORTMSG').checked, intCookieLife);
    
        //Additional Settings (Optional)
        //Note: The corporate message is always loaded from the central settings file.
        setCookie('ZALARIS-AWAYMSG', document.getElementById('AWAYMSG').checked, intCookieLife);
        setCookie('ZALARIS-OPTMSG', document.getElementById('OPTMSG').checked, intCookieLife);
    
        //Auto Load
        setCookie('ZALARIS-AUTOLOAD', document.getElementById('AUTOLOAD').checked, intCookieLife);
        
        //Update the save message
        document.getElementById('status').innerHTML = 'Settings saved at ' + new Date().toLocaleString() + '.';
    }

    //Load cookie data to input fields, update the output areas and feedback to user by outputting a status area update.
    function LoadSettings() {
        //Checkboxes are more of a pain as implicit conversion from text to Boolean does not occur.
        //Instead we use a transitional variable to set each checkbox.
        let isTrueSet;
    
        //Signature details
        document.getElementById('INPUT-FIRSTNAME').value = getCookie('ZALARIS-FIRSTNAME');
        document.getElementById('INPUT-SURNAME').value = getCookie('ZALARIS-SURNAME');
        document.getElementById('INPUT-JOBTITLE').value = getCookie('ZALARIS-JOBTITLE');
        document.getElementById('INPUT-COUNTRY').value = getCookie('ZALARIS-COUNTRY');
    
        //Signature Details (Optional)
        //Note that the international dial code has to be set towards the end of the function.
        document.getElementById('INPUT-MOBILE').value = getCookie('ZALARIS-MOBILE');
        document.getElementById('INPUT-EMAIL').value = getCookie('ZALARIS-EMAIL');
        document.getElementById('INPUT-LINKEDIN').value = getCookie('ZALARIS-LINKEDIN');
        document.getElementById('INPUT-XING').value = getCookie('ZALARIS-XING');
        document.getElementById('INPUT-TWITTER').value = getCookie('ZALARIS-TWITTER');
        document.getElementById('INPUT-SCN').value = getCookie('ZALARIS-SCN');
        document.getElementById('INPUT-HOURS').value = getCookie('ZALARIS-HOURS');
        document.getElementById('INPUT-MEETING-TEXT').value = getCookie('ZALARIS-MEETING-TEXT');
        document.getElementById('INPUT-MEETING-LINK').value = getCookie('ZALARIS-MEETING-LINK');
        document.getElementById('INPUT-SIGNOFF').value = getCookie('ZALARIS-SIGNOFF');
        document.getElementById('INPUT-PREFNAME').value = getCookie('ZALARIS-PREFNAME');
        isTrueSet = (getCookie('ZALARIS-CORP-WEB-SOCIAL-COMBINE') == 'true');
        document.getElementById('INPUT-CORP-WEB-SOCIAL-COMBINE').checked = isTrueSet;
    
        //Out of office details    
        document.getElementById('INPUT-LINEMGR').value = getCookie('ZALARIS-LINEMGR');
        document.getElementById('INPUT-OOOSD-DD').value = getCookie('ZALARIS-OOOSD-DD');
        document.getElementById('INPUT-OOOSD-MMM').value = getCookie('ZALARIS-OOOSD-MMM');
        document.getElementById('INPUT-OOOSD-YYYY').value = getCookie('ZALARIS-OOOSD-YYYY');
        document.getElementById('INPUT-OOOED-DD').value = getCookie('ZALARIS-OOOED-DD');
        document.getElementById('INPUT-OOOED-MMM').value = getCookie('ZALARIS-OOOED-MMM');
        document.getElementById('INPUT-OOOED-YYYY').value = getCookie('ZALARIS-OOOED-YYYY');
    
        //Out of office details (Optional)
        //Note: Leave type doesn't get saved (on purpose)
        if (getCookie('ZALARIS-URGENCYMSG') == 'true') isTrueSet = true;
        else isTrueSet = false;
        document.getElementById('URGENCYMSG').checked = isTrueSet;
        if (getCookie('ZALARIS-SUPPORTMSG') == 'true') isTrueSet = true;
        else isTrueSet = false;
        document.getElementById('SUPPORTMSG').checked = isTrueSet;
    
        //Additional Settings (Optional)
        //Note: The corporate message is always loaded from the central settings file.
        if (getCookie('ZALARIS-OPTMSG') == 'true') isTrueSet = true;
        else isTrueSet = false;
        document.getElementById('OPTMSG').checked = isTrueSet;
        if (getCookie('ZALARIS-AWAYMSG') == 'true') isTrueSet = true;
        else isTrueSet = false;
        document.getElementById('AWAYMSG').checked = isTrueSet;
    
        //Auto load
        if (getCookie('ZALARIS-AUTOLOAD') == 'true') isTrueSet = true;
        else isTrueSet = false;
        document.getElementById('AUTOLOAD').checked = isTrueSet;
    
        //Trigger the country selection.
        //This implicitly triggers the preview rebuild too.
        countrySelectionUpdate();
        //Because the country selection defaults the international dial code, we have to explicitly load the IDC *after* this has occurred.
        document.getElementById('INPUT-DIALCODE').value = getCookie('ZALARIS-DIALCODE');
        //Because we have to do that, the implicit preview is no longer enough, and we have to launch an explicit preview in case the IDC is different to the default.
        Preview();
    
    
        //Update the load message
        document.getElementById('status').innerHTML = 'Settings loaded at ' + new Date().toLocaleString() + '.';
    }

    //Build an e-mail address from user's first name and surname
    //Now allowing user to specify manually and this will drive whether to include or not
    //Change for v2 due to Zalaris rebranding template changes
    /*
    function BuildEmail() {
        strAddress = document.getElementById('INPUT-FIRSTNAME').value.toLowerCase() + "." + document.getElementById('INPUT-SURNAME').value.toLowerCase() + g_jsonData.comms_domain;
        return '<a href="mailto: ' + strAddress + '">' + strAddress + '</a>';
    }
    */

    //Build an e-mail address from manager's name
    function BuildManagerEmail() {
        strAddress = document.getElementById('INPUT-LINEMGR').value.toLowerCase().replace(' ', '.') + g_jsonData.comms_domain;
        return '<a href="mailto: ' + strAddress + '">' + strAddress + '</a>';
    }

    //This function clears the arrays
    function ClearTemplateArrays() {
        g_listTemplatesText.length = 0;
        g_listTemplatesValue.length = 0;
    }

    //This function logically adds to the arrays in matched pairs of data
    function AddTemplate(p_strText, p_strValue) {
        g_listTemplatesText.push(p_strText);
        g_listTemplatesValue.push(p_strValue);
    }

    //Populate the list of out of office options
    function PopulateTemplatesListOutOfOffice() {
        //Use the AddTemplate function to populate the arrays.
        // The first entry in the arrays is the default.
        ClearTemplateArrays();
        AddTemplate('No phone, with signature', 'OOO_STDSIG');
        AddTemplate('Emergency phone, with signature', 'OOO_EMERSIG');
        AddTemplate('No phone, no signature', 'OOO_STD');
        AddTemplate('Emergency phone, no signature', 'OOO_EMER');
    
        //Clear the drop down list
        document.getElementById( 'listTemplates' ).innerHTML = "";
        
        //Add an option to the listTemplates for each array pair
        objSelect = document.getElementById( 'listTemplates' );
        for( objEntry in g_listTemplatesText )
        {
            var op = new Option();
            op.text = g_listTemplatesText[objEntry];
            op.value = g_listTemplatesValue[objEntry];
            objSelect.add( op );        
        };
        //The listTemplates drop down list should now be populated
    }

    //Helper function for concatenating multiple strings into a field for display in a text area.
    function appendToTemplateField(p_strInput1, p_strInput2) {
        return p_strInput1 + p_strInput2 + '\n';
    }

    //Generic set cookie
    function setCookie(p_strCookieName, p_strCookieValue, p_intExpiresAfterNDays) {
        var d = new Date();
        d.setTime(d.getTime() + (p_intExpiresAfterNDays*24*60*60*1000));
        var strExpires = "expires=" + d.toGMTString();
        document.cookie = p_strCookieName + "=" + p_strCookieValue + ";" + strExpires + ";path=/";
    }

    //Generic retrieve cookie
    function getCookie(p_strCookieName) {
        var name = p_strCookieName + "=";
        var decodedCookie = decodeURIComponent(document.cookie);
        var ca = decodedCookie.split(';');
        for(var i = 0; i < ca.length; i++) {
            var c = ca[i];
            while (c.charAt(0) == ' ') {
                c = c.substring(1);
            }
            if (c.indexOf(name) == 0) {
                return c.substring(name.length, c.length);
            }
        }
        return "";
    }

    function parseDate(strddMMMyyyy) {
        var aMonths = {January:0,February:1,March:2,April:3,May:4,June:5,July:6,August:7,September:8,October:9,November:10,December:11};
        var aDate = strddMMMyyyy.split(' ');
        return new Date(aDate[2], aMonths[aDate[1]], aDate[0]);
    }

    //Write the output section
    function WriteOutput() {
        document.getElementById('OUTPUT-HTML').value = document.getElementById('OUTPUT-PREVIEW').innerHTML;
    }

    
    function buildTwitterLink(p_strInput) {
        let strInput = p_strInput.trim();
        //Check if it is an @ handle
        if (strInput.charAt(0) == "@") return "https://twitter.com/" + strInput.substr(1);
        //Check if it begins with https://
        if (strInput.startsWith('https://')) return strInput;
        //If it begind with http:// make it https://
        if (strInput.startsWith('http://')) return "https://" + strInput.substr(7);
        //None of those applied, so we're going to have to assume it is a Twitter handle without the "@" prefix
        return "https://twitter.com/" + strInput;
    }

    //Convert an image to Base64 format
    function getBase64Image(p_img) {
        //Create a canvas and load the image into it
        let objCanvas = document.createElement("canvas");
        objCanvas.width = p_img.width;
        objCanvas.height = p_img.height;
        let ctx = objCanvas.getContext("2d");
        ctx.drawImage(p_img, 0, 0);
        
        //Convert the image to base64 data
        let strDataURL = objCanvas.toDataURL("image/png");
        let strReturn = strDataURL.replace(/^data:image\/(png|jpg);base64,/, "");
        
        //Clean up & return
        objCanvas.remove();
        return strReturn;
    }

    //Export (includes download) the HTML template to an "HTM" file that the user can copy into their
    //Outlook signature folder and potentially associated images, all bundled into a ZIP file.
    function exportSettings() {
        //Ask the user for a signature name
        let strName = window.prompt("Enter a file name","Signature")
    
        //Create ZIP file object
        let objZip = new JSZip();
        
        //add web page
        objZip.file(strName + ".htm", document.getElementById("OUTPUT-PREVIEW").outerHTML);
        
        //Add images
        g_jsonData.images.forEach(function(strImageName)
        {
            let strImgData = getBase64Image(document.getElementById(strImageName));
            objZip.file(strImageName, strImgData, {base64: true});
        });
        
        //Generate ZIP file
        objZip.generateAsync({type:"blob"})
        .then(function(content)
        {
            saveAs(content, strName + ".zip");
        });
    }

    //Preload images onto page ready for download
    function initialiseImagesForDownload() {
        //For each of the specified images the signature could use, put them into the page, in a hidden DIV (id=IMGTEMP)
        //We need to do this at the start before the download occurs.  If we do this during the download,
        //The code can race ahead of the page and try to access images that haven't yet been loaded.
        //While it would be possible to wait on each image load and use callbacks, that approach is more complex,
        //slows down the download process, and just feels like overkill when compared to a simple option like this.
        g_jsonData.images.forEach(function(strImageName)
        {
            document.getElementById("IMGTEMP").innerHTML += '<img id="' + strImageName + '" src="' + strImageName + '">';
        });
    }

    function initialiseTinyMCE() {
        //Not that two event functions have been added to update the preview when the editor is initialised,
        //and when it changes content.
        tinymce.init(
        {
            setup: function(editor)
            {
                editor.on('Change', function(e)
                {
                    //console.log('The TinyMCE Editor has changed.');
                    Preview();
                });
    
                editor.on('init', function(e)
                {
                      //console.log('The TinyMCE Editor has initialised.');
                      Preview();
                  });
            },
            selector: '#INPUT-MESSAGE',
            width: 800,
            height: 300,
            plugins: [
                'advlist autolink link image lists charmap print preview hr anchor pagebreak spellchecker',
                'searchreplace wordcount visualblocks visualchars code fullscreen insertdatetime media nonbreaking',
                'table emoticons template paste help'
            ],
            toolbar: 'undo redo | styleselect | bold italic | alignleft aligncenter alignright alignjustify | ' +
                'bullist numlist outdent indent | link image | print preview media fullpage | ' +
                'forecolor backcolor emoticons | help',
            menubar: 'favs file edit view insert format tools table help',
            font_formats: 'Calibri=calibri; Arial=arial,helvetica,sans-serif; Courier New=courier new,courier,monospace;',
            content_style: 'body { font-size: 10pt; font-family: Calibri; }'
        });
    }

    //Update the rich output and raw HTML output areas
    function Preview() {
    
        //Set the overall style for the output area. 
        //Necessary to override the body default from the BootStrap scaffolding library
        document.getElementById("OUTPUT-PREVIEW").setAttribute("style", g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].previewdiv_style);
    
        //<!-- OUT OF OFFICE: Company Settings Driven -->
        //<!-- Standard thanks for contacting me message when out of office -->
        document.getElementById('INPUT-OOOTHANKS-PREVIEW').innerHTML = g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].msg_ooo_thanks;
        //Open the list
        let strOOOMain = g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].msg_ooo_main_start;
        //Urgency maeesage
        if(document.getElementById("URGENCYMSG").checked == true) strOOOMain += g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].msg_ooo_main_urgent;
        //Close the list
        strOOOMain += g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].msg_ooo_main_end;
        //Support team member message
        if(document.getElementById("SUPPORTMSG").checked == true) strOOOMain += g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].msg_support;
        //Output the message
        document.getElementById('INPUT-OOOMSG-PREVIEW').innerHTML = strOOOMain;
    
    
        //<!-- OUT OF OFFICE: Company Settings Driven -->
        //<!-- Standard out of office statement -->
    
    
        //<!-- SIGNATURE: User Specified -->
        //<!-- Sign off from e-mail, e.g. "Regards, Stephen" -->
        if (document.getElementById('INPUT-SIGNOFF').value.length > 0)
        {
            document.getElementById('INPUT-SIGNOFF-PREVIEW').innerHTML = g_jsonData.signoff_prefix + 
                document.getElementById('INPUT-SIGNOFF').value + '<br><br>' + document.getElementById('INPUT-PREFNAME').value + 
                g_jsonData.signoff_suffix;
        } else document.getElementById('INPUT-SIGNOFF-PREVIEW').innerHTML = '';
        console.log(document.getElementById('INPUT-SIGNOFF-PREVIEW').innerHTML);
    
    
        //<!-- BOTH: User Specified -->
        //<!-- Full Name -->
        let strName =  g_jsonData.name_prefix + document.getElementById('INPUT-FIRSTNAME').value + " ";
        strName += document.getElementById('INPUT-SURNAME').value + g_jsonData.name_suffix;
        document.getElementById('INPUT-NAME-PREVIEW').innerHTML = strName;
    
    
        //<!-- BOTH: User Specified -->
        //<!-- Job Title/Designation -->
        document.getElementById('INPUT-DESIGNATION-PREVIEW').innerHTML = g_jsonData.designation_prefix + document.getElementById('INPUT-JOBTITLE').value + g_jsonData.designation_suffix;
    
    
        //<!-- BOTH: User Specified -->
        //<!-- Mobile Phone Number (+ own social links) -->
        let astrComms = [];
        if (document.getElementById('INPUT-MOBILE').value != '') astrComms.push(g_jsonData.comms_prefix + document.getElementById('INPUT-DIALCODE').value + 
            ' ' + document.getElementById('INPUT-MOBILE').value + g_jsonData.comms_suffix);
        if (document.getElementById('INPUT-EMAIL').value != '') astrComms.push(g_jsonData.comms_prefix + '<a href="mailto:' + document.getElementById('INPUT-EMAIL').value + '">' + document.getElementById('INPUT-EMAIL').value + '</a>' + g_jsonData.comms_suffix);
        if (document.getElementById('INPUT-LINKEDIN').value != '') astrComms.push(g_jsonData.comms_prefix + '<a href="' + document.getElementById('INPUT-LINKEDIN').value + '">LinkedIn</a>' + g_jsonData.comms_suffix);
        if (document.getElementById('INPUT-XING').value != '') astrComms.push(g_jsonData.comms_prefix + '<a href="' + document.getElementById('INPUT-XING').value + '">Xing</a>' + g_jsonData.comms_suffix);
        if (document.getElementById('INPUT-TWITTER').value != '') astrComms.push(g_jsonData.comms_prefix + '<a href="' + buildTwitterLink(document.getElementById('INPUT-TWITTER').value) + '">Twitter</a>' + g_jsonData.comms_suffix);
        if (document.getElementById('INPUT-SCN').value != '') astrComms.push(g_jsonData.comms_prefix + '<a href="' + document.getElementById('INPUT-SCN').value + '">SCN</a>' + g_jsonData.comms_suffix);
        strComms = astrComms.join(g_jsonData.comms_separator);
        if (strComms.length > 0) document.getElementById('INPUT-COMMS-PREVIEW').innerHTML = g_jsonData.comms_start + strComms + g_jsonData.comms_end;
        else document.getElementById('INPUT-COMMS-PREVIEW').innerHTML = "";
    
    
        //<!-- SIGNATURE: User Specified -->
        //<!-- Used to specify atypical working patterns - e.g. part time/shift work/return to work -->
        if (document.getElementById('INPUT-HOURS').value.length > 0 && g_jsonData.templates[document.getElementById("INPUT-TEMPLATE").value].msg_hours)
        {
            document.getElementById('INPUT-HOURS-PREVIEW').innerHTML = g_jsonData.msg_hours_prefix + document.getElementById('INPUT-HOURS').value + g_jsonData.msg_hours_suffix;
        }
        else document.getElementById('INPUT-HOURS-PREVIEW').innerHTML = '';
    
    
        //<!-- BOTH: Country Settings Driven -->
        //<!-- Zalaris Logo -->
        document.getElementById('INPUT-LOGO-PREVIEW').innerHTML = g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].logo;
    
    
        //<!-- BOTH: Company Settings Driven -->
        //<!-- Corporate mission -->
        document.getElementById('INPUT-CORPORATE-MISSION-PREVIEW').innerHTML = g_jsonData.mission;
    
    
        //<!-- SIGNATURE: User Specified -->
        //<!-- Upcoming leave information -->
        let strStart = document.getElementById('INPUT-OOOSD-DD').value + ' ' + document.getElementById('INPUT-OOOSD-MMM').value + ' ' + document.getElementById('INPUT-OOOSD-YYYY').value;
        let strEnd = document.getElementById('INPUT-OOOED-DD').value + ' ' + document.getElementById('INPUT-OOOED-MMM').value + ' ' + document.getElementById('INPUT-OOOED-YYYY').value;
        let strAway;
        let strLeaveType;
        strLeaveType = "Away - ";
        if (document.getElementById('INPUT-TYPEOFLEAVE').value == "Leave") strLeaveType = "Annual leave - ";
        if (document.getElementById('INPUT-TYPEOFLEAVE').value == "Training") strLeaveType = "Training Course - ";
    
        if (document.getElementById('INPUT-OOOSD-DD').value == "") strAway = g_jsonData.msg_away_prefix + strLeaveType + strEnd + g_jsonData.msg_away_suffix;
        else
        {
            if (strStart == strEnd) strAway = g_jsonData.msg_away_prefix + strLeaveType + strEnd + g_jsonData.msg_away_suffix;
            else strAway = g_jsonData.msg_away_prefix + strLeaveType + strStart + g_jsonData.msg_away_period_join + strEnd + g_jsonData.msg_away_suffix;
        }
        document.getElementById('INPUT-MSG-AWAY-PREVIEW').innerHTML = strAway;
        setObjectVisibility("INPUT-MSG-AWAY-PREVIEW",document.getElementById('AWAYMSG').checked);
    
    
        //<!-- BOTH: Country Settings Driven -->
        //<!-- Zalaris Tag Line -->
        document.getElementById('INPUT-TAGLINE-PREVIEW').innerHTML = g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].tagline;
    
    
        //<!-- BOTH: Country Settings Driven -->
        //<!-- Zalaris Web site-->
        //Load in if we've got the countries loaded in now
        if (document.getElementById("INPUT-COUNTRY").value == 'Loading...') document.getElementById('INPUT-WEBSITE-LINK-PREVIEW').innerHTML = '';
        else
        {
            //Hide the web site link section if it has been combined with the communications line
            setObjectVisibility("INPUT-WEBSITE-LINK-PREVIEW",!document.getElementById('INPUT-CORP-WEB-SOCIAL-COMBINE').checked);
            //Set the web link
            document.getElementById('INPUT-WEBSITE-LINK-PREVIEW').innerHTML = g_jsonData.website_prefix + '<a href="' + g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].website_url + '">' + g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].website_name + '</a>';
        }
    
    
        //<!-- BOTH: Company Settings Driven -->
        //<!-- Zalaris Social Media Links -->
        //Load in if we've got the countries loaded in now
        if (document.getElementById("INPUT-COUNTRY").value == 'Loading...') document.getElementById('INPUT-ZALARISLINKS-PREVIEW').innerHTML = '';
        else
        {
            //An array which builds up the elements of the corporate social communication line
            let astrCorpSocial = []
            //If the setting to combine the web site onto the social line is set, include  it
            if (document.getElementById('INPUT-CORP-WEB-SOCIAL-COMBINE').checked)
            {
                let strWebLink = '<a href="' + g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].website_url + '">' + g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].website_name + '</a>';
                astrCorpSocial.push(strWebLink);
            }
            //Read in the company settings for each social media account
            if (g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_social_linkedin.length > 0) astrCorpSocial.push(g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_social_linkedin);
    
            if (g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_social_twitter.length > 0) astrCorpSocial.push(g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_social_twitter);
    
            if (g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_social_facebook.length > 0) astrCorpSocial.push(g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_social_facebook);
    
            let strCorpSocial = g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_social_start;
            strCorpSocial += astrCorpSocial.join(g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_social_separator);
            strCorpSocial += g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_social_end;
            document.getElementById('INPUT-ZALARISLINKS-PREVIEW').innerHTML = strCorpSocial;
        }
        
        //<!-- BOTH: User Specified -->
        //<!-- Special Links -->
        if (document.getElementById('INPUT-MEETING-TEXT').value.length > 0)
        {
            document.getElementById('INPUT-SPECIAL-LINKS').innerHTML = g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_speciallink_start + 
                '📆&nbsp;<a href="' + document.getElementById('INPUT-MEETING-LINK').value + '">' + document.getElementById('INPUT-MEETING-TEXT').value + '</a>' +
                g_jsonData.speciallink_separator.toString() +
                g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_speciallink_end;
        }
        else 
        {
            document.getElementById('INPUT-SPECIAL-LINKS').innerHTML = g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_speciallink_start + 
                g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].company_speciallink_end;
        }
    
        //<!-- BOTH: Country Settings Driven -->
        //<!-- Corporate message -->
        if (document.getElementById('OPTMSG').checked)
        {
            if (tinymce.get("INPUT-MESSAGE").getContent() == null) document.getElementById('INPUT-CORP-MESSAGE-PREVIEW').innerHTML = "";
            else document.getElementById('INPUT-CORP-MESSAGE-PREVIEW').innerHTML = tinymce.get("INPUT-MESSAGE").getContent();
        }
        else document.getElementById('INPUT-CORP-MESSAGE-PREVIEW').innerHTML = "";
        
    
        //<!-- BOTH: Country Settings Driven -->
        //<!-- Legal disclaimer information -->
        if (document.getElementById("INPUT-COUNTRY").value == 'Loading...')
        {
            document.getElementById('INPUT-MAILDISCLAIMER-PREVIEW').innerHTML = '';
        }
        else
        {
            document.getElementById('INPUT-MAILDISCLAIMER-PREVIEW').innerHTML = g_jsonData.countries[document.getElementById("INPUT-COUNTRY").value].legal_info;
        }
    
    
        //>>>>>>> We're going to rebuild some of the information shown to the user too... <<<<<<<
        
    
        //Calculation of the return to work date
        let astrMonths = new Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");
        let dtReturn = parseDate(document.getElementById('INPUT-OOOED-DD').value + ' ' 
            + document.getElementById('INPUT-OOOED-MMM').value + ' ' 
            + document.getElementById('INPUT-OOOED-YYYY').value);
        dtReturn.setDate(dtReturn.getDate() + 1);
        let strReturnDate = dtReturn.getDate() + " " + astrMonths[dtReturn.getMonth()] + " " + dtReturn.getFullYear()
        document.getElementById('RETURNDATE').innerHTML = "Returning to work on " + strReturnDate;
    
    
        //Build the out of office reason and date(s).
        let strReasonandDate;
        //let bUseOneDate = (document.getElementById('INPUT-OOOSD-DD').value == "") || (strStart == strEnd);
    
    
        //Determine the absence dates
        let strAbsenceDates = "none";
        if (document.getElementById('INPUT-OOOED-DD').value.length > 0) strAbsenceDates = "single";
        if (document.getElementById('INPUT-OOOSD-DD').value.length > 0) strAbsenceDates = "range";
        //Look up the correct value from the JSON for the type and dates specified
        strReasonandDate = g_jsonData.leave[document.getElementById('INPUT-TYPEOFLEAVE').value.toLowerCase()][strAbsenceDates];
        //Now update the date placeholders ($1/$2)
        if (strAbsenceDates == "single") strReasonandDate = strReasonandDate.replace(/\$1/g, strReturnDate);
        if (strAbsenceDates == "range") strReasonandDate = strReasonandDate.replace(/\$1/g, strStart).replace(/\$2/g, strEnd);
    
        //Dynamic entries - these are specified in the JSON file, so we have to check they exist and then trigger an update only if they do
        //Out of office template.
        if (!! document.getElementById('X-DATE')) document.getElementById('X-DATE').innerHTML = strReasonandDate;
        if (!! document.getElementById("X-MANAGER")) document.getElementById("X-MANAGER").innerHTML = document.getElementById('INPUT-LINEMGR').value;
        if (!! document.getElementById("X-MANAGEREMAIL")) document.getElementById("X-MANAGEREMAIL").innerHTML = BuildManagerEmail();
        if (!! document.getElementById("X-MOBILE"))
        {
            //If we have a dial code specified, prefix it and a space to the mobile number when we output it
            let strMobile = document.getElementById('INPUT-MOBILE').value;
            let strDialCode = document.getElementById('INPUT-DIALCODE').value;
            if(strDialCode.length == 0) document.getElementById("X-MOBILE").innerHTML = strMobile;
            else document.getElementById("X-MOBILE").innerHTML = strDialCode + " " + strMobile;
        }
        
    
        //Update the raw HTML output text area
        WriteOutput();
    }


    // Public functions

    return {
        Preview: Preview,
        countrySelectionUpdate: countrySelectionUpdate,
        LoadSettings: LoadSettings,
        SaveSettings: SaveSettings,
        templateSelectionUpdate: templateSelectionUpdate,
        exportSettings: exportSettings
    }
})();
