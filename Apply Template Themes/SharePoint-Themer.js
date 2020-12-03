// GTM SharePoint Themer (Add-In) | A Thing by Shawn Luther | Created On: 10/27/2020 | Version: 1.0.0-Beta
(() => {
    // Force Close Script Function.
    function exit(code) {
        const prevOnError = window.onerror
        window.onerror = () => {
            window.onerror = prevOnError
            return true
        }
        // Debug message to throw in the console.
        throw new Error(`Failed to apply 'GTM Default Theme (Blue)' theme, failed with error code '${code || 0}'.`);
    }
    // Default Theme (Blue Theme).
    function DefaultThemeBlue(url, params) {
        // Container to hold rest request.
        let req = new XMLHttpRequest();

        // Pass Request When State Is Loaded.
        req.onreadystatechange = () => {
            if (req.readyState != 4) // Loaded
                return;
            console.log(req.responseText); // Log data response in the console for debugging.
        };
        // Try to run the code excecution script.
        try {
            // Ensure that prepending to URLs are always correct and extra white space and slashes are removed.
            let webBasedUrl = (_spPageContextInfo.webServerRelativeUrl + "//" + url).replace(/\/{2,}/, "/");
            req.open("POST", webBasedUrl, true);
            req.setRequestHeader("Content-Type", "application/json;charset=utf-8");
            req.setRequestHeader("ACCEPT", "application/json; odata.metadata=minimal");
            req.setRequestHeader("x-requestdigest", _spPageContextInfo.formDigestValue);
            req.setRequestHeader("ODATA-VERSION", "4.0");
            req.send(params ? JSON.stringify(params) : void 0);
        }
        // If there are errors, catch and let the user know, then end the script.
        catch {
            // Error Message.
            alert("Applying 'GTM Default Theme (Blue)' has failed because you aren't on the 'Site Contents' page of your SharePoint site. Please navigate to to the correct location to execute your custom SharePoint theme.")

            // Force Quit.
            exit();
        }
    }
    // Get the current URL to inject our theme onto when POST is called.
    let SPSite = document.referrer;

    // Theme Picker Interface.
    let ThemePicker = prompt("Please pick a theme for your new communication site: \n 1.) GTM Default Theme (Blue)\n\n Please type the number of the theme you want to provision and select 'OK' to confirm it.");

    // If '1' is typed, add 'GTM Default Theme (Blue)' to the site collection.
    if (ThemePicker === "1") {
        alert("You selected 1, setting theme 'GTM Default Theme (Blue)' now....");
        // POST For 'GTM Default Theme (Blue)'.
        DefaultThemeBlue("/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility.ApplySiteDesign", { siteDesignId: "6ee74453-5cae-47f1-a7cd-64dbe3c4fb18", "webUrl": SPSite });

        // Let the user know that theme injection was successful, then, reload page to show updated content.
        setTimeout(() => {
            alert("The 'GTM Default Theme (Blue)' has been added to this site collection! Click 'OK' to view the updated site theme.");
            window.location.reload();
        }, 1000);
    } else { // If the user enters an invalid theme number, handle it by prompting them.
        alert("Please enter a valid theme number!");
    }

})();