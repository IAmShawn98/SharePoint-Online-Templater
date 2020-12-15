# Powershell Window Title
$host.ui.RawUI.WindowTitle = "GTM: Shot! (SharePoint Online Templater)"

# Powershell Dependencies.
$wsh = New-Object -ComObject Wscript.Shell

# Let the user know the server is awaiting auth login data.
Write-Host "Waiting For Authentication...."

# Authenticate User If Admin Account Info Passes Credit Check....
if ($host.ui.PromptForCredential("Sign In", "Please sign in using a Microsoft account with valid admin tenant access to continue. If you do not have access to a valid admin account, please contact your network administrator for assistance.", '', "SYSTEM\Administrator")) {}
# Otherwise, let the user know their account info is invalid....
else {
    # Prompt user.
    $wsh.Popup("You have entered invalid administrative credentials. Please provide a valid username and password to continue.")
    # Exit program.
    exit 
}

# Connect to the GTM tenant site URL.
Connect-SPOService -Url https://gtmgrp-admin.sharepoint.com/

# Allow full access to site/script execution.
Set-ExecutionPolicy Unrestricted
Set-ExecutionPolicy Unrestricted -Force

## Start MenuFD
function StartMenu {
    Clear-Host
    Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Write-Host "                                                                                                                       
                                                          ttt:::t                                  
                                                          t:::::t                                  
                                                          t:::::t                                  
                                   ggggggggg   gggggttttttt:::::ttttttt       mmmmmmm    mmmmmmm   
                                  g:::::::::ggg::::gt:::::::::::::::::t     mm:::::::m  m:::::::mm 
                                 g:::::::::::::::::gt:::::::::::::::::t    m::::::::::mm::::::::::m
                                g::::::ggggg::::::ggtttttt:::::::tttttt    m::::::::::::::::::::::m
                                g:::::g     g:::::g       t:::::t          m:::::mmm::::::mmm:::::m
                                g:::::g     g:::::g       t:::::t          m::::m   m::::m   m::::m
                                g:::::g     g:::::g       t:::::t          m::::m   m::::m   m::::m
                                g::::::g    g:::::g       t:::::t    ttttttm::::m   m::::m   m::::m
                                g:::::::ggggg:::::g       t::::::tttt:::::tm::::m   m::::m   m::::m
                                 g::::::::::::::::g       tt::::::::::::::tm::::m   m::::m   m::::m
                                  gg::::::::::::::g         tt:::::::::::ttm::::m   m::::m   m::::m
                                    gggggggg::::::g           ttttttttttt  mmmmmm   mmmmmm   mmmmmm
                                            g:::::g 
                                            g:::::g             (        )     )          
                                            g:::::g             )\ )  ( /(  ( /(   *   )  
                                 g::::::ggg:::::::g             (()/(  )\()) )\())` )  /(  
                                 gg:::::::::::::g               /(_))((_)\ ((_)\  ( )(_)) 
                                 ggg::::::ggg                   (_))   _((_)  ((_)(_(_())  
                                     gggggg                     / __| | || | / _ \|_   _|  
                                                                \__ \ | __ || (_) | | |    
                                                                |___/ |_||_| \___/  |_|
                                                                "  
                                                              
    Write-Host "                            ---------------------------------------------------------------------------" 
    Write-Host "                                                                                                       "                      
    Write-Host "                            ╔═╗┬ ┬┌─┐┬─┐┌─┐╔═╗┌─┐┬┌┐┌┌┬┐  ╔═╗┌┐┌┬  ┬┌┐┌┌─┐  ╔╦╗┌─┐┌┬┐┌─┐┬  ┌─┐┌┬┐┌─┐┬─┐"
    Write-Host "                            ╚═╗├─┤├─┤├┬┘├┤ ╠═╝│ │││││ │   ║ ║││││  ││││├┤    ║ ├┤ │││├─┘│  ├─┤ │ ├┤ ├┬┘"
    Write-Host "                            ╚═╝┴ ┴┴ ┴┴└─└─┘╩  └─┘┴┘└┘ ┴   ╚═╝┘└┘┴─┘┴┘└┘└─┘   ╩ └─┘┴ ┴┴  ┴─┘┴ ┴ ┴ └─┘┴└─"
    Write-Host "
                            ---------------------------------------------------------------------------            " 
    Write-Host "                                                                                                       " 
    Write-Host "                             [1] Create Templates | [2] Delete Templates | [3] View Active Templates  
           "
    Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Write-Host "



                                                                                                                    "
    Write-Host "                                             © 2020 Global Tax Management, Inc.

                                           Create Dynamic GTM Themed SharePoint Templates
                                                                               
"
}
do {
    # Menu Select Functionality.
    # Renders the 'StartMenu' When the Program Runs.
    StartMenu

    # When A User Enters '1', Send Them to the SharePoint Authenticator.
    $MenuSelect = Read-Host " ";

    # Switch Function; If the User Types '1' Send Them to the SP Auth Func.
    switch ($MenuSelect) {
        '1' {
            # Site Script definitions.
            $scriptFile = $PSScriptRoot + "\customScript.json"
            $scriptTitle = "customScript"

            # Push Site Script Into SPOSite With Unique ID.
            $siteScriptId = (Get-Content $scriptFile -Raw | Add-SPOSiteScript -Title $scriptTitle) | Select-Object -First 1 Id

            # Collect SPOSite Design Definitions.
            $designTitle = Read-Host "Template Name"

            # Collect SPOSite Design Template Type.
            $designWebTemplate = Read-Host "Template Type (Enter '68' For Default)"

            # Collect SPOSite Design Description.
            $designDescription = Read-Host "Site Description"

            # Push Design Into SPOSite.
            Add-SPOSiteDesign -Title $designTitle -WebTemplate $designWebTemplate -SiteScripts  $siteScriptId.id -Description $designDescription

        }
        '2' {
            # Collect SPO Site ID For Deletion.
            Write-Host "Enter the site ID of the SP template you wish to delete."
            $SiteID = Read-Host " "

            # Perform Delete Site Action.
            Remove-SPOSiteDesign $SiteID

            $siteScriptId.id

            # Pause to Show Delete Success.
            pause

            # Clear Section UI.
            Clear-Host

            # Go Home.
            StartMenu
        }
        '3' {
            # Show All Active SPO Sites.
            Get-SPOSiteDesign
            
            # Allow Users to Stop and Copy Their ID Before Continiing.
            pause

            # Clear ID Table.
            Clear-Host

            # Go Home.
            StartMenu
        }

    }
}
# Keep This Switch Active Until the User Types '1' to Continue or Presses 'q' to Quit.
until ($input -eq '1')