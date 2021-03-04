# SharePoint Online Templater (SHOT!) By Shawn Luther
Add-Type -AssemblyName System.Windows.Forms

# PowerShell Window Title.
$host.ui.RawUI.WindowTitle = "SharePoint Online Templater (SHOT!)"

# Store Dynamic Year Data.
$Year = Get-Date -UFormat %Y

# Powershell Dependencies.
$wsh = New-Object -ComObject Wscript.Shell

# Admin Tenant Menu.
function ConnectAdminTenant {
Write-Host "

















                        #########################################################################
                        #                                                                       #
                        #                   - SHAREPOINT ONLINE TEMPLATER -                     #
                        #                                                                       #
                        #########################################################################
                        #                                                                       #
                        #             Please Provide Your Admin Tenant URL to Continue          #
                        #                                                                       #
                        # --------------------------------------------------------------------- #
                        #                         - Tenant URL Example  -                       #
                        # --------------------------------------------------------------------- #
                        #                                                                       #
                        #              - https://gtmdemo1-admin.sharepoint.com/ -               #
                        #                                                                       #
                        #########################################################################
"

# Allow full access to site/script execution.
Set-ExecutionPolicy Unrestricted
Set-ExecutionPolicy Unrestricted -Force
}

# Renders 'Admin Tenant Menu'.
ConnectAdminTenant

# Connect to SharePoint Admin Tenant Domain.
$SPDomain = Read-Host " ";
Connect-SPOService -Url $SPDomain

# Start Menu (Home Section).
function StartMenu {
Clear-Host
Write-Host " 
                        Tenant Being Served: $SPDomain

 
                                            
                        ╔══════════════════════════════╗                                          
                            Welcome, $env:USERNAME!                                              
                        ╚══════════════════════════════╝

                        [D] Download Latest Patch | [I] Install Required Modules

                        #############################################################################
                        #---------------------------------------------------------------------------#
                        # - - - - - - - - - - - - SHAREPOINT SITE ACTIONS - - - - - - - - - - - - - #
                        #---------------------------------------------------------------------------#
                        #                 [1] Template Options | [2] Switch Tenants                 #
                        #---------------------------------------------------------------------------#
                        #--         ______       ___   ___      ______       _________            --#
                        #--        /_____/\     /__/\ /__/\    /_____/\     /________/\           --#
                        #--        \::::_\/_    \::\ \\  \ \   \:::_ \ \    \__.::.__\/           --#
                        #--         \:\/___/\    \::\/_\ .\ \   \:\ \ \ \      \::\ \             --#
                        #--          \_::._\:\    \:: ___::\ \   \:\ \ \ \      \::\ \            --#
                        #--            /____\:\    \: \ \\::\ \   \:\_\ \ \      \::\ \           --#
                        #--            \_____\/     \__\/ \::\/    \_____\/       \__\/           --#
                        #--                                                                       --#"             
Write-Host "                        #---------------------------------------------------------------------------#
                        #---------------------------------------------------------------------------#
                        #╔═╗┬ ┬┌─┐┬─┐┌─┐╔═╗┌─┐┬┌┐┌┌┬┐  ╔═╗┌┐┌┬  ┬┌┐┌┌─┐  ╔╦╗┌─┐┌┬┐┌─┐┬  ┌─┐┌┬┐┌─┐┬─┐#
                        #╚═╗├─┤├─┤├┬┘├┤ ╠═╝│ │││││ │   ║ ║││││  ││││├┤    ║ ├┤ │││├─┘│  ├─┤ │ ├┤ ├┬┘#
                        #╚═╝┴ ┴┴ ┴┴└─└─┘╩  └─┘┴┘└┘ ┴   ╚═╝┘└┘┴─┘┴┘└┘└─┘   ╩ └─┘┴ ┴┴  ┴─┘┴ ┴ ┴ └─┘┴└─#
                        #---------------------------------------------------------------------------#"
Write-Host "                        #---------------------------------------------------------------------------#
                        #- - - Create, Backup, and Easily Manage Your Teams SharePoint Sites - - -  #
                        #---------------------------------------------------------------------------#
                        #---------------------------------------------------------------------------#
                        #  - - - - - - - - -  © $Year Global Tax Management, Inc. - - - - - - - - -  #
                        #---------------------------------------------------------------------------#
                        #---------------------------------------------------------------------------#
                        # - - - - - - - - - - -  A Thing By Shawn Luther - - - - - - - - - - - - -  #
                        #---------------------------------------------------------------------------#
                        #############################################################################"     
}
do {
    # Renders 'Start Menu (Home Section)'.
    StartMenu

    # Collect menu data from users.
    $MenuSelect = Read-Host " ";
    Clear-Host

    # Switch Function; If the user selects option '1', let them choose to either create a template or set new theme default.
    switch ($MenuSelect) {
        '1' {
            # Data Collection For Picking Between Site Provisioning and Theme Defaulting.
            Write-Host "[1] Create Site Backup | [2] Push Backup Template | [b] Go Back"
            $MenuSelect = Read-Host " ";

            # Switch Function; Allows user to pick between changing theme defaults or provisioning a brand new site.
            Switch ($MenuSelect) {
                # Create Site Backup.
                '1' {
                    # Collect Site For Use In Backuping Up Data.
                    Clear-Host
                    Write-Host "Go to SharePoint and copy the URL of the site you wish to backup, then paste it here."
                    $SiteBackTarget = Read-Host " "

                    # Get Current Tenant Data.
                    Connect-PnPOnline -Url $SiteBackTarget -UseWebLogin
                    # Backup SharePoint Data in XML Format.
                    Get-PnPProvisioningTemplate -Out backup.xml
                    # Let the user know their data has finished processing.
                    $msgBody = "Your SharePoint site has been saved as 'backup.xml'!"
                    $msgTitle = "SharePoint Site Saved"
                    [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle)
                    Clear-Host

                    # Ask the user if they'd like to apply their newly generated backup or not.
                    Write-Host "Would you like to apply your backup or finish? Type 'Yes' or 'No'."
                    $YesNo = Read-Host " ";

                    # Switch; (Yes; No) Handle Template Backup Action Choices.
                    Switch ($YesNo) {
                        # If 'yes', let the user push a template back up.
                        'yes' {
                            Write-Host "Find the site backup you created and Copy/Paste it here."
                            $BackPath = Read-Host " ";
                            Write-Host "Now, target any existing SharePoint site to apply your backup theme."
                            $SiteOverride = Read-Host " ";
                            Connect-PnPOnline -Url $SiteOverride -UseWebLogin
                            Apply-PnPProvisioningTemplate $BackPath
                        }
                        # If 'no', send the user back home.
                        'no' {
                            # Go Home.
                            StartMenu
                        }
                    }
                }
                '2' {
                    Write-Host "Find the site backup you created and Copy/Paste it here."
                    $BackPath = Read-Host " ";
                    Write-Host "Now, target any existing SharePoint site to apply your backup theme."
                    $SiteOverride = Read-Host " ";
                    Connect-PnPOnline -Url $SiteOverride -UseWebLogin
                    Apply-PnPProvisioningTemplate $BackPath

                    StartMenu
                }
            }
        }

        '2' {
            # Disconnect User From Current Tenant.
            Disconnect-SPOService

            # Renders 'Admin Tenant Menu'.
            ConnectAdminTenant
            
            # Switch Admin Tenant Accounts.
            # Connect to SharePoint Admin Tenant Domain.
            $SPDomain = Read-Host " ";
            Connect-SPOService -Url $SPDomain
        }
        # Download Latest Patch of 'SHOT!'.
        'd' {
            # Let User Know the Latest Patch is Being Downloaded.
            Write-Host "Downloading Latest Patch From Github...."
            Start-Sleep -Seconds 8
            
            # Begin Download.
            Start-Process "https://github.com/IAmShawn98/SharePoint-Online-Templater/archive/main.zip"
            Clear-Host
            
            # Let User Know the Download Is Complete.
            Write-Host "The latest patch has completed downloading, check your downloads for the latest patch zip."
            pause
            
            # Go Home.
            StartMenu
        }
        # Install the Latest Version of SharePoint.
        'i' {
            # Install SharePoint Online Management Shell.
            Install-Module -Name Microsoft.Online.SharePoint.PowerShell -AllowClobber
            Install-Module SharePointPnPPowerShellOnline -AllowClobber
            pause
            
            # Go Home.
            StartMenu
        }

    }
}
# Keep This Switch Active Until the User Types '1' to Continue or Presses 'q' to Quit.
until ($input -eq '1')