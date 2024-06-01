<#
===================================================================================================================================
 _____  _     _        _ _           _   _                _____                         __  __                                   
|  __ \(_)   | |      (_) |         | | (_)              / ____|                       |  \/  |                                  
| |  | |_ ___| |_ _ __ _| |__  _   _| |_ _  ___  _ __   | |  __ _ __ ___  _   _ _ __   | \  / | __ _ _ __   __ _  __ _  ___ _ __ 
| |  | | / __| __| '__| | '_ \| | | | __| |/ _ \| '_ \  | | |_ | '__/ _ \| | | | '_ \  | |\/| |/ _` | '_ \ / _` |/ _` |/ _ \ '__|
| |__| | \__ \ |_| |  | | |_) | |_| | |_| | (_) | | | | | |__| | | | (_) | |_| | |_) | | |  | | (_| | | | | (_| | (_| |  __/ |   
|_____/|_|___/\__|_|  |_|_.__/ \__,_|\__|_|\___/|_| |_|  \_____|_|  \___/ \__,_| .__/  |_|  |_|\__,_|_| |_|\__,_|\__, |\___|_|   
                                                                               | |                                __/ |          
                                                                               |_|                               |___/           
===================================================================================================================================
.SYNOPSIS
This script manages Exchange Online distribution groups by allowing admins or technicians to add or remove members quickly without needing knowledge of EXO PowerShell cmdlets.

.DESCRIPTION
The 'Distribution Group Manager' PowerShell script provides a graphical user interface (GUI) for administrators to manage distribution groups in Exchange Online.
It allows for the addition and removal of users from selected groups. The script checks for the required ExchangeOnlineManagement module and installs it if necessary, then connects to Exchange Online using provided credentials.

.PARAMETER credential
This has been deprecated and is no longer used. The script now prompts for modern auth credentials in the browser when run to support MFA.
The PowerShell credential object used to connect to Exchange Online.

.EXAMPLE
.\DistributionGroupManager.ps1

.NOTES
Author: Ryan Schultz
Last Updated: 2024-06-01
Requires: PowerShell 5.1 or higher, ExchangeOnlineManagement module
#>

##*================================================================
# Set execution policy for app process
##*================================================================

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

##*================================================================
##* Check if EXO management module is installed and connect to EXO
##*================================================================

if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    # The module is not installed, install it
    Write-Host "Exchange Online Management module not found. Installing now..."
    
    # Install the Exchange Online Management module
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
    
    # Import the module after installation
    Import-Module ExchangeOnlineManagement
} else {
    # The module is installed, import the module
    Write-Host "Exchange Online Management module found. Importing..."
    Import-Module ExchangeOnlineManagement
}

# Connect-ExchangeOnline -Credential $credential # Deprecated
Connect-ExchangeOnline

##*===============================================================
##* Generate window UI
##*===============================================================

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Distribution Group Manager'
$form.Size = New-Object System.Drawing.Size(300,400) # Increased form height to 400

$addGroupButton = New-Object System.Windows.Forms.Button
$addGroupButton.Location = New-Object System.Drawing.Point(10,320)
$addGroupButton.Size = New-Object System.Drawing.Size(120,23)
$addGroupButton.Text = 'Add User'
$form.Controls.Add($addGroupButton)

$removeGroupButton = New-Object System.Windows.Forms.Button
$removeGroupButton.Location = New-Object System.Drawing.Point(150,320)
$removeGroupButton.Size = New-Object System.Drawing.Size(120,23)
$removeGroupButton.Text = 'Remove User'
$form.Controls.Add($removeGroupButton)

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(10,10)
$label1.Size = New-Object System.Drawing.Size(280,20)
$label1.Text = 'Select the distribution group(s):'
$form.Controls.Add($label1)

$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
$checkedListBox.Location = New-Object System.Drawing.Point(10,30)
$checkedListBox.Size = New-Object System.Drawing.Size(260,150) # Increased height to display more items
$checkedListBox.CheckOnClick = $true

##*===============================================================
##* Retrieve and list all distribution groups in the tenant
##*===============================================================

$groups = Get-DistributionGroup -ResultSize Unlimited | Select-Object -ExpandProperty Name
foreach ($group in $groups) {
    [void]$checkedListBox.Items.Add($group)
}
$form.Controls.Add($checkedListBox)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10,190)
$label2.Size = New-Object System.Drawing.Size(280,20)
$label2.Text = 'Please enter the user email:'
$form.Controls.Add($label2)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(10,210)
$textBox2.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox2)

##*===============================================================
##* Add user button click actions
##*===============================================================

$addGroupButton.Add_Click({
    $selectedGroups = $checkedListBox.CheckedItems
    $user = $textBox2.Text
    foreach ($group in $selectedGroups) {
        try {
            Add-DistributionGroupMember -Identity $group -Member $user -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to add user '$user' to '$group': $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    [System.Windows.Forms.MessageBox]::Show("User '$user' successfully added to selected groups.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

##*===============================================================
##* Remove user button click actions
##*===============================================================

$removeGroupButton.Add_Click({
    $selectedGroups = $checkedListBox.CheckedItems
    $user = $textBox2.Text
    foreach ($group in $selectedGroups) {
        try {
            Remove-DistributionGroupMember -Identity $group -Member $user -Confirm:$false -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to remove user '$user' from '$group': $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    [System.Windows.Forms.MessageBox]::Show("User '$user' successfully removed from selected groups.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

##*===============================================================
##* Display the window
##*===============================================================
$form.ShowDialog()
