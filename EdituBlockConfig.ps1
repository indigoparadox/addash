#
# EditManagedBookmarks.ps1
# Last Updated: indigoparadox, 2019/04/02
#

$MyDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module $MyDir\FormWrappers.ps1

# Find out which GPO has the managed bookmarks and return the master list.
Function Get-uBlockSettingsGPO {
    Get-GPO -All | ForEach-Object {
        $ublockSettingsKey = Get-GPPrefRegistryValue -Guid $_.Id -Context Computer `
            -Key "HKLM\SOFTWARE\Policies\Google\Chrome\3rdparty\Extensions\cjpalhdlnbpafiamejdnhcphjbkeiagm\policy" `
            -ValueName "adminSettings" -ErrorAction SilentlyContinue
        If( $null -ne $ublockSettingsKey ) {
            $ublockSettingsStruct = $($ublockSettingsKey.Value | Out-String | ConvertFrom-Json -Verbose)
            $ublockSettingsGPO = [PSCustomObject]@{
                Settings = $ublockSettingsStruct
                SettingsRawJson = $ublockSettingsKey.Value | Out-String
                GPOGuid = $_.Id
                GPOName = $_.DisplayName
            }
            Return $ublockSettingsGPO
        }
    }
}

# Simple dialog for editing bookmark name/URL.
Function Show-UBWhitelistEditBox {
    Param(
        [PSObject]$WLInput = ""
    )
    
    $editForm = New-ADDForm -Title $('Editing Whitelist')
    
    # Create the dialog.
    $bmRenameText = Format-TextBox -Name 'Name' -Value $WLInput
    $null = $bmRenameText | New-ADDFormControl -Parent $editForm -Layout 'MainLayout' -LabelText 'Domain'
    
    $null = New-ADDFormPanel -Columns 2 -Name 'ButtonsLayout' | `
        New-ADDFormControl -Parent $editForm -Layout 'MainLayout'
    $null = Format-Button -DialogResult OK -Label 'OK' | `
        New-ADDFormControl -Parent $editForm -Horizontal $true -Layout 'ButtonsLayout'
    $null = Format-Button -DialogResult Cancel -Label 'Cancel' | `
        New-ADDFormControl -Parent $editForm -Horizontal $true -Layout 'ButtonsLayout'

	$formResult = $editForm.ShowDialog()

    # Return text on OK or null on cancel.
    If( $formResult -eq [System.Windows.Forms.DialogResult]::OK ) {
        Return $bmRenameText.Text
    } Else {
        Return $null
    }
}

# Present the bookmarks list. This function recurses, with each iteration modifying the list until the
# "Close" button is pressed, at which point the final list is passed back up to the top.
Function Show-Whitelist {
    Param(
        [PSCustomObject[]]$WLList
    )

    $mbNamesList = @()

    $wlForm = New-ADDForm -Title 'uBlock Whitelisted Domains'

    $Error.Clear()

    # Create the dialog.
	$WLListBox = Format-StringListBox -StringList $WLList
    $null = New-ADDFormControl -Parent $wlForm -Control $WLListBox -Layout 'MainLayout'

    $null = New-ADDFormPanel -Columns 4 -Name 'ButtonsLayout' | `
        New-ADDFormControl -Parent $wlForm -Layout 'MainLayout' -RowDivision 20
    $null = Format-Button -DialogResult OK -Label 'Create' | `
        New-ADDFormControl -Parent $wlForm -Horizontal $true -Layout 'ButtonsLayout'
    $null = Format-Button -DialogResult Retry -Label 'Edit' | `
        New-ADDFormControl -Parent $wlForm -Horizontal $true -Layout 'ButtonsLayout'
    $null = Format-Button -DialogResult Ignore -Label 'Remove' | `
        New-ADDFormControl -Parent $wlForm -Horizontal $true -Layout 'ButtonsLayout'
    $null = Format-Button -DialogResult Cancel -Label 'Close' | `
        New-ADDFormControl -Parent $wlForm -Horizontal $true -Layout 'ButtonsLayout'

	$formResult = $wlForm.ShowDialog()
    $wlIndex = $WLListBox.SelectedIndex

    # Parse the button feedback from the dialog.
    # Note that OK/Retry/Ignore are not the literal button labels. See DialogResult parms above
    # for translations.
    If( $formResult -eq [System.Windows.Forms.DialogResult]::OK ) {
        # Create
        $newItem = Show-UBWhitelistEditBox
        If( $newItem -ne $null ) {
            Write-Debug "Added $newItem to list."
            $WLList += $newItem
        }
        $WLList = $(Show-Whitelist -WLList $WLList)
    } ElseIf( $formResult -eq [System.Windows.Forms.DialogResult]::Retry ) {
        # Edit
        $editedItem = Show-UBWhitelistEditBox -WLInput $WLList[$wlIndex]
        If( $editedItem -ne $null ) {
            Write-Debug "Edited ($wlIndex): $editedItem"
            $WLList[$wlIndex] = $editedItem
        }
        $WLList = $(Show-Whitelist -WLList $WLList)
    } ElseIf( $formResult -eq [System.Windows.Forms.DialogResult]::Ignore ) {
        # Remove
        $removeItem = $WLList[$wlIndex]
        $WLList = $WLList | Where-Object { $_ -ne $removeItem }
        $WLList = $(Show-Whitelist -WLList $WLList)
    } Else {
        # Close
    }

    Return $WLList
}

# Show the resulting JSON for manual verification before committing.
Function Get-SaveBox {
    Param(
        [string]$JsonInput
    )

    $jsonForm = New-ADDForm -Title 'uBlock Settings' -Width 480 -Height 640

    $Error.Clear()

    $jsonText = Format-TextBox -Name 'JSON' -Lines 20 -Value $JsonInput
    $null = $jsonText | New-ADDFormControl -Parent $jsonForm -Layout 'MainLayout' `
        -LabelrowDivision 10 -LabelText 'Do you wish to save the following JSON?'

    $null = New-ADDFormPanel -Columns 2 -Name 'ButtonsLayout' | `
        New-ADDFormControl -Parent $jsonForm -Layout 'MainLayout' -RowDivision 20
    $null = Format-Button -DialogResult Yes -Label 'Yes' | `
        New-ADDFormControl -Parent $jsonForm -Horizontal $true -Layout 'ButtonsLayout'
    $null = Format-Button -DialogResult No -Label 'No' | `
        New-ADDFormControl -Parent $jsonForm -Horizontal $true -Layout 'ButtonsLayout'

	$formResult = $jsonForm.ShowDialog()

    If( $formResult -eq [System.Windows.Forms.DialogResult]::Yes ) {
        Return $true
    } Else {
        Return $false
    }
}

#$AdminCredential = Get-AdminCredential -AdminName 'domain\administrator' -SavePasswords $True -RegistryPath 'Software\MBE'

$ublockSettingsGPO = Get-uBlockSettingsGPO
Write-Host "Old settings: " + $ublockSettingsGPO.SettingsRawJson
# Parse uBlock's weird format with literal "\n"s.
$ublockWL = $ublockSettingsGPO.Settings.netWhitelist -Split( "\n" )
$whiteListNew = Show-Whitelist -WLList $ublockWL

$ublockSettingsGPO.Settings.netWhitelist = $whiteListNew -join( "\n" )
$ubJson = $ublockSettingsGPO.Settings | ConvertTo-Json | Out-String
# Get rid of pretty formatting (newlines, spaces).
$ubJson = $ubJson -replace "`n", ""
$ubJson = $ubJson -replace "    ", ""
$ubJson = $ubJson -replace '\\n', "n"
Write-Host "New settings: " + $ubJson

If( Get-SaveBox -JsonInput $ubJson ) {
#    Set-GPRegistryValue -Guid $ublockSettingsGPO.GPOGuid `
#        -Key "HKCU\Software\Policies\Google\Chrome" `
#        -Type String `
#        -ValueName "ManagedBookmarks" -Value $bmJson
   Out-Dialog -Message "Saved!"
}
