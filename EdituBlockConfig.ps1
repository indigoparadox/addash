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
    
    $bmRenameText = Format-TextBox -Name 'Name' -Value $bmNameIn
    $null = $bmRenameText | New-ADDFormControl -Parent $editForm -Layout 'MainLayout' -LabelText 'Name'
    
    $bmEditText = Format-TextBox -Name 'URL' -Value $bmURLIn
    $null = $bmEditText | New-ADDFormControl -Parent $editForm -Layout 'MainLayout' -LabelText 'URL'
    
    $null = New-ADDFormPanel -Columns 2 -Name 'ButtonsLayout' | `
        New-ADDFormControl -Parent $editForm -Layout 'MainLayout'
    $null = Format-Button -DialogResult OK -Label 'OK' | `
        New-ADDFormControl -Parent $editForm -Horizontal $true -Layout 'ButtonsLayout'
    $null = Format-Button -DialogResult Cancel -Label 'Cancel' | `
        New-ADDFormControl -Parent $editForm -Horizontal $true -Layout 'ButtonsLayout'

	$formResult = $editForm.ShowDialog()

    If( $formResult -eq [System.Windows.Forms.DialogResult]::OK ) {
        Return [PSCustomObject]@{
            name = $bmRenameText.Text
            url = $bmEditText.Text
        }
    } Else {
        Return $null
    }
}

# Present the bookmarks list. This function recurses, with each iteration modifying the list until the
# "Close" button is pressed, at which point the final list is passed back up to the top.
Function Show-BookmarksList {
    Param(
        [PSCustomObject[]]$WLList
    )

    $mbNamesList = @()

    $mbForm = New-ADDForm -Title 'uBlock Whitelisted Domains'

    $Error.Clear()

	$mbListBox = Format-StringListBox -StringList $WLList
    $null = New-ADDFormControl -Parent $mbForm -Control $mbListBox -Layout 'MainLayout'

    $null = New-ADDFormPanel -Columns 4 -Name 'ButtonsLayout' | `
        New-ADDFormControl -Parent $mbForm -Layout 'MainLayout' -RowDivision 20
    $null = Format-Button -DialogResult OK -Label 'Create' | `
        New-ADDFormControl -Parent $mbForm -Horizontal $true -Layout 'ButtonsLayout'
    $null = Format-Button -DialogResult Retry -Label 'Edit' | `
        New-ADDFormControl -Parent $mbForm -Horizontal $true -Layout 'ButtonsLayout'
    $null = Format-Button -DialogResult Ignore -Label 'Remove' | `
        New-ADDFormControl -Parent $mbForm -Horizontal $true -Layout 'ButtonsLayout'
    $null = Format-Button -DialogResult Cancel -Label 'Close' | `
        New-ADDFormControl -Parent $mbForm -Horizontal $true -Layout 'ButtonsLayout'

	$formResult = $mbForm.ShowDialog()
    $mbIndex = $mbListBox.SelectedIndex

    If( $formResult -eq [System.Windows.Forms.DialogResult]::OK ) {
        # Create
        $newMB = Show-BookmarkEditBox
        If( $newMB -ne $null ) {
            Write-Debug "Created ${$newMB.name} to point to: ${$newMB.url}"
            $MBList += $newMB
        }
        $MBList = $(Show-BookmarksList -MBList $MBList)
    } ElseIf( $formResult -eq [System.Windows.Forms.DialogResult]::Retry ) {
        # Edit
        $editedMB = Show-BookmarkEditBox -BMInput $MBList[$mbIndex]
        If( $editedMB -ne $null ) {
            Write-Debug "Edited ($mbIndex) ${$editedMB.name} to point to: ${$editedMB.url}"
            $MBList[$mbIndex].url = $editedMB.url
            $MBList[$mbIndex].name = $editedMB.name
        }
        $MBList = $(Show-BookmarksList -MBList $MBList)
    } ElseIf( $formResult -eq [System.Windows.Forms.DialogResult]::Ignore ) {
        # Remove
        $removeName = $MBList[$mbIndex].name
        $MBList = $MBList | Where-Object { $_.name -ne $removeName }
        $MBList = $(Show-BookmarksList -MBList $MBList)
    } Else {
        # Close
    }

    Return $MBList
}

# Show the resulting JSON for manual verification before committing.
Function Get-SaveBox {
    Param(
        [string]$JsonInput
    )

    $jsonForm = New-ADDForm -Title 'Managed Bookmarks' -Width 480 -Height 640

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
$bookmarksListNew = Show-BookmarksList -WLList $ublockSettingsGPO.Settings.netWhitelist

#$bmJson = $bookmarksListNew | ConvertTo-Json | Out-String
#$bmJson = $bmJson -replace "`n", ""
#$bmJson = $bmJson -replace "    ", ""

#If( Get-SaveBox -JsonInput $bmJson ) {
#    Set-GPRegistryValue -Guid $ublockSettingsGPO.GPOGuid `
#        -Key "HKCU\Software\Policies\Google\Chrome" `
#        -Type String `
#        -ValueName "ManagedBookmarks" -Value $bmJson
#   Out-Dialog -Message "Saved!"
#}

#
