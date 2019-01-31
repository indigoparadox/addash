#
# ADDash.ps1
# A convenient set of tools for doing common tasks with the Active Directory.
#

$LastSelectedIndex = 0

$MyDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module $MyDir\FormWrappers.ps1

Function Invoke-OnDC {
    Param(
        [Parameter( Mandatory=$true, Position=0 )]
        [System.Management.Automation.PSCredential] $AdminCredential,
        [Parameter( Mandatory=$true, Position=1 )]
        [scriptblock] $ScriptBlock
    )

    Return $(Invoke-Command -Credential $AdminCredential -ComputerName "dc01" `
        -ScriptBlock $ScriptBlock)
}

Function Get-RemoteADObject {
    Param(
        [Parameter( Mandatory=$true )]
        [string] $OU,
        [Parameter( Mandatory=$true )]
        [string] $ObjectType,
        [Parameter( Mandatory=$true )]
        [string] $Filter,
        [System.Management.Automation.PSCredential] $AdminCredential
    )

    $Properties = "DistinguishedName","Enabled","Name","SamAccountName","SID","LockedOut"

    If( $ObjectType -eq 'User' ) {
	    $ADDObjects = Invoke-OnDC $AdminCredential { Get-ADUser -SearchBase $Using:OU -Filter $Using:Filter -Properties $Using:Properties } |
            Sort-Object -Property Name
    } ElseIf( $ObjectType -eq 'Computer' ) {
	    $ADDObjects = Invoke-OnDC $AdminCredential { Get-ADComputer -SearchBase $Using:OU -Filter $Using:Filter -Properties $Using:Properties } | 
            Sort-Object -Property Name
    } ElseIf( $ObjectType -eq 'Group' ) {
	    $ADDObjects = Invoke-OnDC $AdminCredential { Get-ADGroup -SearchBase $Using:OU -Filter $Using:Filter } |
            Sort-Object -Property Name
    }

    Return ,$ADDObjects
}

Function Applet-Users {
    Param(
        [System.Management.Automation.PSCredential] $AdminCredential,
        [bool] $ShowDisabled
    )
    
    $UsersFilter = '*'
    If( -not $ShowDisabled ) {
        $UsersFilter = {(Enabled -eq $true)}
    }

	# Build the form.

	$ADDForm = New-ADDForm -Title 'AD Users'
	
	#$ADDUsers = Get-ADObject -SearchBase 'OU=Users,OU=Albany,DC=domain,DC=local' `
    #    -Filter $UsersFilter -Properties DistinguishedName,Enabled,Name,SamAccountName,SID,LockedOut | 
    #    Sort-Object -Property Name
    $Error.Clear()
    $ADDUsers = Get-RemoteADObject -OU 'OU=Users,OU=Albany,DC=domain,DC=local' `
        -ObjectType 'User' -Filter $UsersFilter -AdminCredential $AdminCredential
    If( $ADDUsers -eq $null -or $ADDUsers.Length -eq 0 ) {
        Out-Dialog -Message $Error -DialogType 'Error'
        Return
    }
	$ADDList = Format-ListBox -ObjectList $ADDUsers | `
        New-ADDFormControl -Parent $ADDForm -h 290 -w 300 -Anchor ( `
            [System.Windows.Forms.AnchorStyles]::Left + `
            [System.Windows.Forms.AnchorStyles]::Right + `
            [System.Windows.Forms.AnchorStyles]::Top + `
            [System.Windows.Forms.AnchorStyles]::Bottom)

    Format-Button -DialogResult OK -Label 'Hide Disabled' -LabelFalse 'Show Disabled' -LabelTest $ShowDisabled | `
        New-ADDFormControl -Parent $ADDForm -Anchor None -w 80 -h 60

    Format-Button -DialogResult Retry -Label 'Set Password' | `
        New-ADDFormControl -Parent $ADDForm -Anchor None -w 80 -h 60

    Format-Button -DialogResult Ignore -Label 'Unlock Account' | `
        New-ADDFormControl -Parent $ADDForm -Anchor None -w 80 -h 60

	$ADDResult = $ADDForm.ShowDialog()

    If( $ADDResult -eq [System.Windows.Forms.DialogResult]::OK ) {
		# Toggle disabled.
        If( $ShowDisabled ) {
		    Applet-Users -AdminCredential $AdminCredential -ShowDisabled $false
        } Else {
            Applet-Users -AdminCredential $AdminCredential -ShowDisabled $true
        }
	} ElseIf( $ADDResult -eq [System.Windows.Forms.DialogResult]::Retry ) {
		# Change password.
		$SelectedUser = $ADDUsers[$ADDList.SelectedIndex]
        Input-Changepass -UserName $SelectedUser.SamAccountName -AdminCredential $AdminCredential
        Applet-Users -AdminCredential $AdminCredential -ShowDisabled $ShowDisabled
	} ElseIf( $ADDResult -eq [System.Windows.Forms.DialogResult]::Ignore ) {
		# Unlock account.
		$SelectedUser = $ADDUsers[$ADDList.SelectedIndex]
        $SelectedDN = $SelectedUser.DistinguishedName
		Invoke-OnDC $AdminCredential { Unlock-ADAccount -Identity $Using:SelectedDN }
		Out-Dialog "Unlocked account for $($SelectedUser.Name)" $DInfo
        Applet-Users -AdminCredential $AdminCredential -ShowDisabled $ShowDisabled
	} Else {
        Applet-Choose -AdminCredential $AdminCredential
    }
}

Function Applet-NewUser {
    Param(
        [System.Management.Automation.PSCredential] $AdminCredential
    )
    
	$ADDForm = New-ADDForm -Title 'New AD User'

    $Error.Clear()

    # Title Groups List
    $ADDTitleGroupsD = Get-RemoteADObject -OU 'OU=Titles,OU=Distribution,OU=Groups,OU=Albany,DC=domain,DC=local' `
        -ObjectType 'Group' -Filter '*' -AdminCredential $AdminCredential
    If( $ADDTitleGroupsD -eq $null ) {
        Out-Dialog -Message 'Titles distribution list is empty.' -DialogType 'Error'
        #Return
    }
    $ADDTitleGroupMiddle = $ADDTitleGroupsD.Length
    $ADDTitleGroupsS = Get-RemoteADObject -OU 'OU=Titles,OU=Security,OU=Groups,OU=Albany,DC=domain,DC=local' `
        -ObjectType 'Group' -Filter '*' -AdminCredential $AdminCredential
    If( $ADDTitleGroupsS -eq $null ) {
        Out-Dialog -Message 'Titles security list is empty.' -DialogType 'Error'
        #Return
    }
    $ADDTitleGroups = $ADDTitleGroupsD + $ADDTitleGroupsS
	$ADDTitleList = Format-ListBox -ForceEnabled $true -DropDown $true -ObjectList $ADDTitleGroups
    New-ADDFormControl -Parent $ADDForm -LabelText 'Title Group' -Control $ADDTitleList

    # Department Groups List
    $ADDDeptGroups = Get-RemoteADObject -OU 'OU=Departments,OU=Security,OU=Groups,OU=Albany,DC=domain,DC=local' `
        -ObjectType 'Group' -Filter '*' -AdminCredential $AdminCredential
    If( $ADDDeptGroups -eq $null ) {
        Out-Dialog -Message 'Departments security list is empty.' -DialogType 'Error'
        #Return
    }
	$ADDDeptList = Format-ListBox -ForceEnabled $true -DropDown $true -ObjectList $ADDDeptGroups
    New-ADDFormControl -Parent $ADDForm -LabelText 'Department Group' -Control $ADDDeptList

    $ADDAccountantGroup = Format-Checkbox 'Accountants Group'
    New-ADDFormControl -Parent $ADDForm -Control $ADDAccountantGroup
    
    $ADDDuoGroup = Format-Checkbox 'Citrix Duo Group'
    New-ADDFormControl -Parent $ADDForm -Control $ADDDuoGroup
    
    $ADDNativePrinterGroup = Format-Checkbox 'Citrix Native Printer Group'
    New-ADDFormControl -Parent $ADDForm -Control $ADDNativePrinterGroup
    
    $ADDWomenGroup = Format-Checkbox 'Women Group'
    New-ADDFormControl -Parent $ADDForm -Control $ADDWomenGroup

	$ADDResult = $ADDForm.ShowDialog()

}

Function Applet-Computers {
    Param(
        [System.Management.Automation.PSCredential] $AdminCredential,
        [bool] $ShowDisabled
    )
    
    $ComputersFilter = '*'
    If( -not $ShowDisabled ) {
        $ComputersFilter = {(Enabled -eq $true)}
    }

	# Build the form.

	$ADDForm = New-ADDForm -Title 'AD Computers'
    
    # Grab the list of computers from the DC.
    $Error.Clear()
    $ADDComputers = Get-RemoteADObject -OU 'OU=Computers,OU=Albany,DC=domain,DC=local' `
        -ObjectType 'Computer' -Filter $ComputersFilter -AdminCredential $AdminCredential
    $ADDList = Format-ListBox -ObjectList $ADDComputers
    New-ADDFormControl -Parent $ADDForm -Control $ADDList
    If( $ADDComputers -eq $null -or $ADDList -eq $null -or $ADDComputers.Length -eq 0 ) {
        Out-Dialog -Message $Error -DialogType 'Error'
        Return
    }

    Format-Button -DialogResult OK -Label 'Hide Disabled' -LabelFalse 'Show Disabled' -LabelTest $ShowDisabled | `
        New-ADDFormControl -Parent $ADDForm -Anchor None -w 80 -h 60
    Format-Button -DialogResult Ignore -Label 'Bitlocker Key' | `
        New-ADDFormControl -Parent $ADDForm -Anchor None -w 80 -h 60

	$ADDResult = $ADDForm.ShowDialog()

    If( $ADDResult -eq [System.Windows.Forms.DialogResult]::OK ) {
		# Toggle disabled.
        If( $ShowDisabled ) {
		    Applet-Computers -AdminCredential $AdminCredential -ShowDisabled $false
        } Else {
            Applet-Computers -AdminCredential $AdminCredential -ShowDisabled $true
        }
	} ElseIf( $ADDResult -eq [System.Windows.Forms.DialogResult]::Ignore ) {
		# Bitlocker key.
        $SelectedDN = $ADDComputers[$ADDList.SelectedIndex].DistinguishedName
        $BitLockerObjects = Invoke-OnDC $AdminCredential { Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} `
            -SearchBase $Using:SelectedDN -Properties 'msFVE-RecoveryPassword' }
        $BitLockerObjects | fl 'msFVE-RecoveryPassword' | Out-String | Out-Dialog -DialogType $DInfo

        Applet-Computers -AdminCredential $AdminCredential -ShowDisabled $ShowDisabled
	} Else {
        Applet-Choose -AdminCredential $AdminCredential
    }
    
}

Function Applet-Choose {
    Param(
        [System.Management.Automation.PSCredential] $AdminCredential
    )

	# Build the form.

	$ADDForm = New-ADDForm -Height 170
    $null = Format-Button -DialogResult Retry -Label 'Users' | `
        New-ADDFormControl -Parent $ADDForm -Anchor None -w 80 -h 60
    $null = Format-Button -DialogResult Ignore -Label 'Computers' | `
        New-ADDFormControl -Parent $ADDForm -Anchor None -w 80 -h 60
    $null = Format-Button -DialogResult No -Label 'New User' | `
        New-ADDFormControl -Parent $ADDForm -Anchor None -w 80 -h 60
    $ADDResult = $ADDForm.ShowDialog()

	If( $ADDResult -eq [System.Windows.Forms.DialogResult]::Retry ) {
		# Users.
        $null = Applet-Users -AdminCredential $AdminCredential
	} ElseIf( $ADDResult -eq [System.Windows.Forms.DialogResult]::No ) {
		# New User.
        $null = Applet-NewUser -AdminCredential $AdminCredential
	} ElseIf( $ADDResult -eq [System.Windows.Forms.DialogResult]::Ignore ) {
		# Computers.
        $null = Applet-Computers -AdminCredential $AdminCredential -ShowDisabled $false
    }
}

$AdminCredential = Get-AdminCredential -AdminName 'domain\administrator'

Applet-Choose -AdminCredential $AdminCredential
