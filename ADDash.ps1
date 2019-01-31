#
# Script.ps1
#

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Set-Variable DInfo -Option Constant -Value 'Information'
Set-Variable DError -Option Constant -Value 'Error'

#Set-Variable DROK -Option Constant -Value ([System.Windows.Forms.DialogResult]::OK)
#Set-Variable DRIgnore -Option Constant -Value ([System.Windows.Forms.DialogResult]::Ignore)
#Set-Variable DRAbort -Option Constant -Value ([System.Windows.Forms.DialogResult]::Abort)
#Set-Variable DRCancel -Option Constant -Value ([System.Windows.Forms.DialogResult]::Cancel)
#Set-Variable DRRetry -Option Constant -Value ([System.Windows.Forms.DialogResult]::Retry)

$AddWindowWidth = 290
$AddWindowHeight = 320

$LastSelectedIndex = 0

Function Out-Dialog {
    Param(
        [Parameter(Position=0, ValueFromPipeline=$true, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [Parameter(Position=1)]
        [string]$DialogType
    )
	
	[System.Windows.Forms.MessageBox]::Show( $Message, 'ADDash', 'OK', $DialogType )
}

Function Get-AdminCredential {
    Param(
        [string]$AdminName
    )

	$AdminCredPath = 'J:\ADDash_cred.' + $ENV:COMPUTERNAME + '.txt'
	$AdminPassword = ''
	
    If( Test-Path -Path $AdminCredPath ) {
			$AdminPassword = Get-Content $AdminCredPath | ConvertTo-SecureString
			$AdminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList  $AdminName,$AdminPassword
	} Else {
	    $AdminCredential = Get-Credential -Message 'Please enter the administrative username and password to use.' -UserName $AdminName
	    If( $AdminCredential.Password.Length -lt 1 ) {
            Out-Dialog 'No administrative credential is available.' $DError
		} Else {
		    $AdminCredential.Password | ConvertFrom-SecureString | Out-File $AdminCredPath
	    }
    }

	Return $AdminCredential
}

Function New-ADDForm {
    Param(   
        [int] $Width = $ADDWindowWidth, 
        [int] $Height = $ADDWindowHeight,
        [string] $Title
    )

    If( $Title.Length -lt 1 ) {
        $Title = 'ADDash'
    }

    $ADDForm = New-Object System.Windows.Forms.Form

    $ADDFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    #$ADDFlow.Anchor = `
    #    [System.Windows.Forms.AnchorStyles]::Left + `
    #    [System.Windows.Forms.AnchorStyles]::Right + `
    #    [System.Windows.Forms.AnchorStyles]::Bottom + `
    #    [System.Windows.Forms.AnchorStyles]::Top
    $ADDFlow.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ADDFlow.Name = 'Layout'
    $ADDForm.Controls.Add( $ADDFlow )

	$ADDForm.Size = New-Object System.Drawing.Size( $Width, $Height )
    $ADDForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $ADDForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $ADDForm.Text = $Title
    $ADDForm.Icon = [Drawing.Icon]::ExtractAssociatedIcon( (Get-Command mmc).Path )
    
    Return $ADDForm
}

Function New-ADDFormControl {
    Param(
        [Parameter(Mandatory, Position=0)]
        [ValidateNotNullOrEmpty()]
        [System.Windows.Forms.Form] $Parent,
        [Parameter(ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [System.Windows.Forms.Control] $Control,
        [string] $LabelText
    )

    # Fetch the layout from the parent form.
    $ADDLayouts = $Parent.Controls.Find( 'Layout', $true )
    If( $ADDLayouts -eq $null -or $ADDLayouts.Length -eq 0 ) {
        Out-Dialog -Message 'Unable to find control root.' -DialogType 'Error'
        Return
    }
    $ADDLayout = $ADDLayouts.Get( 0 )
    
        Out-Dialog $LabelText
    
    # Add a label if text was specified.
    If( $LabelText -ne $null ) {
        $ListLabel = New-Object System.Windows.Forms.Label
        $ListLabel.Text = $LabelText
        $ListLabel.Location = New-Object System.Drawing.Point( $x, $y )
        $ListLabel.Size = New-Object System.Drawing.Size( $w, 15 )
        #$y = $y + 15
	    #$Parent.Controls.Add( $ListLabel )
        $ADDLayout.Controls.Add( $ListLabel );
    } else {
        Out-Dialog $LabelText
    }

    $ADDLayout.Controls.Add( $Control );
    
    Return $Parent
}

$UserList_DrawItem = {
    Param(   
        [System.Object] $Sender,
        [System.Windows.Forms.DrawItemEventArgs] $Event
    )

    [System.Windows.Forms.ListBox] $SenderListBox = $Sender
    
    # Suppose Sender type Listbox.
    if( $SenderListBox.Items.Count -eq 0 ) {
        return
    }

    if( $SenderListBox.SelectedIndex -ne $LastSelectedIndex ) {
        $SenderListBox.Invalidate()
        Set-Variable -Name LastSelectedIndex -Value $SenderListBox.SelectedIndex `
            -Scope Global
    }

    # Suppose item type String.
    $ItemLabel = $SenderListBox.Items[$Event.Index]
    If( $SenderListBox.SelectedIndex -eq $Event.Index ) {
        $Color = [System.Drawing.Color]::Cyan
    } ElseIf ( $ItemLabel.Contains( '[Disabled]' ) ) { 
        $Color = [System.Drawing.Color]::LightGray
    } ElseIf ( $ItemLabel.Contains( '[Locked]' ) ) { 
        $Color = [System.Drawing.Color]::Red
    } Else {
        $Color = [System.Drawing.Color]::White
    }

    try {
        $Brush = New-Object System.Drawing.SolidBrush( $Color )
        $Event.Graphics.FillRectangle( $Brush, $Event.Bounds )
    } finally {
        $Brush.Dispose()
    }

    $Event.Graphics.DrawString(
        $ItemLabel,
        $Event.Font,
        [System.Drawing.SystemBrushes]::ControlText,
        (New-Object System.Drawing.PointF( $Event.Bounds.X, $Event.Bounds.Y ))
    )
}

Function Input-Changepass {
    Param(
        [Parameter(Mandatory, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string] $UserName,
        [System.Management.Automation.PSCredential] $AdminCredential
    )
	
    $UserCredential = Get-Credential -Message 'Please enter the new password for this user.' -UserName $UserName
	
    $InputUserName = $UserCredential.UserName
    $InputPassword = $UserCredential.Password
    $Result = Invoke-OnDC $AdminCredential { 
        $Error.Clear()
        Try {
            Set-ADAccountPassword -Reset -Identity $Using:InputUserName -NewPassword $Using:InputPassword
            Return 0
        } Catch {
            Return $Error
        }
    }
    If( $Result -eq 0 ) {
        Out-Dialog 'The password reset was successful.' $DInfo
    } Else {
        Out-Dialog $Result $DError
    }
}

Function Format-Button {
    Param(
        [Parameter(Mandatory, Position=0)]
        [ValidateNotNullOrEmpty()]
        [System.Windows.Forms.DialogResult] $DialogResult,
        [int] $x = 10,
        [int] $y = 10,
        [int] $w = 80,
        [int] $h = 60,
        [bool] $LabelTest = $true,
        [string] $Label,
        [string] $LabelFalse
    )

    $ADDButton = New-Object System.Windows.Forms.Button
	$ADDButton.Location = New-Object System.Drawing.Point( $x, $y )
	$ADDButton.Size = New-Object System.Drawing.Size( $w, $h )
    If( $LabelTest ) {
	    $ADDButton.Text = $Label
    } Else {
        $ADDButton.Text = $LabelFalse
    }
	$ADDButton.DialogResult = $DialogResult
	#$ADDForm.Controls.Add( $ADDShowDisabled )
    
    Return $ADDButton
}

Function Format-Checkbox {
    Param(
        [Parameter( Mandatory=$true, Position=0 )]
        [string] $Label,
        [int] $x = 10,
        [int] $y = 10,
        [int] $w = 190,
        [int] $h = 15
    )
    $ADDCheck = New-Object System.Windows.Forms.Checkbox
    $ADDCheck.Text = $Label
    $ADDCheck.Location = New-Object System.Drawing.Point( $x, $y )
    $ADDCheck.Size = New-Object System.Drawing.Size( $w, $h )

    Return $ADDCheck
}

Function Format-ListBox {
    Param(
        [Parameter(ValueFromPipeline=$true)]
        [Object[]] $ObjectList,
        [int] $x = 10,
        [int] $y = 10,
        [int] $w = ($ADDWindowWidth - 30),
        [int] $h = 200,
        [bool] $DropDown = $false,
        [bool] $ForceEnabled = $false
    )

    If( $DropDown -eq $false ) {
	    $ListBox = New-Object System.Windows.Forms.ListBox

	    $ListBox.Location = New-Object System.Drawing.Point( $x, $y )
	    $ListBox.Size = New-Object System.Drawing.Size( $w, $h )
        $ListBox.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
        $ListBox.Add_DrawItem( $UserList_DrawItem )
    } Else {
	    $ListBox = New-Object System.Windows.Forms.ComboBox

	    $ListBox.Location = New-Object System.Drawing.Point( $x, $y )
	    $ListBox.Size = New-Object System.Drawing.Size( $w, $h )
	    $ListBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        #$ListBox.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
        #$ListBox.Add_DrawItem( $UserList_DrawItem )
    }

    $i = 0
    # If we don't assign to $null, the list items pollute the return stream.
    $null = $ObjectList | ForEach-Object {
        $i++
        $ObjectName = $_.Name

        If( -not $ForceEnabled -and -not $_.Enabled ) {
            $ObjectName += " [Disabled]"
        } ElseIf( $_.LockedOut ) {
            $ObjectName += " [Locked]"
        }

		$ListBox.Items.Add( $ObjectName )

        Write-Progress -Activity "Building Object List" -Status “Adding $ObjectName” `
            -PercentComplete ($i / $ObjectList.Count * 100)
	}

    Write-Progress -Activity "Building Object List" -Completed $true
    
	#$Parent.Controls.Add( $ListBox )
    #New-ADDFormControl -Parent $Parent -Control $ListBox

    Return $ListBox
}

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
        New-ADDFormControl -Parent $ADDForm

    Format-Button -DialogResult OK -Label 'Hide Disabled' -LabelFalse 'Show Disabled' -LabelTest $ShowDisabled | `
        New-ADDFormControl -Parent $ADDForm

    Format-Button -DialogResult Retry -Label 'Set Password' | `
        New-ADDFormControl -Parent $ADDForm

    Format-Button -DialogResult Ignore -Label 'Unlock Account' | `
        New-ADDFormControl -Parent $ADDForm

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
	Format-ListBox -ForceEnabled $true -DropDown $true -ObjectList $ADDTitleGroups | `
        New-ADDFormControl -Parent $ADDForm -LabelText 'Title Group'

    # Department Groups List
    $ADDDeptGroups = Get-RemoteADObject -OU 'OU=Departments,OU=Security,OU=Groups,OU=Albany,DC=domain,DC=local' `
        -ObjectType 'Group' -Filter '*' -AdminCredential $AdminCredential
    If( $ADDDeptGroups -eq $null ) {
        Out-Dialog -Message 'Departments security list is empty.' -DialogType 'Error'
        #Return
    }
	Format-ListBox -ForceEnabled $true -DropDown $true -ObjectList $ADDDeptGroups  | `
        New-ADDFormControl -Parent $ADDForm -LabelText 'Department Group'

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
        New-ADDFormControl -Parent $ADDForm
    Format-Button -DialogResult Ignore -Label 'Bitlocker Key' | New-ADDFormControl -Parent $ADDForm

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
    Format-Button -DialogResult Retry -Label 'Users' | New-ADDFormControl -Parent $ADDForm
    Format-Button -DialogResult Ignore -Label 'Computers' | New-ADDFormControl -Parent $ADDForm
    Format-Button -DialogResult No -Label 'New User' | New-ADDFormControl -Parent $ADDForm
    $ADDResult = $ADDForm.ShowDialog()

	If( $ADDResult -eq [System.Windows.Forms.DialogResult]::Retry ) {
		# Users.
        Applet-Users -AdminCredential $AdminCredential
	} ElseIf( $ADDResult -eq [System.Windows.Forms.DialogResult]::No ) {
		# New User.
        Applet-NewUser -AdminCredential $AdminCredential
	} ElseIf( $ADDResult -eq [System.Windows.Forms.DialogResult]::Ignore ) {
		# Computers.
        Applet-Computers -AdminCredential $AdminCredential -ShowDisabled $false
    }
}

$AdminCredential = Get-AdminCredential -AdminName 'domain\administrator'

Applet-Choose -AdminCredential $AdminCredential
