#
# Script.ps1
#

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Set-Variable DInfo -Option Constant -Value 'Information'
Set-Variable DError -Option Constant -Value 'Error'

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

	$AdminCredPath = 'J:\ADDash_cred.txt'
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
        [int] $Width, 
        [int] $Height,
        [string] $Title
    )

    If( $Title.Length -lt 1 ) {
        $Title = 'ADDash'
    }

    $ADDForm = New-Object System.Windows.Forms.Form

	$ADDForm.Size = New-Object System.Drawing.Size( $Width, $Height )
    $ADDForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $ADDForm.FormBorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $ADDForm.Text = $Title
    $ADDForm.Icon = [Drawing.Icon]::ExtractAssociatedIcon( (Get-Command mmc).Path )
    
    Return $ADDForm
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
	
    $UserCredential = Get-Credential -Message 'Please enter the administrative username and password to use.' -UserName $UserName
	
    $Error.Clear()
    Try {
        Set-ADAccountPassword -Reset -Identity $UserName -NewPassword $UserCredential.Password -Credential $AdminCredential
        Out-Dialog 'The password reset was successful.' $DInfo
    } Catch {
        Out-Dialog $Error $DError
    }

}

Function Applet-Users {
    Param(
        [System.Management.Automation.PSCredential] $AdminCredential,
        [bool] $ShowDisabled
    )
    
    $UsersFilter = '*'
    If( -not $ShowDisabled ) {
        $UsersFilter = {(Enabled -ne $ShowDisabled)}
    }

	# Build the form.

	$ADDForm = New-ADDForm -Width 290 -Height 320 -Title 'AD Users'
	
	$ADDList = New-Object System.Windows.Forms.ListBox
	$ADDList.Location = New-Object System.Drawing.Point( 10, 10 )
	$ADDList.Size = New-Object System.Drawing.Size( 260, 200 )
    $ADDList.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
    $ADDList.Add_DrawItem( $UserList_DrawItem )
	$ADDUsers = Get-ADUser -SearchBase 'OU=Users,OU=Albany,DC=domain,DC=local' `
        -Filter $UsersFilter -Properties DistinguishedName,Enabled,Name,SamAccountName,SID,LockedOut | 
        Sort-Object -Property Name
	$ADDUsers | ForEach-Object {
        $UserName = $_.Name

        If( -not $_.Enabled ) {
            $UserName += " [Disabled]"
        } ElseIf( $_.LockedOut ) {
            $UserName += " [Locked]"
        }

		$ADDList.Items.Add( $UserName )
	}
	$ADDForm.Controls.Add( $ADDList )

    $ADDShowDisabled = New-Object System.Windows.Forms.Button
	$ADDShowDisabled.Location = New-Object System.Drawing.Point( 10, 220 )
	$ADDShowDisabled.Size = New-Object System.Drawing.Size( 80, 60 )
    If( $ShowDisabled ) {
	    $ADDShowDisabled.Text = 'Hide Disabled'
    } Else {
        $ADDShowDisabled.Text = 'Show Disabled'
    }
	$ADDShowDisabled.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$ADDForm.Controls.Add( $ADDShowDisabled )

	$ADDChangePassword = New-Object System.Windows.Forms.Button
	$ADDChangePassword.Location = New-Object System.Drawing.Point( 100, 220 )
	$ADDChangePassword.Size = New-Object System.Drawing.Size( 80, 60 )
	$ADDChangePassword.Text = 'Set Password'
	$ADDChangePassword.DialogResult = [System.Windows.Forms.DialogResult]::Retry
	$ADDForm.Controls.Add( $ADDChangePassword )

	$ADDUnlock = New-Object System.Windows.Forms.Button
	$ADDUnlock.Location = New-Object System.Drawing.Point( 190, 220 )
	$ADDUnlock.Size = New-Object System.Drawing.Size( 80, 60 )
	$ADDUnlock.Text = 'Unlock Account'
	$ADDUnlock.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
	$ADDForm.Controls.Add( $ADDUnlock )

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
		Unlock-ADAccount -Credential $AdminCredential -Identity $SelectedUser.DistinguishedName
		Out-Dialog "Unlocked account for $($SelectedUser.Name)" $DInfo
        Applet-Users -AdminCredential $AdminCredential -ShowDisabled $ShowDisabled
	} Else {
        Applet-Choose -AdminCredential $AdminCredential
    }
}

Function Applet-Computers {
    Param(
        [System.Management.Automation.PSCredential] $AdminCredential,
        [bool] $ShowDisabled
    )
    
    $ComputersFilter = '*'
    If( -not $ShowDisabled ) {
        $ComputersFilter = {(Enabled -ne $ShowDisabled)}
    }

	# Build the form.

	$ADDForm = New-ADDForm -Width 290 -Height 320 -Title 'AD Computers'
    
    $ADDList = New-Object System.Windows.Forms.ListBox
	$ADDList.Location = New-Object System.Drawing.Point( 10, 10 )
	$ADDList.Size = New-Object System.Drawing.Size( 260, 200 )
    $ADDList.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
    $ADDList.Add_DrawItem( $UserList_DrawItem )
	$ADDComputers = Get-ADComputer -SearchBase 'OU=Computers,OU=Albany,DC=domain,DC=local' `
        -Filter $Computersfilter -Properties DistinguishedName,Enabled,Name,SamAccountName,SID,LockedOut | 
        Sort-Object -Property Name
	$ADDComputers | ForEach-Object {
        $ComputerName = $_.Name

        If( -not $_.Enabled ) {
            $ComputerName += " [Disabled]"
        }

    	$ADDList.Items.Add( $ComputerName )
	}
	$ADDForm.Controls.Add( $ADDList )

    $ADDShowDisabled = New-Object System.Windows.Forms.Button
	$ADDShowDisabled.Location = New-Object System.Drawing.Point( 10, 220 )
	$ADDShowDisabled.Size = New-Object System.Drawing.Size( 80, 60 )
    If( $ShowDisabled ) {
	    $ADDShowDisabled.Text = 'Hide Disabled'
    } Else {
        $ADDShowDisabled.Text = 'Show Disabled'
    }
	$ADDShowDisabled.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$ADDForm.Controls.Add( $ADDShowDisabled )

	$ADDBitlocker = New-Object System.Windows.Forms.Button
	$ADDBitlocker.Location = New-Object System.Drawing.Point( 100, 220 )
	$ADDBitlocker.Size = New-Object System.Drawing.Size( 80, 60 )
	$ADDBitlocker.Text = 'Bitlocker Key'
	$ADDBitlocker.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
	$ADDForm.Controls.Add( $ADDBitlocker )

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
        $BitLockerObjects = Get-ADObject -Credential $AdminCredential -Filter {objectclass -eq 'msFVE-RecoveryInformation'} `
            -SearchBase $ADDComputers[$ADDList.SelectedIndex].DistinguishedName -Properties 'msFVE-RecoveryPassword'
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

	$ADDForm = New-ADDForm -Width 290 -Height 110
    
	$ADDUsers = New-Object System.Windows.Forms.Button
	$ADDUsers.Location = New-Object System.Drawing.Point( 10, 10 )
	$ADDUsers.Size = New-Object System.Drawing.Size( 80, 60 )
	$ADDUsers.Text = 'Users'
	$ADDUsers.DialogResult = [System.Windows.Forms.DialogResult]::Retry
	$ADDForm.Controls.Add( $ADDUsers )

	$ADDComputers = New-Object System.Windows.Forms.Button
	$ADDComputers.Location = New-Object System.Drawing.Point( 100, 10 )
	$ADDComputers.Size = New-Object System.Drawing.Size( 80, 60 )
	$ADDComputers.Text = 'Computers'
	$ADDComputers.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
	$ADDForm.Controls.Add( $ADDComputers )

    $ADDResult = $ADDForm.ShowDialog()

	If( $ADDResult -eq [System.Windows.Forms.DialogResult]::Retry ) {
		# Users.
        Applet-Users -AdminCredential $AdminCredential
	} ElseIf( $ADDResult -eq [System.Windows.Forms.DialogResult]::Ignore ) {
		# Computers.
        Applet-Computers -AdminCredential $AdminCredential -ShowDisabled $false
    }
}

$AdminCredential = Get-AdminCredential -AdminName 'domain\administrator'

Applet-Choose -AdminCredential $AdminCredential
