#
# Script.ps1
#

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Set-Variable DInfo -Option Constant -Value 'Information'
Set-Variable DError -Option Constant -Value 'Error'

Function Out-Dialog {
    Param(
        [Parameter(Mandatory, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [Parameter(Mandatory, Position=1)]
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
        [System.Object] $sender, 
        [System.Windows.Forms.DrawItemEventArgs] $e
    )
    
    # Suppose Sender type Listbox.
    if( $sender.Items.Count -eq 0 ) {
        return
    }

    # Suppose item type String.
    $lbItem = $sender.Items[$e.Index]
    If( $sender.SelectedIndex -eq $e.Index ) {
        $color = [System.Drawing.Color]::Cyan
    } ElseIf ( $lbItem.Contains( '[Disabled]' ) ) { 
        $color = [System.Drawing.Color]::LightGray
    } ElseIf ( $lbItem.Contains( '[Locked]' ) ) { 
        $color = [System.Drawing.Color]::Red
    } Else {
        $color = [System.Drawing.Color]::White
    }

    try {
        $brush = new-object System.Drawing.SolidBrush( $color )
        $e.Graphics.FillRectangle( $brush, $e.Bounds )
    } finally {
        $brush.Dispose()
    }

    $e.Graphics.DrawString( $lbItem, $e.Font, [System.Drawing.SystemBrushes]::ControlText, (New-Object System.Drawing.PointF( $e.Bounds.X, $e.Bounds.Y )) )
}

Function Input-Changepass {
    Param(
        [Parameter(Mandatory, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$UserName
    )
	
    

}

Function Applet-Users {
    Param(
        [System.Management.Automation.PSCredential]$AdminCredential
    )

	# Build the form.

	$ADDForm = New-ADDForm -Width 290 -Height 320 -Title 'AD Users'
	
	$ADDList = New-Object System.Windows.Forms.ListBox
	$ADDList.Location = New-Object System.Drawing.Point( 10, 10 )
	$ADDList.Size = New-Object System.Drawing.Size( 260, 200 )
    $ADDList.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
    $ADDList.Add_DrawItem( $UserList_DrawItem )
	$ADDUsers = Get-ADUser -SearchBase 'OU=Users,OU=Albany,DC=domain,DC=local' `
        -Filter * -Properties DistinguishedName,Enabled,Name,SamAccountName,SID,LockedOut | 
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

	$ADDChangePassword = New-Object System.Windows.Forms.Button
	$ADDChangePassword.Location = New-Object System.Drawing.Point( 10, 220 )
	$ADDChangePassword.Size = New-Object System.Drawing.Size( 80, 60 )
	$ADDChangePassword.Text = 'Set Password'
	$ADDChangePassword.DialogResult = [System.Windows.Forms.DialogResult]::Retry
	$ADDForm.Controls.Add( $ADDChangePassword )

	$ADDUnlock = New-Object System.Windows.Forms.Button
	$ADDUnlock.Location = New-Object System.Drawing.Point( 100, 220 )
	$ADDUnlock.Size = New-Object System.Drawing.Size( 80, 60 )
	$ADDUnlock.Text = 'Unlock Account'
	$ADDUnlock.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
	$ADDForm.Controls.Add( $ADDUnlock )

	$ADDResult = $ADDForm.ShowDialog()

	If( $ADDResult -eq [System.Windows.Forms.DialogResult]::Retry ) {
		# Change password.
	} ElseIf( $ADDResult -eq [System.Windows.Forms.DialogResult]::Ignore ) {
		# Unlock account.
		$SelectedUser = $ADDUsers[$ADDList.SelectedIndex]
		Unlock-ADAccount -Credential $AdminCredential -Identity $SelectedUser.DistinguishedName
		Out-Dialog "Unlocked account for $($SelectedUser.Name)" $DInfo
        Applet-Users -AdminCredential $AdminCredential
	} Else {
        Applet-Choose -AdminCredential $AdminCredential
    }
}

Function Applet-Choose {
    Param(
        [System.Management.Automation.PSCredential]$AdminCredential
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
        Applet-ManagePC -AdminCredential $AdminCredential
    }
}

$AdminCredential = Get-AdminCredential -AdminName 'domain\administrator'

Applet-Choose -AdminCredential $AdminCredential
