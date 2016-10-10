#
# Script.ps1
#

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Function Out-ErrorMessage {
    Param(
        [Parameter(Mandatory, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$Message
    )
	
	[System.Windows.Forms.MessageBox]::Show( $Message, 'ADDash', 'OK', 'Error' )
}

Function Out-InfoMessage {
    Param(
        [Parameter(Mandatory, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$Message
    )
	
	[System.Windows.Forms.MessageBox]::Show( $Message, 'ADDash', 'OK', 'Information' )
}

Function Get-AdminCredential {
    Param(
        [string]$AdminName
    )

	$AdminCredPath = 'J:\ADDash_cred.txt'
	$AdminPassword = ''
	
	$AdminCredential = Get-Credential -Message 'Please enter the administrative username and password to use.' -UserName $AdminName
	If( $AdminCredential.Password.Length -lt 1 ) {
		If( Test-Path -Path $AdminCredPath ) {
			$AdminPassword = Get-Content $AdminCredPath | ConvertTo-SecureString
			$AdminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList  $AdminName,$AdminPassword
		} Else {
			Out-ErrorMessage 'No administrative credential is available.'
		}
	} Else {
		$AdminCredential.Password | ConvertFrom-SecureString | Out-File $AdminCredPath
	}

	Return $AdminCredential
}

Function Applet-Changepass {
    Param(
        [System.Management.Automation.PSCredential]$Credential
    )

	# Build the form.

	$ADDForm = New-Object System.Windows.Forms.Form
	$ADDForm.Size = New-Object System.Drawing.Size( 300, 400 )
	
	$ADDList = New-Object System.Windows.Forms.ListBox
	$ADDList.Location = New-Object System.Drawing.Point( 10, 10 )
	$ADDList.Size = New-Object System.Drawing.Size( 260, 200 )
	$ADDUsers = Get-ADUser -SearchBase 'OU=Users,OU=Albany,DC=domain,DC=local' -Filter * | Sort-Object -Property Name
	$ADDUsers | ForEach-Object {
		$ADDList.Items.Add( $_.Name )
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
		Unlock-ADAccount -Credential $Credential -Identity $SelectedUser.DistinguishedName
		Out-InfoMessage( "Unlocked account for " + $SelectedUser.Name )
	}
}

$AdminCredential = Get-AdminCredential -AdminName 'domain\administrator'

Applet-Changepass -Credential $AdminCredential