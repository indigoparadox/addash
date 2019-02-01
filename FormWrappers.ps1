﻿#
# FormWrappers.ps1
# A simple wrapper around Windows Forms with some common utilities thrown in.
#

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Set-Variable DInfo -Option Constant -Value 'Information'
Set-Variable DError -Option Constant -Value 'Error'
Set-Variable AddWindowWidth -Option Constant -Value 290
Set-Variable AddWindowHeight -Option Constant -Value 320
Set-Variable AddWindowMargin -Option Constant -Value 10
Set-Variable ADDControlDefaultSingleLineHeight -Option Constant -Value 15

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
        [string] $Title = 'ADDash',
        [int] $Columns = 1,
        [int] $Rows = 1
    )
    
    $ADDForm = New-Object System.Windows.Forms.Form

    $ADDPanel = New-ADDFormPanel -Name 'MainLayout' -Columns $Columns -Rows $Rows
    $null = $ADDForm.Controls.Add( $ADDPanel )

	$ADDForm.Size = New-Object System.Drawing.Size( $Width, $Height )
    $ADDForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    #$ADDForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $ADDForm.Text = $Title
    $ADDForm.Icon = [Drawing.Icon]::ExtractAssociatedIcon( (Get-Command mmc).Path )
    
    Return $ADDForm
}

Function New-ADDFormPanel {
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string] $Name,
        [int] $Columns = 1,
        [int] $Rows = 1,
        [int] $Division = 100
    )

    $ADDPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $ADDPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ADDPanel.Name = $Name
    $ADDPanel.ColumnCount = $Columns
    $ADDPanel.RowCount = $Rows

    $ADDColumn = New-Object System.Windows.Forms.ColumnStyle( [System.Windows.Forms.SizeType]::Percent, $Division )

    $null = $ADDPanel.ColumnStyles.Add( $ADDColumn )

    Return $ADDPanel
}

Function New-ADDFormControl {
    Param(
        [Parameter(Mandatory, Position=0)]
        [ValidateNotNullOrEmpty()]
        [System.Windows.Forms.Form] $Parent,
        [Parameter(Position=1)]
        [string] $LabelText = $null,
        [Parameter(ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [System.Windows.Forms.Control] $Control,
        [ValidateNotNullOrEmpty()]
        [string] $Layout,
        [bool] $Horizontal = $false,
        [int] $RowDivision = 100,
        [int] $ColumnDivision = 100
    )

    # Fetch the layout from the parent form.
    $ADDLayouts = $Parent.Controls.Find( $Layout, $true )
    If( $ADDLayouts -eq $null -or $ADDLayouts.Length -eq 0 ) {
        Out-Dialog -Message ('Unable to find control root: ' + $Layout) -DialogType 'Error'
        Return
    }
    $ADDLayout = $ADDLayouts.Get( 0 )
    
    # Add a label if text was specified.
    If( $LabelText -ne $null -and $LabelText -ne '' ) {
        #Write-Host ('Using layout label: ' + $LabelText)
        $ListLabel = New-Object System.Windows.Forms.Label
        $ListLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
        $ListLabel.Text = $LabelText

        If( $Horizontal ) {
            $LabelColumn = New-Object System.Windows.Forms.ColumnStyle( [System.Windows.Forms.SizeType]::Percent, $ColumnDivision )
            $null = $ADDLayout.ColumnStyles.Add( $LabelColumn )
        }
        $ListLabelRow = New-Object System.Windows.Forms.RowStyle( [System.Windows.Forms.SizeType]::Percent, $RowDivision )
        $ADDLayout.RowStyles.Add( $ListLabelRow )
        $ADDLayout.Controls.Add( $ListLabelRow )
    }
    
    If( $Horizontal ) {
        $ControlColumn = New-Object System.Windows.Forms.ColumnStyle( [System.Windows.Forms.SizeType]::Percent, $ColumnDivision )
        $null = $ADDLayout.ColumnStyles.Add( $ControlColumn )
    }
    $ControlRow = New-Object System.Windows.Forms.RowStyle( [System.Windows.Forms.SizeType]::Percent, $RowDivision )
    $ADDLayout.RowStyles.Add( $ControlRow )
    $Control.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ADDLayout.Controls.Add( $Control )
    
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
        [bool] $LabelTest = $true,
        [string] $Label,
        [string] $LabelFalse
    )

    $ADDButton = New-Object System.Windows.Forms.Button
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
        [string] $Label
    )
    $ADDCheck = New-Object System.Windows.Forms.Checkbox
    $ADDCheck.Text = $Label

    Return $ADDCheck
}

Function Format-ListBox {
    Param(
        [Parameter(ValueFromPipeline=$true)]
        [Object[]] $ObjectList,
        [bool] $DropDown = $false,
        [bool] $ForceEnabled = $false
    )

    If( $DropDown -eq $false ) {
	    $ListBox = New-Object System.Windows.Forms.ListBox

        $ListBox.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
        $ListBox.Add_DrawItem( $UserList_DrawItem )
    } Else {
	    $ListBox = New-Object System.Windows.Forms.ComboBox

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
