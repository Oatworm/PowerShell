<#

.SYNOPSIS
    Generates a GUI for PowerShell scripts based on the parameters it accepts, then executes the script.
    
.DESCRIPTION
    GUIfy.ps1 uses Get-Help to discover the parameters used by a script designed to accept parameters, then generates a GUI designed to accept values for each parameter.

.EXAMPLE
    GUIfy.ps1

    Creates a dialog box that asks for a script to GUIfy. Once a script is selected, GUIfy then builds a GUI based on the parameters accepted by the selected script.

.EXAMPLE
    GUIfy.ps1 -ScriptPath "C:\Scripts\script.ps1"
    
    Builds a GUI based on the parameters accepted by scrip1.ps1.
   
.PARAMETER ScriptPath
    Path of the script to build a GUI around.

.NOTES
    Author: David Colborne (david@colbornemmx.com)
    Last Modified: September 26, 2018

#>

Param(
    [string]$ScriptPath
)

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Global variables go here
$ypad = 25

# Functions go next
# Event handlers first...
function BrowseButtonOnClick {
[cmdletbinding()]
    Param (
        [System.Windows.Forms.TextBox]$TextBox
    )

    Process {
        $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $FileDialog.DefaultExt = "ps1"
        $FileDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1"
        $FileDialog.ShowDialog()
        $TextBox.Text = $FileDialog.FileName
    }
}

# Forms next...
function FormSelectScript {
[cmdletbinding()]
    [OutputType([string])]
    Param ()

    Process {
        $FormSelect = New-Object system.Windows.Forms.Form
        $FormSelect.AutoScroll = $true
        $FormSelect.AutoSize = $true
        $FormSelect.StartPosition = "CenterScreen"
        $FormSelect.TopMost = $true

        $ButtonSelectOK = New-Object System.Windows.Forms.Button
        $ButtonSelectOK.Text = "OK"
        $ButtonSelectOK.DialogResult = "OK"
        $ButtonSelectCancel = New-Object System.Windows.Forms.Button
        $ButtonSelectCancel.Text = "Cancel"
        $ButtonSelectCancel.DialogResult = "Cancel"

        $folderx = 10
        $foldery = 10

        $FolderLabel = New-Object System.Windows.Forms.Label
        $FolderLabel.Text = "Please select script:"
        $FolderLabel.TextAlign = "BottomLeft"
        $FolderLabel.AutoSize = $true
        $FolderLabel.Location = New-Object System.Drawing.Point($folderx,$foldery)
        $FormSelect.Controls.Add($FolderLabel)

        $foldery+=$ypad

        $FileBox = New-Object System.Windows.Forms.TextBox
        $FileBox.Multiline = $false
        $FileBox.Location = New-Object System.Drawing.Point($folderx,$foldery)
        $FileBox.Size = New-Object System.Drawing.Size(200,20)
        $FormSelect.Controls.Add($FileBox)

        $BrowseButton = New-Object System.Windows.Forms.Button
        $BrowseButton.Text = "Browse..."
        $folderx+=$FileBox.Width
        $folderx+=10
        $foldery--
        $BrowseButton.Location = New-Object System.Drawing.Point($folderx,$foldery)
        $BrowseButton.Add_Click({BrowseButtonOnClick -TextBox $FileBox})
        $FormSelect.Controls.Add($BrowseButton)

        $folderx = 10
        $foldery+=$ypad

        $ButtonSelectOK.Location = New-Object System.Drawing.Point($folderx,$foldery)
        $FormSelect.AcceptButton = $ButtonSelectOK
        $FormSelect.Controls.Add($ButtonSelectOK)

        $folderx+=$ButtonSelectOK.Width
        $folderx+=10

        $ButtonSelectCancel.Location = New-Object System.Drawing.Point($folderx,$foldery)
        $FormSelect.CancelButton = $ButtonSelectCancel
        $FormSelect.Controls.Add($ButtonSelectCancel)

        $DialogResult = $FormSelect.ShowDialog()
    }

    End {
        If ($DialogResult -eq "OK") {
            If ([string]::IsNullOrWhiteSpace($FileBox.Text)) {
                $Suppress = [System.Windows.Forms.MessageBox]::Show("You did not select a script.")
                FormSelectScript
            } Else {
                return $FileBox.Text
            }
        } Else {
            return $DialogResult
        }
    }

}

function FormLoadScript {
[cmdletbinding()]
    Param (
        [string]$ScriptPath
    )

    Begin {
        $ScreenWidth = (Get-WmiObject -Class Win32_DesktopMonitor).ScreenWidth
        $ScreenWidth = [string]$ScreenWidth
        $ScreenWidth = [int]$ScreenWidth
    }

    Process {
        $GUIParameters = @()

        # Get parameters from the script
        $Parameters = get-help $ScriptPath
        ForEach($Parameter in $Parameters.parameters.parameter){
            $GUIParameters += $Parameter
        }

        $Form = New-Object system.Windows.Forms.Form
        $Form.AutoScroll = $true
        $Form.AutoSize = $true
        $Form.StartPosition = "CenterScreen"
        $Form.TopMost = $true

        $ButtonOK = New-Object System.Windows.Forms.Button
        $ButtonOK.Text = "Run Script"
        $ButtonOK.DialogResult = "OK"
        $ButtonCancel = New-Object System.Windows.Forms.Button
        $ButtonCancel.Text = "Cancel"
        $ButtonCancel.DialogResult = "Cancel"

        $Form.Text = Split-Path $ScriptPath -Leaf

        $i = 0
        $x = 10
        $y = 10
        $ypad = 25

        # If there's a description, show it.
        $DescriptionLabel = New-Object System.Windows.Forms.Label
        $DescriptionLabel.TextAlign = "BottomLeft"
        $DescriptionLabel.Location = New-Object System.Drawing.Point($x,$y)
        $DescriptionLabel.AutoSize = $true
        $DescriptionLabel.MaximumSize = New-Object System.Drawing.Size(($ScreenWidth-(3*$x)),0)

        If (![string]::IsNullOrWhiteSpace($Parameters.description.Text)){
            $DescriptionLabel.Text = $Parameters.Description.Text
        } Else {
            $DescriptionLabel.Text = "There is no description available for this script or cmdlet. Either the author of the script did not provide a DESCRIPTION at the start of this script or Update-Help has not been run with administrator access on this machine."
        }

        $Form.Controls.Add($DescriptionLabel)
        $y+=$DescriptionLabel.Size.Height            
        $y+=$ypad

        # Show each parameter
        ForEach($GUIParameter in $GUIParameters){

            $Label = New-Object System.Windows.Forms.Label
            If ($GUIParameter.required -eq $true) {
                $Label.Text = $GUIParameter.name + " [Required]"
            } else {
                $Label.Text = $GUIParameter.name
            }
            $Label.TextAlign = "BottomLeft"
            $Label.AutoSize = $true
            $Font = $Label.Font.Name + "," + $Label.Font.Size + ",style=Bold"
            $Label.Font = $Font
            $Label.Location = New-Object System.Drawing.Point($x,$y)
            $Form.Controls.Add($Label)

            $y+=$ypad

            $TextBox = New-Object System.Windows.Forms.TextBox
            $TextBox.Multiline = $false
            $TextBox.Name = $GUIParameter.name
            $TextBox.Text = $GUIParameter.defaultValue
            $TextBox.Location = New-Object System.Drawing.Point($x,$y)
            $Form.Controls.Add($TextBox)

            If (![string]::IsNullOrWhiteSpace($GUIParameter.description.Text)){
                $Label = New-Object System.Windows.Forms.Label
                $descx=$x+110
                $Label.Text = $GUIParameter.description.Text
                $Label.TextAlign = "BottomLeft"
                $Label.Location = New-Object System.Drawing.Point($descx,$y)
                $Label.AutoSize = $true
                $Label.MaximumSize = New-Object System.Drawing.Size(($ScreenWidth-$descx-(2*$x)),0)
                $Form.Controls.Add($Label)
            }

            If ($Label.Size.Height -gt $ypad) {
                $y+=$Label.Size.Height
            } Else {
                $y+=$ypad
            }

        }

        $y+=$ypad

        $ButtonOK.Location = New-Object System.Drawing.Point($x,$y)
        $Form.AcceptButton = $ButtonOK
        $Form.Controls.Add($ButtonOK)

        $x=$Form.Size.Width-110

        $ButtonCancel.Location = New-Object System.Drawing.Point($x,$y)
        $Form.CancelButton = $ButtonCancel
        $Form.Controls.Add($ButtonCancel)
        
        $DialogResult = $Form.ShowDialog()
    }

    End {
        If ($DialogResult -eq "OK"){
            $Args = @()
            ForEach($Control in $Form.Controls){
                If ($Control.GetType().Name -eq "TextBox"){
                    If (![string]::IsNullOrWhiteSpace($Control.Text)) {
                        $Args += ([string]::Concat("-",$Control.Name),$Control.Text)
                    }
                }
            }
            return $Args
        }
        Else {
            return $DialogResult
        }
    }
}

# And away we go...
If ([string]::IsNullOrWhiteSpace($ScriptPath)) {
    $ScriptPath = FormSelectScript
}

If ($ScriptPath -ne "Cancel") {
    $argumentList = FormLoadScript -ScriptPath $ScriptPath
    $argumentList
    If ($argumentList -ne "Cancel") {
        Invoke-Expression "& `"$ScriptPath`" $argumentList"
    }
}