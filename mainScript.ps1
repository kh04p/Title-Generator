#PACKAGES FOR GUI CREATION
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#ENABLE FLAT UI STYLE
[System.Windows.Forms.Application]::EnableVisualStyles()

#CREATE BASE FORM
$form = New-Object System.Windows.Forms.Form
$form.Text = 'CARTS Title Generator'
$form.Size = New-Object System.Drawing.Size(900,350)
$form.FormBorderStyle = 'Fixed3D'
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'
$form.Add_Load({
	$form.Activate()
})

#DEVICE ID ENTRY
#1. Label for device ID
$labelID = New-Object System.Windows.Forms.Label
$labelID.Location = New-Object System.Drawing.Point(10,20)
$labelID.Size = New-Object System.Drawing.Size(295,30)
$labelID.Text = 'Enter a device ID (MBR01, EDPP03, etc.):'
$labelID.Font = ‘Segoe UI,12’
$form.Controls.Add($labelID)

#2. Text entry field for device ID
$boxID = New-Object System.Windows.Forms.TextBox
$boxID.Location = New-Object System.Drawing.Point(305,20)
$boxID.Size = New-Object System.Drawing.Size(100,25)
$boxID.Multiline = $true
$boxID.AutoSize = $true
$boxID.Font = ‘Segoe UI,12’
$boxID.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($boxID)

#PLATFORM SELECTOR
#1. Label for selector
$labelPlatform = New-Object System.Windows.Forms.Label
$labelPlatform.Location = New-Object System.Drawing.Point(415,20)
$labelPlatform.Size = New-Object System.Drawing.Size(150,30)
$labelPlatform.Text = 'Choose a platform:'
$labelPlatform.Font = ‘Segoe UI,12’
$form.Controls.Add($labelPlatform)

#2. Drop-down menu for selector
$listPlatform = New-Object System.Windows.Forms.ComboBox
$listPlatform.Location = New-Object System.Drawing.Point(565,20)
$listPlatform.Size = New-Object System.Drawing.Size(300,20)
$listPlatform.AutoSize = $true

#3. Add default platforms to drop down menu
$arraySW = Get-Content -Path @("$PSScriptRoot\platforms.txt") | Sort-Object | ForEach-Object {[void] $listPlatform.Items.Add($_)}
$listPlatform.SelectedIndex = 0
$listPlatform.Font = ‘Segoe UI,12’
$listPlatform.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($listPlatform)


#4. Import list of HW and SW issues
$issuesSW = get-content -raw '.\servicesSW.txt' | ConvertFrom-StringData
$issuesHW = get-content -raw '.\servicesHW.txt' | ConvertFrom-StringData

$arraySW = @()
$arraySW_apps = @()
$arrayHW = @()

#Split list of issues into array of individual issues
foreach ($item in $issuesSW.Keys) {
	$itemValue = $issuesSW.Item($item)
	$tempArray = $itemValue.Split(",")
	$serviceObj = [pscustomobject]@{
		Platform = $item
		Issues = $tempArray
	}
	$appObj = [pscustomobject]@{
		Platform = $item
		Issues = ""
	}
	$arraySW += $serviceObj
	$arraySW_apps += $appObj
}

#Pull names of software folders from packagesource for PC/Thin Client software issues
$packagesourcePath = "\\REDACTED\PRD_SMSPackageSource"
$hacPath = "\\REDACTED\PRD_SMSPackageSource\Hearing_Aid"
$packagesourceArray = Get-ChildItem $packagesourcePath -Directory | Where Name -inotmatch 'test' | Where Name -inotmatch 'temp' | Select Name
$hacArray = Get-ChildItem $hacPath -Directory | Where Name -inotmatch 'archive' | Where Name -inotmatch 'test' | Where Name -inotmatch 'temp' | Select Name
[array]$phmArray = Get-Content -Path "$PSScriptRoot\phmSW.txt"
[array]$optArray = Get-Content -Path "$PSScriptRoot\optSW.txt"
[array]$additionalSWArray = Get-Content -Path "$PSScriptRoot\additionalSW.txt"

foreach ($object in $arraySW_apps) {
	if ($object.Platform -eq "Desktop" -or $object.Platform -eq "Hearing Aid PC/Thin Client" -or $object.Platform -eq "Laptop" -or $object.Platform -eq "Optical PC/Thin Client" -or $object.Platform -eq "Pharmacy Thin Client" -or $object.Platform -eq "Pharmacy Workflow (PWF) PC" -or $object.Platform -eq "Thin Client") {
		[array]$tempArraySW = @()
		:outer
        foreach ($program in $packagesourceArray) {
			if ($object.Platform -eq "Hearing Aid PC/Thin Client") {
				foreach ($hacProgram in $hacArray) {
					$tempArraySW += $hacProgram.Name					
				}
				break outer
			}

            if ($object.Platform -eq "Pharmacy Thin Client" -or $object.Platform -eq "Pharmacy Workflow (PWF) PC") {
				foreach ($phmProgram in $phmArray) {
					$tempArraySW += $phmProgram
				}
				break outer
			}

            if ($object.Platform -eq "Optical PC/Thin Client") {
				foreach ($optProgram in $optArray) {
					$tempArraySW += $optProgram
				}
				break outer
			}

			$tempArraySW += $program.Name
		}

		foreach ($sw in $additionalSWArray) {
			$tempArraySW += $sw
		}
		$object.Issues = $tempArraySW
	}
}

#Split list of issues into array of individual issues
foreach ($item in $issuesHW.Keys) {
	$itemValue = $issuesHW.Item($item)
	$tempArray = $itemValue.Split(",")
	$serviceObj = [pscustomobject]@{
		Platform = $item
		Issues = $tempArray
	}
	$arrayHW += $serviceObj
}

#5. Filter Point of Failure list based on Platform Selector
$listPlatform_SelectedIndexChanged = {
	$listFailure.Items.Clear()
	$listFailure.Text = $null

	foreach ($service in $arrayHW) {
		if ($listPlatform.Text -eq $service.Platform -and $listHWSW.Text -eq "Hardware") {
			$tempObj = $arrayHW | Where-Object {$_.Platform -eq $listPlatform.Text}
			$tempObj.Issues | Sort-Object | ForEach-Object {[void] $listFailure.Items.Add($_)}
			$listFailure.SelectedIndex = 0
		}
	}

	foreach ($service in $arraySW) {
		if ($listPlatform.Text -eq $service.Platform -and $listHWSW.Text -eq "Software - Common") {
			$tempObj = $arraySW | Where-Object {$_.Platform -eq $listPlatform.Text}
			$tempObj.Issues | Sort-Object | ForEach-Object {[void] $listFailure.Items.Add($_)}
			$listFailure.SelectedIndex = 0
		}
	}

	foreach ($service in $arraySW_apps) {
		if ($listPlatform.Text -eq $service.Platform -and $listHWSW.Text -eq "Software - Applications") {
			$tempObj = $arraySW_apps | Where-Object {$_.Platform -eq $listPlatform.Text}
			$tempObj.Issues | Sort-Object | ForEach-Object {[void] $listFailure.Items.Add($_)}
			$listFailure.SelectedIndex = 0
		}
	}
}
$listPlatform.add_SelectedIndexChanged($listPlatform_SelectedIndexChanged)



#HARDWARE/SOFTWARE SELECTOR
#1. Label for selector
$labelHWSW = New-Object System.Windows.Forms.Label
$labelHWSW.Location = New-Object System.Drawing.Point(10,60)
$labelHWSW.Size = New-Object System.Drawing.Size(150,30)
$labelHWSW.Text = 'Hardware/Software:'
$labelHWSW.Font = ‘Segoe UI,12’
$form.Controls.Add($labelHWSW)

#2. Drop-down menu for selector
$listHWSW = New-Object System.Windows.Forms.ComboBox
$listHWSW.Location = New-Object System.Drawing.Point(160,60)
$listHWSW.Size = New-Object System.Drawing.Size(170,20)
$listHWSW.AutoSize = $true
$listHWSW.Font = ‘Segoe UI,12’
@("Hardware","Software - Common","Software - Applications") | ForEach-Object {[void] $listHWSW.Items.Add($_)}
$listHWSW.SelectedIndex = 0
$listHWSW.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($listHWSW)
$listHWSW.add_SelectedIndexChanged($listPlatform_SelectedIndexChanged)

#POINT OF FAILURE SELECTOR
#1. Label for selector
$labelFailure = New-Object System.Windows.Forms.Label
$labelFailure.Location = New-Object System.Drawing.Point(340,60)
$labelFailure.Size = New-Object System.Drawing.Size(190,30)
$labelFailure.Text = 'Choose a point of failure:'
$labelFailure.Font = ‘Segoe UI,12’
$form.Controls.Add($labelFailure)

#2. Drop-down menu for selector
$listFailure = New-Object System.Windows.Forms.ComboBox
$listFailure.Location = New-Object System.Drawing.Point(530,60)
$listFailure.Size = New-Object System.Drawing.Size(335,20)
$listFailure.AutoSize = $true

#3. Add default points of failure to drop down menu
foreach ($service in $arrayHW) {
	if ($listPlatform.Text -eq $service.Platform -and $listHWSW.Text -eq "Hardware") {
		$tempObj = $arrayHW | Where-Object {$_.Platform -eq $listPlatform.Text}
		$tempObj.Issues | Sort-Object | ForEach-Object {[void] $listFailure.Items.Add($_)}
		$listFailure.SelectedIndex = 0
	}
}

foreach ($service in $arraySW) {
	if ($listPlatform.Text -eq $service.Platform -and $listHWSW.Text -eq "Software - Common") {
		$tempObj = $arraySW | Where-Object {$_.Platform -eq $listPlatform.Text}
		$tempObj.Issues | Sort-Object | ForEach-Object {[void] $listFailure.Items.Add($_)}
		$listFailure.SelectedIndex = 0
	}
}

$listFailure.Font = ‘Segoe UI,12’
$listFailure.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($listFailure)

#ERROR INPUT
#1. Label for Error text box
$labelErr = New-Object System.Windows.Forms.Label
$labelErr.Location = New-Object System.Drawing.Point(10,100)
$labelErr.Size = New-Object System.Drawing.Size(230,30)
$labelErr.Text = 'Error code or issue description:'
$labelErr.Font = ‘Segoe UI,12’
$form.Controls.Add($labelErr)

#2. Error text box
$boxErr = New-Object System.Windows.Forms.TextBox
$boxErr.Location = New-Object System.Drawing.Point(240,100)
$boxErr.Size = New-Object System.Drawing.Size(625,30)
$boxErr.AutoSize = $false
$boxErr.Multiline = $false
$boxErr.Text = "  "
$boxErr.Font = ‘Segoe UI,12’
$boxErr.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($boxErr)

#RESULT
#1. Label for result text box
$labelRes = New-Object System.Windows.Forms.Label
$labelRes.Location = New-Object System.Drawing.Point(10,220)
$labelRes.Size = New-Object System.Drawing.Size(100,20)
$labelRes.Text = 'Generated Title:'
$labelRes.AutoSize = $true
$labelRes.Font = ‘Segoe UI,12’
$form.Controls.Add($labelRes)

#2. Result text box
$boxRes = New-Object System.Windows.Forms.TextBox
$boxRes.Location = New-Object System.Drawing.Point(135,220)
$boxRes.Size = New-Object System.Drawing.Size(725,30)
$boxRes.AutoSize = $false
$boxRes.Multiline = $false
$boxRes.Font = ‘Segoe UI,11’
$boxRes.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($boxRes)

#CLIPBOARD STATUS - Indicates whether generated title has been copied to system clipboard
$labelClip = New-Object System.Windows.Forms.Label
$labelClip.Size = New-Object System.Drawing.Size(900,70)
$LabelClip.Top = ($listFailure.Height + 240)
$labelClip.Left = ($form.Width - 890)
$LabelClip.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$LabelClip.Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold+[System.Drawing.FontStyle]::Italic)
$form.Controls.Add($LabelClip)

#GENERATE BUTTON

#1. Title Generator
$generate = {
	$arrayUnfiltered = @($boxID.Text, $listPlatform.Text, $listHWSW.Text, $listFailure.Text, $boxErr.Text)
	$arrayFiltered = @()
	$result = "  "
	$emptyFlag = ""

	#a. Filter for empty fields
	foreach ($item in $arrayUnfiltered) {
		#b. If field is empty, put EMPTY! in bold and red as warning
		if ([string]::IsNullOrWhitespace($item)) {
			switch ($item) {
				$listPlatform.Text {
					$listPlatform.Text = "  EMPTY!"
					$listPlatform.ForeColor = [System.Drawing.Color]::OrangeRed
					$listPlatform.Font = [System.Drawing.Font]::new("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
					$emptyFlag = "PLATFORM field"
				}

				$listHWSW.Text {
					$listHWSW.Text = "  EMPTY!"
					$listHWSW.ForeColor = [System.Drawing.Color]::OrangeRed
					$listHWSW.Font = [System.Drawing.Font]::new("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
					$emptyFlag = "HARDWARE/SOFTWARE field"
				}
				$listFailure.Text {
					$listFailure.Text = "  EMPTY!"
					$listFailure.ForeColor = [System.Drawing.Color]::OrangeRed
					$listFailure.Font = [System.Drawing.Font]::new("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
					$emptyFlag = "ISSUE field"
				}
				$boxErr.Text {
					$boxErr.Text = "  EMPTY!"
					$boxErr.ForeColor = [System.Drawing.Color]::OrangeRed
					$boxErr.Font = [System.Drawing.Font]::new("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
					$emptyFlag = "ISSUE DESCRIPTION field"
				}
			}
			continue
		}

		#c. If field is NOT empty, trim whitespace and add to filtered array
		$item = $item.Trim()
		$arrayFiltered += $item
	}
	
	#d. Add all items in filtered array to result string for output
	foreach ($item in $arrayFiltered) {
		$result += "$item"

		#e. If item is not last in array, add slashes in between items
		if ($item -ne $arrayFiltered[-1]) {
			$result += " // "
		}
	}

	#f. Show result string in result text box
	$boxRes.Text = $result
	
	#g. Trim whitespace and send to system clipboard then display confirmation
	$result.Trim() | clip
	if ([string]::IsNullOrWhitespace($emptyFlag)) {
		$labelClip.Text = "`r`nCopied to system clipboard!"
	} else {		
		$labelClip.Text = "Copied to system clipboard!`r`nWARNING: Did you mean to leave the $emptyFlag empty?"
		$labelClip.ForeColor = [System.Drawing.Color]::OrangeRed
	}	
}

#2. Create Button
$buttonGen = New-Object System.Windows.Forms.Button
$buttonGen.Location = New-Object System.Drawing.Point(340,160)
$buttonGen.Size = New-Object System.Drawing.Size(100,30)
$buttonGen.Text = 'GENERATE'
$buttonGen.Font = ‘Segoe UI,12’
$buttonGen.ForeColor = [System.Drawing.Color]::White
$buttonGen.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonGen.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(91,181,91)
$buttonGen.BackColor = [System.Drawing.Color]::FromArgb(91,181,91)
$buttonGen.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::DarkGreen
$buttonGen.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Green
$form.Controls.Add($buttonGen)
$buttonGen.Add_Click($generate) #Adds function to generate button

#RESET BUTTON
#1. Reset all fields and their font styles
$reset = {
	$boxID.Text = ""
	$boxID.Font = ‘Segoe UI,12’
	$boxID.ForeColor = [System.Drawing.Color]::Black

	$boxErr.Text = "  "
	$boxErr.Font = ‘Segoe UI,12’
	$boxErr.ForeColor = [System.Drawing.Color]::Black

	$boxRes.Text = ""
	$boxRes.Font = 'Segoe UI,12'
	$boxRes.ForeColor = [System.Drawing.Color]::Black

	$listPlatform.SelectedIndex = 0
	$listPlatform.Font = 'Segoe UI,12'
	$listPlatform.ForeColor = [System.Drawing.Color]::Black

	$listHWSW.SelectedIndex = 0
	$listHWSW.Font = 'Segoe UI,12'
	$listHWSW.ForeColor = [System.Drawing.Color]::Black

	$listFailure.SelectedIndex = 0
	$listFailure.Font = 'Segoe UI,12'
	$listFailure.ForeColor = [System.Drawing.Color]::Black

	$labelClip.Text = ""
	$labelClip.ForeColor = [System.Drawing.Color]::Black
}

#2. Create Button
$buttonReset = New-Object System.Windows.Forms.Button
$buttonReset.Location = New-Object System.Drawing.Point(460,160)
$buttonReset.Size = New-Object System.Drawing.Size(100,30)
$buttonReset.Text = 'RESET'
$buttonReset.Font = ‘Segoe UI,12’
$buttonReset.ForeColor = [System.Drawing.Color]::White
$buttonReset.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonReset.FlatAppearance.BorderColor = [System.Drawing.Color]::Tomato
$buttonReset.BackColor = [System.Drawing.Color]::Tomato
$buttonReset.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Red
$buttonReset.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::OrangeRed
$form.Controls.Add($buttonReset)
$buttonReset.Add_Click($reset) #Adds function to reset button

# POPUP MESSAGE WHEN USER CLOSES APP
$form.Add_Closing({param($sender,$e)
    $result = [System.Windows.Forms.MessageBox]::Show(`
        "Are you sure you want to exit?", `
        "EXIT WARNING", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
		return
		exit
	}
	
	if ($result -eq [System.Windows.Forms.DialogResult]::No) {
		$e.Cancel = $true
	}
})

# WHAT HAPPENS WHEN USER PRESSES ENTER
$form.AcceptButton = $buttonGen

#SHOW FORM
$Null = $form.ShowDialog()
