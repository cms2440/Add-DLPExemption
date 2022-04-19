Remove-Variable domain,RemoveUsers,RemoveComputers -EA silentlycontinue

#since we're creating accounts across domains, we need an account that can make accounts in all domain
if ((whoami) -match ".adm" -and (whoami) -notmatch "area42\\") {
    Write-Host ("Warning: Not running as elevated AREA42 Admin.  Can only modify computers in " + (whoami).split("\")[0].toUpper() + " domain.")
    $domain = (whoami).split("\")[0].toUpper().replace("S","")
    #Read-Host
    }
    
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
if (-not $myWindowsPrincipal.IsInRole($adminRole)) {
    $scriptpath = $MyInvocation.MyCommand.Definition
    $scriptpaths = "'$scriptPath'"
    Start-Process -FilePath PowerShell.exe -Verb runAs -ArgumentList "& $scriptPaths"
    exit
    }

Import-Module activedirectory

$Comps = @()
$Users = @()

##GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Initialize our style spacings
$string = 'T'
$font = [System.Windows.Forms.Label]::DefaultFont

$size = [System.Windows.Forms.TextRenderer]::MeasureText($string, $font)

$OffsetX = 10
$OffsetY = 20
$CharHeight = $size.Height
#$CharWidth = $size.Width #I dont know why this is wrong, but it is
$CharWidth = 7

$ComboBuffer = 5 #pixels between data entry and next label
$LabelBuffer = 3 #Pixels betwen label and its data entry field
$TextBoxHeight = 20

$Label = New-Object System.Windows.Forms.Label

$Label.Location = New-Object System.Drawing.Point($OffsetX,$OffsetY) 
$Label.Size = New-Object System.Drawing.Size(280,($CharHeight * 3)) 
$Label.Text = "Paste in your list of computernames.  I can handle FQDNs, blank spaces, and duplicates."
    
#Generate our GUI
$form = New-Object System.Windows.Forms.Form 
$form.Text = "Enter ComputerNames"
$form.StartPosition = "CenterScreen"

$form.Controls.Add($Label) 

$OffsetY += $LabelBuffer + ($CharHeight * 3)

$DataEntry = New-Object System.Windows.Forms.RichTextBox
$DataEntry.Location = New-Object System.Drawing.Point($OffsetX,$OffsetY) 
$DataEntry.Size = New-Object System.Drawing.Size(260,($TextBoxHeight * 20))
$DataEntry.Multiline = $true
$DataEntry.ScrollBars = 3
$DataEntry.ShortcutsEnabled = $true

$form.Controls.Add($DataEntry) 

$OffsetY += $TextBoxHeight*20 + $ComboBuffer

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,$OffsetY)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,$OffsetY)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Skip"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$form.Size = New-Object System.Drawing.Size(300,($OffsetY + $TextBoxHeight*4)) 

#Make sure our GUI is up in yo face
$form.Topmost = $True

if ($form.ShowDialog() -eq "Cancel") {
    $RemoveComputers = $false
    }

$comps += $DataEntry.Text.split("`n") | foreach { $_.split(".")[0].trim().split(" ")[0].trim()} | sort | get-unique | Where {$_ -ne ""}

#Now get Users
$OffsetX = 10
$OffsetY = 20
$CharHeight = $size.Height
#$CharWidth = $size.Width #I dont know why this is wrong, but it is
$CharWidth = 7

$ComboBuffer = 5 #pixels between data entry and next label
$LabelBuffer = 3 #Pixels betwen label and its data entry field
$TextBoxHeight = 20

$Label = New-Object System.Windows.Forms.Label

$Label.Location = New-Object System.Drawing.Point($OffsetX,$OffsetY) 
$Label.Size = New-Object System.Drawing.Size(280,($CharHeight * 3)) 
$Label.Text = "Paste in your list of Users.  Make sure they provide an EDIPI in the string, or the string is just the Pre-W2K Logon Name."
    
#Generate our GUI
$form = New-Object System.Windows.Forms.Form 
$form.Text = "Enter Users"
$form.StartPosition = "CenterScreen"

$form.Controls.Add($Label) 

$OffsetY += $LabelBuffer + ($CharHeight * 3)

$DataEntry = New-Object System.Windows.Forms.RichTextBox
$DataEntry.Location = New-Object System.Drawing.Point($OffsetX,$OffsetY) 
$DataEntry.Size = New-Object System.Drawing.Size(260,($TextBoxHeight * 20))
$DataEntry.Multiline = $true
$DataEntry.ScrollBars = 3
$DataEntry.ShortcutsEnabled = $true

$form.Controls.Add($DataEntry) 

$OffsetY += $TextBoxHeight*20 + $ComboBuffer

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,$OffsetY)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,$OffsetY)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Skip"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$form.Size = New-Object System.Drawing.Size(300,($OffsetY + $TextBoxHeight*4)) 

#Make sure our GUI is up in yo face
$form.Topmost = $True

if ($form.ShowDialog() -eq "Cancel") {
    $RemoveUsers = $false
    }

#If we aren't removing Users or Computers, just quit
If ($RemoveComputers -eq $false -and $RemoveUsers -eq $false) {exit}

$Users += $DataEntry.Text.split("`n") | foreach { $_.trim()} | sort | get-unique | Where {$_ -ne ""}

#$domain = "AFMC"
if (!$domain) {
    $Title = "Select Domain"
    $Message = "Select which domain these computers are in"
    $ACC = New-Object System.Management.Automation.Host.ChoiceDescription "&ACC","ACC"
    $AFMC = New-Object System.Management.Automation.Host.ChoiceDescription "&SAFMC","AFMC"
    $Cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Will Abort Script"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($ACC,$AFMC,$Cancel)
    do {
        $result = $host.ui.PromptForChoice($title,$message,$options,0)

        $success = $true
        switch ($result) {
            0 {$domain = "ACC"}
            1 {$domain = "AFMC"}
            2 {exit}
            default {$success = $false}
            }
        } until ($success -eq $true)
    }    

switch ($domain) {
    "ACC" {
        $server = Get-ADDomainController -Server "acc.accroot.ds.af.smil.mil" | select -ExpandProperty hostname
        $parentGroup = "SG-Removable Media Write Access Block Exempt"
        $groupFilter = "GLS_*_Removable Media Write Access Block Exempt"
        $groupOUDN = "OU=_Enterprise,OU=Administrative Groups,OU=Administration,DC=acc,DC=accroot,DC=ds,DC=af,DC=smil,DC=mil"
        }
    "AFMC" {
        $server = Get-ADDomainController -Server "afmc.ds.af.smil.mil" | select -ExpandProperty hostname
        #$server = "ftfa-dc-002v.afmc.ds.af.smil.mil"
        $parentGroup = "SG-Removable Media Write Access Exempt"
        $groupFilter = "GLS_*_Removable Media Write Access Exempt"
        $groupOUDN = "OU=_ENTERPRISE,OU=Administrative Groups,OU=Administration,DC=afmc,DC=ds,DC=af,DC=smil,DC=mil"
        }
    default {break}
    }

#ACC : SG-Removable Media Write Access Block Exempt
#AFMC : SG-Removable Media Write Access Exempt

#Find which Base Group to remove from
$Bases = Get-ADGroup -Server $server -SearchBase $groupOUDN -Filter {Name -like $groupFilter} | foreach {$_.name.split("_")[1]} | sort

$OffsetY = 20
$MaxWidth = 300
            
$form_SelectBase = New-Object System.Windows.Forms.Form 
$form_SelectBase.Text = "Select Base"
$form_SelectBase.StartPosition = "CenterScreen"
                
$Label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Point($OffsetX,$OffsetY)
$Label.Size = New-Object System.Drawing.Point($MaxWidth,$CharHeight) 
$Label.Text = "Which Base are we removing Burn Rights from"
$form_SelectBase.Controls.Add($Label)
                
$OffsetY += $LabelBuffer + $CharHeight

$ComboBox = new-object System.Windows.Forms.ComboBox
$ComboBox.Location = new-object System.Drawing.Size($OffsetX,$OffsetY)
$ComboBox.Size = new-object System.Drawing.Size(($MaxWidth - $CharWidth * 2),$CharHeight)
$ComboBox.DropDownStyle = "DropDownList"

$OffsetY += $TextBoxHeight + $ComboBuffer
                
foreach ($str in (@("None") + $bases)) {
    $ComboBox.Items.Add($str) | Out-Null
    }

$form_SelectBase.Controls.Add($ComboBox)
                
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,$OffsetY)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form_SelectBase.AcceptButton = $OKButton
$form_SelectBase.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,$OffsetY)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form_SelectBase.CancelButton = $CancelButton
$form_SelectBase.Controls.Add($CancelButton)
                
#Dynamically decide how long our window will be
#I honestly don't remember my logic behind the additions, but it works
$form_SelectBase.Size = New-Object System.Drawing.Size(($MaxWidth + ($OffsetX * 2)),($OffsetY + $TextBoxHeight*4)) 

#Make sure our GUI is up in yo face
$form_SelectBase.Topmost = $True

if ($form_SelectBase.ShowDialog() -eq "Cancel") {exit}

switch ($ComboBox.SelectedItem) {
    "None" {
        $group = $parentGroup
        }
    default {
        $group = $groupFilter.Replace("*",$ComboBox.SelectedItem)
        }
    }
$GroupDN = Get-ADGroup -Server $server $group | select -ExpandProperty distinguishedname

$strArray = @()

if ($RemoveComputers -ne $false) {
    $Removed = @()
    $NotMember = @()
    $CompDontExist = @()
    foreach ($comp in $comps) {
        try {
            $co = Get-ADComputer -Server $server $comp -Properties memberof -EA Stop
            #if (($co | select -ExpandProperty memberof) -contains $groupDN) {
            [array]$remGroups = $co.memberof -match ($groupFilter.split("_") | select -Last 1)
            if ($remGroups.count -ne 0) {
                $Removed += $comp
                $remGroups | foreach {Remove-ADGroupMember -Server $server -Identity $_ -Members $co -EA silentlycontinue -Confirm:$false}
                continue
                }
            else {
                $NotMember += $comp
                continue
                }
            }
        Catch {
            $CompDontExist += $comp
            continue
            }
        }

    switch ($Removed.count) {
        0 {break}
        1 { 
            $StrArray += "$Removed has been removed from `"$group`""
            $StrArray += "`nPlease reboot $Removed and force a gpupdate to ensure it gets the appropriate GPOs.`n`n"
            break
            }
        default {
            $StrArray += "The following computers have been removed from `"$group`":"
            $StrArray += $Removed
            $StrArray += "`nPlease reboot these systems and force a gpupdate to ensure they get the appropriate GPOs.`n`n"
            }
        }
    switch ($NotMember.count) {
        0 {break}
        1 { 
            $StrArray += "$NotMember is already not a member of `"$group`"`n`n"
            break
            }
        default {
            $StrArray += "The following computers were already not members of `"$group`":"
            $StrArray += $NotMember
            $StrArray += "`n`n"
            }
        }
    switch ($CompDontExist.count) {
        0 {break}
        1 { 
            $StrArray += "$CompDontExist does not exist in $domain`n`n"
            break
            }
        default {
            $StrArray += "The following computers do not exist in $domain`:"
            $StrArray += $CompDontExist
            $StrArray += "`n`n"
            }
        }
    }

If ($RemoveUsers -ne $false) {
    $Removed = @()
    $NotMember = @()
    $UserDontExist = @()
    :main foreach ($user in $users) {
        Remove-Variable EDIPI,acc -EA SilentlyContinue
        try {
            $account = Get-ADUser $user -Server $server -Properties memberof -EA Stop
            }
        catch {
            $EDIPI = [regex]::Match($user,"[0-9]{10}").value
            if (!$EDIPI) {
                $UserDontExist += $user
                continue main
                }
            $filter = $EDIPI + "*"
            do {
                $exit = $true
                try {
                    [array]$account = Get-ADUser -Server $server -Filter {UserPrincipalName -like $filter -or SAMAccountName -like $filter} -Properties memberof | where {$_.distinguishedname -notlike "*OU=Administration*"}
                    switch ($account.count) {
                        0 { #No user found
                            $UserDontExist += $user
                            continue main
                            }
                        1 { #Found only one user by this EDIPI, we good
                            break
                            }
                        default {#Found multiple accounts
                            #Use GUI to make us select the correct one
                            $OffsetY = 20
                            $MaxWidth = 300
            
                            $form_SelectBase = New-Object System.Windows.Forms.Form 
                            $form_SelectBase.Text = "Select Account"
                            $form_SelectBase.StartPosition = "CenterScreen"
                
                            $Label = New-Object System.Windows.Forms.Label
                            $Label.Location = New-Object System.Drawing.Point($OffsetX,$OffsetY)
                            $Label.Size = New-Object System.Drawing.Point($MaxWidth,$CharHeight) 
                            $Label.Text = "Which Logon name is the correct one? (Hit Cancel to skip)"
                            $form_SelectBase.Controls.Add($Label)
                
                            $OffsetY += $LabelBuffer + $CharHeight

                            $ComboBox = new-object System.Windows.Forms.ComboBox
                            $ComboBox.Location = new-object System.Drawing.Size($OffsetX,$OffsetY)
                            $ComboBox.Size = new-object System.Drawing.Size(($MaxWidth - $CharWidth * 2),$CharHeight)
                            $ComboBox.DropDownStyle = "DropDownList"

                            $OffsetY += $TextBoxHeight + $ComboBuffer
                
                            foreach ($str in $account) {
                                $ComboBox.Items.Add($str.samaccountname) | Out-Null
                                }

                            $form_SelectBase.Controls.Add($ComboBox)
                
                            $OKButton = New-Object System.Windows.Forms.Button
                            $OKButton.Location = New-Object System.Drawing.Point(75,$OffsetY)
                            $OKButton.Size = New-Object System.Drawing.Size(75,23)
                            $OKButton.Text = "OK"
                            $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                            $form_SelectBase.AcceptButton = $OKButton
                            $form_SelectBase.Controls.Add($OKButton)

                            $CancelButton = New-Object System.Windows.Forms.Button
                            $CancelButton.Location = New-Object System.Drawing.Point(150,$OffsetY)
                            $CancelButton.Size = New-Object System.Drawing.Size(75,23)
                            $CancelButton.Text = "Cancel"
                            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                            $form_SelectBase.CancelButton = $CancelButton
                            $form_SelectBase.Controls.Add($CancelButton)
                
                            #Dynamically decide how long our window will be
                            #I honestly don't remember my logic behind the additions, but it works
                            $form_SelectBase.Size = New-Object System.Drawing.Size(($MaxWidth + ($OffsetX * 2)),($OffsetY + $TextBoxHeight*4)) 

                            #Make sure our GUI is up in yo face
                            $form_SelectBase.Topmost = $True

                            if ($form_SelectBase.ShowDialog() -eq "Cancel") {
                                $UserDontExist += $user
                                continue main
                                }
                            $account = get-aduser $ComboBox.SelectedItem -Server $server -Properties memberof
                
                            }
                }
                }
                catch {
                    if ($_.Exception.Message -eq "This operation returned because the timeout period expired") {
                        Write-Host "Error: timeout trying to query account for  $user ($EDIPI)"
                        Write-Host "Will retry $(5-$attempts) more times"
                        $attempts++
                        $exit = $false
                        }
                    else {
                        Write-Host "Error: could not gather user for $user ($EDIPI)"
                        $_.Exception.Message | Write-Host -ForegroundColor Red
                        $_.InvocationInfo.PositionMessage | Write-Host -ForegroundColor Red
                        continue main
                        }
                    }
                } until ($exit -or $attempts -gt 5)
                if ($attempts -gt 5) {
                    Write-Host "Error: Constant timeouts looking for  $user ($EDIPI).  Skipping"
                    continue main
                    }

            }
        
        #if (($account | select -ExpandProperty memberof) -contains $groupDN) {
        [array]$remGroups = $account.memberof | Where {$_ -match "Removable Media Write Access Exempt"}
        if ($remGroups.Count -ne 0) {
            #Remove-ADGroupMember -Server $server -Identity $group -Members $account -EA silentlycontinue -Confirm:$false
            $remGroups | foreach {Remove-ADGroupMember -Server $server -Identity $_ -Members $account -EA silentlycontinue -Confirm:$false}
            $Removed += $user
            continue
            }
        else {
            $NotMember += $user
            }

        }

    switch ($Removed.count) {
        0 {break}
        1 { 
            $StrArray += "$Removed has been removed from `"$group`"`n`n"
            break
            }
        default {
            $StrArray += "The following users have been removed from `"$group`":"
            $StrArray += $Removed
            $StrArray += "`n`n"
            }
        }
    switch ($NotMember.count) {
        0 {break}
        1 { 
            $StrArray += "$NotMember is already not a member of `"$group`"`n`n"
            break
            }
        default {
            $StrArray += "The following users were already not members of `"$group`":"
            $StrArray += $NotMember
            $StrArray += "`n`n"
            }
        }
    switch ($UserDontExist.count) {
        0 {break}
        1 { 
            $StrArray += "We could not find a corresponding user account for $UserDontExist in $domain, based on their EDIPI."
            break
            }
        default {
            $StrArray += "We could not find corresponding user accounts for the following users in $domain, based on their EDIPIs`:"
            $StrArray += $UserDontExist
            }
        }
    }

$strArray | clip
Write-Host "Remedy notes copied to clipboard"
read-host -Prompt "Script Finished.  Press Enter to close window."
