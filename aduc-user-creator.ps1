Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:targetOuPatterns = @(
    "OU=Users,OU=Camden,OU=Mazza Demo",
    "OU=Users,OU=Tinton Falls,OU=Mazza Demo",
    "OU=Users,OU=Philadelphia,OU=Mazza Demo",
    "OU=Users,OU=Campus Parkway,OU=Mazza Demo"
)

$script:ouDisplayMap = @{}
$script:terminateUserMap = @{}

function Get-SamAccountName {
    param(
        [string]$FirstName,
        [string]$LastName
    )

    $firstInitial = $FirstName.Trim().Substring(0, 1)
    $last = $LastName.Trim()
    $sam = ($firstInitial + $last).ToLowerInvariant()
    return ($sam -replace "[^a-z0-9]", "")
}

function Get-ProxyAddresses {
    param(
        [string]$MailNickname,
        [string]$Domain
    )

    $primary = "SMTP:$MailNickname@$Domain"
    return @($primary)
}

function Get-AvailableDomains {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        return $null
    }

    try {
        $forest = Get-ADForest
        $domains = @()
        if ($forest.Domains) {
            $domains += $forest.Domains
        }
        if ($forest.UPNSuffixes) {
            $domains += $forest.UPNSuffixes
        }
        return $domains | Sort-Object -Unique
    }
    catch {
        return $null
    }
}

function Update-Status {
    param(
        [System.Windows.Forms.Label]$Label,
        [string]$Message,
        [System.Drawing.Color]$Color
    )

    $Label.Text = $Message
    $Label.ForeColor = $Color
}

function Get-DisableOuFromUserOu {
    param(
        [string]$UserOuDn
    )

    if ($UserOuDn -match "OU=Users,") {
        return $UserOuDn -replace "OU=Users,", "OU=Disabled Users,"
    }

    return $null
}

function New-RandomPassword {
    param(
        [int]$Length = 16
    )

    $chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%&*"
    $random = New-Object System.Random
    $passwordChars = for ($i = 0; $i -lt $Length; $i++) {
        $chars[$random.Next(0, $chars.Length)]
    }
    return -join $passwordChars
}

function Load-OUs {
    param(
        [System.Windows.Forms.ComboBox]$ComboBox,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $ComboBox.Items.Clear()
    $script:ouDisplayMap = @{}

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $StatusLabel -Message "ActiveDirectory module not found. Enter OU DN manually." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    try {
        $ous = Get-ADOrganizationalUnit -Filter * | Sort-Object Name
    }
    catch {
        Update-Status -Label $StatusLabel -Message "Unable to query OUs. Enter OU DN manually." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    foreach ($ou in $ous) {
        $dn = $ou.DistinguishedName
        $matchesPattern = $false
        $matchedPattern = $null

        foreach ($pattern in $script:targetOuPatterns) {
            if ($dn -like "*$pattern*") {
                $matchesPattern = $true
                $matchedPattern = $pattern
                break
            }
        }

        if ($matchesPattern) {
            $displayName = $matchedPattern -replace "^OU=Users,OU=", ""
            $displayName = $displayName -replace ",OU=Mazza Demo", ""
            $displayName = $displayName -replace "OU=", ""
            $displayName = $displayName -replace ",", " / "

            $script:ouDisplayMap[$displayName] = $dn
            [void]$ComboBox.Items.Add($displayName)
        }
    }

    if ($ComboBox.Items.Count -gt 0) {
        $ComboBox.SelectedIndex = 0
        Update-Status -Label $StatusLabel -Message "Loaded target OUs from Active Directory." -Color ([System.Drawing.Color]::DarkGreen)
    }
    else {
        Update-Status -Label $StatusLabel -Message "No OUs found. Enter OU DN manually." -Color ([System.Drawing.Color]::DarkRed)
    }
}

function Resolve-OuDn {
    param(
        [string]$Selection
    )

    if ($script:ouDisplayMap.ContainsKey($Selection)) {
        return $script:ouDisplayMap[$Selection]
    }

    return $Selection
}

function Get-ThumbnailImage {
    param(
        [string]$Url,
        [int]$Width,
        [int]$Height
    )

    try {
        $webClient = New-Object System.Net.WebClient
        $bytes = $webClient.DownloadData($Url)
        $stream = New-Object System.IO.MemoryStream(,$bytes)
        $image = [System.Drawing.Image]::FromStream($stream)
        $scaled = New-Object System.Drawing.Bitmap($image, $Width, $Height)
        $stream.Dispose()
        $image.Dispose()
        return $scaled
    }
    catch {
        return $null
    }
}

function Show-Panel {
    param(
        [System.Windows.Forms.Panel]$PanelToShow,
        [System.Windows.Forms.Panel]$PanelToHide1,
        [System.Windows.Forms.Panel]$PanelToHide2
    )

    $PanelToShow.Visible = $true
    $PanelToShow.BringToFront()
    if ($PanelToHide1) {
        $PanelToHide1.Visible = $false
    }
    if ($PanelToHide2) {
        $PanelToHide2.Visible = $false
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "J.U.M.P. Jaydien User Management Platform"
$form.Size = New-Object System.Drawing.Size(720, 520)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$font = New-Object System.Drawing.Font("Segoe UI", 10)

$mainMenuPanel = New-Object System.Windows.Forms.Panel
$mainMenuPanel.Location = New-Object System.Drawing.Point(10, 10)
$mainMenuPanel.Size = New-Object System.Drawing.Size(690, 460)

$createPanel = New-Object System.Windows.Forms.Panel
$createPanel.Location = New-Object System.Drawing.Point(10, 10)
$createPanel.Size = New-Object System.Drawing.Size(690, 460)
$createPanel.Visible = $false

$terminatePanel = New-Object System.Windows.Forms.Panel
$terminatePanel.Location = New-Object System.Drawing.Point(10, 10)
$terminatePanel.Size = New-Object System.Drawing.Size(690, 460)
$terminatePanel.Visible = $false

$menuTitleLabel = New-Object System.Windows.Forms.Label
$menuTitleLabel.Text = "Select an action"
$menuTitleLabel.Location = New-Object System.Drawing.Point(20, 20)
$menuTitleLabel.Size = New-Object System.Drawing.Size(300, 30)
$menuTitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)

$createTileButton = New-Object System.Windows.Forms.Button
$createTileButton.Text = "Create User"
$createTileButton.Location = New-Object System.Drawing.Point(80, 80)
$createTileButton.Size = New-Object System.Drawing.Size(230, 260)
$createTileButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$createTileButton.TextImageRelation = "ImageAboveText"
$createTileButton.ImageAlign = "MiddleCenter"
$createTileButton.TextAlign = "BottomCenter"

$terminateTileButton = New-Object System.Windows.Forms.Button
$terminateTileButton.Text = "Terminate User"
$terminateTileButton.Location = New-Object System.Drawing.Point(380, 80)
$terminateTileButton.Size = New-Object System.Drawing.Size(230, 260)
$terminateTileButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$terminateTileButton.TextImageRelation = "ImageAboveText"
$terminateTileButton.ImageAlign = "MiddleCenter"
$terminateTileButton.TextAlign = "BottomCenter"

$createTileButton.Add_Click({
    Show-Panel -PanelToShow $createPanel -PanelToHide1 $mainMenuPanel -PanelToHide2 $terminatePanel
})

$terminateTileButton.Add_Click({
    Show-Panel -PanelToShow $terminatePanel -PanelToHide1 $mainMenuPanel -PanelToHide2 $createPanel
})

$backToMenuFromCreate = New-Object System.Windows.Forms.Button
$backToMenuFromCreate.Text = "Back"
$backToMenuFromCreate.Location = New-Object System.Drawing.Point(20, 20)
$backToMenuFromCreate.Size = New-Object System.Drawing.Size(80, 28)
$backToMenuFromCreate.Font = $font
$backToMenuFromCreate.Add_Click({
    Show-Panel -PanelToShow $mainMenuPanel -PanelToHide1 $createPanel -PanelToHide2 $terminatePanel
})

$backToMenuFromTerminate = New-Object System.Windows.Forms.Button
$backToMenuFromTerminate.Text = "Back"
$backToMenuFromTerminate.Location = New-Object System.Drawing.Point(20, 20)
$backToMenuFromTerminate.Size = New-Object System.Drawing.Size(80, 28)
$backToMenuFromTerminate.Font = $font
$backToMenuFromTerminate.Add_Click({
    Show-Panel -PanelToShow $mainMenuPanel -PanelToHide1 $createPanel -PanelToHide2 $terminatePanel
})

$firstNameLabel = New-Object System.Windows.Forms.Label
$firstNameLabel.Text = "First Name"
$firstNameLabel.Location = New-Object System.Drawing.Point(20, 60)
$firstNameLabel.Size = New-Object System.Drawing.Size(120, 24)
$firstNameLabel.Font = $font

$firstNameTextBox = New-Object System.Windows.Forms.TextBox
$firstNameTextBox.Location = New-Object System.Drawing.Point(180, 58)
$firstNameTextBox.Size = New-Object System.Drawing.Size(460, 24)
$firstNameTextBox.Font = $font

$lastNameLabel = New-Object System.Windows.Forms.Label
$lastNameLabel.Text = "Last Name"
$lastNameLabel.Location = New-Object System.Drawing.Point(20, 100)
$lastNameLabel.Size = New-Object System.Drawing.Size(120, 24)
$lastNameLabel.Font = $font

$lastNameTextBox = New-Object System.Windows.Forms.TextBox
$lastNameTextBox.Location = New-Object System.Drawing.Point(180, 98)
$lastNameTextBox.Size = New-Object System.Drawing.Size(460, 24)
$lastNameTextBox.Font = $font

$domainLabel = New-Object System.Windows.Forms.Label
$domainLabel.Text = "Email/UPN Domain"
$domainLabel.Location = New-Object System.Drawing.Point(20, 140)
$domainLabel.Size = New-Object System.Drawing.Size(150, 24)
$domainLabel.Font = $font

$domainComboBox = New-Object System.Windows.Forms.ComboBox
$domainComboBox.Location = New-Object System.Drawing.Point(180, 138)
$domainComboBox.Size = New-Object System.Drawing.Size(460, 24)
$domainComboBox.Font = $font
$domainComboBox.DropDownStyle = "DropDownList"

$ouLabel = New-Object System.Windows.Forms.Label
$ouLabel.Text = "Target OU"
$ouLabel.Location = New-Object System.Drawing.Point(20, 180)
$ouLabel.Size = New-Object System.Drawing.Size(120, 24)
$ouLabel.Font = $font

$ouComboBox = New-Object System.Windows.Forms.ComboBox
$ouComboBox.Location = New-Object System.Drawing.Point(180, 178)
$ouComboBox.Size = New-Object System.Drawing.Size(360, 24)
$ouComboBox.Font = $font
$ouComboBox.DropDownStyle = "DropDown"

$passwordLabel = New-Object System.Windows.Forms.Label
$passwordLabel.Text = "Password (optional)"
$passwordLabel.Location = New-Object System.Drawing.Point(20, 220)
$passwordLabel.Size = New-Object System.Drawing.Size(150, 24)
$passwordLabel.Font = $font

$passwordTextBox = New-Object System.Windows.Forms.TextBox
$passwordTextBox.Location = New-Object System.Drawing.Point(180, 218)
$passwordTextBox.Size = New-Object System.Drawing.Size(460, 24)
$passwordTextBox.Font = $font
$passwordTextBox.UseSystemPasswordChar = $true

$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Text = "Preview"
$previewLabel.Location = New-Object System.Drawing.Point(20, 260)
$previewLabel.Size = New-Object System.Drawing.Size(120, 24)
$previewLabel.Font = $font

$previewTextBox = New-Object System.Windows.Forms.TextBox
$previewTextBox.Location = New-Object System.Drawing.Point(180, 258)
$previewTextBox.Size = New-Object System.Drawing.Size(460, 90)
$previewTextBox.Multiline = $true
$previewTextBox.ReadOnly = $true
$previewTextBox.Font = $font

$createButton = New-Object System.Windows.Forms.Button
$createButton.Text = "Create User"
$createButton.Location = New-Object System.Drawing.Point(180, 360)
$createButton.Size = New-Object System.Drawing.Size(140, 36)
$createButton.Font = $font
$createButton.Enabled = $false

$createStatusLabel = New-Object System.Windows.Forms.Label
$createStatusLabel.Text = "Ready"
$createStatusLabel.Location = New-Object System.Drawing.Point(20, 410)
$createStatusLabel.Size = New-Object System.Drawing.Size(620, 24)
$createStatusLabel.Font = $font
$createStatusLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray

$terminateOuLabel = New-Object System.Windows.Forms.Label
$terminateOuLabel.Text = "Source Users OU"
$terminateOuLabel.Location = New-Object System.Drawing.Point(20, 60)
$terminateOuLabel.Size = New-Object System.Drawing.Size(150, 24)
$terminateOuLabel.Font = $font

$terminateOuComboBox = New-Object System.Windows.Forms.ComboBox
$terminateOuComboBox.Location = New-Object System.Drawing.Point(180, 58)
$terminateOuComboBox.Size = New-Object System.Drawing.Size(360, 24)
$terminateOuComboBox.Font = $font
$terminateOuComboBox.DropDownStyle = "DropDown"

$userListBox = New-Object System.Windows.Forms.ListBox
$userListBox.Location = New-Object System.Drawing.Point(20, 140)
$userListBox.Size = New-Object System.Drawing.Size(620, 180)
$userListBox.Font = $font
$userListBox.SelectionMode = "MultiExtended"

$terminateButton = New-Object System.Windows.Forms.Button
$terminateButton.Text = "Terminate Selected"
$terminateButton.Location = New-Object System.Drawing.Point(180, 330)
$terminateButton.Size = New-Object System.Drawing.Size(180, 36)
$terminateButton.Font = $font

$terminatePreviewLabel = New-Object System.Windows.Forms.Label
$terminatePreviewLabel.Text = "Preview / Log"
$terminatePreviewLabel.Location = New-Object System.Drawing.Point(20, 380)
$terminatePreviewLabel.Size = New-Object System.Drawing.Size(150, 24)
$terminatePreviewLabel.Font = $font

$terminatePreviewTextBox = New-Object System.Windows.Forms.TextBox
$terminatePreviewTextBox.Location = New-Object System.Drawing.Point(180, 378)
$terminatePreviewTextBox.Size = New-Object System.Drawing.Size(460, 56)
$terminatePreviewTextBox.Multiline = $true
$terminatePreviewTextBox.ReadOnly = $true
$terminatePreviewTextBox.Font = $font

$terminateStatusLabel = New-Object System.Windows.Forms.Label
$terminateStatusLabel.Text = "Ready"
$terminateStatusLabel.Location = New-Object System.Drawing.Point(20, 440)
$terminateStatusLabel.Size = New-Object System.Drawing.Size(620, 24)
$terminateStatusLabel.Font = $font
$terminateStatusLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray

$updatePreview = {
    $first = $firstNameTextBox.Text
    $last = $lastNameTextBox.Text
    $domain = $domainComboBox.Text

    if ([string]::IsNullOrWhiteSpace($first) -or
        [string]::IsNullOrWhiteSpace($last) -or
        [string]::IsNullOrWhiteSpace($domain)) {
        $previewTextBox.Text = ""
        return
    }

    $sam = Get-SamAccountName -FirstName $first -LastName $last
    $mailNickname = $sam
    $upn = "$sam@$domain"
    $proxy = (Get-ProxyAddresses -MailNickname $mailNickname -Domain $domain) -join "; "

    $previewTextBox.Text =
        "samAccountName: $sam`r`n" +
        "userPrincipalName: $upn`r`n" +
        "mailNickname: $mailNickname`r`n" +
        "proxyAddresses: $proxy"
}


function Update-CreateButtonState {
    $first = $firstNameTextBox.Text.Trim()
    $last = $lastNameTextBox.Text.Trim()
    $domain = $domainComboBox.Text.Trim()
    $ouSelection = $ouComboBox.Text.Trim()
    $ouDn = Resolve-OuDn -Selection $ouSelection

    $isValid = -not ([string]::IsNullOrWhiteSpace($first) -or
        [string]::IsNullOrWhiteSpace($last) -or
        [string]::IsNullOrWhiteSpace($domain) -or
        [string]::IsNullOrWhiteSpace($ouDn))

    $createButton.Enabled = $isValid
    if (-not $isValid) {
        Update-Status -Label $createStatusLabel -Message "Fill in all required fields to enable Create User." -Color ([System.Drawing.Color]::DarkRed)
    }
    else {
        Update-Status -Label $createStatusLabel -Message "Ready" -Color ([System.Drawing.Color]::DarkSlateGray)
    }
}

function Load-UsersFromOu {
    $userListBox.Items.Clear()
    $script:terminateUserMap = @{}

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $terminateStatusLabel -Message "ActiveDirectory module not available." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $ouSelection = $terminateOuComboBox.Text.Trim()
    $ouDn = Resolve-OuDn -Selection $ouSelection

    if ([string]::IsNullOrWhiteSpace($ouDn)) {
        Update-Status -Label $terminateStatusLabel -Message "Select or enter an OU distinguished name." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    try {
        $users = Get-ADUser -Filter * -SearchBase $ouDn -SearchScope OneLevel | Sort-Object Name
    }
    catch {
        Update-Status -Label $terminateStatusLabel -Message "Unable to query users in OU." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    foreach ($user in $users) {
        $display = $user.Name
        $script:terminateUserMap[$display] = $user
        [void]$userListBox.Items.Add($display)
    }

    Update-Status -Label $terminateStatusLabel -Message "Loaded $($users.Count) users from OU." -Color ([System.Drawing.Color]::DarkGreen)
    $terminatePreviewTextBox.Text = "Loaded $($users.Count) users from $ouDn"
}

$firstNameTextBox.Add_TextChanged({
    $updatePreview.Invoke()
    Update-CreateButtonState
})
$lastNameTextBox.Add_TextChanged({
    $updatePreview.Invoke()
    Update-CreateButtonState
})
$domainComboBox.Add_SelectedIndexChanged({
    $updatePreview.Invoke()
    Update-CreateButtonState
})
$ouComboBox.Add_TextChanged({
    Update-CreateButtonState
})

$terminateOuComboBox.Add_TextChanged({
    Load-UsersFromOu
})

$createButton.Add_Click({
    $first = $firstNameTextBox.Text.Trim()
    $last  = $lastNameTextBox.Text.Trim()
    $domain = $domainComboBox.Text.Trim()

    $ouSelection = $ouComboBox.Text.Trim()
    $ouDn = Resolve-OuDn -Selection $ouSelection

    if ([string]::IsNullOrWhiteSpace($first) -or
        [string]::IsNullOrWhiteSpace($last) -or
        [string]::IsNullOrWhiteSpace($domain)) {
        Update-Status -Label $createStatusLabel -Message "First name, last name, and domain are required." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    if ([string]::IsNullOrWhiteSpace($ouDn)) {
        Update-Status -Label $createStatusLabel -Message "Select or enter an OU distinguished name." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $createStatusLabel -Message "ActiveDirectory module not available. Cannot create user." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $sam = Get-SamAccountName -FirstName $first -LastName $last
    $mailNickname = $sam
    $upn = "$sam@$domain"
    $proxyAddresses = Get-ProxyAddresses -MailNickname $mailNickname -Domain $domain

    $securePassword = $null
    if (-not [string]::IsNullOrWhiteSpace($passwordTextBox.Text)) {
        $securePassword = ConvertTo-SecureString $passwordTextBox.Text -AsPlainText -Force
    }

    $otherAttributes = @{
        proxyAddresses = $proxyAddresses
        mailNickname   = $mailNickname
    }

    try {
        $params = @{
            Name              = "$first $last"
            GivenName         = $first
            Surname           = $last
            SamAccountName    = $sam
            UserPrincipalName = $upn
            Path              = $ouDn
            Enabled           = $true
            OtherAttributes   = $otherAttributes
        }

        if ($securePassword) {
            $params.AccountPassword = $securePassword
        }
        else {
            $params.Enabled = $false
        }

        New-ADUser @params
        Update-Status -Label $createStatusLabel -Message "User created successfully." -Color ([System.Drawing.Color]::DarkGreen)
    }
    catch {
        Update-Status -Label $createStatusLabel -Message "Failed to create user: $($_.Exception.Message)" -Color ([System.Drawing.Color]::DarkRed)
    }
})

$terminateButton.Add_Click({
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $terminateStatusLabel -Message "ActiveDirectory module not available." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $ouSelection = $terminateOuComboBox.Text.Trim()
    $sourceOuDn = Resolve-OuDn -Selection $ouSelection

    if ([string]::IsNullOrWhiteSpace($sourceOuDn)) {
        Update-Status -Label $terminateStatusLabel -Message "Select or enter a source OU." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $disabledOuDn = Get-DisableOuFromUserOu -UserOuDn $sourceOuDn
    if (-not $disabledOuDn) {
        Update-Status -Label $terminateStatusLabel -Message "Unable to resolve Disabled Users OU. Expected OU=Users, in path." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    if ($userListBox.SelectedItems.Count -eq 0) {
        Update-Status -Label $terminateStatusLabel -Message "Select at least one user to terminate." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $failed = @()
    foreach ($display in $userListBox.SelectedItems) {
        $user = $script:terminateUserMap[$display]
        if (-not $user) {
            $failed += $display
            continue
        }

        try {
            Disable-ADAccount -Identity $user.DistinguishedName
            $newPassword = New-RandomPassword
            $securePassword = ConvertTo-SecureString $newPassword -AsPlainText -Force
            Set-ADAccountPassword -Identity $user.DistinguishedName -Reset -NewPassword $securePassword
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $disabledOuDn
        }
        catch {
            $failed += $display
        }
    }

    if ($failed.Count -gt 0) {
        Update-Status -Label $terminateStatusLabel -Message "Completed with errors. Failed: $($failed -join ', ')" -Color ([System.Drawing.Color]::DarkRed)
        $terminatePreviewTextBox.Text = "Failed to terminate: $($failed -join ', ')"
    }
    else {
        Update-Status -Label $terminateStatusLabel -Message "Selected users terminated and moved to Disabled Users OU." -Color ([System.Drawing.Color]::DarkGreen)
        $terminatePreviewTextBox.Text = "Terminated and moved: $($userListBox.SelectedItems -join ', ')"
    }
})

$createPanel.Controls.AddRange(@(
    $backToMenuFromCreate,
    $firstNameLabel,
    $firstNameTextBox,
    $lastNameLabel,
    $lastNameTextBox,
    $domainLabel,
    $domainComboBox,
    $ouLabel,
    $ouComboBox,
    $passwordLabel,
    $passwordTextBox,
    $previewLabel,
    $previewTextBox,
    $createButton,
    $createStatusLabel
))

$terminatePanel.Controls.AddRange(@(
    $backToMenuFromTerminate,
    $terminateOuLabel,
    $terminateOuComboBox,
    $userListBox,
    $terminateButton,
    $terminatePreviewLabel,
    $terminatePreviewTextBox,
    $terminateStatusLabel
))

$mainMenuPanel.Controls.AddRange(@(
    $menuTitleLabel,
    $createTileButton,
    $terminateTileButton
))

$form.Controls.Add($mainMenuPanel)
$form.Controls.Add($createPanel)
$form.Controls.Add($terminatePanel)

$form.Add_Shown({
    $form.Activate()

    $domainValues = Get-AvailableDomains
    if ($domainValues -and $domainValues.Count -gt 0) {
        foreach ($domain in $domainValues) {
            [void]$domainComboBox.Items.Add($domain)
        }
        $domainComboBox.SelectedIndex = 0
        Update-Status -Label $createStatusLabel -Message "Domains loaded from Active Directory." -Color ([System.Drawing.Color]::DarkGreen)
    }
    else {
        Update-Status -Label $createStatusLabel -Message "Unable to load domains from Active Directory." -Color ([System.Drawing.Color]::DarkRed)
    }

    Load-OUs -ComboBox $ouComboBox -StatusLabel $createStatusLabel
    Load-OUs -ComboBox $terminateOuComboBox -StatusLabel $terminateStatusLabel

    $createThumbnail = Get-ThumbnailImage -Url "https://static.thenounproject.com/png/1929283-200.png" -Width 120 -Height 120
    if ($createThumbnail) {
        $createTileButton.Image = $createThumbnail
    }

    $terminateThumbnail = Get-ThumbnailImage -Url "https://static.vecteezy.com/system/resources/thumbnails/032/403/811/small_2x/black-human-silhouette-with-red-cross-deleted-or-blocked-web-user-interface-with-offline-warning-vector.jpg" -Width 120 -Height 120
    if ($terminateThumbnail) {
        $terminateTileButton.Image = $terminateThumbnail
    }
})

[void]$form.ShowDialog()
