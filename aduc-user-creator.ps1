Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:targetOuPatterns = @(
    "OU=Users,OU=Camden,OU=Mazza Demo",
    "OU=Users,OU=Tinton Falls,OU=Mazza Demo",
    "OU=Users,OU=Philadelphia,OU=Mazza Demo",
    "OU=Users,OU=Campus Parkway,OU=Mazza Demo"
)

$script:ouDisplayMap = @{}

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
        [string]$Message,
        [System.Drawing.Color]$Color
    )

    $statusLabel.Text = $Message
    $statusLabel.ForeColor = $Color
}

function Load-OUs {
    $ouComboBox.Items.Clear()
    $script:ouDisplayMap = @{}

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Message "ActiveDirectory module not found. Enter OU DN manually." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    try {
        $ous = Get-ADOrganizationalUnit -Filter * | Sort-Object Name
    }
    catch {
        Update-Status -Message "Unable to query OUs. Enter OU DN manually." -Color ([System.Drawing.Color]::DarkRed)
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
            [void]$ouComboBox.Items.Add($displayName)
        }
    }

    if ($ouComboBox.Items.Count -gt 0) {
        $ouComboBox.SelectedIndex = 0
        Update-Status -Message "Loaded target OUs from Active Directory." -Color ([System.Drawing.Color]::DarkGreen)
    }
    else {
        Update-Status -Message "No OUs found. Enter OU DN manually." -Color ([System.Drawing.Color]::DarkRed)
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "ADUC User Creator"
$form.Size = New-Object System.Drawing.Size(640, 460)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$font = New-Object System.Drawing.Font("Segoe UI", 10)

$firstNameLabel = New-Object System.Windows.Forms.Label
$firstNameLabel.Text = "First Name"
$firstNameLabel.Location = New-Object System.Drawing.Point(20, 20)
$firstNameLabel.Size = New-Object System.Drawing.Size(120, 24)
$firstNameLabel.Font = $font

$firstNameTextBox = New-Object System.Windows.Forms.TextBox
$firstNameTextBox.Location = New-Object System.Drawing.Point(160, 18)
$firstNameTextBox.Size = New-Object System.Drawing.Size(420, 24)
$firstNameTextBox.Font = $font

$lastNameLabel = New-Object System.Windows.Forms.Label
$lastNameLabel.Text = "Last Name"
$lastNameLabel.Location = New-Object System.Drawing.Point(20, 60)
$lastNameLabel.Size = New-Object System.Drawing.Size(120, 24)
$lastNameLabel.Font = $font

$lastNameTextBox = New-Object System.Windows.Forms.TextBox
$lastNameTextBox.Location = New-Object System.Drawing.Point(160, 58)
$lastNameTextBox.Size = New-Object System.Drawing.Size(420, 24)
$lastNameTextBox.Font = $font

$domainLabel = New-Object System.Windows.Forms.Label
$domainLabel.Text = "Email/UPN Domain"
$domainLabel.Location = New-Object System.Drawing.Point(20, 100)
$domainLabel.Size = New-Object System.Drawing.Size(140, 24)
$domainLabel.Font = $font

$domainComboBox = New-Object System.Windows.Forms.ComboBox
$domainComboBox.Location = New-Object System.Drawing.Point(160, 98)
$domainComboBox.Size = New-Object System.Drawing.Size(420, 24)
$domainComboBox.Font = $font
$domainComboBox.DropDownStyle = "DropDownList"

$ouLabel = New-Object System.Windows.Forms.Label
$ouLabel.Text = "Target OU"
$ouLabel.Location = New-Object System.Drawing.Point(20, 140)
$ouLabel.Size = New-Object System.Drawing.Size(120, 24)
$ouLabel.Font = $font

$ouComboBox = New-Object System.Windows.Forms.ComboBox
$ouComboBox.Location = New-Object System.Drawing.Point(160, 138)
$ouComboBox.Size = New-Object System.Drawing.Size(320, 24)
$ouComboBox.Font = $font
$ouComboBox.DropDownStyle = "DropDown"

$loadOuButton = New-Object System.Windows.Forms.Button
$loadOuButton.Text = "Load OUs"
$loadOuButton.Location = New-Object System.Drawing.Point(490, 136)
$loadOuButton.Size = New-Object System.Drawing.Size(90, 28)
$loadOuButton.Font = $font
$loadOuButton.Add_Click({ Load-OUs })

$passwordLabel = New-Object System.Windows.Forms.Label
$passwordLabel.Text = "Password (optional)"
$passwordLabel.Location = New-Object System.Drawing.Point(20, 180)
$passwordLabel.Size = New-Object System.Drawing.Size(140, 24)
$passwordLabel.Font = $font

$passwordTextBox = New-Object System.Windows.Forms.TextBox
$passwordTextBox.Location = New-Object System.Drawing.Point(160, 178)
$passwordTextBox.Size = New-Object System.Drawing.Size(420, 24)
$passwordTextBox.Font = $font
$passwordTextBox.UseSystemPasswordChar = $true

$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Text = "Preview"
$previewLabel.Location = New-Object System.Drawing.Point(20, 220)
$previewLabel.Size = New-Object System.Drawing.Size(120, 24)
$previewLabel.Font = $font

$previewTextBox = New-Object System.Windows.Forms.TextBox
$previewTextBox.Location = New-Object System.Drawing.Point(160, 218)
$previewTextBox.Size = New-Object System.Drawing.Size(420, 100)
$previewTextBox.Multiline = $true
$previewTextBox.ReadOnly = $true
$previewTextBox.Font = $font

$createButton = New-Object System.Windows.Forms.Button
$createButton.Text = "Create User"
$createButton.Location = New-Object System.Drawing.Point(160, 330)
$createButton.Size = New-Object System.Drawing.Size(140, 36)
$createButton.Font = $font

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready"
$statusLabel.Location = New-Object System.Drawing.Point(20, 380)
$statusLabel.Size = New-Object System.Drawing.Size(560, 24)
$statusLabel.Font = $font
$statusLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray

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

$firstNameTextBox.Add_TextChanged($updatePreview)
$lastNameTextBox.Add_TextChanged($updatePreview)
$domainComboBox.Add_SelectedIndexChanged($updatePreview)

$createButton.Add_Click({
    $first = $firstNameTextBox.Text.Trim()
    $last  = $lastNameTextBox.Text.Trim()
    $domain = $domainComboBox.Text.Trim()

    $ouSelection = $ouComboBox.Text.Trim()
    $ouDn = $ouSelection
    if ($script:ouDisplayMap.ContainsKey($ouSelection)) {
        $ouDn = $script:ouDisplayMap[$ouSelection]
    }

    if ([string]::IsNullOrWhiteSpace($first) -or
        [string]::IsNullOrWhiteSpace($last) -or
        [string]::IsNullOrWhiteSpace($domain)) {
        Update-Status -Message "First name, last name, and domain are required." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    if ([string]::IsNullOrWhiteSpace($ouDn)) {
        Update-Status -Message "Select or enter an OU distinguished name." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Message "ActiveDirectory module not available. Cannot create user." -Color ([System.Drawing.Color]::DarkRed)
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
            Name            = "$first $last"
            GivenName       = $first
            Surname         = $last
            SamAccountName  = $sam
            UserPrincipalName = $upn
            Path            = $ouDn
            Enabled         = $true
            OtherAttributes = $otherAttributes
        }

        if ($securePassword) {
            $params.AccountPassword = $securePassword
        }
        else {
            $params.Enabled = $false
        }

        New-ADUser @params
        Update-Status -Message "User created successfully." -Color ([System.Drawing.Color]::DarkGreen)
    }
    catch {
        Update-Status -Message "Failed to create user: $($_.Exception.Message)" -Color ([System.Drawing.Color]::DarkRed)
    }
})

$form.Controls.AddRange(@(
    $firstNameLabel,
    $firstNameTextBox,
    $lastNameLabel,
    $lastNameTextBox,
    $domainLabel,
    $domainComboBox,
    $ouLabel,
    $ouComboBox,
    $loadOuButton,
    $passwordLabel,
    $passwordTextBox,
    $previewLabel,
    $previewTextBox,
    $createButton,
    $statusLabel
))

$form.Add_Shown({
    $form.Activate()

    $domainValues = Get-AvailableDomains
    if ($domainValues -and $domainValues.Count -gt 0) {
        foreach ($domain in $domainValues) {
            [void]$domainComboBox.Items.Add($domain)
        }
        $domainComboBox.SelectedIndex = 0
        Update-Status -Message "Domains loaded from Active Directory." -Color ([System.Drawing.Color]::DarkGreen)
    }
    else {
        Update-Status -Message "Unable to load domains from Active Directory." -Color ([System.Drawing.Color]::DarkRed)
    }

    Load-OUs
})

[void]$form.ShowDialog()
