# Windows Forms start screen (UI). One window for connection and report options, so an interactive run needs no
# console input. Validation lives in Test-TarStartFormInput (Connection.ps1, cross-platform, unit-tested); this file
# only builds controls and moves values in and out. Loaded everywhere, executed only when Test-TarGuiAvailable.

function Test-TarGuiAvailable {
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    if (-not $IsWindows -or -not [Environment]::UserInteractive) { return $false }
    try { Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction Stop; $true } catch { $false }
}

function Show-TarMessageBox {
    # Yes/No pop-up on top of the console window.
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][string]$Text, [bool]$Default = $true)
    $owner = [System.Windows.Forms.Form]::new()
    try {
        $owner.TopMost = $true
        $defaultButton = if ($Default) { [System.Windows.Forms.MessageBoxDefaultButton]::Button1 } else { [System.Windows.Forms.MessageBoxDefaultButton]::Button2 }
        $answer = [System.Windows.Forms.MessageBox]::Show($owner, $Text, 'Teams Access Report', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question, $defaultButton)
        $answer -eq [System.Windows.Forms.DialogResult]::Yes
    } finally { $owner.Dispose() }
}

function Get-TarFormUi {
    # Event handlers receive the sender ($this); the control map is stored in Form.Tag.
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)]$Control)
    if ($Control -is [System.Windows.Forms.Form]) { $Control.Tag } else { $Control.FindForm().Tag }
}

function Add-TarFormRow {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Table, [string]$Label, [Parameter(Mandatory)]$Control, $Extra)
    $row = $Table.RowCount
    $Table.RowCount++
    $null = $Table.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    if ($Label) {
        $l = [System.Windows.Forms.Label]@{ Text = $Label; AutoSize = $true; Anchor = 'Left'; Margin = [System.Windows.Forms.Padding]::new(3, 7, 10, 3) }
        $Table.Controls.Add($l, 0, $row)
    }
    $Table.Controls.Add($Control, 1, $row)
    if ($Extra) { $Table.Controls.Add($Extra, 2, $row) } else { $Table.SetColumnSpan($Control, 2) }
}

function New-TarFormTable {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an in-memory UI object only.')]
    [CmdletBinding()]
    param()
    $t = [System.Windows.Forms.TableLayoutPanel]@{ ColumnCount = 3; AutoSize = $true; AutoSizeMode = 'GrowAndShrink'; Dock = 'Fill'; Padding = [System.Windows.Forms.Padding]::new(6) }
    $null = $t.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    $null = $t.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    $null = $t.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    $t
}

function New-TarStartForm {
    # Builds the form without showing it (tests construct it on Windows). Returns the control map (also in Form.Tag).
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates in-memory UI objects only.')]
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([AllowEmptyCollection()][object[]]$Profiles = @(), [AllowEmptyCollection()][object[]]$Certificates = @(), [hashtable]$Defaults = @{}, [switch]$ConnectionOnly)
    $textWidth = 380
    $ui = @{ Profiles = @($Profiles); CertRows = @($Certificates); Defaults = $Defaults; ConnectionOnly = [bool]$ConnectionOnly; Result = $null }
    $form = [System.Windows.Forms.Form]@{
        Text = $(if ($ConnectionOnly) { 'Teams Access Report - connection profile' } else { 'Teams Access Report' })
        AutoSize = $true; AutoSizeMode = 'GrowAndShrink'; FormBorderStyle = 'FixedDialog'; MaximizeBox = $false; MinimizeBox = $false
        StartPosition = 'CenterScreen'; TopMost = $true; AutoScaleMode = 'Dpi'; Font = [System.Drawing.SystemFonts]::MessageBoxFont; Padding = [System.Windows.Forms.Padding]::new(8)
    }
    $ui.Form = $form; $form.Tag = $ui
    $root = [System.Windows.Forms.FlowLayoutPanel]@{ FlowDirection = 'TopDown'; AutoSize = $true; AutoSizeMode = 'GrowAndShrink'; WrapContents = $false; Dock = 'Fill' }
    $form.Controls.Add($root)

    # --- Connection ---
    $grpConn = [System.Windows.Forms.GroupBox]@{ Text = 'Connection'; AutoSize = $true; AutoSizeMode = 'GrowAndShrink' }
    $tc = New-TarFormTable
    $grpConn.Controls.Add($tc)
    $ui.Profile = [System.Windows.Forms.ComboBox]@{ DropDownStyle = 'DropDownList'; Width = $textWidth }
    $null = $ui.Profile.Items.Add('<New connection>')
    foreach ($p in $ui.Profiles) { $null = $ui.Profile.Items.Add("$($p.Name)  -  $(if ($p.TenantQuery) { $p.TenantQuery } else { $p.TenantId })  ($($p.AuthMode))") }
    Add-TarFormRow $tc 'Saved connection' $ui.Profile
    $ui.Tenant = [System.Windows.Forms.TextBox]@{ Width = $textWidth }
    $ui.TenantCheck = [System.Windows.Forms.Button]@{ Text = 'Check'; AutoSize = $true }
    Add-TarFormRow $tc 'Tenant' $ui.Tenant $ui.TenantCheck
    $ui.TenantStatus = [System.Windows.Forms.Label]@{ Text = 'e.g. contoso, contoso.com, admin@contoso.com or the tenant ID'; AutoSize = $true; ForeColor = [System.Drawing.SystemColors]::GrayText }
    Add-TarFormRow $tc '' $ui.TenantStatus
    $modes = [System.Windows.Forms.FlowLayoutPanel]@{ AutoSize = $true; WrapContents = $false; Margin = [System.Windows.Forms.Padding]::new(0) }
    $ui.AppOnly = [System.Windows.Forms.RadioButton]@{ Text = 'App-only (certificate)'; AutoSize = $true; Checked = $true }
    $ui.Delegated = [System.Windows.Forms.RadioButton]@{ Text = 'Delegated (sign in as yourself)'; AutoSize = $true }
    $modes.Controls.AddRange(@($ui.AppOnly, $ui.Delegated))
    Add-TarFormRow $tc 'Sign-in' $modes
    $ui.ClientId = [System.Windows.Forms.TextBox]@{ Width = $textWidth }
    Add-TarFormRow $tc 'App (client) ID' $ui.ClientId
    $ui.Certificate = [System.Windows.Forms.ComboBox]@{ DropDownStyle = 'DropDown'; Width = $textWidth }
    foreach ($c in $ui.CertRows) { $null = $ui.Certificate.Items.Add("$($c.Subject)  -  expires $($c.Expires)  -  $($c.Store)") }
    Add-TarFormRow $tc 'Certificate' $ui.Certificate
    $ui.PnPClientId = [System.Windows.Forms.TextBox]@{ Width = $textWidth }
    Add-TarFormRow $tc 'PnP app ID (optional)' $ui.PnPClientId
    $saveRow = [System.Windows.Forms.FlowLayoutPanel]@{ AutoSize = $true; WrapContents = $false; Margin = [System.Windows.Forms.Padding]::new(0) }
    $ui.Save = [System.Windows.Forms.CheckBox]@{ Text = 'Save as profile'; AutoSize = $true; Checked = $true }
    $ui.ProfileName = [System.Windows.Forms.TextBox]@{ Width = 200 }
    $saveRow.Controls.AddRange(@($ui.Save, $ui.ProfileName))
    Add-TarFormRow $tc '' $saveRow
    $root.Controls.Add($grpConn)

    # --- Report options ---
    $grpRun = [System.Windows.Forms.GroupBox]@{ Text = 'Report'; AutoSize = $true; AutoSizeMode = 'GrowAndShrink'; Visible = -not $ConnectionOnly }
    $tr = New-TarFormTable
    $grpRun.Controls.Add($tr)
    $ui.SharePoint = [System.Windows.Forms.CheckBox]@{ Text = 'Include SharePoint permissions, claim resolution and drift (phases 2-4)'; AutoSize = $true }
    Add-TarFormRow $tr '' $ui.SharePoint
    $ui.Associations = [System.Windows.Forms.CheckBox]@{ Text = "Include internal users' teams in other tenants (app-only)"; AutoSize = $true }
    Add-TarFormRow $tr '' $ui.Associations
    $ui.Channels = [System.Windows.Forms.CheckBox]@{ Text = 'Choose channels per team (otherwise all channels)'; AutoSize = $true }
    Add-TarFormRow $tr '' $ui.Channels
    $ui.Output = [System.Windows.Forms.TextBox]@{ Width = $textWidth }
    $ui.Browse = [System.Windows.Forms.Button]@{ Text = 'Browse...'; AutoSize = $true }
    Add-TarFormRow $tr 'Output folder' $ui.Output $ui.Browse
    $root.Controls.Add($grpRun)

    # --- Errors and buttons ---
    $ui.Error = [System.Windows.Forms.Label]@{ AutoSize = $true; MaximumSize = [System.Drawing.Size]::new(560, 0); ForeColor = [System.Drawing.Color]::Firebrick }
    $root.Controls.Add($ui.Error)
    $buttons = [System.Windows.Forms.FlowLayoutPanel]@{ FlowDirection = 'RightToLeft'; AutoSize = $true; Anchor = 'Right'; WrapContents = $false }
    $ui.Start = [System.Windows.Forms.Button]@{ Text = $(if ($ConnectionOnly) { 'Save' } else { 'Start' }); AutoSize = $true }
    $ui.Cancel = [System.Windows.Forms.Button]@{ Text = 'Cancel'; AutoSize = $true; DialogResult = 'Cancel' }
    $buttons.Controls.AddRange(@($ui.Cancel, $ui.Start))
    $root.Controls.Add($buttons)
    $form.AcceptButton = $ui.Start; $form.CancelButton = $ui.Cancel

    # --- Events (each handler resolves the control map from the sender; failures show in the error label) ---
    $ui.Profile.Add_SelectedIndexChanged({ $u = Get-TarFormUi $this; try { Set-TarStartFormProfile -Ui $u } catch { $u.Error.Text = $_.Exception.Message } })
    $ui.AppOnly.Add_CheckedChanged({ Update-TarStartFormState -Ui (Get-TarFormUi $this) })
    $ui.Save.Add_CheckedChanged({ Update-TarStartFormState -Ui (Get-TarFormUi $this) })
    $ui.TenantCheck.Add_Click({ Invoke-TarStartFormCheck -Ui (Get-TarFormUi $this) })
    $ui.Browse.Add_Click({
            $u = Get-TarFormUi $this
            try {
                $dlg = [System.Windows.Forms.FolderBrowserDialog]@{ Description = 'Folder for the report (a new subfolder is fine)'; UseDescriptionForTitle = $true }
                if ($dlg.ShowDialog($u.Form) -eq 'OK') { $u.Output.Text = Join-Path $dlg.SelectedPath (Split-Path $u.Output.Text -Leaf) }
            } catch { $u.Error.Text = "Folder picker unavailable ($($_.Exception.Message)); type the path instead." }
        })
    $ui.Start.Add_Click({ $null = Invoke-TarStartFormSubmit -Ui (Get-TarFormUi $this) })

    # --- Initial values ---
    $ui.Output.Text = [string]$Defaults['OutputPath']
    $ui.PnPClientId.Text = [string]$Defaults['PnPClientId']
    Set-TarStartFormOption -Ui $ui -Options $null
    $newest = @($ui.Profiles | Sort-Object { [string]$_.UpdatedAt } -Descending | Select-Object -First 1)
    $ui.Profile.SelectedIndex = if ($newest.Count) { [array]::IndexOf($ui.Profiles, $newest[0]) + 1 } else { 0 }
    Set-TarStartFormProfile -Ui $ui
    $ui
}

function Set-TarStartFormOption {
    # Profile options first, explicit command-line switches (Defaults) win.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Sets in-memory UI state only.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ui, $Options)
    foreach ($pair in @(@('SharePoint', 'IncludeSharePoint'), @('Associations', 'IncludeUserAssociations'), @('Channels', 'SelectChannels'))) {
        $value = if ($Ui.Defaults.ContainsKey($pair[1])) { [bool]$Ui.Defaults[$pair[1]] } elseif ($Options) { [bool](Get-TarProperty $Options $pair[1]) } else { $false }
        $Ui[$pair[0]].Checked = $value
    }
}

function Set-TarStartFormProfile {
    # Fills the connection fields from the selected profile; index 0 = new connection (fields cleared).
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Sets in-memory UI state only.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ui)
    $i = $Ui.Profile.SelectedIndex
    $p = if ($i -gt 0) { $Ui.Profiles[$i - 1] } else { $null }
    $Ui.Tenant.Text = if ($p) { [string]$(if ($p.TenantQuery) { $p.TenantQuery } else { $p.TenantId }) } else { '' }
    $Ui.TenantStatus.Text = if ($p) { "Tenant ID $($p.TenantId)" } else { 'e.g. contoso, contoso.com, admin@contoso.com or the tenant ID' }
    $Ui.AppOnly.Checked = -not $p -or $p.AuthMode -eq 'AppOnly'
    $Ui.Delegated.Checked = -not $Ui.AppOnly.Checked
    $Ui.ClientId.Text = if ($p) { [string]$p.ClientId } else { '' }
    $Ui.Certificate.SelectedIndex = -1
    $Ui.Certificate.Text = ''
    $thumb = if ($p) { [string]$p.CertificateThumbprint } else { '' }
    # @() around the whole if: an if expression unrolls an empty array to $null.
    $match = @(if ($thumb -and $Ui.CertRows.Count) { 0..($Ui.CertRows.Count - 1) | Where-Object { $Ui.CertRows[$_].Id -eq $thumb } })
    if ($match.Count) { $Ui.Certificate.SelectedIndex = $match[0] }
    elseif ($thumb) { $Ui.Certificate.Text = $thumb }
    elseif (-not $p -and $Ui.CertRows.Count -eq 1) { $Ui.Certificate.SelectedIndex = 0 }
    if ($p -and (Get-TarProperty $p 'PnPClientId')) { $Ui.PnPClientId.Text = [string]$p.PnPClientId }
    $Ui.Save.Text = if ($p) { 'Update profile' } else { 'Save as profile' }
    $Ui.Save.Checked = $true
    $Ui.ProfileName.Text = if ($p) { [string]$p.Name } else { '' }
    Set-TarStartFormOption -Ui $Ui -Options $(if ($p) { Get-TarProperty $p 'Options' } else { $null })
    $Ui.Error.Text = ''
    Update-TarStartFormState -Ui $Ui
}

function Update-TarStartFormState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Sets in-memory UI state only.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ui)
    $app = $Ui.AppOnly.Checked
    $Ui.ClientId.Enabled = $app
    $Ui.Certificate.Enabled = $app
    $Ui.Associations.Enabled = $app
    if (-not $app) { $Ui.Associations.Checked = $false }
    $Ui.ProfileName.Enabled = $Ui.Save.Checked
    if ($Ui.ConnectionOnly) { $Ui.Save.Checked = $true; $Ui.Save.Enabled = $false; $Ui.ProfileName.Enabled = $true }
}

function Get-TarStartFormState {
    # Control values -> plain hashtable for Test-TarStartFormInput.
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)][hashtable]$Ui)
    $ci = $Ui.Certificate.SelectedIndex
    $thumb = if ($ci -ge 0 -and $Ui.Certificate.Text -eq [string]$Ui.Certificate.Items[$ci]) { $Ui.CertRows[$ci].Id } else { $Ui.Certificate.Text }
    $pi = $Ui.Profile.SelectedIndex
    @{
        SelectedProfile = $(if ($pi -gt 0) { $Ui.Profiles[$pi - 1].Name } else { '' })
        Tenant = $Ui.Tenant.Text; AuthMode = $(if ($Ui.AppOnly.Checked) { 'AppOnly' } else { 'Delegated' })
        ClientId = $Ui.ClientId.Text; CertificateThumbprint = $thumb; PnPClientId = $Ui.PnPClientId.Text
        SaveProfile = $Ui.Save.Checked; ProfileName = $Ui.ProfileName.Text
        IncludeSharePoint = $Ui.SharePoint.Checked; IncludeUserAssociations = $Ui.Associations.Checked; SelectChannels = $Ui.Channels.Checked
        OutputPath = $Ui.Output.Text; ConnectionOnly = $Ui.ConnectionOnly
    }
}

function Invoke-TarStartFormCheck {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ui)
    $Ui.Form.UseWaitCursor = $true
    try {
        $t = Resolve-TarTenantOrPassThrough -Tenant $Ui.Tenant.Text
        $Ui.TenantStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        $Ui.TenantStatus.Text = "OK: $($t.Query) = tenant $($t.TenantId)$(if ($t.Region) { " (region $($t.Region))" })"
        if (-not $Ui.ProfileName.Text) { $Ui.ProfileName.Text = ($t.Query -split '\.')[0] }
    } catch {
        $Ui.TenantStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $Ui.TenantStatus.Text = $_.Exception.Message
    } finally { $Ui.Form.UseWaitCursor = $false }
}

function Invoke-TarStartFormSubmit {
    # Start/Save button: validate; on errors keep the window open and list them, otherwise close with OK.
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][hashtable]$Ui)
    $Ui.Form.UseWaitCursor = $true
    try {
        $r = Test-TarStartFormInput -State (Get-TarStartFormState -Ui $Ui)
        if ($r.Errors.Count) { $Ui.Error.Text = ((@($r.Errors.Values)) -join [Environment]::NewLine); return $false }
        $Ui.Error.Text = ''
        $Ui.Result = $r
        $Ui.Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $true
    } catch { $Ui.Error.Text = $_.Exception.Message; $false }
    finally { $Ui.Form.UseWaitCursor = $false }
}

function Show-TarStartForm {
    # Shows the start screen. Returns the validated result (see Test-TarStartFormInput) or $null when cancelled.
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([hashtable]$Defaults = @{}, [switch]$ConnectionOnly)
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $ui = New-TarStartForm -Profiles @(Read-TarProfileStore) -Certificates @(Get-TarCertificateCandidate | Where-Object { $_.HasPrivateKey -and $_.Valid }) -Defaults $Defaults -ConnectionOnly:$ConnectionOnly
    try {
        Write-TarLog 'Start screen opened (it may be behind other windows).' -Level DEBUG
        $answer = $ui.Form.ShowDialog()
        if ($answer -ne [System.Windows.Forms.DialogResult]::OK -or -not $ui.Result) { return }
        $ui.Result
    } finally { $ui.Form.Dispose() }
}
