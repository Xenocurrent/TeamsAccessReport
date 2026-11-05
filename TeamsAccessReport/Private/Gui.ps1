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
    # Short display text (no 'CN=' prefix) and a wider drop-down so the store is never cut off.
    foreach ($c in $ui.CertRows) { $null = $ui.Certificate.Items.Add("$($c.Subject -replace '^CN=', '')  (expires $($c.Expires), $($c.Store))") }
    $ui.Certificate.DropDownWidth = [Math]::Max($textWidth, 560)
    Add-TarFormRow $tc 'Certificate' $ui.Certificate
    $ui.PnPClientId = [System.Windows.Forms.TextBox]@{ Width = $textWidth }
    $ui.PnPCreate = [System.Windows.Forms.Button]@{ Text = 'Create...'; AutoSize = $true }
    Add-TarFormRow $tc 'PnP app ID' $ui.PnPClientId $ui.PnPCreate
    $ui.PnPHint = [System.Windows.Forms.Label]@{ AutoSize = $true; MaximumSize = [System.Drawing.Size]::new($textWidth + 100, 0); ForeColor = [System.Drawing.SystemColors]::GrayText }
    Add-TarFormRow $tc '' $ui.PnPHint
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
    $ui.PnPCreate.Add_Click({ $u = Get-TarFormUi $this; try { Invoke-TarStartFormPnPCreate -Ui $u } catch { $u.Error.Text = $_.Exception.Message } })
    $ui.SharePoint.Add_CheckedChanged({ Update-TarStartFormState -Ui (Get-TarFormUi $this) })
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
    $Ui.PnPHint.ForeColor = [System.Drawing.SystemColors]::GrayText
    $Ui.PnPHint.Text = if ($app) { 'Not needed for app-only: the app above is used for SharePoint too.' }
    elseif ($Ui.PnPClientId.Text.Trim()) { 'Used for SharePoint with delegated sign-in.' }
    elseif ($Ui.SharePoint.Checked) { "Required for SharePoint with delegated sign-in. Click 'Create...' for a guided setup." }
    else { "Only needed if you include SharePoint. 'Create...' sets it up." }
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

# --- Multi-select list with check boxes (teams, channels) ---------------------------------------------------

function New-TarCheckListForm {
    # Builds (does not show) a check-box list: filter, Select all / Clear, "n of m selected", OK enabled when n > 0.
    # Check state lives in $ui.Checked (HashSet of Id), so filtering never loses a selection.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates in-memory UI objects only.')]
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string[]]$Columns, [string[]]$CheckedId = @())
    $ui = @{ Rows = @($Rows); Columns = $Columns; Checked = [System.Collections.Generic.HashSet[string]]::new([string[]]@($CheckedId)); Syncing = $false }
    # Cell texts are computed once: filtering thousands of teams per keystroke must not re-read every property.
    $ui.Cells = @(foreach ($r in $ui.Rows) { , [string[]]@($Columns | ForEach-Object { [string](Get-TarProperty $r $_) }) })
    $form = [System.Windows.Forms.Form]@{
        Text = $Title; Width = 760; Height = 520; StartPosition = 'CenterScreen'; TopMost = $true; MinimizeBox = $false
        AutoScaleMode = 'Dpi'; Font = [System.Drawing.SystemFonts]::MessageBoxFont; Padding = [System.Windows.Forms.Padding]::new(8)
    }
    $ui.Form = $form; $form.Tag = $ui
    $top = [System.Windows.Forms.FlowLayoutPanel]@{ Dock = 'Top'; AutoSize = $true; WrapContents = $false }
    $top.Controls.Add([System.Windows.Forms.Label]@{ Text = 'Filter'; AutoSize = $true; Margin = [System.Windows.Forms.Padding]::new(3, 7, 6, 3) })
    $ui.Filter = [System.Windows.Forms.TextBox]@{ Width = 300 }
    $ui.All = [System.Windows.Forms.Button]@{ Text = 'Select all'; AutoSize = $true }
    $ui.None = [System.Windows.Forms.Button]@{ Text = 'Clear'; AutoSize = $true }
    $top.Controls.AddRange(@($ui.Filter, $ui.All, $ui.None))
    $ui.List = [System.Windows.Forms.ListView]@{ Dock = 'Fill'; View = 'Details'; CheckBoxes = $true; FullRowSelect = $true; GridLines = $true; MultiSelect = $false; HideSelection = $false }
    foreach ($c in $Columns) { $null = $ui.List.Columns.Add($c) }
    $bottom = [System.Windows.Forms.TableLayoutPanel]@{ Dock = 'Bottom'; AutoSize = $true; ColumnCount = 2 }
    $null = $bottom.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    $null = $bottom.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    $ui.Count = [System.Windows.Forms.Label]@{ AutoSize = $true; Anchor = 'Left' }
    $buttons = [System.Windows.Forms.FlowLayoutPanel]@{ AutoSize = $true; WrapContents = $false; Anchor = 'Right' }
    $ui.Ok = [System.Windows.Forms.Button]@{ Text = 'OK'; AutoSize = $true; DialogResult = 'OK' }
    $ui.Cancel = [System.Windows.Forms.Button]@{ Text = 'Cancel'; AutoSize = $true; DialogResult = 'Cancel' }
    $buttons.Controls.AddRange(@($ui.Ok, $ui.Cancel))
    $bottom.Controls.Add($ui.Count, 0, 0); $bottom.Controls.Add($buttons, 1, 0)
    # Fill first, then the docked edges (WinForms docks in reverse z-order).
    $form.Controls.AddRange(@($ui.List, $top, $bottom))
    $form.AcceptButton = $ui.Ok; $form.CancelButton = $ui.Cancel

    $ui.Filter.Add_TextChanged({ Update-TarCheckListView -Ui (Get-TarFormUi $this) })
    $ui.All.Add_Click({ $u = Get-TarFormUi $this; Set-TarCheckListAll -Ui $u -Checked $true })
    $ui.None.Add_Click({ $u = Get-TarFormUi $this; Set-TarCheckListAll -Ui $u -Checked $false })
    $ui.List.Add_ItemChecked({
            $u = Get-TarFormUi $this
            if ($u.Syncing) { return }
            $id = [string]$_.Item.Tag
            if ($_.Item.Checked) { $null = $u.Checked.Add($id) } else { $null = $u.Checked.Remove($id) }
            Update-TarCheckListCount -Ui $u
        })
    # A click anywhere on the row toggles it (not only on the small box).
    $ui.List.Add_MouseClick({
            $hit = $this.HitTest($_.Location)
            if ($hit.Item -and $hit.Location -ne [System.Windows.Forms.ListViewHitTestLocations]::StateImage) { $hit.Item.Checked = -not $hit.Item.Checked }
        })
    Update-TarCheckListView -Ui $ui
    $ui
}

function Update-TarCheckListView {
    # Shows rows whose visible columns contain the filter text (case-insensitive); keeps check state.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Sets in-memory UI state only.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ui)
    $needle = $Ui.Filter.Text.Trim()
    $Ui.Syncing = $true
    $Ui.List.BeginUpdate()
    try {
        $Ui.List.Items.Clear()
        $items = [System.Collections.Generic.List[System.Windows.Forms.ListViewItem]]::new()
        for ($i = 0; $i -lt $Ui.Rows.Count; $i++) {
            $r = $Ui.Rows[$i]; $values = $Ui.Cells[$i]
            if ($needle -and -not @($values | Where-Object { $_.IndexOf($needle, [StringComparison]::OrdinalIgnoreCase) -ge 0 }).Count) { continue }
            $item = [System.Windows.Forms.ListViewItem]::new($values)
            $item.Tag = [string]$r.Id
            $item.Checked = $Ui.Checked.Contains([string]$r.Id)
            $items.Add($item)
        }
        $Ui.List.Items.AddRange($items.ToArray())
        $Ui.List.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
    } finally { $Ui.List.EndUpdate(); $Ui.Syncing = $false }
    Update-TarCheckListCount -Ui $Ui
}

function Set-TarCheckListAll {
    # Select all / Clear act on the visible (filtered) rows only.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Sets in-memory UI state only.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ui, [Parameter(Mandatory)][bool]$Checked)
    foreach ($item in $Ui.List.Items) {
        if ($Checked) { $null = $Ui.Checked.Add([string]$item.Tag) } else { $null = $Ui.Checked.Remove([string]$item.Tag) }
    }
    Update-TarCheckListView -Ui $Ui
}

function Update-TarCheckListCount {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Sets in-memory UI state only.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ui)
    $hidden = $Ui.Rows.Count - $Ui.List.Items.Count
    $Ui.Count.Text = "$($Ui.Checked.Count) of $($Ui.Rows.Count) selected" + $(if ($hidden) { " ($hidden hidden by filter)" } else { '' })
    $Ui.Ok.Enabled = $Ui.Checked.Count -gt 0
}

function Show-TarCheckList {
    # Emits the checked rows (in their original order), nothing when cancelled.
    [CmdletBinding()]
    [OutputType([psobject])]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string[]]$Columns, [string[]]$CheckedId = @())
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $ui = New-TarCheckListForm -Rows $Rows -Title $Title -Columns $Columns -CheckedId $CheckedId
    try {
        if ($ui.Form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $Rows | Where-Object { $ui.Checked.Contains([string]$_.Id) }
    } finally { $ui.Form.Dispose() }
}

# --- PnP app guide ------------------------------------------------------------------------------------------

function Get-TarPnPGuideText {
    # Pure: the guide text (tested on every platform).
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$Tenant, [Parameter(Mandatory)][bool]$PnPInstalled)
    $nl = [Environment]::NewLine
    $lines = @(
        'Why: SharePoint scanning with delegated sign-in uses PnP PowerShell, and PnP requires its own Entra app'
        '(mandatory since 9 September 2024). App-only does not need it - the certificate app already covers SharePoint.'
        ''
        "What 'Create app now' does in $($Tenant):"
        '  1. Opens a PowerShell window and a browser sign-in.'
        "  2. Registers the app 'TeamsAccessReport-PnP' with ONE delegated permission: SharePoint AllSites.FullControl"
        '     (needed to read permissions), redirect URI http://localhost. No secrets, no certificate.'
        '  3. Opens the admin-consent page - accept it (an admin must consent).'
        '  4. Fills in the app ID here; it is saved with the profile when you click Start.'
        ''
        'Requirements: an account that may register apps (for example Application Administrator or Cloud Application'
        'Administrator) and PnP.PowerShell installed.'
    )
    if (-not $PnPInstalled) { $lines += @('', 'PnP.PowerShell is NOT installed. Run in PowerShell, then click Create again:', '  Install-Module PnP.PowerShell -Scope CurrentUser') }
    $lines -join $nl
}

function New-TarPnPGuideForm {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates in-memory UI objects only.')]
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)][string]$Tenant, [Parameter(Mandatory)][bool]$PnPInstalled)
    $ui = @{}
    $form = [System.Windows.Forms.Form]@{
        Text = 'PnP app for SharePoint (delegated)'; AutoSize = $true; AutoSizeMode = 'GrowAndShrink'; FormBorderStyle = 'FixedDialog'
        MaximizeBox = $false; MinimizeBox = $false; StartPosition = 'CenterParent'; TopMost = $true; AutoScaleMode = 'Dpi'
        Font = [System.Drawing.SystemFonts]::MessageBoxFont; Padding = [System.Windows.Forms.Padding]::new(12)
    }
    $ui.Form = $form; $form.Tag = $ui
    $root = [System.Windows.Forms.FlowLayoutPanel]@{ FlowDirection = 'TopDown'; AutoSize = $true; WrapContents = $false; Dock = 'Fill' }
    $ui.Text = [System.Windows.Forms.Label]@{ AutoSize = $true; MaximumSize = [System.Drawing.Size]::new(640, 0); Text = (Get-TarPnPGuideText -Tenant $Tenant -PnPInstalled $PnPInstalled) }
    $buttons = [System.Windows.Forms.FlowLayoutPanel]@{ FlowDirection = 'RightToLeft'; AutoSize = $true; WrapContents = $false; Anchor = 'Right'; Margin = [System.Windows.Forms.Padding]::new(0, 12, 0, 0) }
    $ui.Create = [System.Windows.Forms.Button]@{ Text = 'Create app now'; AutoSize = $true; DialogResult = 'OK'; Enabled = $PnPInstalled }
    $ui.Docs = [System.Windows.Forms.Button]@{ Text = 'Open PnP documentation'; AutoSize = $true }
    $ui.Cancel = [System.Windows.Forms.Button]@{ Text = 'Cancel'; AutoSize = $true; DialogResult = 'Cancel' }
    $buttons.Controls.AddRange(@($ui.Cancel, $ui.Create, $ui.Docs))
    $root.Controls.AddRange(@($ui.Text, $buttons))
    $form.Controls.Add($root)
    $form.AcceptButton = $ui.Create; $form.CancelButton = $ui.Cancel
    $ui.Docs.Add_Click({ Start-Process 'https://pnp.github.io/powershell/articles/registerapplication.html' })
    $ui
}

function Invoke-TarStartFormPnPCreate {
    # 'Create...' button: guide -> registration in a separate window -> fill the field.
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ui)
    if (-not $Ui.Tenant.Text.Trim()) { $Ui.Error.Text = 'Enter the tenant first.'; return }
    $tenant = (Resolve-TarTenantOrPassThrough -Tenant $Ui.Tenant.Text).Query
    $guide = New-TarPnPGuideForm -Tenant $tenant -PnPInstalled (Test-TarPnPInstalled)
    try { if ($guide.Form.ShowDialog($Ui.Form) -ne [System.Windows.Forms.DialogResult]::OK) { return } } finally { $guide.Form.Dispose() }
    $Ui.Form.UseWaitCursor = $true
    $Ui.PnPHint.ForeColor = [System.Drawing.SystemColors]::ControlText
    $Ui.PnPHint.Text = 'Registering... finish the sign-in and the consent in the browser / PowerShell window.'
    $Ui.Form.Enabled = $false
    try {
        $id = New-TarPnPApp -Tenant $tenant -WaitCallback { [System.Windows.Forms.Application]::DoEvents() } -Confirm:$false
        $Ui.PnPClientId.Text = $id
        Update-TarStartFormState -Ui $Ui
        $Ui.PnPHint.ForeColor = [System.Drawing.Color]::DarkGreen
        $Ui.PnPHint.Text = "Created TeamsAccessReport-PnP ($id). It is saved with the profile when you click Start."
        $Ui.Error.Text = ''
    } catch {
        Update-TarStartFormState -Ui $Ui
        $Ui.Error.Text = $_.Exception.Message
    } finally { $Ui.Form.Enabled = $true; $Ui.Form.UseWaitCursor = $false; $Ui.Form.Activate() }
}
