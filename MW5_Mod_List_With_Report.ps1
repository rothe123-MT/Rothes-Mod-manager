Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =========================
# Canonical asset keys (manifest /Game/... vs pak .../Content/Game/...)
# =========================
function Get-CanonicalAssetKey([string]$raw) {
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    $s = $raw.Trim().Trim('"').Replace('\', '/').ToLowerInvariant()
    $s = $s -replace '^\./+', ''

    # mod.json manifest: /Game/...
    if ($s -match '^/game/(.+)$') {
        $rest = $matches[1] -replace '\.(uasset|uexp|umap)$', ''
        return ('game/' + $rest)
    }

    # UnrealPak listing: .../content/game/...
    if ($s -match '/content/game/(.+)$') {
        $rest = $matches[1] -replace '\.(uasset|uexp|umap)$', ''
        return ('game/' + $rest)
    }

    # Lines that are already game/... without /content/
    if ($s -match '(?:^|/)game/(.+\.(uasset|uexp|umap))') {
        $full = $matches[1]
        $rest = $full -replace '\.(uasset|uexp|umap)$', ''
        return ('game/' + $rest)
    }

    return $null
}

function Add-AssetOwner($assetMap, [string]$key, $mod) {
    if (-not $key) { return }
    if (-not $assetMap.ContainsKey($key)) { $assetMap[$key] = @() }
    foreach ($m in $assetMap[$key]) {
        if ($m.FolderKey -eq $mod.FolderKey) { return }
    }
    $assetMap[$key] += $mod
}

function Format-DisplayAssetPath([string]$canonicalKey) {
    if ($canonicalKey -match '^game/(.+)$') {
        return ('.../Game/' + $matches[1] + '.uasset')
    }
    if ($canonicalKey) {
        return ('.../' + $canonicalKey + '.uasset')
    }
    return '.../unknown.uasset'
}

function Build-ConflictViewText($forMod, $assetMap) {
    $fk = $forMod.FolderKey
    $otherByKey = @{}
    foreach ($assetKey in $assetMap.Keys) {
        $owners = $assetMap[$assetKey]
        if ($owners.Count -le 1) { continue }
        $hasM = $false
        foreach ($o in $owners) { if ($o.FolderKey -eq $fk) { $hasM = $true; break } }
        if (-not $hasM) { continue }
        foreach ($o in $owners) {
            if ($o.FolderKey -ne $fk) { $otherByKey[$o.FolderKey] = $o }
        }
    }
    $others = @($otherByKey.Values | Sort-Object { $_.Name })
    if ($others.Count -eq 0) { return '(No conflicting enabled mods.)' }

    $out = New-Object System.Collections.Generic.List[string]
    foreach ($o in $others) {
        $whereOWins = New-Object System.Collections.Generic.List[string]
        $whereMyModWins = New-Object System.Collections.Generic.List[string]
        $whereOtherModWins = New-Object System.Collections.Generic.List[string]
        foreach ($assetKey in $assetMap.Keys) {
            $owners = $assetMap[$assetKey]
            if ($owners.Count -le 1) { continue }
            $hasM = $false
            $hasO = $false
            foreach ($x in $owners) {
                if ($x.FolderKey -eq $fk) { $hasM = $true }
                if ($x.FolderKey -eq $o.FolderKey) { $hasO = $true }
            }
            if (-not $hasM -or -not $hasO) { continue }
            $sorted = @($owners | Sort-Object LoadOrder)
            $winner = $sorted[-1]
            $disp = Format-DisplayAssetPath $assetKey
            if ($winner.FolderKey -eq $o.FolderKey) {
                $whereOWins.Add($disp)
            } elseif ($winner.FolderKey -eq $fk) {
                $whereMyModWins.Add($disp)
            } else {
                $whereOtherModWins.Add(($disp + '  (load-order winner: ' + $winner.Name + ')'))
            }
        }
        $whereOWins = @($whereOWins | Sort-Object -Unique)
        $whereMyModWins = @($whereMyModWins | Sort-Object -Unique)
        $whereOtherModWins = @($whereOtherModWins | Sort-Object -Unique)

        $warn = [char]0x26A0
        if ($whereOWins.Count -gt 0) {
            $out.Add(($warn + ' "' + $o.Name + '" is overriding mod content'))
            foreach ($line in $whereOWins) { $out.Add(('   ' + $line)) }
            $out.Add('')
        }
        if ($whereMyModWins.Count -gt 0) {
            $out.Add(($warn + ' "' + $forMod.Name + '" wins load order over "' + $o.Name + '" for:'))
            foreach ($line in $whereMyModWins) { $out.Add(('   ' + $line)) }
            $out.Add('')
        }
        if ($whereOtherModWins.Count -gt 0) {
            $out.Add(($warn + ' Another enabled mod wins load order for assets you share with "' + $o.Name + '":'))
            foreach ($line in $whereOtherModWins) { $out.Add(('   ' + $line)) }
            $out.Add('')
        }
    }
    $text = ($out -join "`r`n").TrimEnd()
    if ([string]::IsNullOrWhiteSpace($text)) { return '(No overlapping assets with this mod.)' }
    return $text
}

function Show-ConflictViewForm($title, $bodyText, $ownerForm) {
    $vf = New-Object System.Windows.Forms.Form
    $vf.Text = $title
    $vf.Size = New-Object System.Drawing.Size(920, 640)
    if ($ownerForm) {
        $vf.Owner = $ownerForm
        $vf.StartPosition = 'CenterParent'
    } else {
        $vf.StartPosition = 'CenterScreen'
    }
    $vf.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $vf.MinimizeBox = $false
    $vf.MaximizeBox = $true

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Fill'
    $panel.Padding = New-Object System.Windows.Forms.Padding(10)
    $panel.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)

    $rtb = New-Object System.Windows.Forms.RichTextBox
    $rtb.Dock = 'Fill'
    $rtb.ReadOnly = $true
    $rtb.BorderStyle = 'None'
    $rtb.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $rtb.ForeColor = [System.Drawing.Color]::White
    $rtb.Font = New-Object System.Drawing.Font('Consolas', 10)
    $rtb.DetectUrls = $false

    $warnCh = [char]0x26A0
    $lines = $bodyText -split "`r`n|`n"
    foreach ($ln in $lines) {
        if ($ln -match ('^' + [regex]::Escape($warnCh) + '\s*"(.+)"\s+is overriding mod content')) {
            $rtb.SelectionColor = [System.Drawing.Color]::Gold
            $rtb.AppendText(([char]0x26A0).ToString() + ' ')
            $rtb.SelectionColor = [System.Drawing.Color]::White
            $rtb.AppendText('"')
            $rtb.SelectionColor = [System.Drawing.Color]::LightSkyBlue
            $rtb.AppendText($matches[1])
            $rtb.SelectionColor = [System.Drawing.Color]::White
            $rtb.AppendText('" is overriding mod content')
        } elseif ($ln -match ('^' + [regex]::Escape($warnCh) + '\s*"(.+)"\s+wins load order over\s+"(.+)"\s+for:')) {
            $rtb.SelectionColor = [System.Drawing.Color]::Gold
            $rtb.AppendText(([char]0x26A0).ToString() + ' ')
            $rtb.SelectionColor = [System.Drawing.Color]::White
            $rtb.AppendText('"')
            $rtb.SelectionColor = [System.Drawing.Color]::LightGreen
            $rtb.AppendText($matches[1])
            $rtb.SelectionColor = [System.Drawing.Color]::White
            $rtb.AppendText('" wins load order over "')
            $rtb.SelectionColor = [System.Drawing.Color]::LightSkyBlue
            $rtb.AppendText($matches[2])
            $rtb.SelectionColor = [System.Drawing.Color]::White
            $rtb.AppendText('" for:')
        } elseif ($ln -match ('^' + [regex]::Escape($warnCh) + '\s+Another enabled mod wins load order for assets you share with\s+"(.+)":')) {
            $rtb.SelectionColor = [System.Drawing.Color]::Gold
            $rtb.AppendText(([char]0x26A0).ToString() + ' ')
            $rtb.SelectionColor = [System.Drawing.Color]::FromArgb(220, 220, 160)
            $rtb.AppendText('Another enabled mod wins load order for assets you share with "')
            $rtb.SelectionColor = [System.Drawing.Color]::LightSkyBlue
            $rtb.AppendText($matches[1])
            $rtb.SelectionColor = [System.Drawing.Color]::FromArgb(220, 220, 160)
            $rtb.AppendText('":')
        } elseif ($ln -match '^\s{3}\.\.\.') {
            $rtb.SelectionColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
            $rtb.AppendText($ln)
        } else {
            $rtb.SelectionColor = [System.Drawing.Color]::White
            $rtb.AppendText($ln)
        }
        $rtb.AppendText("`n")
    }

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = 'Close'
    $btn.Dock = 'Bottom'
    $btn.Height = 36
    $btn.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatStyle = 'Flat'
    $btn.Add_Click({ $vf.Close() })

    $panel.Controls.Add($rtb)
    $vf.Controls.Add($btn)
    $vf.Controls.Add($panel)
    $vf.ShowDialog() | Out-Null
}

# =========================
# CONFIG & UnrealPak
# =========================
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$UnrealPak = Join-Path $ScriptDir "UnrealPak.exe"
$ConfigFile = Join-Path $ScriptDir "MW5_Mod_Settings.json"
$ReportFile = Join-Path $ScriptDir "MW5_Mod_List.txt"
$PlaylistsFile = Join-Path $ScriptDir "MW5_Mod_Playlists.json"
$UiSettingsFile = Join-Path $ScriptDir "MW5_UI_Settings.json"

function Load-UiSettings {
    $s = $null
    if (Test-Path $UiSettingsFile) {
        try { $s = Get-Content $UiSettingsFile -Raw | ConvertFrom-Json } catch {}
    }
    if (-not $s) {
        $s = [PSCustomObject]@{
            theme = @{
                background = "#F0F0F0"
                foreground = "#111111"
                accent     = "#2D89EF"
                gridAlt    = "#E8E8E8"
            }
            logoPath = ""
        }
    }
    return $s
}

function Save-UiSettings($s) {
    try { ($s | ConvertTo-Json -Depth 10) | Set-Content $UiSettingsFile -Force -Encoding UTF8 } catch {}
}

function ColorFromHex([string]$hex, [System.Drawing.Color]$fallback) {
    try {
        if ([string]::IsNullOrWhiteSpace($hex)) { return $fallback }
        $h = $hex.Trim()
        if ($h.StartsWith("#")) { $h = $h.Substring(1) }
        if ($h.Length -eq 6) {
            $r = [Convert]::ToInt32($h.Substring(0,2),16)
            $g = [Convert]::ToInt32($h.Substring(2,2),16)
            $b = [Convert]::ToInt32($h.Substring(4,2),16)
            return [System.Drawing.Color]::FromArgb($r,$g,$b)
        }
    } catch {}
    return $fallback
}

function HexFromColor([System.Drawing.Color]$c) {
    return ('#{0:X2}{1:X2}{2:X2}' -f $c.R, $c.G, $c.B)
}

function Apply-ThemeToControl($ctrl, $colors) {
    if (-not $ctrl) { return }
    try {
        if ($ctrl -is [System.Windows.Forms.DataGridView]) {
            $ctrl.BackgroundColor = $colors.Background
            $ctrl.GridColor = $colors.Accent
            $ctrl.DefaultCellStyle.BackColor = $colors.Background
            $ctrl.DefaultCellStyle.ForeColor = $colors.Foreground
            $ctrl.AlternatingRowsDefaultCellStyle.BackColor = $colors.GridAlt
            $ctrl.ColumnHeadersDefaultCellStyle.BackColor = $colors.Accent
            $ctrl.ColumnHeadersDefaultCellStyle.ForeColor = $colors.Background
            $ctrl.EnableHeadersVisualStyles = $false
        } elseif ($ctrl -is [System.Windows.Forms.Button]) {
            $ctrl.BackColor = $colors.Accent
            $ctrl.ForeColor = $colors.Background
            $ctrl.FlatStyle = 'Flat'
        } elseif ($ctrl -is [System.Windows.Forms.RichTextBox] -or $ctrl -is [System.Windows.Forms.TextBox]) {
            $ctrl.BackColor = $colors.Background
            $ctrl.ForeColor = $colors.Foreground
        } elseif ($ctrl -is [System.Windows.Forms.Form] -or $ctrl -is [System.Windows.Forms.TabPage] -or $ctrl -is [System.Windows.Forms.Panel] -or $ctrl -is [System.Windows.Forms.TabControl]) {
            $ctrl.BackColor = $colors.Background
            $ctrl.ForeColor = $colors.Foreground
        } elseif ($ctrl -is [System.Windows.Forms.Label]) {
            $ctrl.BackColor = [System.Drawing.Color]::Transparent
            $ctrl.ForeColor = $colors.Foreground
        } else {
            if ($ctrl.PSObject.Properties.Name -contains "BackColor") { $ctrl.BackColor = $colors.Background }
            if ($ctrl.PSObject.Properties.Name -contains "ForeColor") { $ctrl.ForeColor = $colors.Foreground }
        }
    } catch {}

    try {
        foreach ($child in $ctrl.Controls) { Apply-ThemeToControl $child $colors }
    } catch {}
}

function Load-Playlists {
    $items = @()
    Write-Host "Playlists file path: $PlaylistsFile" -ForegroundColor Yellow
    Write-Host "Script directory: $ScriptDir" -ForegroundColor Yellow
    
    # Ensure script directory exists
    if (-not (Test-Path $ScriptDir)) {
        Write-Host "Script directory not found: $ScriptDir" -ForegroundColor Red
        try {
            New-Item -ItemType Directory -Path $ScriptDir -Force | Out-Null
            Write-Host "Created script directory: $ScriptDir" -ForegroundColor Green
        } catch {
            Write-Host "Failed to create script directory: $ScriptDir" -ForegroundColor Red
        }
    }
    
    if (Test-Path $PlaylistsFile) {
        try {
            $data = Get-Content $PlaylistsFile -Raw | ConvertFrom-Json
            if ($data.playlists) { $items = @($data.playlists) }
            elseif ($data -is [System.Array]) { $items = @($data) }
        } catch {
            Write-Host "Error loading playlists file: $PlaylistsFile" -ForegroundColor Red
        }
    } else {
        Write-Host "Playlists file not found: $PlaylistsFile" -ForegroundColor Red
        # Create empty playlists file
        try {
            $emptyData = @{ playlists = @() } | ConvertTo-Json -Depth 20
            $emptyData | Set-Content $PlaylistsFile -Force -Encoding UTF8
            Write-Host "Created empty playlists file: $PlaylistsFile" -ForegroundColor Green
        } catch {
            Write-Host "Failed to create playlists file: $PlaylistsFile" -ForegroundColor Red
        }
    }
    # Filter out null/empty entries if JSON got partially edited/corrupted
    return @($items | Where-Object { $null -ne $_ })
}

function Save-Playlists($items) {
    $payload = @{ playlists = @($items) } | ConvertTo-Json -Depth 20
    $payload | Set-Content $PlaylistsFile -Force -Encoding UTF8
}

function Normalize-PlaylistName([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    $t = $s.Trim()
    # collapse repeated whitespace to one space
    return ([regex]::Replace($t, '\s+', ' '))
}

function Get-SafeFileName([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return "playlist" }
    $n = $name.Trim()
    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) {
        $n = $n.Replace($c, '_')
    }
    $n = [regex]::Replace($n, '\s+', ' ').Trim()
    if ($n.Length -eq 0) { return "playlist" }
    return $n
}

function Ensure-PlaylistIds($items) {
    $changed = $false
    foreach ($p in @($items)) {
        if ($null -eq $p) { continue }
        if (-not $p.id -or [string]::IsNullOrWhiteSpace([string]$p.id)) {
            try { $p | Add-Member -NotePropertyName id -NotePropertyValue ([guid]::NewGuid().ToString()) -Force } catch {}
            $changed = $true
        }
    }
    if ($changed) { Save-Playlists $items }
    return @($items)
}

$Settings = @{}
if (Test-Path $ConfigFile) {
    try { $Settings = Get-Content $ConfigFile -Raw | ConvertFrom-Json } catch {}
}

if (-not (Test-Path $UnrealPak)) {
    [System.Windows.Forms.MessageBox]::Show("UnrealPak.exe not found!", "Error", "OK", "Error")
    exit
}

# =========================
# FOLDER SELECTION
# =========================
function Pick-Folder($title, $lastPath, $defaultExample) {
    $f = New-Object System.Windows.Forms.FolderBrowserDialog
    $f.Description = $title + "`n`nDefault path example: $defaultExample"
    if ($lastPath -and (Test-Path $lastPath)) { $f.SelectedPath = $lastPath }
    if ($f.ShowDialog() -eq "OK") { return $f.SelectedPath }
    else { exit }
}

$modsPath = $Settings.LocalModsPath
if ($modsPath -and (Test-Path $modsPath)) {
    if ([System.Windows.Forms.MessageBox]::Show("Use previous Local Mods Folder?`n`n$modsPath", "Reuse Folder?", "YesNo", "Question") -eq "No") {
        $modsPath = Pick-Folder "Select your Local Mods Folder" $modsPath "D:\SteamLibrary\steamapps\common\MechWarrior 5 Mercenaries\MW5Mercs\mods"
    }
} else {
    $modsPath = Pick-Folder "Select your Local Mods Folder" $modsPath "D:\SteamLibrary\steamapps\common\MechWarrior 5 Mercenaries\MW5Mercs\mods"
}

$workshopPath = $Settings.WorkshopPath
if ($workshopPath -and (Test-Path $workshopPath)) {
    if ([System.Windows.Forms.MessageBox]::Show("Use previous Steam Workshop Folder?`n`n$workshopPath", "Reuse Folder?", "YesNo", "Question") -eq "No") {
        $workshopPath = Pick-Folder "Select your Steam Workshop Folder" $workshopPath "D:\SteamLibrary\steamapps\workshop\content\784080"
    }
} else {
    $workshopPath = Pick-Folder "Select your Steam Workshop Folder" $workshopPath "D:\SteamLibrary\steamapps\workshop\content\784080"
}

$Settings.LocalModsPath = $modsPath
$Settings.WorkshopPath = $workshopPath
$Settings | ConvertTo-Json | Set-Content $ConfigFile -Force

# =========================
# LOAD modlist.json
# =========================
$ModListFile = Join-Path $modsPath "modlist.json"
$EnabledMods = @{}
$GameVersion = $null

if (Test-Path $ModListFile) {
    try {
        $modlist = Get-Content $ModListFile -Raw | ConvertFrom-Json
        if ($modlist.gameVersion) {
            $GameVersion = $modlist.gameVersion
        }
        if ($modlist.modStatus) {
            $modlist.modStatus.PSObject.Properties | ForEach-Object {
                $EnabledMods[$_.Name] = [bool]$_.Value.bEnabled
            }
        }
    } catch { Write-Host "Warning: Could not parse modlist.json" -ForegroundColor Yellow }
}

# =========================
# GET MODS
# =========================
function Get-Mods($path) {
    $mods = @()
    if (-not (Test-Path $path)) { return $mods }
    
    # Determine if this is workshop path
    $isWorkshop = $path -eq $workshopPath

    Get-ChildItem $path -Directory | ForEach-Object {
        $folderName = $_.Name
        $modFolder = $_.FullName
        $jsonFile = Join-Path $modFolder "mod.json"
        $pakFiles = @(Get-ChildItem $modFolder -Recurse -Filter *.pak -File -ErrorAction SilentlyContinue)

        if ($pakFiles.Count -eq 0 -or -not (Test-Path $jsonFile)) { return }

        $load = 0
        $displayName = $folderName
        $manifest = @()

        try {
            $json = Get-Content $jsonFile -Raw | ConvertFrom-Json

            if ($null -ne $json.loadOrder) { $load = [int]$json.loadOrder }
            elseif ($null -ne $json.defaultLoadOrder) { $load = [int]$json.defaultLoadOrder }

            if ($json.displayName -and $json.displayName.Trim() -ne "") {
                $displayName = $json.displayName.Trim()
            }

            if ($json.manifest) {
                foreach ($entry in $json.manifest) {
                    if ($entry -and $entry.ToString().Trim() -ne "") {
                        $manifest += $entry.ToString().Trim()
                    }
                }
            }
        } catch {}

        $isEnabled = if ($EnabledMods.ContainsKey($folderName)) { $EnabledMods[$folderName] } else { $false }

        $mods += [PSCustomObject]@{
            Name        = $displayName
            FolderKey   = $folderName
            ModJsonPath = $jsonFile
            ModFolder   = $modFolder
            LoadOrder   = $load
            OriginalLoadOrder = $load  # Store original load order
            Enabled     = $isEnabled
            PakFiles    = $pakFiles
            Manifest    = $manifest
            IsWorkshop  = $isWorkshop  # Track if this is a workshop mod
        }
    }
    return $mods
}

$mods = @()
$mods += Get-Mods $modsPath
$mods += Get-Mods $workshopPath

if ($mods.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No valid mods found.", "No Mods")
    exit
}

# =========================
# CONFLICT SCANNING (enabled mods only - manifest + all .pak UnrealPak -List)
# =========================
$assetMap = @{}

foreach ($mod in ($mods | Where-Object { $_.Enabled })) {
    foreach ($mPath in $mod.Manifest) {
        $key = Get-CanonicalAssetKey $mPath
        Add-AssetOwner $assetMap $key $mod
    }

    foreach ($pakFile in $mod.PakFiles) {
        try {
            $output = & $UnrealPak $pakFile.FullName -List 2>$null
            foreach ($line in $output) {
                if ($line -match "^\s*(.+\.(uasset|uexp|umap))") {
                    $asset = $matches[1].Trim()
                    $key = Get-CanonicalAssetKey $asset
                    if (-not $key) {
                        $key = ($asset -replace '\\', '/').ToLowerInvariant()
                    }
                    Add-AssetOwner $assetMap $key $mod
                }
            }
        } catch {}
    }
}

# Per-folder conflict details - only mods that LOSE load order on a shared asset (omit winners rows)
$conflictResults = @{}
foreach ($asset in $assetMap.Keys) {
    $owners = $assetMap[$asset]
    if ($owners.Count -le 1) { continue }

    $sorted = @($owners | Sort-Object LoadOrder)
    $winner = $sorted[-1]
    foreach ($m in $owners) {
        if ($m.FolderKey -eq $winner.FolderKey) { continue }
        $otherNames = $owners | Where-Object { $_.FolderKey -ne $m.FolderKey } | ForEach-Object { $_.Name }
        if (-not $conflictResults.ContainsKey($m.FolderKey)) { $conflictResults[$m.FolderKey] = @() }
        $conflictResults[$m.FolderKey] += [PSCustomObject]@{
            Asset   = $asset
            Winner  = $winner.Name
            IsWinner = $false
            Others  = ($otherNames -join ", ")
        }
    }
}

# =========================
# FUNCTION: Create/Update Report
# =========================
function Update-ModReport {
    $w = 80
    $lineEq = ('=' * $w)
    $lineHy = ('-' * $w)
    $modList = @($mods | Sort-Object LoadOrder)
    $nTotal = $modList.Count
    $nEnabled = @($modList | Where-Object { $_.Enabled }).Count
    $nConflict = @($conflictResults.Keys).Count

    $report = New-Object System.Collections.Generic.List[string]
    $report.Add('')
    $report.Add($lineEq)
    $report.Add('  MW5 MERCENARIES - MOD LIST & CONFLICT REPORT')
    $report.Add($lineEq)
    $report.Add('')
    $report.Add(('  Generated      : {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    $report.Add(('  Total mods     : {0}' -f $nTotal))
    $report.Add(('  Enabled        : {0}' -f $nEnabled))
    $report.Add(('  Losing overlaps: {0}  (CONFLICT rows - loses load order on 1+ shared asset)' -f $nConflict))
    $report.Add('')
    $report.Add($lineHy)
    $report.Add('  HOW TO READ')
    $report.Add($lineHy)
    $report.Add('  - Scans ENABLED mods only (manifest in mod.json + UnrealPak -List per .pak).')
    $report.Add('  - CONFLICT here = this mod loses to another enabled mod on at least one file.')
    $report.Add('  - Mods that always win load order on shared files do not appear as CONFLICT.')
    $report.Add('  - Main table columns match the app (left to right). Full conflict text is under CONFLICT DETAILS.')
    $report.Add('')
    $report.Add($lineEq)
    $report.Add('  MOD LIST - same column order as the form')
    $report.Add($lineEq)
    $report.Add('')

    $cwName = 42
    $cwLoad = 10
    $cwStat = 10
    $cwEn = 9
    $cwDet = 8

    $report.Add('  ' + 'Mod name'.PadRight($cwName) + ' ' + 'Load order'.PadLeft($cwLoad) + ' ' + 'Status'.PadRight($cwStat) + ' ' + 'Enabled'.PadRight($cwEn) + ' ' + 'Details'.PadRight($cwDet))
    $report.Add('  ' + ('-' * $cwName) + ' ' + ('-' * $cwLoad) + ' ' + ('-' * $cwStat) + ' ' + ('-' * $cwEn) + ' ' + ('-' * $cwDet))

    foreach ($mod in $modList) {
        $status = if ($mod.Enabled -and $conflictResults.ContainsKey($mod.FolderKey)) { 'CONFLICT' } else { 'Clean' }
        $enabledText = if ($mod.Enabled) { 'ENABLED' } else { 'DISABLED' }
        $detailsCol = if ($mod.Enabled -and $conflictResults.ContainsKey($mod.FolderKey)) { 'View' } else { '-' }
        $nameCol = if ([string]::IsNullOrEmpty($mod.Name)) { '' } elseif ($mod.Name.Length -le $cwName) { $mod.Name } else { $mod.Name.Substring(0, $cwName - 3) + '...' }
        $loadStr = [string]$mod.LoadOrder
        $report.Add('  ' + $nameCol.PadRight($cwName) + ' ' + $loadStr.PadLeft($cwLoad) + ' ' + $status.PadRight($cwStat) + ' ' + $enabledText.PadRight($cwEn) + ' ' + $detailsCol.PadRight($cwDet))
    }

    $report.Add('')
    $report.Add($lineEq)
    $report.Add('  CONFLICT DETAILS - same text as the View window (Folder: mod folder id)')
    $report.Add($lineEq)
    $report.Add('')

    if ($nConflict -eq 0) {
        $report.Add('  (No mods lose load order on a shared asset - no conflict detail blocks.)')
        $report.Add('')
    }

    $cNum = 0
    foreach ($mod in $modList) {
        if (-not ($mod.Enabled -and $conflictResults.ContainsKey($mod.FolderKey))) { continue }
        $cNum++

        $report.Add($lineHy)
        $report.Add(('  Conflict detail {0} of {1}  |  {2}  |  Folder: {3}' -f $cNum, $nConflict, $mod.Name, $mod.FolderKey))
        $report.Add($lineHy)
        $report.Add('')
        $detail = Build-ConflictViewText $mod $assetMap
        foreach ($dl in ($detail -split "`r`n|`n")) {
            if ($dl.Length -eq 0) {
                $report.Add('')
            } else {
                $report.Add('  ' + $dl)
            }
        }
        $report.Add('')
    }

    $report.Add($lineEq)
    $report.Add(('  End of report - {0} mods listed.' -f $nTotal))
    $report.Add($lineEq)
    $report.Add('')

    ($report -join "`r`n") | Out-File -FilePath $ReportFile -Encoding utf8 -Force
    Write-Host "Report created/updated: $ReportFile" -ForegroundColor Green
}

Update-ModReport

# =========================
# UI
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "MW5 Conflict Scanner + Mod Manager"
$form.Size = New-Object System.Drawing.Size(1300, 750)
$form.StartPosition = "CenterScreen"

$ui = Load-UiSettings
$themeBg = ColorFromHex ([string]$ui.theme.background) ([System.Drawing.Color]::FromArgb(240,240,240))
$themeFg = ColorFromHex ([string]$ui.theme.foreground) ([System.Drawing.Color]::FromArgb(17,17,17))
$themeAccent = ColorFromHex ([string]$ui.theme.accent) ([System.Drawing.Color]::FromArgb(45,137,239))
$themeGridAlt = ColorFromHex ([string]$ui.theme.gridAlt) ([System.Drawing.Color]::FromArgb(232,232,232))
$themeColors = [PSCustomObject]@{ Background=$themeBg; Foreground=$themeFg; Accent=$themeAccent; GridAlt=$themeGridAlt }

$logo = New-Object System.Windows.Forms.PictureBox
$logo.Location = New-Object System.Drawing.Point(1260, 10)
$logo.SizeMode = "AutoSize"
$logo.BackColor = [System.Drawing.Color]::Transparent
try {
    if ($ui.logoPath -and (Test-Path ([string]$ui.logoPath))) { $logo.Image = [System.Drawing.Image]::FromFile([string]$ui.logoPath) }
} catch {}

$themeBtn = New-Object System.Windows.Forms.Button
$themeBtn.Text = "Theme / Logo"
$themeBtn.Size = New-Object System.Drawing.Size(140, 30)
$themeBtn.Location = New-Object System.Drawing.Point(10, 702)

# Game Version Dropdown
$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Text = "MW5 Game Version:"
$versionLabel.Size = New-Object System.Drawing.Size(120, 20)
$versionLabel.Location = New-Object System.Drawing.Point(160, 707)

$versionDropdown = New-Object System.Windows.Forms.ComboBox
$versionDropdown.Size = New-Object System.Drawing.Size(150, 30)
$versionDropdown.Location = New-Object System.Drawing.Point(285, 702)
$versionDropdown.DropDownStyle = "DropDown"  # Changed from "DropDownList" to "DropDown" to allow typing

# Add version label
$addVersionLabel = New-Object System.Windows.Forms.Label
$addVersionLabel.Text = "Type new version and press Enter"
$addVersionLabel.Size = New-Object System.Drawing.Size(200, 15)
$addVersionLabel.Location = New-Object System.Drawing.Point(440, 707)
$addVersionLabel.ForeColor = [System.Drawing.Color]::Gray

# Load and populate game versions
$GameVersionsFile = Join-Path $ScriptDir "MW5_Game_Versions.json"
$GameVersions = @()
if (Test-Path $GameVersionsFile) {
    try {
        $GameVersions = Get-Content $GameVersionsFile -Raw | ConvertFrom-Json
        if (-not ($GameVersions -is [System.Array])) { $GameVersions = @($GameVersions) }
    } catch {}
}

# Ensure current version is in the list
if ($GameVersion -and $GameVersion -notin $GameVersions) {
    $GameVersions = @($GameVersions) + @($GameVersion) | Sort-Object
}

# Add default versions if list is empty
if ($GameVersions.Count -eq 0) {
    $GameVersions = @("1.13.378", "1.12.378", "1.11.378")
}

foreach ($ver in $GameVersions) {
    $versionDropdown.Items.Add($ver) | Out-Null
}

# Select current version if available
if ($GameVersion) {
    $versionDropdown.Text = $GameVersion  # Use Text instead of SelectedIndex for DropDown mode
} elseif ($GameVersions.Count -gt 0) {
    $versionDropdown.Text = $GameVersions[0]
}

# Add version management (handle both selection and typing)
$versionDropdown.Add_TextChanged({
    $selectedVersion = $versionDropdown.Text.Trim()
    if ($selectedVersion -and $selectedVersion -ne "") {
        # Add to versions list if not already present
        if ($selectedVersion -notin $GameVersions) {
            $GameVersions += $selectedVersion
            $GameVersions = $GameVersions | Sort-Object
            # Refresh dropdown items
            $versionDropdown.Items.Clear()
            foreach ($ver in $GameVersions) {
                $versionDropdown.Items.Add($ver) | Out-Null
            }
            # Set the text back to maintain user input
            $versionDropdown.Text = $selectedVersion
        }
        
        # Save to versions file
        $GameVersions | ConvertTo-Json | Set-Content $GameVersionsFile -Force
        
        # Update modlist.json if it exists
        if (Test-Path $ModListFile) {
            try {
                $modlist = Get-Content $ModListFile -Raw | ConvertFrom-Json
                $modlist.gameVersion = $selectedVersion
                $modlist | ConvertTo-Json -Depth 10 | Set-Content $ModListFile -Force
                Write-Host "Game version updated to: $selectedVersion" -ForegroundColor Green
            } catch {
                Write-Host "Warning: Could not update game version in modlist.json" -ForegroundColor Yellow
            }
        }
    }
})

# Also handle selection from dropdown
$versionDropdown.Add_SelectedIndexChanged({
    if ($versionDropdown.SelectedItem) {
        $versionDropdown.Text = $versionDropdown.SelectedItem.ToString()
    }
})

$themeBtn.Add_Click({
        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = "Theme + Logo"
        $dlg.Size = New-Object System.Drawing.Size(520, 320)
        $dlg.StartPosition = "CenterParent"
        $dlg.Owner = $form

        $bgBtn = New-Object System.Windows.Forms.Button
        $bgBtn.Text = "Pick background"
        $bgBtn.Location = New-Object System.Drawing.Point(12, 20)
        $bgBtn.Size = New-Object System.Drawing.Size(160, 30)

        $fgBtn = New-Object System.Windows.Forms.Button
        $fgBtn.Text = "Pick foreground"
        $fgBtn.Location = New-Object System.Drawing.Point(12, 60)
        $fgBtn.Size = New-Object System.Drawing.Size(160, 30)

        $acBtn = New-Object System.Windows.Forms.Button
        $acBtn.Text = "Pick accent"
        $acBtn.Location = New-Object System.Drawing.Point(12, 100)
        $acBtn.Size = New-Object System.Drawing.Size(160, 30)

        $altBtn = New-Object System.Windows.Forms.Button
        $altBtn.Text = "Pick grid alt row"
        $altBtn.Location = New-Object System.Drawing.Point(12, 140)
        $altBtn.Size = New-Object System.Drawing.Size(160, 30)

        $setLogoBtn = New-Object System.Windows.Forms.Button
        $setLogoBtn.Text = "Set logo image..."
        $setLogoBtn.Location = New-Object System.Drawing.Point(12, 190)
        $setLogoBtn.Size = New-Object System.Drawing.Size(160, 30)

        $clearLogoBtn = New-Object System.Windows.Forms.Button
        $clearLogoBtn.Text = "Clear logo"
        $clearLogoBtn.Location = New-Object System.Drawing.Point(12, 230)
        $clearLogoBtn.Size = New-Object System.Drawing.Size(160, 30)

        $preview = New-Object System.Windows.Forms.Panel
        $preview.Location = New-Object System.Drawing.Point(200, 20)
        $preview.Size = New-Object System.Drawing.Size(290, 120)
        $preview.BorderStyle = "FixedSingle"

        $previewLbl = New-Object System.Windows.Forms.Label
        $previewLbl.Location = New-Object System.Drawing.Point(10, 10)
        $previewLbl.Size = New-Object System.Drawing.Size(260, 20)
        $previewLbl.Text = "Preview text"

        $previewBtn = New-Object System.Windows.Forms.Button
        $previewBtn.Location = New-Object System.Drawing.Point(10, 40)
        $previewBtn.Size = New-Object System.Drawing.Size(160, 30)
        $previewBtn.Text = "Accent button"

        $preview.Controls.Add($previewLbl)
        $preview.Controls.Add($previewBtn)

        $ok = New-Object System.Windows.Forms.Button
        $ok.Text = "Save"
        $ok.Location = New-Object System.Drawing.Point(390, 240)
        $ok.Size = New-Object System.Drawing.Size(100, 30)

        $cancel = New-Object System.Windows.Forms.Button
        $cancel.Text = "Cancel"
        $cancel.Location = New-Object System.Drawing.Point(280, 240)
        $cancel.Size = New-Object System.Drawing.Size(100, 30)

        $tmp = [PSCustomObject]@{ Background=$themeColors.Background; Foreground=$themeColors.Foreground; Accent=$themeColors.Accent; GridAlt=$themeColors.GridAlt; LogoPath=[string]$ui.logoPath }
        $cd = New-Object System.Windows.Forms.ColorDialog

        function Update-Preview {
            $preview.BackColor = $tmp.Background
            $previewLbl.ForeColor = $tmp.Foreground
            $previewBtn.BackColor = $tmp.Accent
            $previewBtn.ForeColor = $tmp.Background
        }
        Update-Preview

        $bgBtn.Add_Click({ $cd.Color = $tmp.Background; if ($cd.ShowDialog() -eq "OK") { $tmp.Background = $cd.Color; Update-Preview } })
        $fgBtn.Add_Click({ $cd.Color = $tmp.Foreground; if ($cd.ShowDialog() -eq "OK") { $tmp.Foreground = $cd.Color; Update-Preview } })
        $acBtn.Add_Click({ $cd.Color = $tmp.Accent; if ($cd.ShowDialog() -eq "OK") { $tmp.Accent = $cd.Color; Update-Preview } })
        $altBtn.Add_Click({ $cd.Color = $tmp.GridAlt; if ($cd.ShowDialog() -eq "OK") { $tmp.GridAlt = $cd.Color; Update-Preview } })

        $setLogoBtn.Add_Click({
                $ofd = New-Object System.Windows.Forms.OpenFileDialog
                $ofd.Title = "Select logo image"
                $ofd.Filter = "Image files (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif|All files (*.*)|*.*"
                $ofd.Multiselect = $false
                if ($ofd.ShowDialog() -ne "OK") { return }
                $tmp.LogoPath = $ofd.FileName
            })
        $clearLogoBtn.Add_Click({ $tmp.LogoPath = "" })

        $ok.Add_Click({
                $ui.theme.background = HexFromColor $tmp.Background
                $ui.theme.foreground = HexFromColor $tmp.Foreground
                $ui.theme.accent = HexFromColor $tmp.Accent
                $ui.theme.gridAlt = HexFromColor $tmp.GridAlt
                $ui.logoPath = [string]$tmp.LogoPath
                Save-UiSettings $ui

                $themeColors.Background = $tmp.Background
                $themeColors.Foreground = $tmp.Foreground
                $themeColors.Accent = $tmp.Accent
                $themeColors.GridAlt = $tmp.GridAlt

                try {
                    if ($logo.Image) { $logo.Image.Dispose() }
                    $logo.Image = $null
                    if ($ui.logoPath -and (Test-Path ([string]$ui.logoPath))) {
                        $logo.Image = [System.Drawing.Image]::FromFile([string]$ui.logoPath)
                    }
                } catch {}

                Apply-ThemeToControl $form $themeColors
                $dlg.Close()
            })
        $cancel.Add_Click({ $dlg.Close() })

        $dlg.Controls.Add($bgBtn)
        $dlg.Controls.Add($fgBtn)
        $dlg.Controls.Add($acBtn)
        $dlg.Controls.Add($altBtn)
        $dlg.Controls.Add($setLogoBtn)
        $dlg.Controls.Add($clearLogoBtn)
        $dlg.Controls.Add($preview)
        $dlg.Controls.Add($ok)
        $dlg.Controls.Add($cancel)
        $dlg.ShowDialog() | Out-Null
    })

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(10, 10)
$tabs.Size = New-Object System.Drawing.Size(1270, 690)

$tabMods = New-Object System.Windows.Forms.TabPage
$tabMods.Text = "Grid"

$tabPlaylists = New-Object System.Windows.Forms.TabPage
$tabPlaylists.Text = "Playlists"

$tabs.TabPages.Add($tabMods) | Out-Null
$tabs.TabPages.Add($tabPlaylists) | Out-Null

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(10, 10)
$grid.Size = New-Object System.Drawing.Size(1240, 590)
$grid.AutoSizeColumnsMode = "Fill"
$grid.AllowUserToAddRows = $false
$grid.ReadOnly = $false
$grid.SelectionMode = "FullRowSelect"

$grid.Columns.Add("Name", "Mod Name") | Out-Null
$grid.Columns.Add("Load", "Load Order") | Out-Null
$grid.Columns.Add("OriginalLoad", "Original Load Order") | Out-Null
$grid.Columns["OriginalLoad"].ReadOnly = $true  # Make Original Load Order column read-only
$grid.Columns.Add("Conflict", "Status") | Out-Null

$comboColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$comboColumn.Name = "Enabled"
$comboColumn.HeaderText = "Enabled"
$comboColumn.Width = 120
$comboColumn.Items.Add("Enabled") | Out-Null
$comboColumn.Items.Add("Disabled") | Out-Null
$grid.Columns.Add($comboColumn) | Out-Null

$grid.Columns.Add("Details", "Details") | Out-Null

# Add Mass Change checkbox column at the end
$massChangeColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$massChangeColumn.Name = "MassChange"
$massChangeColumn.HeaderText = "Mass Enable Change"
$massChangeColumn.Width = 80
$grid.Columns.Add($massChangeColumn) | Out-Null

foreach ($mod in ($mods | Sort-Object LoadOrder)) {
    $i = $grid.Rows.Add()
    $row = $grid.Rows[$i]

    # Add SW- prefix for workshop mods
    $displayName = if ($mod.IsWorkshop) { "SW-" + $mod.Name } else { $mod.Name }
    $row.Cells[0].Value = $displayName
    $row.Cells[1].Value = $mod.LoadOrder
    $row.Cells[2].Value = $mod.OriginalLoadOrder  # Original load order
    $row.Cells[4].Value = if ($mod.Enabled) { "Enabled" } else { "Disabled" }

    if ($mod.Enabled -and $conflictResults.ContainsKey($mod.FolderKey)) {
        $row.Cells[3].Value = "CONFLICT"
        $row.Cells[5].Value = "View"
        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightCoral
    } else {
        $row.Cells[3].Value = "Clean"
        $row.Cells[5].Value = "-"
        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
    }

    # Store the original mod object for later reference
    $row.Tag = $mod
}

$grid.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        
        # Check if double-click is on Load Order column (index 1)
        if ($e.ColumnIndex -eq 1) {
            $row = $grid.Rows[$e.RowIndex]
            $mod = $row.Tag
            if (-not $mod) { return }
            
            $originalOrder = $mod.OriginalLoadOrder
            $currentOrder = [int]$row.Cells[1].Value  # Get current value from grid, not mod object
            
            if ($originalOrder -eq $currentOrder) {
                [System.Windows.Forms.MessageBox]::Show("Load order is already at original value: $originalOrder", "No Change Needed", "OK", "Information")
                return
            }
            
            $result = [System.Windows.Forms.MessageBox]::Show("Restore '$($mod.Name)' load order to original value?`n`nCurrent: $currentOrder`nOriginal: $originalOrder", "Confirm Restore", "YesNo", "Question")
            if ($result -ne "Yes") { return }
            
            try {
                # Update mod.json file
                $json = Get-Content $mod.ModJsonPath -Raw | ConvertFrom-Json
                
                # Handle different property names for load order
                if ($json.PSObject.Properties.Name -contains "loadOrder") {
                    $json.loadOrder = $originalOrder
                } elseif ($json.PSObject.Properties.Name -contains "defaultLoadOrder") {
                    $json.defaultLoadOrder = $originalOrder
                } else {
                    # Add loadOrder property if it doesn't exist
                    $json | Add-Member -NotePropertyName "loadOrder" -NotePropertyValue $originalOrder
                }
                
                $json | ConvertTo-Json -Depth 10 | Set-Content $mod.ModJsonPath -Force
                
                # Update mod object and grid
                $mod.LoadOrder = $originalOrder
                $row.Cells[1].Value = $originalOrder
                
                Write-Host "Restored load order for $($mod.Name) to $originalOrder" -ForegroundColor Green
                [System.Windows.Forms.MessageBox]::Show("Load order restored to $originalOrder`n`nSave changes to apply permanently.", "Success", "OK", "Information")
            } catch {
                Write-Host "Error restoring load order: $_" -ForegroundColor Red
                [System.Windows.Forms.MessageBox]::Show("Error restoring load order: $_", "Error", "OK", "Error")
            }
        }
    })

$grid.Add_CellClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        if ($e.ColumnIndex -ne 5) { return }  # Updated to column index 5 for Details
        $row = $grid.Rows[$e.RowIndex]
        if ($row.Cells[5].Value -ne "View") { return }
        $mod = $row.Tag
        if (-not $mod.Enabled -or -not $conflictResults.ContainsKey($mod.FolderKey)) { return }
        $body = Build-ConflictViewText $mod $assetMap
        if ($body.Length -gt 200000) { $body = $body.Substring(0, 200000) + "`r`n... (truncated)" }
        Show-ConflictViewForm "Mod Conflicts: $($mod.Name)" $body $form
    })

    # Add CellValueChanged event handler for Mass Enable Change functionality
    $grid.Add_CellValueChanged({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        if ($e.ColumnIndex -ne 6) { return }  # Only handle Mass Enable Change column (index 6)
        
        $row = $grid.Rows[$e.RowIndex]
        $mod = $row.Tag
        $folderKey = [string]$mod.FolderKey
        
        if ($row.Cells[6].Value -eq $true) {
            # Checkbox is checked - toggle the enabled state
            if ($mod.Enabled) {
                # Mod was originally enabled - disable it
                $row.Cells[4].Value = "Disabled"
                $EnabledMods[$folderKey] = $false
            } else {
                # Mod was originally disabled - enable it
                $row.Cells[4].Value = "Enabled"
                $EnabledMods[$folderKey] = $true
            }
        } else {
            # Checkbox is unchecked - revert to original mod state
            $row.Cells[4].Value = if ($mod.Enabled) { "Enabled" } else { "Disabled" }
            $EnabledMods[$folderKey] = $mod.Enabled
        }
        
        $grid.Refresh()
    })

$saveBtn = New-Object System.Windows.Forms.Button
$saveBtn.Text = "Save to modlist.json"
$saveBtn.Size = New-Object System.Drawing.Size(200, 40)
$saveBtn.Location = New-Object System.Drawing.Point(10, 610)
$saveBtn.BackColor = [System.Drawing.Color]::LightBlue
$saveBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

# Add Revert All Changed Load Orders button
$revertAllBtn = New-Object System.Windows.Forms.Button
$revertAllBtn.Text = "Revert All Changed Load Orders"
$revertAllBtn.Size = New-Object System.Drawing.Size(200, 40)
$revertAllBtn.Location = New-Object System.Drawing.Point(430, 610)
$revertAllBtn.BackColor = [System.Drawing.Color]::LightYellow
$revertAllBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)


# Add Reset Playlists button
$resetPlaylistsBtn = New-Object System.Windows.Forms.Button
$resetPlaylistsBtn.Text = "Reset Playlists File"
$resetPlaylistsBtn.Size = New-Object System.Drawing.Size(200, 40)
$resetPlaylistsBtn.Location = New-Object System.Drawing.Point(940, 605)
$resetPlaylistsBtn.BackColor = [System.Drawing.Color]::Orange
$resetPlaylistsBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$resetPlaylistsBtn.Add_Click({
    $res = [System.Windows.Forms.MessageBox]::Show("Are you sure? This will remove all playlists.", "Reset Playlists", "YesNo", "Warning")
    if ($res -eq "Yes") {
        try {
            Remove-Item $PlaylistsFile -Force -ErrorAction SilentlyContinue
            Write-Host "Playlists file deleted: $PlaylistsFile" -ForegroundColor Green
            $script:playlists = @()
            Write-Host "Playlists array reset to empty" -ForegroundColor Yellow
        } catch {
            Write-Host "Error deleting playlists file: $_" -ForegroundColor Red
        }
    }
})
$updateReportBtn = New-Object System.Windows.Forms.Button
$updateReportBtn.Text = "Update Report"
$updateReportBtn.Size = New-Object System.Drawing.Size(200, 40)
$updateReportBtn.Location = New-Object System.Drawing.Point(650, 610)
$updateReportBtn.BackColor = [System.Drawing.Color]::LightGreen
$updateReportBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$updateReportBtn.Add_Click({
    Update-ModReport
    [System.Windows.Forms.MessageBox]::Show("Report updated!`n`nMW5_Mod_List.txt has been regenerated.", "Success", "OK", "Information")
})
$revertAllBtn.Add_Click({
        $changedMods = @()
        
        # Find all mods with changed load orders
        foreach ($row in $grid.Rows) {
            $mod = $row.Tag
            if (-not $mod) { continue }
            
            $currentOrder = [int]$row.Cells[1].Value  # Get current value from grid
            if ($currentOrder -ne $mod.OriginalLoadOrder) {
                $changedMods += @{
                    Mod = $mod
                    Row = $row
                    Current = $currentOrder
                    Original = $mod.OriginalLoadOrder
                }
            }
        }
        
        if ($changedMods.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No load orders have been changed.", "No Changes", "OK", "Information")
            return
        }
        
        $changedList = $changedMods | ForEach-Object { "• $($_.Mod.Name): $($_.Current) → $($_.Original)" } | Out-String
        
        $result = [System.Windows.Forms.MessageBox]::Show("Revert $($changedMods.Count) mods to original load order?`n`n$changedList", "Confirm Revert All", "YesNo", "Question")
        if ($result -ne "Yes") { return }
        
        $successCount = 0
        $errorCount = 0
        
        foreach ($item in $changedMods) {
            try {
                # Update mod.json file
                $json = Get-Content $item.Mod.ModJsonPath -Raw | ConvertFrom-Json
                
                # Handle different property names for load order
                if ($json.PSObject.Properties.Name -contains "loadOrder") {
                    $json.loadOrder = $item.Original
                } elseif ($json.PSObject.Properties.Name -contains "defaultLoadOrder") {
                    $json.defaultLoadOrder = $item.Original
                } else {
                    # Add loadOrder property if it doesn't exist
                    $json | Add-Member -NotePropertyName "loadOrder" -NotePropertyValue $item.Original
                }
                
                $json | ConvertTo-Json -Depth 10 | Set-Content $item.Mod.ModJsonPath -Force
                
                # Update mod object and grid
                $item.Mod.LoadOrder = $item.Original
                $item.Row.Cells[1].Value = $item.Original
                
                $successCount++
                Write-Host "Restored load order for $($item.Mod.Name) to $($item.Original)" -ForegroundColor Green
            } catch {
                $errorCount++
                Write-Host "Error restoring load order for $($item.Mod.Name): $_" -ForegroundColor Red
            }
        }
        
        $message = "Load order revert completed!`n`nSuccessfully reverted: $successCount mods"
        if ($errorCount -gt 0) {
            $message += "`nErrors: $errorCount mods"
        }
        $message += "`n`nSave changes to apply permanently."
        
        [System.Windows.Forms.MessageBox]::Show($message, "Revert Complete", "OK", "Information")
    })

$saveBtn.Add_Click({
        # Only update mods that actually changed
        $changedMods = @()
        
        for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
            $row = $grid.Rows[$i]
            $mod = $row.Tag
            if (-not $mod) { continue }
            
            $newLoad = [int]$row.Cells[1].Value
            $newEnabled = ($row.Cells[4].Value -eq "Enabled")
            
            # Check if this mod actually changed
            $loadChanged = ($newLoad -ne $mod.OriginalLoadOrder)
            $enabledChanged = ($newEnabled -ne $mod.Enabled)
            
            if ($loadChanged -or $enabledChanged) {
                $changedMods += @{
                    Mod = $mod
                    Row = $row
                    NewLoad = $newLoad
                    NewEnabled = $newEnabled
                }
                
                # Update mod.json if changed
                try {
                    $json = Get-Content $mod.ModJsonPath -Raw | ConvertFrom-Json
                    if ($json.PSObject.Properties.Name -contains 'loadOrder') {
                        $json.loadOrder = $newLoad
                    } elseif ($json.PSObject.Properties.Name -contains 'defaultLoadOrder') {
                        $json.defaultLoadOrder = $newLoad
                    } else {
                        $json | Add-Member -NotePropertyName 'loadOrder' -NotePropertyValue $newLoad -Force
                    }
                    $json | ConvertTo-Json -Depth 20 | Set-Content $mod.ModJsonPath -Force -Encoding UTF8
                } catch { 
                    Write-Host "Failed updating mod.json for $($mod.Name)" -ForegroundColor Red
                }
                
                # Update mod object
                $mod.LoadOrder = $newLoad
                $mod.Enabled = $newEnabled
            }
        }
        
        # Update modlist.json only with changed mods
        if ($changedMods.Count -gt 0) {
            foreach ($item in $changedMods) {
                $EnabledMods[$item.Mod.FolderKey] = $item.NewEnabled
            }
            
            # Read existing modlist.json to preserve structure
            $existingModlist = $null
            if (Test-Path $ModListFile) {
                try {
                    $existingModlist = Get-Content $ModListFile -Raw | ConvertFrom-Json
                } catch {
                    Write-Host "Warning: Could not read existing modlist.json" -ForegroundColor Yellow
                }
            }
            
            # Update only changed mods in existing structure
            if ($existingModlist) {
                # Create new object with existing structure
                $modlistContent = @{
                    gameVersion = $existingModlist.gameVersion
                    modStatus = $existingModlist.modStatus
                }
                # Update changed mods
                foreach ($item in $changedMods) {
                    $modlistContent.modStatus | Add-Member -NotePropertyName $item.Mod.FolderKey -NotePropertyValue @{ bEnabled = $item.NewEnabled } -Force
                }
            } else {
                # Fallback: create new structure if existing is unreadable
                $modlistContent = @{
                    gameVersion = "1.13.x"
                    modStatus = $EnabledMods.GetEnumerator() | ForEach-Object { @{ bEnabled = $_.Value } }
                }
            }
            
            $modlistContent | ConvertTo-Json -Depth 10 | Set-Content $ModListFile -Force -Encoding UTF8
            
            [System.Windows.Forms.MessageBox]::Show("Updated $($changedMods.Count) mod changes in modlist.json!`n`nClick 'Update Report' to generate new report.", "Success", "OK", "Information")
        } else {
            [System.Windows.Forms.MessageBox]::Show("No changes to save.", "No Changes", "OK", "Information")
        }
    })

$updateReportBtn = New-Object System.Windows.Forms.Button
$updateReportBtn.Text = "Update Report"
$updateReportBtn.Size = New-Object System.Drawing.Size(200, 40)
$updateReportBtn.Location = New-Object System.Drawing.Point(650, 610)
$updateReportBtn.BackColor = [System.Drawing.Color]::LightGreen
$updateReportBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$updateReportBtn.Add_Click({
        Update-ModReport
        [System.Windows.Forms.MessageBox]::Show("Report updated!`n`nMW5_Mod_List.txt has been regenerated.", "Success", "OK", "Information")
    })

$tabMods.Controls.Add($grid)
$tabMods.Controls.Add($saveBtn)
$tabMods.Controls.Add($revertAllBtn)
$tabMods.Controls.Add($updateReportBtn)

# =========================
# Playlists tab
# =========================
$playlists = Ensure-PlaylistIds (Load-Playlists)

$plLabel = New-Object System.Windows.Forms.Label
$plLabel.Text = "Playlist name"
$plLabel.Location = New-Object System.Drawing.Point(10, 15)
$plLabel.Size = New-Object System.Drawing.Size(140, 20)

$plName = New-Object System.Windows.Forms.TextBox
$plName.Location = New-Object System.Drawing.Point(10, 38)
$plName.Size = New-Object System.Drawing.Size(360, 24)
$plName.Text = "Known working order"

$plSaveBtn = New-Object System.Windows.Forms.Button
$plSaveBtn.Text = "Save current enabled + load order"
$plSaveBtn.Location = New-Object System.Drawing.Point(390, 35)
$plSaveBtn.Size = New-Object System.Drawing.Size(280, 30)

$plGrid = New-Object System.Windows.Forms.DataGridView
$plGrid.Location = New-Object System.Drawing.Point(10, 80)
$plGrid.Size = New-Object System.Drawing.Size(1240, 520)
$plGrid.AutoSizeColumnsMode = "Fill"
$plGrid.AllowUserToAddRows = $false
$plGrid.ReadOnly = $true
$plGrid.SelectionMode = "FullRowSelect"
$plGrid.MultiSelect = $false

$plGrid.Columns.Add("PlName", "Playlist") | Out-Null
$plGrid.Columns.Add("PlCreated", "Created") | Out-Null
$plGrid.Columns.Add("PlEnabledCount", "Enabled Mods") | Out-Null
$plGrid.Columns.Add("PlNotes", "Notes") | Out-Null

function Refresh-PlaylistGrid {
    $plGrid.Rows.Clear()
    $plGrid.Refresh()
    
    # Debug: Show what playlists we have
    Write-Host "Refresh-PlaylistGrid called. Playlists in memory: $($playlists.Count)" -ForegroundColor Cyan
    
    # IMPORTANT: Do NOT reload from file - use only in-memory data
    # This prevents accidentally restoring deleted playlists
    
    # Ensure playlists array is not null
    if ($null -eq $playlists) {
        $playlists = @()
        Write-Host "Playlists was null, initialized to empty array" -ForegroundColor Red
    }
    
    # Remove any null entries and duplicates
    $cleanPlaylists = @($playlists | Where-Object { $null -ne $_ } | Sort-Object { $_.name } -Unique)
    
    Write-Host "Clean playlists count: $($cleanPlaylists.Count)" -ForegroundColor Cyan
    
    # Only add playlists that are actually in memory
    foreach ($p in $cleanPlaylists) {
        if ($null -eq $p) { continue }
        $i = $plGrid.Rows.Add()
        $row = $plGrid.Rows[$i]
        $row.Cells[0].Value = [string]$p.name
        $row.Cells[1].Value = [string]$p.created
        $row.Cells[2].Value = if ($p.mods) { @($p.mods).Count } else { 0 }
        $row.Cells[3].Value = [string]$p.notes
        $row.Tag = $p
        Write-Host "Added playlist to grid: $($p.name) with $(if ($p.mods) { @($p.mods).Count } else { 0 }) mods" -ForegroundColor Gray
    }
    
    Write-Host "Grid refresh complete. Final row count: $($plGrid.Rows.Count)" -ForegroundColor Yellow
    $plGrid.Refresh()
}

function Select-PlaylistRowById([string]$id) {
    if ([string]::IsNullOrWhiteSpace($id)) { return }
    try {
        $plGrid.ClearSelection()
        for ($i = 0; $i -lt $plGrid.Rows.Count; $i++) {
            $row = $plGrid.Rows[$i]
            $p = $row.Tag
            if ($p -and ([string]$p.id -eq $id)) {
                $row.Selected = $true
                try { $plGrid.FirstDisplayedScrollingRowIndex = $i } catch {}
                return
            }
        }
    } catch {}
}

$plApplyBtn = New-Object System.Windows.Forms.Button
$plApplyBtn.Text = "Apply selected playlist to grid"
$plApplyBtn.Location = New-Object System.Drawing.Point(10, 605)
$plApplyBtn.Size = New-Object System.Drawing.Size(220, 30)

$plViewModsBtn = New-Object System.Windows.Forms.Button
$plViewModsBtn.Text = "See mods in playlist"
$plViewModsBtn.Location = New-Object System.Drawing.Point(240, 605)
$plViewModsBtn.Size = New-Object System.Drawing.Size(220, 30)

$plExportBtn = New-Object System.Windows.Forms.Button
$plExportBtn.Text = "Export selected"
$plExportBtn.Location = New-Object System.Drawing.Point(470, 605)
$plExportBtn.Size = New-Object System.Drawing.Size(140, 30)

$plImportBtn = New-Object System.Windows.Forms.Button
$plImportBtn.Text = "Import from file..."
$plImportBtn.Location = New-Object System.Drawing.Point(620, 605)
$plImportBtn.Size = New-Object System.Drawing.Size(160, 30)

$plDeleteBtn = New-Object System.Windows.Forms.Button
$plDeleteBtn.Text = "Delete selected"
$plDeleteBtn.Location = New-Object System.Drawing.Point(790, 605)
$plDeleteBtn.Size = New-Object System.Drawing.Size(140, 30)


$plSaveBtn.Add_Click({
        $name = Normalize-PlaylistName $plName.Text
        if ([string]::IsNullOrWhiteSpace($name)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a playlist name.", "Missing name", "OK", "Warning") | Out-Null
            return
        }

        $defaultName = "Known working order"
        $defaultNameNorm = Normalize-PlaylistName $defaultName
        $selectedPlaylist = $null
        if ($plGrid.SelectedRows.Count -gt 0) { $selectedPlaylist = $plGrid.SelectedRows[0].Tag }

        $enabledRows = @()
        Write-Host "Scanning grid for enabled mods... Total rows: $($grid.Rows.Count)" -ForegroundColor Cyan
        for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
            $row = $grid.Rows[$i]
            $mod = $row.Tag
            $isEnabled = ($row.Cells[4].Value -eq "Enabled")
            Write-Host ("Row " + $i + ": " + $mod.Name + " - Enabled column value: '" + $row.Cells[4].Value + "' - IsEnabled: " + $isEnabled) -ForegroundColor Gray
            if (-not $isEnabled) { continue }
            $enabledRows += [PSCustomObject]@{
                FolderKey  = $mod.FolderKey
                Name       = $mod.Name
                LoadOrder  = [int]$row.Cells[1].Value
                Enabled    = $true
            }
        }
        
        Write-Host "Found $($enabledRows.Count) enabled mods for playlist creation" -ForegroundColor Green

        # Decide which playlist to modify (if any)
        $existingIdx = -1
        $mode = "new"
        $targetName = $name
        $targetNameNorm = Normalize-PlaylistName $targetName

        # If user leaves default name and a playlist is selected, treat Save as "modify selected"
        if ($selectedPlaylist -and ((Normalize-PlaylistName ([string]$name)) -eq $defaultNameNorm)) {
            $targetId = [string]$selectedPlaylist.id
            if (-not [string]::IsNullOrWhiteSpace($targetId)) {
                for ($i = 0; $i -lt $playlists.Count; $i++) {
                    if ([string]$playlists[$i].id -eq $targetId) { $existingIdx = $i; break }
                }
            }
            if ($existingIdx -lt 0) {
                # fallback: locate by name if id missing
                for ($i = 0; $i -lt $playlists.Count; $i++) {
                    if ((Normalize-PlaylistName ([string]$playlists[$i].name)) -eq (Normalize-PlaylistName ([string]$selectedPlaylist.name))) { $existingIdx = $i; break }
                }
            }

            if ($existingIdx -ge 0) {
                $targetName = Normalize-PlaylistName ([string]$playlists[$existingIdx].name)
                $targetNameNorm = Normalize-PlaylistName $targetName
                $res = [System.Windows.Forms.MessageBox]::Show("Modify selected playlist '$targetName'?`n`nYes = Overwrite selected`nNo = Add/Update into selected`nCancel = Do nothing", "Modify selected playlist", "YesNoCancel", "Question")
                if ($res -eq "Cancel") { return }
                if ($res -eq "Yes") { $mode = "overwrite" } else { $mode = "append" }
            } else {
                # Nothing found to modify; fall back to name-based save
                $selectedPlaylist = $null
            }
        }

        # If not modifying selected: save by name. If name already exists, prompt overwrite vs add/update.
        if ($existingIdx -lt 0) {
            for ($i = 0; $i -lt $playlists.Count; $i++) {
                if ((Normalize-PlaylistName ([string]$playlists[$i].name)) -eq $targetNameNorm) { $existingIdx = $i; break }
            }
            if ($existingIdx -ge 0) {
                $res = [System.Windows.Forms.MessageBox]::Show("Playlist '$targetName' already exists.`n`nYes = Overwrite playlist`nNo = Add/Update mods in playlist`nCancel = Do nothing", "Save playlist", "YesNoCancel", "Question")
                if ($res -eq "Cancel") { return }
                if ($res -eq "Yes") { $mode = "overwrite" } else { $mode = "append" }
            }
        }

        $now = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        $plObj = $null
        if ($mode -eq "append") {
            $plObj = $playlists[$existingIdx]
            if (-not $plObj.id) {
                try { $plObj | Add-Member -NotePropertyName id -NotePropertyValue ([guid]::NewGuid().ToString()) -Force } catch {}
            }
            $existingMods = @()
            if ($plObj.mods) { $existingMods = @($plObj.mods) }
            $map = @{}
            foreach ($m in $existingMods) {
                if ($m.FolderKey) { $map[[string]$m.FolderKey] = $m }
            }
            foreach ($m in @($enabledRows)) {
                $k = [string]$m.FolderKey
                if ($map.ContainsKey($k)) {
                    $map[$k].LoadOrder = [int]$m.LoadOrder
                    $map[$k].Name = [string]$m.Name
                    $map[$k].Enabled = $true
                } else {
                    $existingMods += $m
                    $map[$k] = $m
                }
            }
            $plObj.name = $targetName
            $plObj.created = $now
            $plObj.notes = "Added/updated enabled mods + load order snapshot"
            $plObj.mods = @($existingMods | Sort-Object LoadOrder)
            
            # Update original load orders for this playlist
            if (-not $plObj.originalLoadOrders) {
                $plObj.originalLoadOrders = @{}
            }
            foreach ($mod in $existingMods) {
                $plObj.originalLoadOrders[$mod.FolderKey] = $mod.LoadOrder
            }
        } else {
            $plObj = [PSCustomObject]@{
                id      = ([guid]::NewGuid().ToString())
                name    = $targetName
                created = $now
                notes   = "Enabled mods + load order snapshot"
                mods    = @($enabledRows | Sort-Object LoadOrder)
                originalLoadOrders = @{}
            }
            
            # Store original load orders for this playlist
            foreach ($mod in $enabledRows) {
                $plObj.originalLoadOrders[$mod.FolderKey] = $mod.LoadOrder
            }
        }

        if ($existingIdx -ge 0) { $playlists[$existingIdx] = $plObj }
        else { $playlists = @($playlists) + @($plObj) }

        Save-Playlists $playlists
        Refresh-PlaylistGrid
        Select-PlaylistRowById ([string]$plObj.id)
        [System.Windows.Forms.MessageBox]::Show("Saved playlist '$targetName'.", "Playlist saved", "OK", "Information") | Out-Null
    })

$plApplyBtn.Add_Click({
        if ($plGrid.SelectedRows.Count -eq 0) { return }
        $p = $plGrid.SelectedRows[0].Tag
        if (-not $p -or -not $p.mods) { return }

        $map = @{}
        foreach ($m in @($p.mods)) {
            if ($m.FolderKey) { $map[[string]$m.FolderKey] = $m }
        }

        for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
            $row = $grid.Rows[$i]
            $mod = $row.Tag
            $k = [string]$mod.FolderKey
            if ($map.ContainsKey($k)) {
                $row.Cells[4].Value = "Enabled"
                # Use original load order from playlist if available, otherwise current load order
                if ($p.originalLoadOrders -and $p.originalLoadOrders.PSObject.Properties.Name -contains $k) {
                    $row.Cells[1].Value = [int]$p.originalLoadOrders.$k
                } else {
                    $row.Cells[1].Value = [int]$map[$k].LoadOrder
                }
                $EnabledMods[$k] = $true
            } else {
                $row.Cells[4].Value = "Disabled"
                $EnabledMods[$k] = $false
            }
        }
        
        # Force grid refresh to make enabled status visible immediately
        $grid.Refresh()
        $grid.Invalidate()

        [System.Windows.Forms.MessageBox]::Show("Applied playlist '$($p.name)' to grid.`n`nNow click 'Save to modlist.json' to write mod.json + modlist.json.", "Playlist applied", "OK", "Information") | Out-Null
    })

$plViewModsBtn.Add_Click({
    if ($plGrid.SelectedRows.Count -eq 0) { 
        [System.Windows.Forms.MessageBox]::Show("Select a playlist first.", "No selection", "OK", "Warning") | Out-Null
        return
    }
    
    $p = $plGrid.SelectedRows[0].Tag
    if (-not $p) { 
        [System.Windows.Forms.MessageBox]::Show("Selected playlist has no data.", "Error", "OK", "Error") | Out-Null
        return
    }
    
    if (-not $p.mods -or $p.mods.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("This playlist has no mods.", "Empty playlist", "OK", "Information") | Out-Null
        return
    }

    $vf = New-Object System.Windows.Forms.Form
    $vf.Text = "Edit playlist: $([string]$p.name)"
    $vf.Size = New-Object System.Drawing.Size(1200, 760)
    $vf.StartPosition = 'CenterParent'
    $vf.Owner = $form

    $hdr = New-Object System.Windows.Forms.Label
    $hdr.Location = New-Object System.Drawing.Point(12, 12)
    $hdr.Size = New-Object System.Drawing.Size(1160, 40)
    $hdr.Text = "Check/uncheck mods to add/remove them from the playlist. Edit load order, then click Save changes."
    $hdr.AutoSize = $false

    $vg = New-Object System.Windows.Forms.DataGridView
    $vg.Location = New-Object System.Drawing.Point(12, 60)
    $vg.Size = New-Object System.Drawing.Size(1160, 600)
    $vg.AutoSizeColumnsMode = "Fill"
    $vg.AllowUserToAddRows = $false
    $vg.ReadOnly = $false
    $vg.SelectionMode = "FullRowSelect"

    $chkCol = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $chkCol.Name = "InPl"
    $chkCol.HeaderText = "In playlist"
    $chkCol.Width = 80
    $vg.Columns.Add($chkCol) | Out-Null
    $vg.Columns.Add("Name", "Mod Name") | Out-Null
    $vg.Columns.Add("Load", "Load Order") | Out-Null
    $vg.Columns.Add("Folder", "Folder Key") | Out-Null

    # Build lookup of playlist mods
    $plMap = @{}
    if ($p.mods) {
        foreach ($m in @($p.mods)) {
            if ($m -and $m.FolderKey) { $plMap[[string]$m.FolderKey] = $m }
        }
    }

    # Fill editor rows from all available mods
    $modsSorted = @($mods | Sort-Object LoadOrder, Name)
    foreach ($mod in $modsSorted) {
        $i = $vg.Rows.Add()
        $r = $vg.Rows[$i]
        $k = [string]$mod.FolderKey
        $in = $plMap.ContainsKey($k)
        $r.Cells[0].Value = $in
        $r.Cells[1].Value = [string]$mod.Name
        if ($in) { 
            $r.Cells[2].Value = [int]$plMap[$k].LoadOrder 
        } else { 
            $r.Cells[2].Value = [int]$mod.LoadOrder 
        }
        $r.Cells[3].Value = $k
    }

    $save = New-Object System.Windows.Forms.Button
    $save.Text = "Save changes to playlist"
    $save.Location = New-Object System.Drawing.Point(12, 670)
    $save.Size = New-Object System.Drawing.Size(220, 34)

    $close = New-Object System.Windows.Forms.Button
    $close.Text = "Close"
    $close.Location = New-Object System.Drawing.Point(250, 670)
    $close.Size = New-Object System.Drawing.Size(120, 34)

    $save.Add_Click({
        $newMods = @()
        for ($ri = 0; $ri -lt $vg.Rows.Count; $ri++) {
            $row = $vg.Rows[$ri]
            $inPl = [bool]$row.Cells[0].Value
            if (-not $inPl) { continue }
            $folder = [string]$row.Cells[3].Value
            if ([string]::IsNullOrWhiteSpace($folder)) { continue }
            $nm = [string]$row.Cells[1].Value
            $lo = 0
            try { $lo = [int]$row.Cells[2].Value } catch { $lo = 0 }
            $newMods += [PSCustomObject]@{
                FolderKey = $folder
                Name      = $nm
                LoadOrder = $lo
                Enabled   = $true
            }
        }

        $p.created = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        $p.notes = "Edited in playlist editor"
        $p.mods = @($newMods | Sort-Object LoadOrder)
        
        if (-not $p.originalLoadOrders) {
            $p.originalLoadOrders = @{}
        }
        foreach ($mod in $newMods) {
            $p.originalLoadOrders[$mod.FolderKey] = $mod.LoadOrder
        }

        Save-Playlists $playlists
        Refresh-PlaylistGrid
        [System.Windows.Forms.MessageBox]::Show("Saved changes to playlist '$([string]$p.name)'.", "Playlist updated", "OK", "Information") | Out-Null
        $vf.Close()
    })

    $close.Add_Click({ $vf.Close() })

    $vf.Controls.Add($hdr)
    $vf.Controls.Add($vg)
    $vf.Controls.Add($save)
    $vf.Controls.Add($close)
    
    $vf.ShowDialog() | Out-Null
})

$plExportBtn.Add_Click({
        if ($plGrid.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Select a playlist to export.", "No selection", "OK", "Information") | Out-Null
            return
        }
        $p = $plGrid.SelectedRows[0].Tag
        if (-not $p) { return }

        $plNameForFile = Get-SafeFileName (Normalize-PlaylistName ([string]$p.name))
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Title = "Export playlist"
        $sfd.InitialDirectory = $ScriptDir
        $sfd.FileName = ($plNameForFile + ".mw5playlist.json")
        $sfd.Filter = "MW5 playlist (*.mw5playlist.json)|*.mw5playlist.json|JSON (*.json)|*.json|All files (*.*)|*.*"
        $sfd.OverwritePrompt = $true

        if ($sfd.ShowDialog() -ne "OK") { return }

        $exportObj = [PSCustomObject]@{
            name    = (Normalize-PlaylistName ([string]$p.name))
            created = [string]$p.created
            notes   = [string]$p.notes
            mods    = @()
        }
        if ($p.mods) { $exportObj.mods = @($p.mods) }

        try {
            ($exportObj | ConvertTo-Json -Depth 20) | Set-Content $sfd.FileName -Force -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Exported playlist to:`n`n$($sfd.FileName)", "Export complete", "OK", "Information") | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to export playlist.`n`n$($_.Exception.Message)", "Export failed", "OK", "Error") | Out-Null
        }
    })

$plImportBtn.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title = "Import playlist"
        $ofd.InitialDirectory = $ScriptDir
        $ofd.Filter = "MW5 playlist (*.mw5playlist.json)|*.mw5playlist.json|JSON (*.json)|*.json|All files (*.*)|*.*"
        $ofd.Multiselect = $false

        if ($ofd.ShowDialog() -ne "OK") { return }

        $file = $ofd.FileName
        $base = [System.IO.Path]::GetFileNameWithoutExtension($file)
        # handle double-extension like .mw5playlist.json -> base becomes "name.mw5playlist"
        if ($base.ToLowerInvariant().EndsWith(".mw5playlist")) {
            $base = $base.Substring(0, $base.Length - ".mw5playlist".Length)
        }
        $importName = Normalize-PlaylistName $base
        if ([string]::IsNullOrWhiteSpace($importName)) { $importName = "Imported playlist" }

        try {
            $raw = Get-Content $file -Raw
            $data = $raw | ConvertFrom-Json

            $modsImported = @()
            if ($data.mods) { $modsImported = @($data.mods) }
            elseif ($data.playlists -and @($data.playlists).Count -gt 0 -and $data.playlists[0].mods) { $modsImported = @($data.playlists[0].mods) }

            $plObj = [PSCustomObject]@{
                id      = ([guid]::NewGuid().ToString())
                name    = $importName
                created = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                notes   = "Imported from file: $([System.IO.Path]::GetFileName($file))"
                mods    = @($modsImported | Sort-Object LoadOrder)
            }

            # if name exists, prompt overwrite vs add/update
            $existingIdx = -1
            $importNameNorm = Normalize-PlaylistName $importName
            for ($i = 0; $i -lt $playlists.Count; $i++) {
                if ((Normalize-PlaylistName ([string]$playlists[$i].name)) -eq $importNameNorm) { $existingIdx = $i; break }
            }
            if ($existingIdx -ge 0) {
                $res = [System.Windows.Forms.MessageBox]::Show("A playlist named '$importName' already exists.`n`nYes = Overwrite existing`nNo = Add/Update into existing`nCancel = Do nothing", "Import playlist", "YesNoCancel", "Question")
                if ($res -eq "Cancel") { return }
                if ($res -eq "No") {
                    # merge mods into existing
                    $existing = $playlists[$existingIdx]
                    $existingMods = @()
                    if ($existing.mods) { $existingMods = @($existing.mods) }
                    $map = @{}
                    foreach ($m in $existingMods) { if ($m.FolderKey) { $map[[string]$m.FolderKey] = $m } }
                    foreach ($m in @($plObj.mods)) {
                        $k = [string]$m.FolderKey
                        if ([string]::IsNullOrWhiteSpace($k)) { continue }
                        if ($map.ContainsKey($k)) {
                            $map[$k].LoadOrder = [int]$m.LoadOrder
                            $map[$k].Name = [string]$m.Name
                            $map[$k].Enabled = $true
                        } else {
                            $existingMods += $m
                            $map[$k] = $m
                        }
                    }
                    $existing.created = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    $existing.notes = "Imported (merged) from file: $([System.IO.Path]::GetFileName($file))"
                    $existing.mods = @($existingMods | Sort-Object LoadOrder)
                    $plObj = $existing
                    $playlists[$existingIdx] = $existing
                } else {
                    # overwrite existing: keep its id
                    $existing = $playlists[$existingIdx]
                    $plObj.id = [string]$existing.id
                    $playlists[$existingIdx] = $plObj
                }
            } else {
                $playlists = @($playlists) + @($plObj)
            }

            Save-Playlists $playlists
            Refresh-PlaylistGrid
            Select-PlaylistRowById ([string]$plObj.id)
            [System.Windows.Forms.MessageBox]::Show("Imported playlist '$importName'.", "Import complete", "OK", "Information") | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to import playlist.`n`n$($_.Exception.Message)", "Import failed", "OK", "Error") | Out-Null
        }
    })

$plDeleteBtn.Add_Click({
        if ($plGrid.SelectedRows.Count -eq 0) { return }
        $p = $plGrid.SelectedRows[0].Tag
        if (-not $p) { return }
        $name = [string]$p.name
        $id = [string]$p.id
        $res = [System.Windows.Forms.MessageBox]::Show("Delete playlist '$name'?", "Confirm delete", "YesNo", "Warning")
        if ($res -ne "Yes") { return }
        
        # Remove the playlist from the array
        $newPlaylists = @()
        $deleted = $false
        foreach ($pl in @($playlists)) {
            if (-not $deleted -and ([string]$pl.id -eq $id -or ([string]::IsNullOrWhiteSpace($id) -and [string]$pl.name -eq $name))) {
                $deleted = $true
                continue  # Skip this playlist (delete it)
            }
            $newPlaylists += $pl
        }
        
        $playlists = $newPlaylists
        
        # Verify deletion worked
        Write-Host "Deleted playlist: $name. Remaining playlists: $($playlists.Count)" -ForegroundColor Yellow
        
        # Save to file
        Save-Playlists $playlists
        
        # Verify file was updated
        try {
            $verifyData = Get-Content $PlaylistsFile -Raw | ConvertFrom-Json
            Write-Host "File now contains: $($verifyData.playlists.Count) playlists" -ForegroundColor Green
        } catch {
            Write-Host "Error verifying file save" -ForegroundColor Red
        }
        
        # Force complete refresh with delay to prevent race conditions
        Start-Sleep -Milliseconds 200
        $plGrid.Rows.Clear()
        $plGrid.Refresh()
        Start-Sleep -Milliseconds 100
        Refresh-PlaylistGrid
        $plGrid.Refresh()
    })

$tabPlaylists.Controls.Add($plLabel)
$tabPlaylists.Controls.Add($plName)
$tabPlaylists.Controls.Add($plSaveBtn)
$tabPlaylists.Controls.Add($plApplyBtn)
$tabPlaylists.Controls.Add($plViewModsBtn)
$tabPlaylists.Controls.Add($plExportBtn)
$tabPlaylists.Controls.Add($plImportBtn)
$tabPlaylists.Controls.Add($plDeleteBtn)
$tabPlaylists.Controls.Add($resetPlaylistsBtn)
$tabPlaylists.Controls.Add($plGrid)

Refresh-PlaylistGrid

$form.Controls.Add($logo)
$form.Controls.Add($themeBtn)
$form.Controls.Add($versionLabel)
$form.Controls.Add($versionDropdown)
$form.Controls.Add($addVersionLabel)
$form.Controls.Add($tabs)

Apply-ThemeToControl $form $themeColors
$form.ShowDialog() | Out-Null
