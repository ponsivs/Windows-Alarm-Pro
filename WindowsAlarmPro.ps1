using namespace System.Speech.Synthesis
<#
.SYNOPSIS
    Windows Alarm Pro - Professional Alarm System for Windows OS
.DESCRIPTION
    A comprehensive GUI-based alarm system with IGRF branded voice announcements
    Features day, week, month-wise recurring alarms, 10-minute advance warnings
    Uses only built-in Windows functions with no external dependencies
.AUTHOR
    IGRF Pvt. Ltd.
.VERSION
    1.0
.COPYRIGHT
    © 2026 IGRF Pvt. Ltd. - https://igrf.co.in/en/software/
#>

#region Assembly Loading
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Speech
#endregion
# Add this after your existing Add-Type statements
Add-Type -AssemblyName System.Core


#region Global Variables
$Global:VoiceRunspaces = New-Object System.Collections.ArrayList
$Global:AlarmsModified = 0
$Global:LastRefreshCount = 0
$Global:Alarms = @()
$Global:SpeechJobLock = New-Object System.Object
$Global:SpeechRunspaces = @{}
$Global:ActiveSpeechJobs = New-Object System.Collections.ArrayList
$Global:AlarmTimer = New-Object System.Windows.Forms.Timer
$Global:AlarmTimer.Interval = 1000  # Check every second
$Global:CurrentCulture = [System.Globalization.CultureInfo]::CurrentCulture

# ===== FIXED: Robust script path detection for both .ps1 and .exe =====
$Global:ScriptPath = $null
try {
    # Method 1: For .ps1 scripts
    if ($MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne $null) {
        $Global:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    # Method 2: For PowerShell ISE or other environments
    elseif ($PSScriptRoot -and $PSScriptRoot -ne "") {
        $Global:ScriptPath = $PSScriptRoot
    }
    # Method 3: For compiled EXE
    elseif ([Environment]::GetCommandLineArgs() -and [Environment]::GetCommandLineArgs().Count -gt 0) {
        $exePath = [Environment]::GetCommandLineArgs()[0]
        $Global:ScriptPath = Split-Path -Parent $exePath
    }
    # Method 4: Fallback to current directory
    else {
        $Global:ScriptPath = (Get-Location).Path
    }
    
    # Ensure path is valid
    if (-not $Global:ScriptPath -or $Global:ScriptPath -eq "") {
        $Global:ScriptPath = (Get-Location).Path
    }
    
    # Create a function to log the path detection (defined after Write-AlarmLog)
} catch {
    $Global:ScriptPath = (Get-Location).Path
}

$Global:LastRefresh = Get-Date
$Global:MainForm = $null
$Global:NotifyIcon = $null
$Global:ContextMenu = $null
$Global:IsAlarmActive = $false
$Global:CurrentAlarmForm = $null
$Global:CurrentAdvanceForm = $null
$Global:CustomSounds = @()
$Global:IsClosing = $false
$Global:AvailableMaleVoices = @()
$Global:AvailableFemaleVoices = @()
$Global:AvailableVoices = @()
$Global:StatusTimer = $null
$Global:CurrentPlayingSound = $null
$Global:SoundPlaybackTimer = $null
$Global:SoundPlayer = $null
$Global:SoundPlayerLock = New-Object System.Object
$Global:AlarmCheckLock = New-Object System.Object
$Global:Config = $null
$Global:LogMaxSize = 5MB
$Global:MaxLogFiles = 10
$Global:BalloonShown = $false
$Global:ManualMinimize = $false
$Global:ManualClose = $false
$Global:ApplicationRunning = $false
$Global:SoundCancellation = $false
$Global:TestMode = $false
$Global:LastUsedVoiceIndex = @{
    Male = 0
    Female = 0
}
$Global:BalloonTimer = $null
$Global:LogoPath = $null
$Global:IconPath = $null
#endregion

# Define Write-AlarmLog function FIRST before it's used
function Write-AlarmLog {
    param(
        [string]$Message,
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message`n"
    
    # CRITICAL FIX: Ensure AppDataPath is valid
    $logDir = $null
    if ($Global:AppDataPath -and (Test-Path $Global:AppDataPath)) {
        $logDir = $Global:AppDataPath
    } else {
        # Use temp directory as absolute fallback
        $logDir = Join-Path $env:TEMP "IGRF_WindowsAlarmPro"
        if (-not (Test-Path $logDir)) {
            try {
                New-Item -ItemType Directory -Path $logDir -Force -ErrorAction SilentlyContinue | Out-Null
            } catch {
                # If even temp fails, don't try to log
                return
            }
        }
    }
    
    $logFile = Join-Path $logDir "alarm_log.txt"
    
    try {
        # Ensure directory exists one more time
        $parentDir = Split-Path -Parent $logFile
        if (-not (Test-Path $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force -ErrorAction SilentlyContinue | Out-Null
        }
        
        Add-Content -Path $logFile -Value $logEntry -ErrorAction SilentlyContinue
    } catch {
        # Silently fail - logging is not critical
    }
}

# ===== NEW: Debug logging function to help diagnose path issues =====
function Write-DebugInfo {
    Write-AlarmLog -Message "========== DEBUG INFO ==========" -Level "Info"
    Write-AlarmLog -Message "ScriptPath: $Global:ScriptPath" -Level "Info"
    Write-AlarmLog -Message "LogoPath: $Global:LogoPath" -Level "Info"
    Write-AlarmLog -Message "Logo exists: $(if ($Global:LogoPath) { Test-Path $Global:LogoPath } else { $false })" -Level "Info"
    Write-AlarmLog -Message "IconPath: $Global:IconPath" -Level "Info"
    Write-AlarmLog -Message "Icon exists: $(if ($Global:IconPath) { Test-Path $Global:IconPath } else { $false })" -Level "Info"
    Write-AlarmLog -Message "AppDataPath: $Global:AppDataPath" -Level "Info"
    Write-AlarmLog -Message "Current Directory: $(Get-Location).Path" -Level "Info"
    Write-AlarmLog -Message "=================================" -Level "Info"
}

# Now initialize AppDataPath
$Global:AppDataPath = $null
try {
    # Try to use APPDATA first
    $Global:AppDataPath = "$env:APPDATA\IGRF_WindowsAlarmPro"
    if (-not (Test-Path $Global:AppDataPath)) {
        New-Item -ItemType Directory -Path $Global:AppDataPath -Force -ErrorAction Stop | Out-Null
    }
    
    # Test write access
    $testFile = Join-Path $Global:AppDataPath "test.tmp"
    [System.IO.File]::WriteAllText($testFile, "test")
    if (Test-Path $testFile) {
        Remove-Item $testFile -Force
    }
    Write-AlarmLog -Message "Using AppData path: $Global:AppDataPath" -Level "Info"
} catch {
    # Fallback to Temp directory if APPDATA fails
    Write-AlarmLog -Message "AppData path failed, using Temp directory" -Level "Warning"
    $Global:AppDataPath = Join-Path $env:TEMP "IGRF_WindowsAlarmPro"
    if (-not (Test-Path $Global:AppDataPath)) {
        New-Item -ItemType Directory -Path $Global:AppDataPath -Force | Out-Null
    }
}

# Add this function for proper volume control
function Set-SystemVolume {
    param([int]$Volume)
    
    try {
        $player = New-Object -ComObject "WMPlayer.OCX"
        $player.settings.volume = $Volume
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($player) | Out-Null
    } catch {
        # Silently fail - volume control not critical
    }
}

# Create custom sounds directory
$Global:CustomSoundsPath = Join-Path $Global:AppDataPath "CustomSounds"
if (-not (Test-Path $Global:CustomSoundsPath)) {
    New-Item -ItemType Directory -Path $Global:CustomSoundsPath -Force | Out-Null
}
#endregion

# Add this function for proper volume control
function Set-SystemVolume {
    param([int]$Volume)
    
    try {
        $player = New-Object -ComObject "WMPlayer.OCX"
        $player.settings.volume = $Volume
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($player) | Out-Null
    } catch {
        # Silently fail - volume control not critical
    }
}

# AppDataPath should already be initialized above
# Just ensure custom sounds directory exists
if ($Global:AppDataPath) {
    $Global:CustomSoundsPath = Join-Path $Global:AppDataPath "CustomSounds"
    if (-not (Test-Path $Global:CustomSoundsPath)) {
        New-Item -ItemType Directory -Path $Global:CustomSoundsPath -Force | Out-Null
    }
} else {
    # Fallback
    $Global:CustomSoundsPath = Join-Path $env:TEMP "IGRF_WindowsAlarmPro\CustomSounds"
    if (-not (Test-Path $Global:CustomSoundsPath)) {
        New-Item -ItemType Directory -Path $Global:CustomSoundsPath -Force | Out-Null
    }
}
#endregion

#region External Resources - Load from script folder (NO EMBEDDING)
# This loads logo.png and icon.ico directly from the script folder
# IMPORTANT: Place logo.png and icon.ico in the same folder as the .ps1 script or .exe file

function Initialize-Resources {
    try {
        Write-AlarmLog -Message "Initializing resources..." -Level "Info"
        Write-AlarmLog -Message "Script path: $Global:ScriptPath" -Level "Info"
        
        # Reset paths
        $Global:LogoPath = $null
        $Global:IconPath = $null
        
        if ($Global:ScriptPath -and (Test-Path $Global:ScriptPath)) {
            $Global:LogoPath = Join-Path $Global:ScriptPath "logo.png"
            $Global:IconPath = Join-Path $Global:ScriptPath "icon.ico"
            
            Write-AlarmLog -Message "Looking for logo at: $Global:LogoPath" -Level "Info"
            Write-AlarmLog -Message "Looking for icon at: $Global:IconPath" -Level "Info"
        } else {
            Write-AlarmLog -Message "Script path invalid or not found, trying current directory" -Level "Warning"
            $Global:LogoPath = Join-Path (Get-Location).Path "logo.png"
            $Global:IconPath = Join-Path (Get-Location).Path "icon.ico"
        }
        
        # Verify logo exists
        if ($Global:LogoPath -and (Test-Path $Global:LogoPath)) {
            Write-AlarmLog -Message "Logo found at: $Global:LogoPath" -Level "Info"
        } else {
            Write-AlarmLog -Message "Logo not found at: $Global:LogoPath - will use text-only header" -Level "Warning"
            $Global:LogoPath = $null
        }
        
        # Verify icon exists
        if ($Global:IconPath -and (Test-Path $Global:IconPath)) {
            Write-AlarmLog -Message "Icon found at: $Global:IconPath" -Level "Info"
        } else {
            Write-AlarmLog -Message "Icon not found at: $Global:IconPath - will use system default" -Level "Warning"
            $Global:IconPath = $null
        }
        
        # ===== ADD THIS LINE: Call debug function to log resource info =====
        Write-DebugInfo
        
    } catch {
        Write-AlarmLog -Message "Failed to initialize resources: $_" -Level "Error"
        $Global:LogoPath = $null
        $Global:IconPath = $null
    }
}

# Initialize resources
Initialize-Resources
#endregion

# Load configuration
function Load-Configuration {
    $configPath = Join-Path $Global:AppDataPath "config.xml"
    $defaultConfig = @{
        LogRetentionDays = 30
        CheckInterval = 5000
        AdvanceWarningMinutes = 10
        DefaultSound = "Windows Default"
        DefaultVoice = "IGRF-Bhramma (Male)"
        MinimizeToTray = $true
        ShowNotifications = $true
        Language = $Global:CurrentCulture.Name
        AlarmVolume = 100
        SoundRepeats = 3
    }
            
    if (Test-Path $configPath) {
        try {
            $Global:Config = Import-Clixml -Path $configPath
            
            # Check each default property and add if missing
            foreach ($key in $defaultConfig.Keys) {
                $propertyExists = $false
                foreach ($prop in $Global:Config.PSObject.Properties) {
                    if ($prop.Name -eq $key) {
                        $propertyExists = $true
                        break
                    }
                }
                if (-not $propertyExists) {
                    Write-AlarmLog -Message "Adding missing config property: $key" -Level "Info"
                    $Global:Config | Add-Member -MemberType NoteProperty -Name $key -Value $defaultConfig[$key] -Force
                }
            }
            
            # Remove deprecated properties if they exist
            $deprecatedProps = @("ShowSecondsInStatus", "EmailEnabled", "EmailAddress", "EmailServer", "SmtpUsername")
            foreach ($prop in $deprecatedProps) {
                $propertyExists = $false
                foreach ($p in $Global:Config.PSObject.Properties) {
                    if ($p.Name -eq $prop) {
                        $propertyExists = $true
                        break
                    }
                }
                if ($propertyExists) {
                    Write-AlarmLog -Message "Removing deprecated config property: $prop" -Level "Info"
                    $Global:Config.PSObject.Properties.Remove($prop)
                }
            }
            
        } catch {
            Write-AlarmLog -Message "Error loading config: $_" -Level "Warning"
            $Global:Config = [PSCustomObject]$defaultConfig
        }
    } else {
        $Global:Config = [PSCustomObject]$defaultConfig
        Save-Configuration
    }
}

#region Helper Functions

# Helper function to center a form over its owner
function Center-FormOverOwner {
    param(
        [System.Windows.Forms.Form]$Form,
        [System.Windows.Forms.Form]$Owner
    )
    
    if ($Owner -and !$Owner.IsDisposed) {
        $form.StartPosition = "Manual"
        $form.Location = New-Object System.Drawing.Point(
            [Math]::Max(0, $Owner.Location.X + ($Owner.Width - $form.Width) / 2),
            [Math]::Max(0, $Owner.Location.Y + ($Owner.Height - $form.Height) / 2)
        )
    } else {
        $form.StartPosition = "CenterScreen"
    }
}

function Get-LogoImage {
    # Check if we have a valid logo path
    if ($Global:LogoPath -and (Test-Path $Global:LogoPath)) {
        try {
            Write-AlarmLog -Message "Attempting to load logo from: $Global:LogoPath" -Level "Info"
            
            # Use FileStream to avoid file locking issues
            $fs = New-Object System.IO.FileStream($Global:LogoPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
            $img = [System.Drawing.Image]::FromStream($fs)
            $fs.Close()
            
            Write-AlarmLog -Message "Logo loaded successfully" -Level "Info"
            return $img
        } catch {
            Write-AlarmLog -Message "Failed to load logo image: $_" -Level "Warning"
            return $null
        }
    }
    
    # No logo available
    Write-AlarmLog -Message "No custom logo available at: $Global:LogoPath" -Level "Info"
    return $null
}

function Get-AppIcon {
    # Check if we have a valid icon path
    if ($Global:IconPath -and (Test-Path $Global:IconPath)) {
        try {
            Write-AlarmLog -Message "Attempting to load icon from: $Global:IconPath" -Level "Info"
            
            # For icon files, we can use Icon.ExtractAssociatedIcon
            $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Global:IconPath)
            
            Write-AlarmLog -Message "Icon loaded successfully" -Level "Info"
            return $icon
        } catch {
            Write-AlarmLog -Message "Failed to load custom icon: $_" -Level "Warning"
            Write-AlarmLog -Message "Using system default icon instead" -Level "Info"
            return [System.Drawing.SystemIcons]::Application
        }
    }
    
    # Fallback to system icon
    Write-AlarmLog -Message "No custom icon available at: $Global:IconPath, using system default" -Level "Info"
    return [System.Drawing.SystemIcons]::Application
}

function Create-Label {
    param($Text, $X, $Y, $Width, $Height, $FontSize = 9, $Bold = $false, $Color = "#2C3E50", $Alignment = "MiddleLeft", $Name = "")
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, $Height)
    $label.Font = New-Object System.Drawing.Font("Segoe UI", $FontSize, $(if($Bold){[System.Drawing.FontStyle]::Bold}else{[System.Drawing.FontStyle]::Regular}))
    $label.ForeColor = $Color
    $label.TextAlign = $Alignment
    if ($Name -ne "") { $label.Name = $Name }
    return $label
}

function Create-TextBox {
    param($X, $Y, $Width, $Height, $Name = "")
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point($X, $Y)
    $textBox.Size = New-Object System.Drawing.Size($Width, $Height)
    $textBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    if ($Name -ne "") { $textBox.Name = $Name }
    return $textBox
}

function Create-Button {
    param($Text, $X, $Y, $Width, $Height, $Color, $TextColor = "White", $Bold = $false, $Name = "")
    
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, $Height)
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 9, $(if($Bold){[System.Drawing.FontStyle]::Bold}else{[System.Drawing.FontStyle]::Regular}))
    $button.BackColor = $Color
    $button.ForeColor = $TextColor
    $button.FlatStyle = "Flat"
    $button.FlatAppearance.BorderSize = 0
    $button.FlatAppearance.MouseOverBackColor = $Color
    $button.Cursor = "Hand"
    if ($Name -ne "") { $button.Name = $Name }
    return $button
}

function Load-CustomSounds {
    $Global:CustomSounds = @()
    if (Test-Path $Global:CustomSoundsPath) {
        $soundFiles = Get-ChildItem -Path $Global:CustomSoundsPath -Include *.wav -Recurse
        foreach ($file in $soundFiles) {
            $Global:CustomSounds += [PSCustomObject]@{
                Name = $file.Name
                Path = $file.FullName
                Type = $file.Extension
                Size = $file.Length
            }
        }
    }
}

function Initialize-NotifyIcon {
    # CRITICAL FIX: Prevent multiple notify icons
    if ($Global:NotifyIcon -and !$Global:NotifyIcon.IsDisposed) {
        Write-AlarmLog -Message "Notify icon already exists, skipping initialization" -Level "Info"
        return
    }
    
    # Clean up any existing notify icon
    if ($Global:NotifyIcon) {
        try {
            $Global:NotifyIcon.Visible = $false
            $Global:NotifyIcon.Dispose()
        } catch {}
        $Global:NotifyIcon = $null
    }
    
    if ($Global:ContextMenu) {
        try {
            $Global:ContextMenu.Dispose()
        } catch {}
        $Global:ContextMenu = $null
    }
    
    Write-AlarmLog -Message "Creating new notify icon" -Level "Info"
    $Global:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
    
    $appIcon = Get-AppIcon
    if ($appIcon -ne $null) {
        try {
            $Global:NotifyIcon.Icon = $appIcon
            Write-AlarmLog -Message "Notify icon set successfully" -Level "Info"
        } catch {
            Write-AlarmLog -Message "Failed to set notify icon: $_" -Level "Warning"
            $Global:NotifyIcon.Icon = [System.Drawing.SystemIcons]::Application
        }
    } else {
        Write-AlarmLog -Message "No icon available for notify icon, using default" -Level "Info"
        $Global:NotifyIcon.Icon = [System.Drawing.SystemIcons]::Application
    }
    
    $Global:NotifyIcon.Text = "Windows Alarm Pro"
    $Global:NotifyIcon.Visible = $true
    
    # Create context menu
    $Global:ContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    
    # Create menu items
    $showItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $showItem.Text = "Show Window"
    $showItem.Add_Click({
        Show-MainWindow
    })
    
    $settingsItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $settingsItem.Text = "Settings"
    $settingsItem.Add_Click({
        Show-SettingsDialog
    })
    
    $exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitItem.Text = "Exit"
    $exitItem.Add_Click({
        $Global:ManualClose = $true
        $Global:ApplicationRunning = $false
        $Global:IsClosing = $true
        Stop-CurrentAlarm
        Save-Alarms
        Save-Configuration
        Cleanup-Resources
        [System.Windows.Forms.Application]::Exit()
    })
    
    # Add items to context menu
    $menuItems = @($showItem, $settingsItem, $exitItem)
    $Global:ContextMenu.Items.AddRange($menuItems)
    
    $Global:NotifyIcon.ContextMenuStrip = $Global:ContextMenu
    
    $Global:NotifyIcon.Add_MouseClick({
        param($sender, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            Toggle-WindowVisibility
        } elseif ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            $Global:NotifyIcon.ContextMenuStrip.Show([System.Windows.Forms.Cursor]::Position)
        }
    })
    
    $Global:NotifyIcon.Add_MouseDoubleClick({
        param($sender, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            Show-MainWindow
        }
    })
    
    $Global:NotifyIcon.Add_BalloonTipClicked({
        Show-MainWindow
    })
}

function Cleanup-Resources {
    try {
        if ($Global:AlarmTimer) {
            $Global:AlarmTimer.Stop()
            $Global:AlarmTimer.Dispose()
            $Global:AlarmTimer = $null
        }
        
        if ($Global:StatusTimer) {
            $Global:StatusTimer.Stop()
            $Global:StatusTimer.Dispose()
            $Global:StatusTimer = $null
        }
        
        if ($Global:SoundPlaybackTimer) {
            $Global:SoundPlaybackTimer.Stop()
            $Global:SoundPlaybackTimer.Dispose()
            $Global:SoundPlaybackTimer = $null
        }
        
        # FIX: Properly dispose BalloonTimer
        if ($Global:BalloonTimer) {
            try {
                $Global:BalloonTimer.Stop()
                $Global:BalloonTimer.Dispose()
            } catch {}
            $Global:BalloonTimer = $null
        }
        
        if ($Global:SoundPlayer) {
            $Global:SoundPlayer.Stop()
            $Global:SoundPlayer.Dispose()
            $Global:SoundPlayer = $null
        }
        
        if ($Global:NotifyIcon) {
            $Global:NotifyIcon.Visible = $false
            $Global:NotifyIcon.Dispose()
            $Global:NotifyIcon = $null
        }
        
        if ($Global:ContextMenu) {
            $Global:ContextMenu.Dispose()
            $Global:ContextMenu = $null
        }
        
        if ($Global:CurrentAlarmForm -and !$Global:CurrentAlarmForm.IsDisposed) {
            $Global:CurrentAlarmForm.Close()
            $Global:CurrentAlarmForm.Dispose()
            $Global:CurrentAlarmForm = $null
        }
        
        if ($Global:CurrentAdvanceForm -and !$Global:CurrentAdvanceForm.IsDisposed) {
            $Global:CurrentAdvanceForm.Close()
            $Global:CurrentAdvanceForm.Dispose()
            $Global:CurrentAdvanceForm = $null
        }
        
        # Clean up temp resource files
		try {
			if ($Global:TempLogoPath -and (Test-Path $Global:TempLogoPath)) {
				Remove-Item $Global:TempLogoPath -Force -ErrorAction SilentlyContinue
			}
			if ($Global:TempIconPath -and (Test-Path $Global:TempIconPath)) {
				Remove-Item $Global:TempIconPath -Force -ErrorAction SilentlyContinue
			}
		} catch {
			# Silently fail during cleanup
		}
    } catch {
        # Silently fail during cleanup
    }
}

function Show-MainWindow {
    if ($Global:MainForm -and !$Global:MainForm.IsDisposed) {
        $Global:MainForm.Show()
        $Global:MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $Global:MainForm.BringToFront()
        $Global:MainForm.Activate()
        $Global:MainForm.TopMost = $true
        $Global:MainForm.TopMost = $false
        $Global:MainForm.Focus()
        $Global:MainForm.Refresh()
        
        [System.Windows.Forms.Application]::DoEvents() | Out-Null
        Update-Status "Window restored"
        $Global:NotifyIcon.Visible = $true
    }
}

function Toggle-WindowVisibility {
    if ($Global:MainForm -and !$Global:MainForm.IsDisposed) {
        if ($Global:MainForm.Visible) {
            $Global:ManualMinimize = $true
            $Global:MainForm.Hide()
            $Global:MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
            $Global:ManualMinimize = $false
        } else {
            Show-MainWindow
        }
    }
}

function Show-SettingsDialog {
    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = "Windows Alarm Pro - Settings"
    $settingsForm.Size = New-Object System.Drawing.Size(620, 680)
    $settingsForm.StartPosition = "CenterScreen"
    $settingsForm.FormBorderStyle = "FixedDialog"
    $settingsForm.MaximizeBox = $false
    $settingsForm.MinimizeBox = $false
    $settingsForm.BackColor = "#F0F4F8"
    $settingsForm.ShowIcon = $true
    
    $appIcon = Get-AppIcon
	if ($appIcon -ne $null) {
		try {
			$settingsForm.Icon = $appIcon
			Write-AlarmLog -Message "Settings dialog icon set successfully" -Level "Info"
		} catch {
			Write-AlarmLog -Message "Failed to set settings dialog icon: $_" -Level "Warning"
		}
	} else {
		Write-AlarmLog -Message "No icon available for settings dialog" -Level "Info"
	}
    
    # ============================================================================
    # GENERAL SETTINGS GROUP
    # ============================================================================
    $generalGroup = New-Object System.Windows.Forms.GroupBox
    $generalGroup.Text = "General Settings"
    $generalGroup.Location = New-Object System.Drawing.Point(20, 20)
    $generalGroup.Size = New-Object System.Drawing.Size(565, 260)
    $generalGroup.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    
    $yPos = 25
    
    # Row 1: Check Interval
    $lblCheck = Create-Label -Text "Check Interval (ms):" -X 15 -Y $yPos -Width 140 -Height 25 -FontSize 9
    $generalGroup.Controls.Add($lblCheck)
    
    $numCheck = New-Object System.Windows.Forms.NumericUpDown
    $numCheck.Location = New-Object System.Drawing.Point(160, $yPos)
    $numCheck.Size = New-Object System.Drawing.Size(100, 25)
    $numCheck.Minimum = 1000
    $numCheck.Maximum = 30000
    $checkValue = [Math]::Max(1000, [Math]::Min(30000, $Global:Config.CheckInterval))
    $numCheck.Value = $checkValue
    $numCheck.Increment = 1000
    $numCheck.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $generalGroup.Controls.Add($numCheck)
    
    $yPos += 30
    
    # Row 2: Advance Warning
    $lblAdvance = Create-Label -Text "Advance Warning (min):" -X 15 -Y $yPos -Width 140 -Height 25 -FontSize 9
    $generalGroup.Controls.Add($lblAdvance)
    
    $numAdvance = New-Object System.Windows.Forms.NumericUpDown
    $numAdvance.Location = New-Object System.Drawing.Point(160, $yPos)
    $numAdvance.Size = New-Object System.Drawing.Size(100, 25)
    $numAdvance.Minimum = 1
    $numAdvance.Maximum = 60
    $advanceValue = [Math]::Max(1, [Math]::Min(60, $Global:Config.AdvanceWarningMinutes))
    $numAdvance.Value = $advanceValue
    $numAdvance.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $generalGroup.Controls.Add($numAdvance)
    
    $yPos += 30
    
    # Row 3: Log Retention
    $lblLogDays = Create-Label -Text "Log Retention (days):" -X 15 -Y $yPos -Width 140 -Height 25 -FontSize 9
    $generalGroup.Controls.Add($lblLogDays)
    
    $numLogDays = New-Object System.Windows.Forms.NumericUpDown
    $numLogDays.Location = New-Object System.Drawing.Point(160, $yPos)
    $numLogDays.Size = New-Object System.Drawing.Size(100, 25)
    $numLogDays.Minimum = 1
    $numLogDays.Maximum = 365
    $logValue = [Math]::Max(1, [Math]::Min(365, $Global:Config.LogRetentionDays))
    $numLogDays.Value = $logValue
    $numLogDays.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $generalGroup.Controls.Add($numLogDays)
    
    $yPos += 30
    
    # Row 4: Volume
    $lblVolume = Create-Label -Text "Volume (%):" -X 15 -Y $yPos -Width 140 -Height 25 -FontSize 9
    $generalGroup.Controls.Add($lblVolume)
    
    $numVolume = New-Object System.Windows.Forms.NumericUpDown
    $numVolume.Location = New-Object System.Drawing.Point(160, $yPos)
    $numVolume.Size = New-Object System.Drawing.Size(100, 25)
    $numVolume.Minimum = 1
    $numVolume.Maximum = 100
    $volumeValue = [Math]::Max(1, [Math]::Min(100, $Global:Config.AlarmVolume))
    $numVolume.Value = $volumeValue
    $numVolume.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $generalGroup.Controls.Add($numVolume)
    
    $yPos += 35
    
    # Row 5: Sound Repeats
    $lblSoundRepeats = Create-Label -Text "Sound Repeats:" -X 15 -Y $yPos -Width 100 -Height 25 -FontSize 9
    $generalGroup.Controls.Add($lblSoundRepeats)

    $numSoundRepeats = New-Object System.Windows.Forms.NumericUpDown
    $numSoundRepeats.Location = New-Object System.Drawing.Point(120, $yPos)
    $numSoundRepeats.Size = New-Object System.Drawing.Size(60, 25)
    $numSoundRepeats.Minimum = 1
    $numSoundRepeats.Maximum = 10
    $soundRepeatsValue = if ($Global:Config.PSObject.Properties.Name -contains "SoundRepeats") { 
        [Math]::Max(1, [Math]::Min(10, $Global:Config.SoundRepeats)) 
    } else { 3 }
    $numSoundRepeats.Value = $soundRepeatsValue
    $numSoundRepeats.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $generalGroup.Controls.Add($numSoundRepeats)
    
    $yPos += 35
    
    # Row 6: Minimize to tray checkbox
    $chkMinimize = New-Object System.Windows.Forms.CheckBox
    $chkMinimize.Text = "Minimize to system tray"
    $chkMinimize.Location = New-Object System.Drawing.Point(15, $yPos)
    $chkMinimize.Size = New-Object System.Drawing.Size(180, 25)
    $chkMinimize.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkMinimize.Checked = $Global:Config.MinimizeToTray
    $chkMinimize.TextAlign = "MiddleLeft"
    $generalGroup.Controls.Add($chkMinimize)
    
    $yPos += 25
    
    # Row 7: Show popup notifications checkbox
    $chkNotifications = New-Object System.Windows.Forms.CheckBox
    $chkNotifications.Text = "Show popup notifications"
    $chkNotifications.Location = New-Object System.Drawing.Point(15, $yPos)
    $chkNotifications.Size = New-Object System.Drawing.Size(200, 25)
    $chkNotifications.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkNotifications.Checked = $Global:Config.ShowNotifications
    $chkNotifications.TextAlign = "MiddleLeft"
    $generalGroup.Controls.Add($chkNotifications)
    
    $settingsForm.Controls.Add($generalGroup)
    
    # ============================================================================
    # DEFAULT SETTINGS GROUP
    # ============================================================================
    $defaultGroup = New-Object System.Windows.Forms.GroupBox
    $defaultGroup.Text = "Default Settings"
    $defaultGroup.Location = New-Object System.Drawing.Point(20, 295)
    $defaultGroup.Size = New-Object System.Drawing.Size(565, 120)
    $defaultGroup.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    
    $yPos = 25
    
    # First row: Default Ring and Default Voice
    $lblDefaultDuration = Create-Label -Text "Default Ring:" -X 15 -Y $yPos -Width 100 -Height 25 -FontSize 9
    $defaultGroup.Controls.Add($lblDefaultDuration)
    
    $cmbDefaultDuration = New-Object System.Windows.Forms.ComboBox
    $cmbDefaultDuration.Location = New-Object System.Drawing.Point(120, $yPos)
    $cmbDefaultDuration.Size = New-Object System.Drawing.Size(150, 25)
    $cmbDefaultDuration.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cmbDefaultDuration.DropDownStyle = "DropDownList"
    $Global:RingDurations.Keys | ForEach-Object { $cmbDefaultDuration.Items.Add($_) }
    $cmbDefaultDuration.SelectedItem = "Until Acknowledged"
    $defaultGroup.Controls.Add($cmbDefaultDuration)
    
    $lblDefaultVoice = Create-Label -Text "Default Voice:" -X 300 -Y $yPos -Width 100 -Height 25 -FontSize 9
    $defaultGroup.Controls.Add($lblDefaultVoice)
    
    $cmbDefaultVoice = New-Object System.Windows.Forms.ComboBox
    $cmbDefaultVoice.Location = New-Object System.Drawing.Point(400, $yPos)
    $cmbDefaultVoice.Size = New-Object System.Drawing.Size(140, 25)
    $cmbDefaultVoice.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cmbDefaultVoice.DropDownStyle = "DropDownList"
    $cmbDefaultVoice.Items.Add("IGRF-Bhramma (Male)")
    $cmbDefaultVoice.Items.Add("IGRF-Saraswathi (Female)")
    $cmbDefaultVoice.Items.Add("No Voice")
    $cmbDefaultVoice.SelectedItem = $Global:Config.DefaultVoice
    $defaultGroup.Controls.Add($cmbDefaultVoice)
    
    $yPos += 30
    
    # Second row: Default Repeat
    $lblDefaultRecurrence = Create-Label -Text "Default Repeat:" -X 15 -Y $yPos -Width 100 -Height 25 -FontSize 9
    $defaultGroup.Controls.Add($lblDefaultRecurrence)
    
    $cmbDefaultRecurrence = New-Object System.Windows.Forms.ComboBox
    $cmbDefaultRecurrence.Location = New-Object System.Drawing.Point(120, $yPos)
    $cmbDefaultRecurrence.Size = New-Object System.Drawing.Size(150, 25)
    $cmbDefaultRecurrence.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cmbDefaultRecurrence.DropDownStyle = "DropDownList"
    $Global:RecurrenceTypes.Keys | ForEach-Object { $cmbDefaultRecurrence.Items.Add($_) }
    $cmbDefaultRecurrence.SelectedItem = "One Time"
    $defaultGroup.Controls.Add($cmbDefaultRecurrence)
    
    $yPos += 30
    
    # Third row: Informational note
    $lblSoundInfo = Create-Label -Text "Note: Sound repeats can be configured in General Settings" -X 15 -Y $yPos -Width 350 -Height 20 -FontSize 8 -Color "#7F8C8D"
    $defaultGroup.Controls.Add($lblSoundInfo)
    
    $settingsForm.Controls.Add($defaultGroup)
    
    # ============================================================================
    # BUTTONS - Properly sized and centered
    # ============================================================================
    $buttonY = 435
    
    # Calculate center position for buttons (form width = 620)
    $btnWidth = 100
    $btnSpacing = 20
    $totalButtonsWidth = ($btnWidth * 2) + $btnSpacing
    $startX = (620 - $totalButtonsWidth) / 2
    
    $btnSave = Create-Button -Text "SAVE" -X $startX -Y $buttonY -Width $btnWidth -Height 35 -Color "#27AE60" -Bold $true
    $btnSave.Add_Click({
        
        # Save general settings
        $Global:Config.CheckInterval = [int]$numCheck.Value
        $Global:Config.AdvanceWarningMinutes = [int]$numAdvance.Value
        $Global:Config.LogRetentionDays = [int]$numLogDays.Value
        $Global:Config.AlarmVolume = [int]$numVolume.Value
        $Global:Config.MinimizeToTray = $chkMinimize.Checked
        $Global:Config.ShowNotifications = $chkNotifications.Checked
        $Global:Config.DefaultVoice = $cmbDefaultVoice.SelectedItem
        $Global:Config.SoundRepeats = [int]$numSoundRepeats.Value
        
        # Update timer interval
        $Global:AlarmTimer.Interval = $Global:Config.CheckInterval
        
        # Save and clean up
        Save-Configuration
        Cleanup-OldLogs
        $settingsForm.Close()
        Update-Status "Settings saved"
        Write-AlarmLog -Message "Settings saved via dialog" -Level "Info"
    })
    
    $btnCancel = Create-Button -Text "CANCEL" -X ($startX + $btnWidth + $btnSpacing) -Y $buttonY -Width $btnWidth -Height 35 -Color "#E74C3C"
    $btnCancel.Add_Click({ 
        Write-AlarmLog -Message "Settings dialog cancelled" -Level "Info"
        $settingsForm.Close() 
    })
    
    $settingsForm.Controls.Add($btnSave)
    $settingsForm.Controls.Add($btnCancel)
    
    # Show the dialog
    $settingsForm.ShowDialog() | Out-Null
}

function Get-VoiceName {
    param([string]$VoiceType)
    
    if ($VoiceType -eq "IGRF-Bhramma (Male)") {
        if ($Global:AvailableMaleVoices -and $Global:AvailableMaleVoices.Count -gt 0) {
            # Rotate through available male voices
            $index = $Global:LastUsedVoiceIndex.Male
            $voiceName = $Global:AvailableMaleVoices[$index]
            
            # Update index for next time (round-robin)
            $Global:LastUsedVoiceIndex.Male = ($index + 1) % $Global:AvailableMaleVoices.Count
            
            # FIXED: Use ${index} to avoid # being interpreted as comment
            Write-AlarmLog -Message "Selected male voice #${index}: $voiceName" -Level "Info"
            return $voiceName
        }
    } elseif ($VoiceType -eq "IGRF-Saraswathi (Female)") {
        if ($Global:AvailableFemaleVoices -and $Global:AvailableFemaleVoices.Count -gt 0) {
            # Rotate through available female voices
            $index = $Global:LastUsedVoiceIndex.Female
            $voiceName = $Global:AvailableFemaleVoices[$index]
            
            # Update index for next time (round-robin)
            $Global:LastUsedVoiceIndex.Female = ($index + 1) % $Global:AvailableFemaleVoices.Count
            
            # FIXED: Use ${index} to avoid # being interpreted as comment
            Write-AlarmLog -Message "Selected female voice #${index}: $voiceName" -Level "Info"
            return $voiceName
        }
    }
    
    return $null
}

function Cleanup-OldLogs {
    $logFile = Join-Path $Global:AppDataPath "alarm_log.txt"
    if (Test-Path $logFile) {
        $fileInfo = Get-Item $logFile
        if ($fileInfo.Length -gt $Global:LogMaxSize) {
            $backupFile = Join-Path $Global:AppDataPath "alarm_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            Move-Item -Path $logFile -Destination $backupFile -Force
            
            $oldLogs = Get-ChildItem -Path $Global:AppDataPath -Filter "alarm_log_*.txt" | Sort-Object LastWriteTime -Descending
            if ($oldLogs.Count -gt $Global:MaxLogFiles) {
                $oldLogs | Select-Object -Skip $Global:MaxLogFiles | Remove-Item -Force
            }
        }
    }
}

function Cleanup-SpeechJobs {
    # Clean up completed speech runspaces and terminate any that are still running
    try {
        if ($Global:SpeechRunspaces) {
            $completed = @()
            foreach ($jobId in $Global:SpeechRunspaces.Keys) {
                $job = $Global:SpeechRunspaces[$jobId]
                try {
                    # Handle PowerShell runspaces
                    if ($job.ContainsKey('PowerShell') -and $job.PowerShell) {
                        if ($job.Handle -and $job.Handle.IsCompleted) {
                            $job.PowerShell.EndInvoke($job.Handle)
                            $job.PowerShell.Dispose()
                            $completed += $jobId
                        } elseif ($Global:SoundCancellation) {
                            try {
                                $job.PowerShell.Stop()
                                $job.PowerShell.Dispose()
                            } catch {}
                            $completed += $jobId
                        }
                    }
                } catch {
                    Write-AlarmLog -Message "Error cleaning up job ${jobId}: $_" -Level "Warning"
                    try { 
                        if ($job.ContainsKey('PowerShell') -and $job.PowerShell) { 
                            $job.PowerShell.Dispose() 
                        }
                    } catch {}
                    $completed += $jobId
                }
            }
            foreach ($jobId in $completed) {
                $Global:SpeechRunspaces.Remove($jobId)
            }
        }
        
        # Also clean up any background jobs (legacy)
        $jobs = Get-Job -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Voice*" -or $_.Name -like "*Advance*" }
        foreach ($job in $jobs) {
            try {
                if ($job.State -eq "Running") {
                    $job.StopJob() | Out-Null
                }
                $job | Remove-Job -Force -ErrorAction SilentlyContinue
            } catch {}
        }
    } catch {
        Write-AlarmLog -Message "Error cleaning up speech jobs: $_" -Level "Warning"
    }
}

function Save-Configuration {
    $configPath = Join-Path $Global:AppDataPath "config.xml"
    $Global:Config | Export-Clixml -Path $configPath -Force
}



# Initialize
Load-Configuration
$Global:AlarmTimer.Interval = $Global:Config.CheckInterval

# Enhanced voice detection
function Initialize-Voices {
    try {
        $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
        $availableVoices = $synth.GetInstalledVoices()
        
        $Global:AvailableMaleVoices = @()
        $Global:AvailableFemaleVoices = @()
        $Global:AvailableVoices = @()
        
        foreach ($voice in $availableVoices) {
            if ($voice.Enabled) {
                $voiceInfo = $voice.VoiceInfo
                $Global:AvailableVoices += $voiceInfo
                
                if ($voiceInfo.Gender -eq [System.Speech.Synthesis.VoiceGender]::Male) {
                    $Global:AvailableMaleVoices += $voiceInfo.Name
                }
                elseif ($voiceInfo.Gender -eq [System.Speech.Synthesis.VoiceGender]::Female) {
                    $Global:AvailableFemaleVoices += $voiceInfo.Name
                }
            }
        }
        
        $synth.Dispose()
        
        # Log voice availability
        Write-AlarmLog -Message "Found $($Global:AvailableMaleVoices.Count) male voices and $($Global:AvailableFemaleVoices.Count) female voices" -Level "Info"
        
    } catch {
        Write-AlarmLog -Message "Voice initialization failed: $_" -Level "Warning"
        $Global:AvailableMaleVoices = @()
        $Global:AvailableFemaleVoices = @()
        $Global:AvailableVoices = @()
    }
}

# Initialize voices
Initialize-Voices

# Sound options
$Global:SoundOptions = @{
    "Windows Default" = "Asterisk"
    "Windows Beep" = "Exclamation"
    "Windows Critical" = "Hand"
    "Windows Notification" = "Notification"
    "Classic Bell" = "Bell"
    "Continuous Ring" = "Ring"
    "Alert Tone" = "Alert"
    "Soft Chime" = "Chime"
    "Custom Sound" = "Custom"
    "No Sound" = "None"
}

# Recurrence patterns
$Global:RecurrenceTypes = @{
    "One Time" = "Once"
    "Daily" = "Daily"
    "Weekly" = "Weekly"
    "Weekdays (Mon-Fri)" = "Weekdays"
    "Weekends" = "Weekends"
    "Monthly" = "Monthly"
    "Yearly" = "Yearly"
    "Custom Days" = "Custom"
}

# Ring duration options
$Global:RingDurations = @{
    "30 Seconds" = 30
    "1 Minute" = 60
    "2 Minutes" = 120
    "3 Minutes" = 180
    "5 Minutes" = 300
    "Until Acknowledged" = 0
}
#endregion

#region Main Form Functions
function Show-MainForm {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    
    $formWidth = 1000
    $formHeight = 670
    
    $Global:MainForm = New-Object System.Windows.Forms.Form
    $Global:MainForm.Text = "Windows Alarm Pro v1.0 - IGRF Pvt. Ltd."
    $Global:MainForm.Size = New-Object System.Drawing.Size($formWidth, $formHeight)
    $Global:MainForm.StartPosition = "CenterScreen"
    $Global:MainForm.FormBorderStyle = "FixedSingle"
    $Global:MainForm.MaximizeBox = $false
    $Global:MainForm.BackColor = "#F0F4F8"
    $Global:MainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	
	# Add keyboard shortcuts
	$Global:MainForm.KeyPreview = $true
	$Global:MainForm.Add_KeyDown({
		param($sender, $e)
		
		# Ctrl+N - New alarm (focus name field)
		if ($e.Control -and $e.KeyCode -eq "N") {
			$txtName = $Global:MainForm.Controls.Find("txtName", $true)[0]
			if ($txtName) { $txtName.Focus() }
			$e.SuppressKeyPress = $true
		}
		# Ctrl+S - Open settings
		elseif ($e.Control -and $e.KeyCode -eq "S") {
			Show-SettingsDialog
			$e.SuppressKeyPress = $true
		}
		# Ctrl+E - Export alarms
		elseif ($e.Control -and $e.KeyCode -eq "E") {
			Export-Alarms
			$e.SuppressKeyPress = $true
		}
		# Ctrl+I - Import alarms
		elseif ($e.Control -and $e.KeyCode -eq "I") {
			Import-Alarms
			$e.SuppressKeyPress = $true
		}
		# F5 - Refresh list
		elseif ($e.KeyCode -eq "F5") {
			Refresh-AlarmList
			$e.SuppressKeyPress = $true
		}
		# F1 - Help/About
		elseif ($e.KeyCode -eq "F1") {
			[System.Windows.Forms.MessageBox]::Show(
				"Windows Alarm Pro v1.0`n`n© 2026 IGRF Pvt. Ltd.`nhttps://igrf.co.in/en/software/`n`nKeyboard Shortcuts:`nCtrl+N - New Alarm`nCtrl+S - Settings`nCtrl+E - Export`nCtrl+I - Import`nF5 - Refresh`nF1 - Help",
				"About Windows Alarm Pro",
				"OK",
				"Information"
			)
			$e.SuppressKeyPress = $true
		}
	})
    
    $appIcon = Get-AppIcon
	if ($appIcon -ne $null) {
		try {
			$Global:MainForm.Icon = $appIcon
			Write-AlarmLog -Message "Main form icon set successfully" -Level "Info"
		} catch {
			Write-AlarmLog -Message "Failed to set main form icon: $_" -Level "Warning"
		}
	} else {
		Write-AlarmLog -Message "No icon available for main form" -Level "Info"
	}
    
    Initialize-NotifyIcon
    
	$Global:MainForm.Add_Resize({
        if ($Global:ManualMinimize) {
            return
        }
        
        if ($Global:MainForm.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
            if ($Global:Config.MinimizeToTray) {
                $Global:MainForm.Hide()
                
                # CRITICAL FIX: Ensure only one notify icon is visible
                if ($Global:NotifyIcon -and !$Global:NotifyIcon.IsDisposed) {
                    $Global:NotifyIcon.Visible = $true
                } else {
                    Initialize-NotifyIcon
                }
                
                # Use a timer to show balloon tip with a slight delay
                if ($Global:BalloonTimer) {
                    try { 
                        $Global:BalloonTimer.Stop() 
                        $Global:BalloonTimer.Dispose()
                    } catch {}
                    $Global:BalloonTimer = $null
                }
                
                $Global:BalloonTimer = New-Object System.Windows.Forms.Timer
                $Global:BalloonTimer.Interval = 500
                $Global:BalloonTimer.Add_Tick({
                    $Global:BalloonTimer.Stop()
                    try {
                        if ($Global:MainForm -and !$Global:MainForm.IsDisposed -and !$Global:MainForm.Visible -and
                            $Global:NotifyIcon -and !$Global:NotifyIcon.IsDisposed) {
                            $Global:NotifyIcon.ShowBalloonTip(5000, "Windows Alarm Pro", 
                                "Application is running in system tray. Left-click icon to show/hide, double-click to restore.", 
                                [System.Windows.Forms.ToolTipIcon]::Info)
                            Write-AlarmLog -Message "Balloon tip shown from Resize" -Level "Info"
                        }
                    } catch {
                        Write-AlarmLog -Message "Balloon tip failed in Resize: $_" -Level "Warning"
                    }
                })
                $Global:BalloonTimer.Start()
                
                Write-AlarmLog -Message "Application minimized to tray (minimize button)"
            }
        }
    })
    
    # NEW CODE: Handle close button (X) to minimize to tray
    $Global:MainForm.Add_FormClosing({
        param($sender, $e)
        
        # Check if this is a manual close (via Exit menu or actual application exit)
        if ($Global:ManualClose) {
            # Allow the form to close
            return
        }
        
        # Check if we're actually closing the application
        if ($e.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) {
            # Cancel the close event
            $e.Cancel = $true
            
            # Minimize to tray instead
            $Global:ManualMinimize = $true
            $Global:MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
            $Global:MainForm.Hide()
            $Global:ManualMinimize = $false
            
            # Show notification that app is still running
            if ($Global:NotifyIcon -and !$Global:NotifyIcon.IsDisposed) {
                $Global:NotifyIcon.ShowBalloonTip(3000, "Windows Alarm Pro", 
                    "Application is still running in the system tray. Right-click the icon for options.", 
                    [System.Windows.Forms.ToolTipIcon]::Info)
            }
            
            Write-AlarmLog -Message "Window closed via X button - minimized to tray" -Level "Info"
        }
    })
    
    $Global:MainForm.Add_Shown({
    # CRITICAL FIX: Ensure no duplicate notify icons
		if ($Global:NotifyIcon -and !$Global:NotifyIcon.IsDisposed) {
			# Already have a notify icon, just ensure it's visible
			$Global:NotifyIcon.Visible = $true
		} else {
			Initialize-NotifyIcon
		}
		
		$Global:MainForm.Activate()
		Update-DateTimeDisplay
		Check-MissedAlarms
	})
    
    # Header Panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Size = New-Object System.Drawing.Size(($formWidth - 20), 80)
    $headerPanel.Location = New-Object System.Drawing.Point(10, 10)
    $headerPanel.BackColor = "#2C3E50"
    $headerPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $headerPanel.Name = "headerPanel"
    
    $logo = Get-LogoImage
    if ($logo) {
        $pictureBox = New-Object System.Windows.Forms.PictureBox
        $pictureBox.Image = $logo
        $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $pictureBox.Size = New-Object System.Drawing.Size(60, 60)
        $pictureBox.Location = New-Object System.Drawing.Point(15, 10)
        $pictureBox.Name = "pictureBox"
        $headerPanel.Controls.Add($pictureBox)
        
        $companyLabel = New-Object System.Windows.Forms.Label
        $companyLabel.Text = "Windows Alarm Pro"
        $companyLabel.ForeColor = "#ECF0F1"
        $companyLabel.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
        $companyLabel.Size = New-Object System.Drawing.Size(350, 35)
        $companyLabel.Location = New-Object System.Drawing.Point(85, 15)
        $companyLabel.Name = "companyLabel"
        $headerPanel.Controls.Add($companyLabel)
        
        $taglineLabel = New-Object System.Windows.Forms.Label
        $taglineLabel.Text = "Professional Time Management Solution"
        $taglineLabel.ForeColor = "#BDC3C7"
        $taglineLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $taglineLabel.Size = New-Object System.Drawing.Size(250, 20)
        $taglineLabel.Location = New-Object System.Drawing.Point(85, 50)
        $taglineLabel.Name = "taglineLabel"
        $headerPanel.Controls.Add($taglineLabel)
    } else {
        $companyLabel = New-Object System.Windows.Forms.Label
        $companyLabel.Text = "Windows Alarm Pro"
        $companyLabel.ForeColor = "#ECF0F1"
        $companyLabel.Font = New-Object System.Drawing.Font("Segoe UI", 22, [System.Drawing.FontStyle]::Bold)
        $companyLabel.Size = New-Object System.Drawing.Size(450, 45)
        $companyLabel.Location = New-Object System.Drawing.Point(20, 18)
        $companyLabel.Name = "companyLabel"
        $headerPanel.Controls.Add($companyLabel)
    }
    
    $infoPanel = New-Object System.Windows.Forms.Panel
    $infoPanel.Size = New-Object System.Drawing.Size(300, 60)
    $infoPanel.Location = New-Object System.Drawing.Point(($formWidth - 330), 10)
    $infoPanel.BackColor = "#34495E"
    $infoPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $infoPanel.Name = "infoPanel"
    
    $devLabel = New-Object System.Windows.Forms.Label
    $devLabel.Text = "IGRF Pvt. Ltd."
    $devLabel.ForeColor = "#3498DB"
    $devLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $devLabel.Size = New-Object System.Drawing.Size(130, 25)
    $devLabel.Location = New-Object System.Drawing.Point(10, 5)
    $devLabel.Name = "devLabel"
    $infoPanel.Controls.Add($devLabel)
    
    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "Version 1.0 (2026)"
    $versionLabel.ForeColor = "#ECF0F1"
    $versionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $versionLabel.Size = New-Object System.Drawing.Size(110, 20)
    $versionLabel.Location = New-Object System.Drawing.Point(10, 30)
    $versionLabel.Name = "versionLabel"
    $infoPanel.Controls.Add($versionLabel)
    
    $linkLabel = New-Object System.Windows.Forms.LinkLabel
    $linkLabel.Text = "IGRF-DITI-SOFTWARE"
    $linkLabel.LinkColor = "#3498DB"
    $linkLabel.ActiveLinkColor = "#2980B9"
    $linkLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Underline)
    $linkLabel.Size = New-Object System.Drawing.Size(130, 20)
    $linkLabel.Location = New-Object System.Drawing.Point(140, 33)
    $linkLabel.Name = "linkLabel"
    $linkLabel.Add_LinkClicked({
        try {
            Start-Process "https://igrf.co.in/en/software/"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Unable to open browser. Please visit: https://igrf.co.in/en/software/", "Link Error", "OK", "Information")
        }
    })
    $infoPanel.Controls.Add($linkLabel)
    
    $headerPanel.Controls.Add($infoPanel)
    $Global:MainForm.Controls.Add($headerPanel)
    
	# Left Panel - Settings
	$leftPanel = New-Object System.Windows.Forms.Panel
	$leftPanel.Size = New-Object System.Drawing.Size(400, 480)
	$leftPanel.Location = New-Object System.Drawing.Point(10, 100)
	$leftPanel.BackColor = "#ECF0F1"
	$leftPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
	$leftPanel.Name = "leftPanel"

	$settingsTitle = Create-Label -Text "ALARM SETTINGS" -X 15 -Y 10 -Width 370 -Height 25 -FontSize 14 -Bold $true -Color "#2C3E50" -Alignment "MiddleLeft" -Name "lblSettingsTitle"
	$leftPanel.Controls.Add($settingsTitle)

	# Description Row
	$lblName = Create-Label -Text "Description:" -X 15 -Y 45 -Width 80 -Height 22 -FontSize 9 -Name "lblName"
	$leftPanel.Controls.Add($lblName)

	$txtName = Create-TextBox -X 100 -Y 45 -Width 280 -Height 22 -Name "txtName"
	$txtName.MaxLength = 100
	$leftPanel.Controls.Add($txtName)

	# Date Row
	$lblDate = Create-Label -Text "Date:" -X 15 -Y 75 -Width 80 -Height 22 -FontSize 9 -Name "lblDate"
	$leftPanel.Controls.Add($lblDate)

	$dtpDate = New-Object System.Windows.Forms.DateTimePicker
	$dtpDate.Location = New-Object System.Drawing.Point(100, 75)
	$dtpDate.Size = New-Object System.Drawing.Size(280, 22)
	$dtpDate.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	$dtpDate.MinDate = (Get-Date).Date
	$dtpDate.Value = (Get-Date).Date
	$dtpDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
	$dtpDate.Name = "dtpDate"
	$leftPanel.Controls.Add($dtpDate)

	# Time Row
	$lblTime = Create-Label -Text "Time:" -X 15 -Y 105 -Width 80 -Height 22 -FontSize 9 -Name "lblTime"
	$leftPanel.Controls.Add($lblTime)

	$dtpTime = New-Object System.Windows.Forms.DateTimePicker
	$dtpTime.Location = New-Object System.Drawing.Point(100, 105)
	$dtpTime.Size = New-Object System.Drawing.Size(280, 22)
	$dtpTime.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	$dtpTime.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
	$dtpTime.ShowUpDown = $true
	$dtpTime.Value = Get-Date -Minute (Get-Date).Minute -Second 0
	$dtpTime.Name = "dtpTime"
	$leftPanel.Controls.Add($dtpTime)

	# Recurrence Row
	$lblRecurrence = Create-Label -Text "Repeat:" -X 15 -Y 135 -Width 80 -Height 22 -FontSize 9 -Name "lblRecurrence"
	$leftPanel.Controls.Add($lblRecurrence)

	$cmbRecurrence = New-Object System.Windows.Forms.ComboBox
	$cmbRecurrence.Location = New-Object System.Drawing.Point(100, 135)
	$cmbRecurrence.Size = New-Object System.Drawing.Size(280, 22)
	$cmbRecurrence.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	$cmbRecurrence.DropDownStyle = "DropDownList"
	$Global:RecurrenceTypes.Keys | ForEach-Object { $cmbRecurrence.Items.Add($_) }
	$cmbRecurrence.SelectedIndex = 0
	$cmbRecurrence.Name = "cmbRecurrence"
	$leftPanel.Controls.Add($cmbRecurrence)

	# Category Row
	$lblCategory = Create-Label -Text "Category:" -X 15 -Y 160 -Width 80 -Height 22 -FontSize 9 -Name "lblCategory"
	$leftPanel.Controls.Add($lblCategory)

	$cmbCategory = New-Object System.Windows.Forms.ComboBox
	$cmbCategory.Location = New-Object System.Drawing.Point(100, 160)
	$cmbCategory.Size = New-Object System.Drawing.Size(280, 22)
	$cmbCategory.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	$cmbCategory.DropDownStyle = "DropDownList"
	$cmbCategory.Items.AddRange(@("General", "Work", "Personal", "Meeting", "Medication", "Birthday", "Reminder"))
	$cmbCategory.SelectedIndex = 0
	$cmbCategory.Name = "cmbCategory"
	$leftPanel.Controls.Add($cmbCategory)

	# Weekday Panel (only visible for Custom Days)
	$weekdayPanel = New-Object System.Windows.Forms.Panel
	$weekdayPanel.Location = New-Object System.Drawing.Point(100, 190)  # 30px below Category
	$weekdayPanel.Size = New-Object System.Drawing.Size(280, 25)
	$weekdayPanel.Visible = $false
	$weekdayPanel.Name = "weekdayPanel"
	$weekdayPanel.BackColor = "#ECF0F1"

	$days = @("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
	$xPos = 0
	for ($i = 0; $i -lt 7; $i++) {
		$chk = New-Object System.Windows.Forms.CheckBox
		$chk.Text = $days[$i]
		$chk.Tag = $days[$i]
		$chk.Location = New-Object System.Drawing.Point($xPos, 2)
		$chk.Size = New-Object System.Drawing.Size(45, 20)
		$chk.Font = New-Object System.Drawing.Font("Segoe UI", 8)
		$chk.Name = "chk$($days[$i])"
		$weekdayPanel.Controls.Add($chk)
		$xPos += 45
	}
	$leftPanel.Controls.Add($weekdayPanel)

	# Voice Row - Fixed Y coordinate (depends on weekday panel visibility)
	$lblVoice = Create-Label -Text "Voice:" -X 15 -Y 190 -Width 80 -Height 22 -FontSize 9 -Name "lblVoice"
	$leftPanel.Controls.Add($lblVoice)

	$cmbVoice = New-Object System.Windows.Forms.ComboBox
	$cmbVoice.Location = New-Object System.Drawing.Point(100, 190)
	$cmbVoice.Size = New-Object System.Drawing.Size(200, 22)
	$cmbVoice.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	$cmbVoice.DropDownStyle = "DropDownList"

	$cmbVoice.Items.Add("IGRF-Bhramma (Male)")
	$cmbVoice.Items.Add("IGRF-Saraswathi (Female)")
	$cmbVoice.Items.Add("No Voice")

	if ($Global:Config.DefaultVoice -eq "IGRF-Bhramma (Male)" -and $Global:AvailableMaleVoices.Count -gt 0) {
		$cmbVoice.SelectedIndex = 0
	} elseif ($Global:Config.DefaultVoice -eq "IGRF-Saraswathi (Female)" -and $Global:AvailableFemaleVoices.Count -gt 0) {
		$cmbVoice.SelectedIndex = 1
	} else {
		$cmbVoice.SelectedIndex = 2
	}

	$cmbVoice.Name = "cmbVoice"
	$leftPanel.Controls.Add($cmbVoice)

	# Voice Test Button
	$btnTestVoice = Create-Button -Text "Test" -X 310 -Y 190 -Width 70 -Height 22 -Color "#3498DB" -Name "btnTestVoice"
	$btnTestVoice.Add_Click({
		# Set test mode to bypass cancellation
		$Global:TestMode = $true
		
		# Get controls
		$cmbVoice = $Global:MainForm.Controls.Find("cmbVoice", $true)[0]
		
		# Check if voice is selected
		if ($cmbVoice.SelectedItem -eq "No Voice") {
			[System.Windows.Forms.MessageBox]::Show("Please select a voice option (Bhramma or Saraswathi) to test.", "Voice Test", "OK", "Warning")
			$Global:TestMode = $false
			return
		}
		
		# Simple test message
		$textToSpeak = "This is a voice test from IGRF"
		
		# Call speak function (no message box)
		Speak-Text -Text $textToSpeak -VoiceOption $cmbVoice.SelectedItem
		
		# Reset test mode
		$Global:TestMode = $false
	})
	$leftPanel.Controls.Add($btnTestVoice)

	# Ring Duration Row (moved up)
	$lblDuration = Create-Label -Text "Ring for:" -X 15 -Y 215 -Width 80 -Height 22 -FontSize 9 -Name "lblDuration"
	$leftPanel.Controls.Add($lblDuration)

	$cmbDuration = New-Object System.Windows.Forms.ComboBox
	$cmbDuration.Location = New-Object System.Drawing.Point(100, 215)
	$cmbDuration.Size = New-Object System.Drawing.Size(280, 22)
	$cmbDuration.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	$cmbDuration.DropDownStyle = "DropDownList"
	$Global:RingDurations.Keys | ForEach-Object { $cmbDuration.Items.Add($_) }
	$cmbDuration.SelectedIndex = 5  # Until Acknowledged
	$cmbDuration.Name = "cmbDuration"
	$leftPanel.Controls.Add($cmbDuration)

	# Options - Checkboxes
	$chkAdvance = New-Object System.Windows.Forms.CheckBox
	$chkAdvance.Text = "Notify 10 minutes before"
	$chkAdvance.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	$chkAdvance.Location = New-Object System.Drawing.Point(100, 245)
	$chkAdvance.Size = New-Object System.Drawing.Size(220, 22)
	$chkAdvance.Checked = $true
	$chkAdvance.Name = "chkAdvance"
	$leftPanel.Controls.Add($chkAdvance)

	$chkPopup = New-Object System.Windows.Forms.CheckBox
	$chkPopup.Text = "Show popup notification"
	$chkPopup.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	$chkPopup.Location = New-Object System.Drawing.Point(100, 270)
	$chkPopup.Size = New-Object System.Drawing.Size(220, 22)
	$chkPopup.Checked = $Global:Config.ShowNotifications
	$chkPopup.Name = "chkPopup"
	$leftPanel.Controls.Add($chkPopup)

	# Action Buttons - All in one row (center aligned)
	# Left panel width is 400px, buttons total width calculation
	$btnY = 325
	$buttonSpacing = 10  # Increased spacing between buttons for better visual separation

	# Button widths: ADD=90, CLEAR=70, RESET=70
	$btnAddWidth = 90
	$btnClearWidth = 70
	$btnResetWidth = 70

	# Calculate total width of all buttons plus spacing
	$totalButtonsWidth = $btnAddWidth + $btnClearWidth + $btnResetWidth + ($buttonSpacing * 2)

	# Calculate starting X position to center the group
	$centerStartX = (400 - $totalButtonsWidth) / 2
	if ($centerStartX -lt 15) { $centerStartX = 15 }  # Ensure it doesn't go out of bounds

	# ADD ALARM button
	$btnAdd = Create-Button -Text "ADD ALARM" -X $centerStartX -Y $btnY -Width $btnAddWidth -Height 30 -Color "#27AE60" -Bold $true -Name "btnAdd"
	$btnAdd.Add_Click({
		$txtName = $Global:MainForm.Controls.Find("txtName", $true)[0]
		$dtpDate = $Global:MainForm.Controls.Find("dtpDate", $true)[0]
		$dtpTime = $Global:MainForm.Controls.Find("dtpTime", $true)[0]
		$cmbRecurrence = $Global:MainForm.Controls.Find("cmbRecurrence", $true)[0]
		$cmbVoice = $Global:MainForm.Controls.Find("cmbVoice", $true)[0]
		$cmbDuration = $Global:MainForm.Controls.Find("cmbDuration", $true)[0]
		$chkAdvance = $Global:MainForm.Controls.Find("chkAdvance", $true)[0]
		$chkPopup = $Global:MainForm.Controls.Find("chkPopup", $true)[0]
		$weekdayPanel = $Global:MainForm.Controls.Find("weekdayPanel", $true)[0]
		$cmbCategory = $Global:MainForm.Controls.Find("cmbCategory", $true)[0]
		
		$alarmName = $txtName.Text.Trim()
		
		if ($alarmName -eq "") {
			[System.Windows.Forms.MessageBox]::Show("Please enter an alarm description.", "Input Required", "OK", "Warning")
			return
		}
		
		$selectedDays = @()
		if ($weekdayPanel.Visible) {
			foreach ($control in $weekdayPanel.Controls) {
				if ($control -is [System.Windows.Forms.CheckBox] -and $control.Checked) {
					$selectedDays += $control.Tag
				}
			}
		}
		
		if ($cmbRecurrence.SelectedItem -eq "Custom Days" -and $selectedDays.Count -eq 0) {
			[System.Windows.Forms.MessageBox]::Show("Please select at least one day for custom recurrence.", "Validation Error", "OK", "Warning")
			return
		}
		
		$durationValue = $Global:RingDurations[$cmbDuration.SelectedItem]
		
		$datePart = $dtpDate.Value.Date
		$timePart = $dtpTime.Value.TimeOfDay
		$alarmDateTime = Get-Date -Year $datePart.Year -Month $datePart.Month -Day $datePart.Day `
									 -Hour $timePart.Hours -Minute $timePart.Minutes -Second 0
		
		if ($cmbRecurrence.SelectedItem -eq "One Time" -and $alarmDateTime -le (Get-Date)) {
			[System.Windows.Forms.MessageBox]::Show("Cannot set alarm for past time. Please select a future time.", "Invalid Time", "OK", "Warning")
			return
		}
		
		Add-Alarm -Name $alarmName `
				 -DateTime $alarmDateTime `
				 -Recurrence $cmbRecurrence.SelectedItem `
				 -Voice $cmbVoice.SelectedItem `
				 -RingDuration $durationValue `
				 -AdvanceNotification $chkAdvance.Checked `
				 -PopupNotification $chkPopup.Checked `
				 -SelectedDays $selectedDays `
				 -Category $cmbCategory.SelectedItem
		
		$txtName.Clear()
		Refresh-AlarmList
		Update-Status "Alarm added for $($alarmDateTime.ToString('dd-MMM-yy hh:mm tt'))"
		Write-AlarmLog -Message "Added alarm: $alarmName at $alarmDateTime"
		
		# FIX: Voice announcement for "alarm set successfully" - ONLY if voice is selected
		if ($cmbVoice.SelectedItem -ne "No Voice") {
			# Set test mode to bypass cancellation
			$Global:TestMode = $true
			Speak-Text -Text "Alarm set successfully" -VoiceOption $cmbVoice.SelectedItem
			$Global:TestMode = $false
		}
	})
	$leftPanel.Controls.Add($btnAdd)

	# CLEAR button
	$btnClear = Create-Button -Text "CLEAR" -X ($centerStartX + $btnAddWidth + $buttonSpacing) -Y $btnY -Width $btnClearWidth -Height 30 -Color "#E67E22" -Name "btnClear"
	$btnClear.Add_Click({
		$txtName = $Global:MainForm.Controls.Find("txtName", $true)[0]
		$txtName.Clear()
	})
	$leftPanel.Controls.Add($btnClear)

	# RESET button
	$btnReset = Create-Button -Text "RESET" -X ($centerStartX + $btnAddWidth + $btnClearWidth + ($buttonSpacing * 2)) -Y $btnY -Width $btnResetWidth -Height 30 -Color "#3498DB" -Name "btnReset"
	$btnReset.Add_Click({
		$dtpDate = $Global:MainForm.Controls.Find("dtpDate", $true)[0]
		$dtpTime = $Global:MainForm.Controls.Find("dtpTime", $true)[0]
		$cmbRecurrence = $Global:MainForm.Controls.Find("cmbRecurrence", $true)[0]
		$cmbVoice = $Global:MainForm.Controls.Find("cmbVoice", $true)[0]
		$cmbDuration = $Global:MainForm.Controls.Find("cmbDuration", $true)[0]
		$chkAdvance = $Global:MainForm.Controls.Find("chkAdvance", $true)[0]
		$chkPopup = $Global:MainForm.Controls.Find("chkPopup", $true)[0]
		$cmbCategory = $Global:MainForm.Controls.Find("cmbCategory", $true)[0]
		
		$dtpDate.Value = (Get-Date).Date
		$dtpTime.Value = Get-Date -Minute (Get-Date).Minute -Second 0
		$cmbRecurrence.SelectedIndex = 0
		$cmbCategory.SelectedIndex = 0
		if ($Global:Config.DefaultVoice -eq "IGRF-Bhramma (Male)" -and $Global:AvailableMaleVoices.Count -gt 0) {
			$cmbVoice.SelectedIndex = 0
		} elseif ($Global:Config.DefaultVoice -eq "IGRF-Saraswathi (Female)" -and $Global:AvailableFemaleVoices.Count -gt 0) {
			$cmbVoice.SelectedIndex = 1
		} else {
			$cmbVoice.SelectedIndex = 2
		}
		$cmbDuration.SelectedIndex = 5
		$chkAdvance.Checked = $true
		$chkPopup.Checked = $Global:Config.ShowNotifications
	})
	$leftPanel.Controls.Add($btnReset)	
	$Global:MainForm.Controls.Add($leftPanel)
	$Global:MainForm.Controls.Add($leftPanel)
    
    # Right Panel - Alarms List
    $rightPanel = New-Object System.Windows.Forms.Panel
    $rightPanel.Size = New-Object System.Drawing.Size(570, 480)
    $rightPanel.Location = New-Object System.Drawing.Point(420, 100)
    $rightPanel.BackColor = "#ECF0F1"
    $rightPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $rightPanel.Name = "rightPanel"
    
    $listTitle = Create-Label -Text "SCHEDULED ALARMS" -X 15 -Y 10 -Width 400 -Height 25 -FontSize 14 -Bold $true -Color "#2C3E50" -Name "lblListTitle"
    $rightPanel.Controls.Add($listTitle)
    
    $lblCount = Create-Label -Text "Total: 0" -X 460 -Y 12 -Width 100 -Height 22 -FontSize 9 -Color "#7F8C8D" -Alignment "MiddleRight" -Name "lblCount"
    $rightPanel.Controls.Add($lblCount)
    
    $lstAlarms = New-Object System.Windows.Forms.ListView
    $lstAlarms.Location = New-Object System.Drawing.Point(15, 40)
    $lstAlarms.Size = New-Object System.Drawing.Size(540, 350)
    $lstAlarms.View = "Details"
    $lstAlarms.FullRowSelect = $true
    $lstAlarms.GridLines = $true
    $lstAlarms.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lstAlarms.MultiSelect = $false
    $lstAlarms.Name = "lstAlarms"
    $lstAlarms.UseCompatibleStateImageBehavior = $false
    $lstAlarms.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    
    $lstAlarms.Columns.Add("Description", 120)
	$lstAlarms.Columns.Add("Date/Time", 120)
	$lstAlarms.Columns.Add("Category", 70)  # New column
	$lstAlarms.Columns.Add("Repeat", 60)
	$lstAlarms.Columns.Add("Sound", 60)
	$lstAlarms.Columns.Add("Voice", 50)
	$lstAlarms.Columns.Add("Notify", 45)
    
    $rightPanel.Controls.Add($lstAlarms)
    
		# Action Buttons - Two rows with STOP centered in second row
		$btnY1 = 400
		$btnY2 = 435
		$buttonSpacing = 88
		$startX = 15

		# ROW 1: DELETE, ENABLE ALL, IMPORT, EXPORT, REFRESH (center aligned)
		$btnY1 = 400
		$btnY2 = 435
		$buttonSpacing = 88
		# Calculate positions for 5 buttons to be centered in the 570px wide panel
		$totalWidth = 5 * $buttonSpacing  # Total width of all buttons (5 * 88 = 440)
		$centerStartX = (570 - $totalWidth) / 2  # Right panel width is 570
		if ($centerStartX -lt 15) { $centerStartX = 15 }  # Ensure it doesn't go out of bounds

		$btnDelete = Create-Button -Text "DELETE" -X $centerStartX -Y $btnY1 -Width 80 -Height 30 -Color "#E74C3C" -Name "btnDelete"
		$btnDelete.Add_Click({
			$lstAlarms = $Global:MainForm.Controls.Find("lstAlarms", $true)[0]
			
			if ($lstAlarms.SelectedItems.Count -gt 0) {
				$result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete this alarm?", "Confirm Delete", "YesNo", "Question")
				if ($result -eq "Yes") {
					$selectedItem = $lstAlarms.SelectedItems[0]
					$alarmId = $selectedItem.Tag.Id
					$alarmName = $selectedItem.Tag.Name
					$Global:Alarms = @($Global:Alarms | Where-Object { $_.Id -ne $alarmId })
					$Global:AlarmsModified++
					Refresh-AlarmList
					Update-Status "Alarm deleted"
					Write-AlarmLog -Message "Deleted alarm: $alarmName"
					Save-Alarms
				}
			}
		})
		$rightPanel.Controls.Add($btnDelete)

		$btnEnableAll = Create-Button -Text "ENABLE ALL" -X ($centerStartX + $buttonSpacing) -Y $btnY1 -Width 80 -Height 30 -Color "#27AE60" -Name "btnEnableAll"
		$btnEnableAll.Add_Click({
			foreach ($alarm in $Global:Alarms) {
				$alarm.Enabled = $true
			}
			$Global:AlarmsModified++
			Refresh-AlarmList
			Update-Status "All alarms enabled"
			Write-AlarmLog -Message "All alarms enabled"
			Save-Alarms
		})
		$rightPanel.Controls.Add($btnEnableAll)

		$btnImport = Create-Button -Text "IMPORT" -X ($centerStartX + $buttonSpacing * 2) -Y $btnY1 -Width 80 -Height 30 -Color "#9B59B6" -Name "btnImport"
		$btnImport.Add_Click({
			Import-Alarms
		})
		$rightPanel.Controls.Add($btnImport)

		$btnExport = Create-Button -Text "EXPORT" -X ($centerStartX + $buttonSpacing * 3) -Y $btnY1 -Width 80 -Height 30 -Color "#3498DB" -Name "btnExport"
		$btnExport.Add_Click({
			Export-Alarms
		})
		$rightPanel.Controls.Add($btnExport)

		$btnRefresh = Create-Button -Text "REFRESH" -X ($centerStartX + $buttonSpacing * 4) -Y $btnY1 -Width 80 -Height 30 -Color "#2C3E50" -Name "btnRefresh"
		$btnRefresh.Add_Click({
			Refresh-AlarmList
		})
		$rightPanel.Controls.Add($btnRefresh)

		# ROW 2: STOP button centered
		$btnStop = Create-Button -Text "STOP" -X 245 -Y $btnY2 -Width 80 -Height 30 -Color "#E74C3C" -Bold $true -Name "btnStop"
		$btnStop.Add_Click({
			Stop-CurrentAlarm
			Write-AlarmLog -Message "Alarm stopped manually"
		})
		$rightPanel.Controls.Add($btnStop)
    
    $Global:MainForm.Controls.Add($rightPanel)
    
    # Status Bar Panel
    $statusPanel = New-Object System.Windows.Forms.Panel
    $statusPanel.Size = New-Object System.Drawing.Size(($formWidth - 20), 35)
    $statusPanel.Location = New-Object System.Drawing.Point(10, 590)
    $statusPanel.BackColor = "#34495E"
    $statusPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $statusPanel.Name = "statusPanel"
    
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Ready - Monitoring for upcoming deadlines"
    $lblStatus.ForeColor = "#ECF0F1"
    $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblStatus.Size = New-Object System.Drawing.Size(550, 30)
    $lblStatus.Location = New-Object System.Drawing.Point(10, 3)
    $lblStatus.Name = "lblStatus"
    $lblStatus.TextAlign = "MiddleLeft"
    $statusPanel.Controls.Add($lblStatus)
    
    $lblDateTime = New-Object System.Windows.Forms.Label
    $lblDateTime.Name = "lblDateTime"
    $lblDateTime.ForeColor = "#3498DB"
    $lblDateTime.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblDateTime.Size = New-Object System.Drawing.Size(350, 30)
    $lblDateTime.Location = New-Object System.Drawing.Point(($formWidth - 380), 3)
    $lblDateTime.TextAlign = "MiddleRight"
    $lblDateTime.AutoSize = $false
    $statusPanel.Controls.Add($lblDateTime)
    
    $Global:StatusTimer = New-Object System.Windows.Forms.Timer
    $Global:StatusTimer.Interval = 1000
    $Global:StatusTimer.Add_Tick({
        Update-DateTimeDisplay
    })
    $Global:StatusTimer.Start()
    
    $Global:MainForm.Controls.Add($statusPanel)
    
    Update-DateTimeDisplay
    
	$cmbRecurrence.Add_SelectedIndexChanged({
		# Get the weekday panel first with null check
		$weekdayPanel = $null
		$weekdayPanelArray = $Global:MainForm.Controls.Find("weekdayPanel", $true)
		if ($weekdayPanelArray -and $weekdayPanelArray.Count -gt 0) {
			$weekdayPanel = $weekdayPanelArray[0]
			$weekdayPanel.Visible = ($this.SelectedItem -eq "Custom Days")
		}
		
		# Check if custom days is selected
		$isCustomDays = ($this.SelectedItem -eq "Custom Days")
		
		# Base Y positions (original positions from initial layout)
		$baseVoiceY = 190
		$baseDurationY = 215
		$baseAdvanceY = 245	
		$basePopupY = 270
		$baseButtonsY = 325
		
		# Weekday panel height is 25px
		$shiftAmount = 0
		if ($isCustomDays) {
			# When custom days is selected, shift all controls below Category down by 25px
			$shiftAmount = 25
		}
		
		# Calculate new Y positions
		$voiceY = $baseVoiceY + $shiftAmount
		$durationY = $baseDurationY + $shiftAmount
		$advanceY = $baseAdvanceY + $shiftAmount
		$popupY = $basePopupY + $shiftAmount
		$buttonsY = $baseButtonsY + $shiftAmount
		
		# Button dimensions for centering
		$btnAddWidth = 90
		$btnClearWidth = 70
		$btnResetWidth = 70
		$buttonSpacing = 10
		
		# Calculate center start X (same for both modes)
		$totalWidth = $btnAddWidth + $btnClearWidth + $btnResetWidth + ($buttonSpacing * 2)
		$startX = (400 - $totalWidth) / 2
		if ($startX -lt 15) { $startX = 15 }
		
		# Update each control individually with error handling and null checks
		
		# Voice label
		try {
			$lblVoiceArray = $Global:MainForm.Controls.Find("lblVoice", $true)
			if ($lblVoiceArray -and $lblVoiceArray.Count -gt 0 -and $lblVoiceArray[0]) {
				$lblVoiceArray[0].Location = New-Object System.Drawing.Point(15, $voiceY)
			}
		} catch { }
		
		# Voice combo box
		try {
			$cmbVoiceArray = $Global:MainForm.Controls.Find("cmbVoice", $true)
			if ($cmbVoiceArray -and $cmbVoiceArray.Count -gt 0 -and $cmbVoiceArray[0]) {
				$cmbVoiceArray[0].Location = New-Object System.Drawing.Point(100, $voiceY)
			}
		} catch { }
		
		# Test voice button
		try {
			$btnTestVoiceArray = $Global:MainForm.Controls.Find("btnTestVoice", $true)
			if ($btnTestVoiceArray -and $btnTestVoiceArray.Count -gt 0 -and $btnTestVoiceArray[0]) {
				$btnTestVoiceArray[0].Location = New-Object System.Drawing.Point(310, $voiceY)
			}
		} catch { }
		
		# Ring duration label
		try {
			$lblDurationArray = $Global:MainForm.Controls.Find("lblDuration", $true)
			if ($lblDurationArray -and $lblDurationArray.Count -gt 0 -and $lblDurationArray[0]) {
				$lblDurationArray[0].Location = New-Object System.Drawing.Point(15, $durationY)
			}
		} catch { }
		
		# Ring duration combo box
		try {
			$cmbDurationArray = $Global:MainForm.Controls.Find("cmbDuration", $true)
			if ($cmbDurationArray -and $cmbDurationArray.Count -gt 0 -and $cmbDurationArray[0]) {
				$cmbDurationArray[0].Location = New-Object System.Drawing.Point(100, $durationY)
			}
		} catch { }
		
		# Advance notification checkbox
		try {
			$chkAdvanceArray = $Global:MainForm.Controls.Find("chkAdvance", $true)
			if ($chkAdvanceArray -and $chkAdvanceArray.Count -gt 0 -and $chkAdvanceArray[0]) {
				$chkAdvanceArray[0].Location = New-Object System.Drawing.Point(100, $advanceY)
			}
		} catch { }
		
		# Popup notification checkbox
		try {
			$chkPopupArray = $Global:MainForm.Controls.Find("chkPopup", $true)
			if ($chkPopupArray -and $chkPopupArray.Count -gt 0 -and $chkPopupArray[0]) {
				$chkPopupArray[0].Location = New-Object System.Drawing.Point(100, $popupY)
			}
		} catch { }
		
		# ADD ALARM button
		try {
			$btnAddArray = $Global:MainForm.Controls.Find("btnAdd", $true)
			if ($btnAddArray -and $btnAddArray.Count -gt 0 -and $btnAddArray[0]) {
				$btnAddArray[0].Location = New-Object System.Drawing.Point($startX, $buttonsY)
			}
		} catch { }
		
		# CLEAR button
		try {
			$btnClearArray = $Global:MainForm.Controls.Find("btnClear", $true)
			if ($btnClearArray -and $btnClearArray.Count -gt 0 -and $btnClearArray[0]) {
				$btnClearArray[0].Location = New-Object System.Drawing.Point($startX + $btnAddWidth + $buttonSpacing, $buttonsY)
			}
		} catch { }
		
		# RESET button
		try {
			$btnResetArray = $Global:MainForm.Controls.Find("btnReset", $true)
			if ($btnResetArray -and $btnResetArray.Count -gt 0 -and $btnResetArray[0]) {
				$btnResetArray[0].Location = New-Object System.Drawing.Point($startX + $btnAddWidth + $btnClearWidth + ($buttonSpacing * 2), $buttonsY)
			}
		} catch { }
		
		# Refresh the left panel
		try {
			$leftPanelArray = $Global:MainForm.Controls.Find("leftPanel", $true)
			if ($leftPanelArray -and $leftPanelArray.Count -gt 0 -and $leftPanelArray[0]) {
				$leftPanelArray[0].Refresh()
			}
		} catch { }
	})
    
    Load-SavedAlarms
    Refresh-AlarmList
    
    $Global:AlarmTimer.Add_Tick({ Check-Alarms })
    $Global:AlarmTimer.Start()
    
    Write-AlarmLog -Message "Application started"
    $Global:MainForm.Show()
    return
}

function Update-DateTimeDisplay {
    if ($Global:MainForm -and !$Global:MainForm.IsDisposed) {
        $lblDateTime = $Global:MainForm.Controls.Find("lblDateTime", $true)[0]
        if ($lblDateTime) {
            $currentTime = Get-Date
            # Format: DD MMM YYYY HH:MM:SS AM/PM - ALWAYS shows seconds
            $lblDateTime.Text = $currentTime.ToString("dd MMM yyyy  hh:mm:ss tt")
        }
    }
}

function Check-MissedAlarms {
    $currentTime = Get-Date
    $missedAlarms = $Global:Alarms | Where-Object { 
        $_.Enabled -and 
        -not $_.Alerted -and 
        $_.DateTime -le $currentTime
    }
    
    foreach ($alarm in $missedAlarms) {
        Write-AlarmLog -Message "Missed alarm detected: $($alarm.Name) at $($alarm.DateTime)" -Level "Warning"
        
        if ($alarm.Recurrence -ne "One Time") {
            Set-NextOccurrence -Alarm $alarm
            Write-AlarmLog -Message "Recurring alarm rescheduled to: $($alarm.DateTime)" -Level "Info"
        } else {
            $alarm.Alerted = $true
            Write-AlarmLog -Message "One-time alarm marked as alerted" -Level "Info"
        }
    }
    
    if ($missedAlarms.Count -gt 0) {
        Save-Alarms
        Refresh-AlarmList
    }
}

function Stop-CurrentAlarm {
    Write-AlarmLog -Message "Stop-CurrentAlarm called from main window STOP button" -Level "Info"
    
    $Global:SoundCancellation = $true
    $Global:IsAlarmActive = $false
    
    # Stop sound playback timer
    if ($Global:SoundPlaybackTimer) {
        try {
            Write-AlarmLog -Message "Stopping sound playback timer" -Level "Info"
            $Global:SoundPlaybackTimer.Stop()
            $Global:SoundPlaybackTimer.Dispose()
        } catch {}
        $Global:SoundPlaybackTimer = $null
    }
    
    # Stop sound player
    if ($Global:SoundPlayer) {
        try {
            $Global:SoundPlayer.Stop()
            $Global:SoundPlayer.Dispose()
        } catch {}
        $Global:SoundPlayer = $null
    }
    
    # Forcefully terminate all PowerShell runspaces and background jobs
    if ($Global:SpeechRunspaces) {
        Write-AlarmLog -Message "Terminating $($Global:SpeechRunspaces.Count) speech runspaces" -Level "Info"
        $jobsToKill = @($Global:SpeechRunspaces.Keys)
        foreach ($jobId in $jobsToKill) {
            try {
                $job = $Global:SpeechRunspaces[$jobId]
                if ($job.ContainsKey('PowerShell') -and $job.PowerShell) {
                    try {
                        if ($job.Handle -and -not $job.Handle.IsCompleted) {
                            $job.PowerShell.Stop()
                        }
                        $job.PowerShell.Dispose()
                    } catch {}
                }
                if ($job.ContainsKey('Job') -and $job.Job) {
                    try {
                        if ($job.Job.State -eq "Running") {
                            Write-AlarmLog -Message "Stopping background job: $($job.Job.Name)" -Level "Info"
                            $job.Job.StopJob() | Out-Null
                        }
                        $job.Job | Remove-Job -Force -ErrorAction SilentlyContinue
                    } catch {}
                }
                $Global:SpeechRunspaces.Remove($jobId)
            } catch {
                Write-AlarmLog -Message "Error terminating job ${jobId}: $_" -Level "Warning"
            }
        }
    }
    
    # Also terminate any background jobs
    $jobs = Get-Job | Where-Object { $_.Name -like "*Voice*" -or $_.Name -like "*Advance*" }
    foreach ($job in $jobs) {
        try {
            Write-AlarmLog -Message "Stopping background job: $($job.Name)" -Level "Info"
            $job.StopJob() | Out-Null
            $job | Remove-Job -Force -ErrorAction SilentlyContinue
        } catch {}
    }
    
    # Close any open forms
    if ($Global:CurrentAlarmForm -and !$Global:CurrentAlarmForm.IsDisposed) {
        Write-AlarmLog -Message "Closing alarm notification form" -Level "Info"
        $Global:CurrentAlarmForm.Close()
        $Global:CurrentAlarmForm.Dispose()
        $Global:CurrentAlarmForm = $null
    }
    if ($Global:CurrentAdvanceForm -and !$Global:CurrentAdvanceForm.IsDisposed) {
        Write-AlarmLog -Message "Closing advance notification form" -Level "Info"
        $Global:CurrentAdvanceForm.Close()
        $Global:CurrentAdvanceForm.Dispose()
        $Global:CurrentAdvanceForm = $null
    }
    
    # Allow a moment for cleanup
    Start-Sleep -Milliseconds 100
    
    $Global:SoundCancellation = $false
    Update-Status "Alarm stopped manually"
    Write-AlarmLog -Message "Stop-CurrentAlarm completed - all sounds and voices terminated" -Level "Info"
}

function Stop-AllSounds {
    Write-AlarmLog -Message "Stop-AllSounds called - terminating all audio" -Level "Info"
    
    # Set cancellation flag
    $Global:SoundCancellation = $true
    
    # CRITICAL: Stop and dispose the sound playback timer FIRST
    if ($Global:SoundPlaybackTimer) {
        try {
            Write-AlarmLog -Message "Stopping sound playback timer" -Level "Info"
            $Global:SoundPlaybackTimer.Stop()
            $Global:SoundPlaybackTimer.Dispose()
        } catch {
            Write-AlarmLog -Message "Error stopping timer: $_" -Level "Warning"
        }
        $Global:SoundPlaybackTimer = $null
    }
    
    # Stop sound player
    if ($Global:SoundPlayer) {
        try {
            Write-AlarmLog -Message "Stopping sound player" -Level "Info"
            $Global:SoundPlayer.Stop()
            $Global:SoundPlayer.Dispose()
        } catch {
            Write-AlarmLog -Message "Error stopping player: $_" -Level "Warning"
        }
        $Global:SoundPlayer = $null
    }
    
    # Terminate all runspaces
    if ($Global:SpeechRunspaces) {
        Write-AlarmLog -Message "Terminating $($Global:SpeechRunspaces.Count) speech runspaces" -Level "Info"
        $jobsToKill = @($Global:SpeechRunspaces.Keys)
        foreach ($jobId in $jobsToKill) {
            try {
                $job = $Global:SpeechRunspaces[$jobId]
                if ($job.PowerShell) {
                    try {
                        if ($job.Handle -and -not $job.Handle.IsCompleted) {
                            $job.PowerShell.Stop()
                        }
                        $job.PowerShell.Dispose()
                    } catch {}
                }
                $Global:SpeechRunspaces.Remove($jobId)
            } catch {
                Write-AlarmLog -Message "Error terminating job ${jobId}: $_" -Level "Warning"
            }
        }
    }
    
    Write-AlarmLog -Message "All sounds and voices stopped" -Level "Info"
}

# ===== FIXED: Using runspaces instead of background jobs for better EXE compatibility =====
function Start-VoiceAnnouncement {
    param(
        [PSCustomObject]$Alarm,
        [bool]$IsAdvance = $false
    )
    
    if ($Alarm.Voice -eq "No Voice") { 
        Write-AlarmLog -Message "Voice announcement skipped: No Voice selected" -Level "Info"
        return 
    }
    
    $repeatCount = if ($IsAdvance) { 3 } else { 1 }  # Advance: 3 cycles, Main: 1 cycle
    $speakCount = if ($IsAdvance) { 15 } else { 7 }  # Advance: 15 times per cycle, Main: 7 times
    
    Write-AlarmLog -Message "Starting voice announcement for: $($Alarm.Name) ($repeatCount cycles of $speakCount, IsAdvance=$IsAdvance)" -Level "Info"
    
    # Create a unique ID for this voice job
    $voiceId = [System.Guid]::NewGuid().ToString()
    
    # Create PowerShell runspace (better for compiled EXE than background jobs)
    $powerShell = [System.Management.Automation.PowerShell]::Create()
    
    # Add the script to the runspace
    [void]$powerShell.AddScript({
        param($alarmName, $voiceOption, $voiceId, $isAdvance, $repeatCount, $speakCount, $appDataPath)
        
        # Function to write to log
        function Write-VoiceLog {
            param([string]$Message, [string]$Level = "Info")
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logEntry = "[$timestamp] [$Level] $Message`n"
            $logFile = Join-Path $appDataPath "alarm_log.txt"
            try {
                $parentDir = Split-Path -Parent $logFile
                if (-not (Test-Path $parentDir)) {
                    New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
                }
                Add-Content -Path $logFile -Value $logEntry -ErrorAction SilentlyContinue
            } catch {}
        }
        
        try {
            if ($voiceOption -eq "No Voice") { 
                Write-VoiceLog -Message "Voice skipped: No Voice selected for ID: $voiceId" -Level "Info"
                return 
            }
            
            Write-VoiceLog -Message "Starting voice announcement for: $alarmName ($repeatCount cycles of $speakCount, IsAdvance=$isAdvance) with ID: $voiceId" -Level "Info"
            
            for ($cycle = 1; $cycle -le $repeatCount; $cycle++) {
                Write-VoiceLog -Message "Starting voice cycle $cycle of $repeatCount for: $alarmName" -Level "Info"
                
                for ($i = 1; $i -le $speakCount; $i++) {
                    $speaker = $null
                    try {
                        $speaker = New-Object System.Speech.Synthesis.SpeechSynthesizer
                        $speaker.Volume = 100
                        $speaker.Rate = 0
                        
                        if ($voiceOption -eq "IGRF-Bhramma (Male)") {
                            try { 
                                $speaker.SelectVoiceByHints([System.Speech.Synthesis.VoiceGender]::Male) 
                                Write-VoiceLog -Message "Selected male voice by gender for cycle $cycle, iteration $i" -Level "Info"
                            } catch { 
                                Write-VoiceLog -Message "Could not select male voice by gender, using default" -Level "Warning"
                            }
                        } elseif ($voiceOption -eq "IGRF-Saraswathi (Female)") {
                            try { 
                                $speaker.SelectVoiceByHints([System.Speech.Synthesis.VoiceGender]::Female)
                                Write-VoiceLog -Message "Selected female voice by gender for cycle $cycle, iteration $i" -Level "Info"
                            } catch { 
                                Write-VoiceLog -Message "Could not select female voice by gender, using default" -Level "Warning"
                            }
                        }
                        
                        Write-VoiceLog -Message "Voice cycle ${cycle}, iteration ${i} of ${speakCount}: speaking '${alarmName}' for voice ID: ${voiceId}" -Level "Info"
                        $speaker.Speak($alarmName)
                        
                    } catch {
                        Write-VoiceLog -Message "Error in voice announcement iteration ${i}: $_" -Level "Error"
                    } finally {
                        if ($speaker) { $speaker.Dispose() }
                    }
                    
                    if ($i -lt $speakCount) { Start-Sleep -Milliseconds 800 }
                }
                
                if ($cycle -lt $repeatCount) { Start-Sleep -Seconds 5 }
            }
            Write-VoiceLog -Message "Voice announcement completed for: $alarmName with voice ID: $voiceId" -Level "Info"
        } catch {
            Write-VoiceLog -Message "Error in voice announcement: $_" -Level "Error"
        }
    })
    
    # Add parameters
    [void]$powerShell.AddParameters(@{
        alarmName = $Alarm.Name
        voiceOption = $Alarm.Voice
        voiceId = $voiceId
        isAdvance = $IsAdvance
        repeatCount = $repeatCount
        speakCount = $speakCount
        appDataPath = $Global:AppDataPath
    })
    
    # Start the runspace asynchronously
    $handle = $powerShell.BeginInvoke()
    
    # Store for cleanup
    $Global:SpeechRunspaces[$voiceId] = @{
        PowerShell = $powerShell
        Handle = $handle
        Job = $null
        Type = "Voice"
        AlarmId = $Alarm.Id
    }
    
    $Global:VoiceRunspaces.Add($voiceId) > $null 2>&1
    
    Write-AlarmLog -Message "Voice announcement started in runspace with ID: $voiceId" -Level "Info"
}

#region Alarm Functions
function Add-Alarm {
    param(
        [string]$Name,
        [datetime]$DateTime,
        [string]$Recurrence,
        [string]$Voice,
        [int]$RingDuration,
        [bool]$AdvanceNotification,
        [bool]$PopupNotification,
        [string[]]$SelectedDays,
        [string]$Category = "General"
    )
    
    $alarm = [PSCustomObject]@{
        Id = [System.Guid]::NewGuid().ToString()
        Name = $Name
        DateTime = $DateTime
        Recurrence = $Recurrence
        Voice = $Voice
        Sound = "No Sound"
        CustomSoundPath = $null
        RingDuration = $RingDuration
        AdvanceNotification = $AdvanceNotification
        PopupNotification = $PopupNotification
        SelectedDays = $SelectedDays
        Alerted = $false
        AdvanceAlerted = $false
        Enabled = $true
        CreatedDate = Get-Date
        SnoozeCount = 0
        OriginalRecurrence = $Recurrence
        OriginalDateTime = $DateTime
        Category = $Category
        BaseTimeOfDay = $DateTime.TimeOfDay
        OriginalDate = $DateTime.Date
    }
    
    if ($Global:Alarms -isnot [System.Collections.IList]) {
        $Global:Alarms = @($Global:Alarms)
    }
    
    $Global:Alarms += $alarm
    $Global:AlarmsModified++
    Save-Alarms
}

function Refresh-AlarmList {
    $form = $Global:MainForm
    if (-not $form -or $form.IsDisposed) { return }
    
    $lstAlarms = $form.Controls.Find("lstAlarms", $true)[0]
    $lblCount = $form.Controls.Find("lblCount", $true)[0]
    
    $lstAlarms.BeginUpdate()
    $lstAlarms.Items.Clear()
    
    $enabledCount = 0
    $currentTime = Get-Date
    
    foreach ($alarm in ($Global:Alarms | Where-Object { $_.Enabled } | Sort-Object DateTime)) {
        $enabledCount++
        
        # Ensure Category has a default value
        if ([string]::IsNullOrEmpty($alarm.Category)) {
            $alarm.Category = "General"
        }
        
        $item = New-Object System.Windows.Forms.ListViewItem($alarm.Name)
        
        # Add subitems in the exact order of columns: Description, Date/Time, Category, Repeat, Sound, Voice, Notify
        $item.SubItems.Add($alarm.DateTime.ToString("dd-MMM-yy hh:mm tt"))
        $item.SubItems.Add($alarm.Category)
        $item.SubItems.Add($alarm.Recurrence)
        
        $item.SubItems.Add("No Sound")  # Always show "No Sound" since sound is removed
        
        $voiceDisplay = if ($alarm.Voice -eq "IGRF-Bhramma (Male)") { "Bhramma" } 
                       elseif ($alarm.Voice -eq "IGRF-Saraswathi (Female)") { "Saraswathi" } 
                       else { "None" }
        $item.SubItems.Add($voiceDisplay)
        
        # Fix: Don't use if statement directly in method call
        if ($alarm.AdvanceNotification) {
            $item.SubItems.Add("Yes")
        } else {
            $item.SubItems.Add("No")
        }
        
        $item.Tag = $alarm
        
        # After creating the item and adding subitems, verify count
        if ($item.SubItems.Count -ne 7) {
            Write-AlarmLog -Message "CRITICAL: Item has $($item.SubItems.Count) subitems, expected 7" -Level "Error"
            # Fix by adding missing subitems
            while ($item.SubItems.Count -lt 7) {
                $item.SubItems.Add("")
            }
        }
        
        $timeDiff = ($alarm.DateTime - $currentTime).TotalMinutes
        if ($timeDiff -le 10 -and $timeDiff -gt 0) {
            $item.BackColor = "#FCF3CF"
        } elseif ($timeDiff -le 0 -and -not $alarm.Alerted) {
            $item.BackColor = "#FADBD8"
        } elseif ($alarm.Alerted) {
            $item.BackColor = "#D5F5E3"
        }
        
        # Verify we have the right number of subitems
        if ($item.SubItems.Count -eq 7) {
            $lstAlarms.Items.Add($item)
        } else {
            Write-AlarmLog -Message "Error: Item has $($item.SubItems.Count) subitems, expected 7" -Level "Error"
        }
    }
    
    $lstAlarms.EndUpdate()
    $lblCount.Text = "Total: $($Global:Alarms.Count) | Active: $enabledCount"
    Update-NextAlarmStatus
}

function Check-Alarms {
    if ([System.Threading.Monitor]::TryEnter($Global:AlarmCheckLock, 100)) {
        try {
            $currentTime = Get-Date
            
            foreach ($alarm in $Global:Alarms) {
                if (-not $alarm.Enabled -or $Global:IsClosing) { continue }
                
                # Check main alarm (40 second window with 5-second buffer before)
				if (-not $alarm.Alerted -and $currentTime -ge $alarm.DateTime.AddSeconds(-5) -and $currentTime -le $alarm.DateTime.AddSeconds(35)) {
					
					# Additional check: If it's before the actual alarm time, wait until actual time
					if ($currentTime -lt $alarm.DateTime) {
						# Calculate milliseconds until actual alarm time
						$msUntilAlarm = ($alarm.DateTime - $currentTime).TotalMilliseconds
						if ($msUntilAlarm -gt 0 -and $msUntilAlarm -le 5000) {  # Within 5 seconds before
							Write-AlarmLog -Message "Early detection for: $($alarm.Name) - waiting $($msUntilAlarm)ms for exact time" -Level "Info"
							Start-Sleep -Milliseconds $msUntilAlarm
						}
					}
					
					Write-AlarmLog -Message "TRIGGERING ALARM: $($alarm.Name) at $(Get-Date)" -Level "Info"
					
					$Global:IsAlarmActive = $true
					$Global:SoundCancellation = $false
					
					# CRITICAL FIX: Stop any existing sounds first
					Stop-AllSounds
					Start-Sleep -Milliseconds 100  # Give time for cleanup
					
					# STEP 1: Show popup dialog FIRST (this will block until user responds)
					if ($alarm.PopupNotification -and $Global:Config.ShowNotifications) {
						Write-AlarmLog -Message "Showing popup for: $($alarm.Name)" -Level "Info"
						
						# Show the alarm notification dialog (this will block until user responds)
						Show-AlarmNotification -Alarm $alarm
						
					} else {
						Write-AlarmLog -Message "Popup not shown: PopupNotification=$($alarm.PopupNotification), ShowNotifications=$($Global:Config.ShowNotifications)" -Level "Warning"
						
						# VOICE ANNOUNCEMENT ONLY - No sound
						if ($alarm.Voice -ne "No Voice") {
							Write-AlarmLog -Message "Starting alarm speech for: $($alarm.Name) (7 times)" -Level "Info"
							Start-VoiceAnnouncement -Alarm $alarm -IsAdvance $false
						}
					}
					
					# Update alarm state
					if ($alarm.Recurrence -ne "One Time") {
						Set-NextOccurrence -Alarm $alarm
						$alarm.Alerted = $false
						$alarm.AdvanceAlerted = $false
					} else {
						$alarm.Alerted = $true
					}
					
					Save-Alarms
					Refresh-AlarmList
				}
                
                # Check advance notification (10 minutes before)
                if ($alarm.AdvanceNotification -and -not $alarm.AdvanceAlerted -and -not $alarm.Alerted) {
				$advanceTime = $alarm.DateTime.AddMinutes(-$Global:Config.AdvanceWarningMinutes)
				$timeUntilAlarm = ($alarm.DateTime - $currentTime).TotalMinutes
				
				Write-AlarmLog -Message "Advance check for '$($alarm.Name)': Current=$currentTime, AdvanceTime=$advanceTime, TimeUntil=$timeUntilAlarm min" -Level "Info"
				
				# Check if we're within the advance warning window
				if ($currentTime -ge $advanceTime -and $currentTime -lt $alarm.DateTime) {
					Write-AlarmLog -Message "Advance notification condition met for: $($alarm.Name)" -Level "Info"
					
					$warningExists = $false
					if ($Global:CurrentAdvanceForm -and !$Global:CurrentAdvanceForm.IsDisposed) {
						$warningExists = $true
						Write-AlarmLog -Message "Advance notification already exists, skipping" -Level "Info"
					}
					
					if (-not $warningExists -and $alarm.PopupNotification -and $Global:Config.ShowNotifications) {
						Write-AlarmLog -Message "Showing advance notification for: $($alarm.Name)" -Level "Info"
						
						# Show the advance notification dialog - this blocks until user closes it
						Show-AdvanceNotification -Alarm $alarm
						
						# IMPORTANT: DO NOT call voice announcement here - it's already handled in Show-AdvanceNotification
						# The form's Shown event handles voice and sound automatically
						
					} else {
						Write-AlarmLog -Message "Not showing advance notification: warningExists=$warningExists, PopupNotification=$($alarm.PopupNotification), ShowNotifications=$($Global:Config.ShowNotifications)" -Level "Info"
					}
					
					$alarm.AdvanceAlerted = $true
					Save-Alarms
					Write-AlarmLog -Message "Advance notification marked as alerted for: $($alarm.Name)"
				} else {
					Write-AlarmLog -Message "Advance notification condition NOT met for: $($alarm.Name) (Current not in advance window)" -Level "Info"
				}
			}
            }
            
            # Only refresh if alarms have changed (track changes with a counter)
            if ($Global:AlarmsModified -and $Global:AlarmsModified -ne $Global:LastRefreshCount) {
                Refresh-AlarmList
                $Global:LastRefreshCount = $Global:AlarmsModified
            }            
        }
        catch {
            Write-AlarmLog -Message "Error in Check-Alarms: $_" -Level "Error"
        }
        finally {
            [System.Threading.Monitor]::Exit($Global:AlarmCheckLock)
        }
    }
}

function Start-SoundLoop {
    param([PSCustomObject]$Alarm)
    
    try {
        if ($Global:SoundPlaybackTimer) {
            $Global:SoundPlaybackTimer.Stop()
            $Global:SoundPlaybackTimer.Dispose()
            $Global:SoundPlaybackTimer = $null
        }
        
        $Global:SoundPlaybackTimer = New-Object System.Windows.Forms.Timer
        $Global:SoundPlaybackTimer.Interval = 4000
        $Global:SoundPlaybackTimer.Add_Tick({
            # CRITICAL: Check if we should continue playing
            if ($Global:IsAlarmActive -and -not $Global:SoundCancellation) {
                Play-Sound -SoundName $Alarm.Sound -CustomSoundPath $Alarm.CustomSoundPath -Duration 1 -Volume $Global:Config.AlarmVolume
            } else {
                # If we shouldn't be playing, stop and dispose the timer
                Write-AlarmLog -Message "Sound loop timer stopping - IsAlarmActive=$($Global:IsAlarmActive), SoundCancellation=$($Global:SoundCancellation)" -Level "Info"
                if ($Global:SoundPlaybackTimer) {
                    $Global:SoundPlaybackTimer.Stop()
                    $Global:SoundPlaybackTimer.Dispose()
                    $Global:SoundPlaybackTimer = $null
                }
            }
        })
        $Global:SoundPlaybackTimer.Start()
        Play-Sound -SoundName $Alarm.Sound -CustomSoundPath $Alarm.CustomSoundPath -Duration 1 -Volume $Global:Config.AlarmVolume
    } catch {
        Write-AlarmLog -Message "Error in Start-SoundLoop: $_" -Level "Error"
    }
}

function Show-AdvanceNotification {
    param([PSCustomObject]$Alarm)
    
    if ($Global:CurrentAdvanceForm -and !$Global:CurrentAdvanceForm.IsDisposed) {
        $Global:CurrentAdvanceForm.Close()
        $Global:CurrentAdvanceForm.Dispose()
        $Global:CurrentAdvanceForm = $null
    }
    
    try {
        # Create the warning form
        $advanceForm = New-Object System.Windows.Forms.Form
        $Global:CurrentAdvanceForm = $advanceForm
        $advanceForm.Text = "Windows Alarm Pro - 10 Minute Warning"
        $advanceForm.Size = New-Object System.Drawing.Size(600, 450)
        $advanceForm.TopMost = $true
        $advanceForm.FormBorderStyle = "FixedDialog"
        $advanceForm.ControlBox = $false
        $advanceForm.BackColor = "#F39C12"
        $advanceForm.ShowInTaskbar = $true
        $advanceForm.KeyPreview = $true
        $advanceForm.Owner = $Global:MainForm
        
        # Center over main window
        Center-FormOverOwner -Form $advanceForm -Owner $Global:MainForm
        
        $appIcon = Get-AppIcon
		if ($appIcon -ne $null) {
			try {
				$advanceForm.Icon = $appIcon
				Write-AlarmLog -Message "Advance notification icon set successfully" -Level "Info"
			} catch {
				Write-AlarmLog -Message "Failed to set advance notification icon: $_" -Level "Warning"
			}
		} else {
			Write-AlarmLog -Message "No icon available for advance notification" -Level "Info"
		}
        
        # Flag to track if user has acknowledged
        $userAcknowledged = $false
        
        $advanceForm.Add_Shown({ 
            $this.Activate()
            $this.Focus()
            
            # Store alarm reference for timers
            $localAlarm = $Alarm
            
            # Start voice announcement for advance notification
            Write-AlarmLog -Message "Starting advance voice announcement for: $($localAlarm.Name)" -Level "Info"
            
            if ($localAlarm.Voice -ne "No Voice") {
                # Use a small delay to ensure form is fully shown before voice starts
                Start-Sleep -Milliseconds 500
                
                # Use the common Start-VoiceAnnouncement function
                Start-VoiceAnnouncement -Alarm $localAlarm -IsAdvance $true
            } else {
                Write-AlarmLog -Message "Voice not selected for advance notification: $($localAlarm.Name)" -Level "Info"
            }
        })
        
        $advanceForm.Add_KeyDown({
            param($sender, $e)
            if ($e.KeyCode -eq "Enter") {
                # User acknowledged - stop all sounds and voice
                $userAcknowledged = $true
                $Global:SoundCancellation = $true
                Stop-AllSounds
                $advanceForm.Close()
            }
        })
        
        # Warning Icon
        $iconPictureBox = New-Object System.Windows.Forms.PictureBox
        $logoImage = Get-LogoImage
        if ($logoImage) {
            $iconPictureBox.Image = $logoImage
            $iconPictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
            $iconPictureBox.Size = New-Object System.Drawing.Size(100, 100)
            $iconPictureBox.Location = New-Object System.Drawing.Point(250, 15)
            $iconPictureBox.BackColor = [System.Drawing.Color]::Transparent
            $advanceForm.Controls.Add($iconPictureBox)
        } else {
            $iconLabel = New-Object System.Windows.Forms.Label
            $iconLabel.Text = "IGRF"
            $iconLabel.ForeColor = "White"
            $iconLabel.Font = New-Object System.Drawing.Font("Segoe UI", 24, [System.Drawing.FontStyle]::Bold)
            $iconLabel.Size = New-Object System.Drawing.Size(100, 50)
            $iconLabel.Location = New-Object System.Drawing.Point(250, 40)
            $iconLabel.TextAlign = "MiddleCenter"
            $advanceForm.Controls.Add($iconLabel)
        }
        
        # Title
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "10 MINUTE WARNING"
        $titleLabel.ForeColor = "White"
        $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
        $titleLabel.Size = New-Object System.Drawing.Size(580, 40)
        $titleLabel.Location = New-Object System.Drawing.Point(10, 120)
        $titleLabel.TextAlign = "MiddleCenter"
        $advanceForm.Controls.Add($titleLabel)
        
        # Alarm Message
        $dateTimeStr = $Alarm.DateTime.ToString("hh:mm tt")
        $messageLabel = New-Object System.Windows.Forms.Label
        $messageLabel.Text = "$($Alarm.Name)`n`nWill trigger at $dateTimeStr"
        $messageLabel.ForeColor = "White"
        $messageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14)
        $messageLabel.Size = New-Object System.Drawing.Size(580, 80)
        $messageLabel.Location = New-Object System.Drawing.Point(10, 165)
        $messageLabel.TextAlign = "MiddleCenter"
        $advanceForm.Controls.Add($messageLabel)
        
        # Button Panel
        $buttonPanel = New-Object System.Windows.Forms.Panel
        $buttonPanel.Size = New-Object System.Drawing.Size(580, 80)
        $buttonPanel.Location = New-Object System.Drawing.Point(10, 250)
        $buttonPanel.BackColor = [System.Drawing.Color]::Transparent
        
        # OK Button
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Size = New-Object System.Drawing.Size(200, 50)
        $okButton.Location = New-Object System.Drawing.Point(190, 15)
        $okButton.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
        $okButton.BackColor = "#2C3E50"
        $okButton.ForeColor = "White"
        $okButton.FlatStyle = "Flat"
        $okButton.FlatAppearance.BorderSize = 2
        $okButton.FlatAppearance.BorderColor = "#FFFFFF"
        $okButton.FlatAppearance.MouseOverBackColor = "#34495E"
        $okButton.Cursor = "Hand"
        $okButton.UseVisualStyleBackColor = $false
        $okButton.Add_Click({ 
            Write-AlarmLog -Message "Advance notification OK button clicked - stopping all sounds" -Level "Info"
            
            # User acknowledged - stop all sounds and voice
            $userAcknowledged = $true
            $Global:SoundCancellation = $true
            
            # Stop all sounds and voice
            Stop-AllSounds
            
            # Force terminate any running PowerShell runspaces for voice
            Cleanup-SpeechJobs
            
            # Close the form
            $advanceForm.Close() 
        })
        $buttonPanel.Controls.Add($okButton)
        
        $advanceForm.Controls.Add($buttonPanel)
        
        # IGRF watermark
        $watermarkLabel = New-Object System.Windows.Forms.Label
        $watermarkLabel.Text = "IGRF Pvt. Ltd. - Professional Alarm System"
        $watermarkLabel.ForeColor = [System.Drawing.Color]::FromArgb(150, 255, 255, 255)
        $watermarkLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
        $watermarkLabel.Size = New-Object System.Drawing.Size(600, 20)
        $watermarkLabel.Location = New-Object System.Drawing.Point(0, 410)
        $watermarkLabel.TextAlign = "MiddleCenter"
        $advanceForm.Controls.Add($watermarkLabel)
        
        # Form closed handler
        $advanceForm.Add_FormClosed({
            $Global:CurrentAdvanceForm = $null
            Write-AlarmLog -Message "Advance notification closed for: $($Alarm.Name)" -Level "Info"
        })
        
        Write-AlarmLog -Message "Showing advance notification form for: $($Alarm.Name)" -Level "Info"
        
        # Hide main form
        if ($Global:MainForm -and !$Global:MainForm.IsDisposed -and $Global:MainForm.Visible) {
            $Global:MainForm.Hide()
        }
        
        # Show the form modally
        $advanceForm.ShowDialog()
        
        # Restore main form
        if ($Global:MainForm -and !$Global:MainForm.IsDisposed) {
            if (-not $Global:MainForm.Visible) {
                $Global:MainForm.Show()
                $Global:MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
                $Global:MainForm.BringToFront()
            }
        }
        
    } catch {
        Write-AlarmLog -Message "Error showing advance notification: $_" -Level "Error"
    }
}

function Show-AlarmNotification {
    param([PSCustomObject]$Alarm)
    
    if ($Global:CurrentAlarmForm -and !$Global:CurrentAlarmForm.IsDisposed) {
        $Global:CurrentAlarmForm.Close()
        $Global:CurrentAlarmForm.Dispose()
        $Global:CurrentAlarmForm = $null
    }
    
    try {
        # Create the form
        $Global:CurrentAlarmForm = New-Object System.Windows.Forms.Form
        $Global:CurrentAlarmForm.Text = "Windows Alarm Pro - IGRF Alarm"
        $Global:CurrentAlarmForm.Size = New-Object System.Drawing.Size(600, 450)  # Clean size without extra space
        $Global:CurrentAlarmForm.TopMost = $true
        $Global:CurrentAlarmForm.TopLevel = $true
        $Global:CurrentAlarmForm.FormBorderStyle = "FixedDialog"
        $Global:CurrentAlarmForm.ControlBox = $false
        $Global:CurrentAlarmForm.BackColor = "#E74C3C"
        $Global:CurrentAlarmForm.ShowInTaskbar = $true
        $Global:CurrentAlarmForm.KeyPreview = $true
        $Global:CurrentAlarmForm.Owner = $Global:MainForm
        
        # Center over main window
        Center-FormOverOwner -Form $Global:CurrentAlarmForm -Owner $Global:MainForm
        
        $appIcon = Get-AppIcon
		if ($appIcon -ne $null) {
			try {
				$Global:CurrentAlarmForm.Icon = $appIcon
				Write-AlarmLog -Message "Alarm notification icon set successfully" -Level "Info"
			} catch {
				Write-AlarmLog -Message "Failed to set alarm notification icon: $_" -Level "Warning"
			}
		} else {
			Write-AlarmLog -Message "No icon available for alarm notification" -Level "Info"
		}
        
        # Flag to track if user has acknowledged or snoozed
        $userResponded = $false
        
        # Form shown event - VOICE ONLY (NO SOUND)
        $Global:CurrentAlarmForm.Add_Shown({
            $Global:CurrentAlarmForm.Activate()
            $Global:CurrentAlarmForm.Focus()
            $Global:CurrentAlarmForm.BringToFront()
            $Global:CurrentAlarmForm.Refresh()
            [System.Windows.Forms.Application]::DoEvents()
            
            # Store alarm reference
            $localAlarm = $Alarm
            
            # Use a small delay to ensure form is fully shown before voice starts
            Start-Sleep -Milliseconds 500
            
            # VOICE ANNOUNCEMENT
            Write-AlarmLog -Message "Starting alarm voice announcement from popup for: $($localAlarm.Name)" -Level "Info"
            
            if ($localAlarm.Voice -ne "No Voice") {
                Start-VoiceAnnouncement -Alarm $localAlarm -IsAdvance $false
            } else {
                Write-AlarmLog -Message "Voice not selected for: $($localAlarm.Name)" -Level "Info"
            }
        })
        
        # Key handlers
        $Global:CurrentAlarmForm.Add_KeyDown({
            param($sender, $e)
            if ($e.KeyCode -eq "Enter") {
                # Stop all sounds and voice
                $Global:SoundCancellation = $true
                $userResponded = $true
                Stop-AllSounds
                $Global:CurrentAlarmForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $Global:CurrentAlarmForm.Close()
                Acknowledge-Alarm -Alarm $Alarm
            }
        })
        
        # Icon/Logo
        $logoImage = Get-LogoImage
        if ($logoImage) {
            $iconPictureBox = New-Object System.Windows.Forms.PictureBox
            $iconPictureBox.Image = $logoImage
            $iconPictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
            $iconPictureBox.Size = New-Object System.Drawing.Size(80, 80)
            $iconPictureBox.Location = New-Object System.Drawing.Point(260, 20)
            $iconPictureBox.BackColor = [System.Drawing.Color]::Transparent
            $Global:CurrentAlarmForm.Controls.Add($iconPictureBox)
        } else {
            $iconLabel = New-Object System.Windows.Forms.Label
            $iconLabel.Text = "⏰"
            $iconLabel.ForeColor = "White"
            $iconLabel.Font = New-Object System.Drawing.Font("Segoe UI", 48, [System.Drawing.FontStyle]::Bold)
            $iconLabel.Size = New-Object System.Drawing.Size(80, 70)
            $iconLabel.Location = New-Object System.Drawing.Point(260, 20)
            $iconLabel.TextAlign = "MiddleCenter"
            $Global:CurrentAlarmForm.Controls.Add($iconLabel)
        }
        
        # Title - Clean without symbols
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "ALARM TRIGGERED"
        $titleLabel.ForeColor = "White"
        $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
        $titleLabel.Size = New-Object System.Drawing.Size(580, 40)
        $titleLabel.Location = New-Object System.Drawing.Point(10, 110)
        $titleLabel.TextAlign = "MiddleCenter"
        $Global:CurrentAlarmForm.Controls.Add($titleLabel)
        
        # Alarm details - Clean without symbols
        $dateTimeStr = $Alarm.DateTime.ToString("dd MMM yyyy  hh:mm:ss tt")
        
        # Message Panel
        $messagePanel = New-Object System.Windows.Forms.Panel
        $messagePanel.Size = New-Object System.Drawing.Size(540, 100)
        $messagePanel.Location = New-Object System.Drawing.Point(30, 155)
        $messagePanel.BackColor = [System.Drawing.Color]::FromArgb(50, 255, 255, 255)
        $Global:CurrentAlarmForm.Controls.Add($messagePanel)
        
        $messageLabel = New-Object System.Windows.Forms.Label
        $messageLabel.Text = "$($Alarm.Name)`n$dateTimeStr`n$($Alarm.Recurrence)  Snoozes: $($Alarm.SnoozeCount)"
        $messageLabel.ForeColor = "White"
        $messageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $messageLabel.Size = New-Object System.Drawing.Size(520, 80)
        $messageLabel.Location = New-Object System.Drawing.Point(10, 10)
        $messageLabel.TextAlign = "MiddleCenter"
        $messagePanel.Controls.Add($messageLabel)
        
        # Button Panel - Two buttons only
        $buttonPanel = New-Object System.Windows.Forms.Panel
        $buttonPanel.Size = New-Object System.Drawing.Size(580, 80)
        $buttonPanel.Location = New-Object System.Drawing.Point(10, 270)
        $buttonPanel.BackColor = [System.Drawing.Color]::Transparent

        # Button Y position (centered vertically in panel)
        $buttonY = 15

        # Acknowledge Button
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "ACKNOWLEDGE"
        $okButton.Size = New-Object System.Drawing.Size(200, 50)
        $okButton.Location = New-Object System.Drawing.Point(90, $buttonY)
        $okButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        $okButton.BackColor = "#2ECC71"
        $okButton.ForeColor = "White"
        $okButton.FlatStyle = "Flat"
        $okButton.FlatAppearance.BorderSize = 2
        $okButton.FlatAppearance.BorderColor = "#FFFFFF"
        $okButton.FlatAppearance.MouseOverBackColor = "#27AE60"
        $okButton.Cursor = "Hand"
        $okButton.UseVisualStyleBackColor = $false
        $okButton.Add_Click({ 
            Write-AlarmLog -Message "Alarm notification ACKNOWLEDGE button clicked - stopping all sounds" -Level "Info"
            
            # User acknowledged - stop all sounds and voice
            $userResponded = $true
            $Global:SoundCancellation = $true
            $Global:IsAlarmActive = $false  # CRITICAL: Reset alarm active flag
            
            # Stop and dispose the sound playback timer
            if ($Global:SoundPlaybackTimer) {
                try {
                    $Global:SoundPlaybackTimer.Stop()
                    $Global:SoundPlaybackTimer.Dispose()
                } catch {}
                $Global:SoundPlaybackTimer = $null
            }
            
            # Stop all sounds and voice
            Stop-AllSounds
            
            # Force terminate any running PowerShell runspaces
            Cleanup-SpeechJobs
            
            # Set dialog result and close
            $Global:CurrentAlarmForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $Global:CurrentAlarmForm.Close()
            
            # Call Acknowledge-Alarm after form is closed
            Acknowledge-Alarm -Alarm $Alarm
        })
        $buttonPanel.Controls.Add($okButton)

        # Snooze Button (5 minutes fixed)
        $snoozeButton = New-Object System.Windows.Forms.Button
        $snoozeButton.Text = "SNOOZE 5 MIN"
        $snoozeButton.Size = New-Object System.Drawing.Size(200, 50)
        $snoozeButton.Location = New-Object System.Drawing.Point(310, $buttonY)
        $snoozeButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        $snoozeButton.BackColor = "#F39C12"
        $snoozeButton.ForeColor = "White"
        $snoozeButton.FlatStyle = "Flat"
        $snoozeButton.FlatAppearance.BorderSize = 2
        $snoozeButton.FlatAppearance.BorderColor = "#FFFFFF"
        $snoozeButton.FlatAppearance.MouseOverBackColor = "#E67E22"
        $snoozeButton.Cursor = "Hand"
        $snoozeButton.UseVisualStyleBackColor = $false
        $snoozeButton.Add_Click({ 
            Write-AlarmLog -Message "Alarm notification SNOOZE button clicked - stopping all sounds and closing dialog" -Level "Info"
            
            # User snoozed - stop all sounds and voice
            $userResponded = $true
            $Global:SoundCancellation = $true
            $Global:IsAlarmActive = $false  # CRITICAL: Reset alarm active flag
            
            # Stop all sounds and voice
            Stop-AllSounds
            
            # Force terminate any running PowerShell runspaces
            Cleanup-SpeechJobs
            
            # Stop and dispose the sound loop timer if it exists
            if ($Global:SoundPlaybackTimer) {
                try {
                    $Global:SoundPlaybackTimer.Stop()
                    $Global:SoundPlaybackTimer.Dispose()
                } catch {}
                $Global:SoundPlaybackTimer = $null
            }
            
            # Fixed 5 minutes snooze
            $minutes = 5
            
            # Store alarm reference for snoozing
            $alarmToSnooze = $Alarm
            $snoozeMinutes = $minutes
            
            # Close the form IMMEDIATELY - dialog disappears
            $Global:CurrentAlarmForm.DialogResult = [System.Windows.Forms.DialogResult]::Retry
            $Global:CurrentAlarmForm.Close()
            
            # Call Snooze-Alarm after form is closed
            Snooze-Alarm -Alarm $alarmToSnooze -Minutes $snoozeMinutes
        })
        $buttonPanel.Controls.Add($snoozeButton)

        $Global:CurrentAlarmForm.Controls.Add($buttonPanel)
        
        # IGRF watermark - Clean text only
        $watermarkLabel = New-Object System.Windows.Forms.Label
        $watermarkLabel.Text = "IGRF Pvt. Ltd. - Professional Alarm System"
        $watermarkLabel.ForeColor = [System.Drawing.Color]::FromArgb(150, 255, 255, 255)
        $watermarkLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
        $watermarkLabel.Size = New-Object System.Drawing.Size(580, 20)
        $watermarkLabel.Location = New-Object System.Drawing.Point(10, 390)
        $watermarkLabel.TextAlign = "MiddleCenter"
        $Global:CurrentAlarmForm.Controls.Add($watermarkLabel)
        
        # Form closed handler
        $Global:CurrentAlarmForm.Add_FormClosed({
            $Global:CurrentAlarmForm = $null
            # Only stop sounds if user hasn't responded yet (closed via X button)
            if (-not $userResponded) {
                $Global:SoundCancellation = $true
                Stop-AllSounds
                Cleanup-SpeechJobs
                Write-AlarmLog -Message "Alarm notification closed without user action" -Level "Info"
            }
        })
        
        Write-AlarmLog -Message "Showing alarm notification form for: $($Alarm.Name)" -Level "Info"
        
        # Hide main form
        if ($Global:MainForm -and !$Global:MainForm.IsDisposed -and $Global:MainForm.Visible) {
            $Global:MainForm.Hide()
        }
        
        # Show the form modally
        $result = $Global:CurrentAlarmForm.ShowDialog()
        
        # Restore main form
        if ($Global:MainForm -and !$Global:MainForm.IsDisposed) {
            if (-not $Global:MainForm.Visible) {
                $Global:MainForm.Show()
                $Global:MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
                $Global:MainForm.BringToFront()
            }
        }
        
    } catch {
        Write-AlarmLog -Message "Error showing alarm notification: $_" -Level "Error"
    }
}

function Acknowledge-Alarm {
    param([PSCustomObject]$Alarm)
    
    # Prevent multiple acknowledgments
    if ($Global:IsAlarmActive -eq $false) {
        Write-AlarmLog -Message "Alarm already acknowledged, skipping duplicate" -Level "Info"
        return
    }
    
    Write-AlarmLog -Message "Acknowledge-Alarm called for: $($Alarm.Name)" -Level "Info"
    
    $Global:IsAlarmActive = $false
    $Global:SoundCancellation = $true
    
    # CRITICAL: Stop and dispose the sound playback timer
    if ($Global:SoundPlaybackTimer) {
        try {
            Write-AlarmLog -Message "Stopping sound playback timer in Acknowledge-Alarm" -Level "Info"
            $Global:SoundPlaybackTimer.Stop()
            $Global:SoundPlaybackTimer.Dispose()
        } catch {
            Write-AlarmLog -Message "Error stopping timer in Acknowledge-Alarm: $_" -Level "Warning"
        }
        $Global:SoundPlaybackTimer = $null
    }
    
    # Stop sound player
    if ($Global:SoundPlayer) {
        try {
            $Global:SoundPlayer.Stop()
            $Global:SoundPlayer.Dispose()
        } catch {}
        $Global:SoundPlayer = $null
    }
    
    # Forcefully terminate all PowerShell runspaces
    if ($Global:SpeechRunspaces) {
        $jobsToKill = @($Global:SpeechRunspaces.Keys)
        foreach ($jobId in $jobsToKill) {
            try {
                $job = $Global:SpeechRunspaces[$jobId]
                if ($job.PowerShell) {
                    try {
                        if ($job.Handle -and -not $job.Handle.IsCompleted) {
                            $job.PowerShell.Stop()
                        }
                        $job.PowerShell.Dispose()
                    } catch {}
                }
                $Global:SpeechRunspaces.Remove($jobId)
            } catch {
                Write-AlarmLog -Message "Error terminating job ${jobId}: $_" -Level "Warning"
            }
        }
    }
    
    Write-AlarmLog -Message "Alarm acknowledged: $($Alarm.Name)"
    
    if ($Global:CurrentAlarmForm -and !$Global:CurrentAlarmForm.IsDisposed) {
        $Global:CurrentAlarmForm.Close()
    }
    
    $Global:SoundCancellation = $false
}

function Snooze-Alarm {
    param(
        [PSCustomObject]$Alarm,
        [int]$Minutes = 5
    )
    
    # Prevent multiple snoozes
    if ($Global:IsAlarmActive -eq $false) {
        Write-AlarmLog -Message "Alarm already inactive, skipping snooze" -Level "Info"
        return
    }
    
    Write-AlarmLog -Message "Snooze-Alarm called for: $($Alarm.Name) for $Minutes minutes" -Level "Info"
    
    # Update alarm time
    $Alarm.DateTime = $Alarm.DateTime.AddMinutes($Minutes)
    $Alarm.Alerted = $false
    $Alarm.AdvanceAlerted = $false
    $Alarm.SnoozeCount++
    $Global:IsAlarmActive = $false
    $Global:SoundCancellation = $true
    
    # Stop sound playback timer
    if ($Global:SoundPlaybackTimer) {
        try {
            Write-AlarmLog -Message "Stopping sound playback timer in Snooze-Alarm" -Level "Info"
            $Global:SoundPlaybackTimer.Stop()
            $Global:SoundPlaybackTimer.Dispose()
        } catch {}
        $Global:SoundPlaybackTimer = $null
    }
    
    # Stop sound player
    if ($Global:SoundPlayer) {
        try {
            $Global:SoundPlayer.Stop()
            $Global:SoundPlayer.Dispose()
        } catch {}
        $Global:SoundPlayer = $null
    }
    
    # Forcefully terminate all PowerShell runspaces and background jobs
    if ($Global:SpeechRunspaces) {
        Write-AlarmLog -Message "Terminating $($Global:SpeechRunspaces.Count) speech runspaces during snooze" -Level "Info"
        $jobsToKill = @($Global:SpeechRunspaces.Keys)
        foreach ($jobId in $jobsToKill) {
            try {
                $job = $Global:SpeechRunspaces[$jobId]
                if ($job.ContainsKey('PowerShell') -and $job.PowerShell) {
                    try {
                        if ($job.Handle -and -not $job.Handle.IsCompleted) {
                            $job.PowerShell.Stop()
                        }
                        $job.PowerShell.Dispose()
                    } catch {}
                }
                if ($job.ContainsKey('Job') -and $job.Job) {
                    try {
                        if ($job.Job.State -eq "Running") {
                            Write-AlarmLog -Message "Stopping background job: $($job.Job.Name)" -Level "Info"
                            $job.Job.StopJob() | Out-Null
                        }
                        $job.Job | Remove-Job -Force -ErrorAction SilentlyContinue
                    } catch {}
                }
                $Global:SpeechRunspaces.Remove($jobId)
            } catch {
                Write-AlarmLog -Message "Error terminating job ${jobId}: $_" -Level "Warning"
            }
        }
    }
    
    # Also terminate any background jobs that might have been missed
    $jobs = Get-Job | Where-Object { $_.Name -like "*Voice*" -or $_.Name -like "*Advance*" }
    foreach ($job in $jobs) {
        try {
            Write-AlarmLog -Message "Stopping background job: $($job.Name)" -Level "Info"
            $job.StopJob() | Out-Null
            $job | Remove-Job -Force -ErrorAction SilentlyContinue
        } catch {}
    }
    
    Write-AlarmLog -Message "Alarm snoozed: $($Alarm.Name) for $Minutes minutes (snooze #$($Alarm.SnoozeCount)) - new time: $($Alarm.DateTime.ToString('hh:mm tt'))"
    
    if ($Global:CurrentAlarmForm -and !$Global:CurrentAlarmForm.IsDisposed) {
        Write-AlarmLog -Message "Closing alarm notification form" -Level "Info"
        $Global:CurrentAlarmForm.Close()
    }
    
    Refresh-AlarmList
    Update-Status "Alarm snoozed until $($Alarm.DateTime.ToString('hh:mm tt'))"
    Save-Alarms
    
    $Global:SoundCancellation = $false
    
    # Log confirmation that snooze is complete and voice will restart at new time
    Write-AlarmLog -Message "Snooze complete - voice announcement will restart at $($Alarm.DateTime.ToString('hh:mm:ss tt'))" -Level "Info"
}


function Speak-Text {
    param(
        [string]$Text,
        [string]$VoiceOption
    )
    
    try {
        if ($VoiceOption -eq "No Voice") { 
            Write-AlarmLog -Message "No voice selected" -Level "Warning"
            return $false 
        }
        
        # Don't check SoundCancellation during test mode
        if ($Global:SoundCancellation -and -not $Global:TestMode) { 
            return $false 
        }
        
        Write-AlarmLog -Message "Speaking text: '$Text' using $VoiceOption" -Level "Info"
        
        $speaker = New-Object System.Speech.Synthesis.SpeechSynthesizer
        
        # Set volume to maximum for testing
        $speaker.Volume = 100
        $speaker.Rate = 0
        
        # Log available voices
        Write-AlarmLog -Message "Available voices: $($Global:AvailableVoices.Count)" -Level "Info"
        
        # Try to select the appropriate voice
        $voiceSelected = $false
        
        if ($VoiceOption -eq "IGRF-Bhramma (Male)") {
            # Try to select male voice by gender first (most reliable)
            try {
                $speaker.SelectVoiceByHints([System.Speech.Synthesis.VoiceGender]::Male)
                Write-AlarmLog -Message "Selected male voice by gender" -Level "Info"
                $voiceSelected = $true
            } catch {
                Write-AlarmLog -Message "Failed to select male voice by gender, trying specific voices" -Level "Warning"
                
                # Fallback to specific male voices if available
                if ($Global:AvailableMaleVoices.Count -gt 0) {
                    try {
                        $voiceName = Get-VoiceName -VoiceType $VoiceOption
                        if ($voiceName) {
                            $speaker.SelectVoice($voiceName)
                            Write-AlarmLog -Message "Selected male voice: $voiceName" -Level "Info"
                            $voiceSelected = $true
                        }
                    } catch {
                        Write-AlarmLog -Message "Failed to select specific male voice" -Level "Warning"
                    }
                }
            }
        }
        elseif ($VoiceOption -eq "IGRF-Saraswathi (Female)") {
            # Try to select female voice by gender first (most reliable)
            try {
                $speaker.SelectVoiceByHints([System.Speech.Synthesis.VoiceGender]::Female)
                Write-AlarmLog -Message "Selected female voice by gender" -Level "Info"
                $voiceSelected = $true
            } catch {
                Write-AlarmLog -Message "Failed to select female voice by gender, trying specific voices" -Level "Warning"
                
                # Fallback to specific female voices if available
                if ($Global:AvailableFemaleVoices.Count -gt 0) {
                    try {
                        $voiceName = Get-VoiceName -VoiceType $VoiceOption
                        if ($voiceName) {
                            $speaker.SelectVoice($voiceName)
                            Write-AlarmLog -Message "Selected female voice: $voiceName" -Level "Info"
                            $voiceSelected = $true
                        }
                    } catch {
                        Write-AlarmLog -Message "Failed to select specific female voice" -Level "Warning"
                    }
                }
            }
        }
        
        # If no voice was selected, try to select any available voice
        if (-not $voiceSelected -and $Global:AvailableVoices.Count -gt 0) {
            try {
                $speaker.SelectVoice($Global:AvailableVoices[0].Name)
                Write-AlarmLog -Message "Selected first available voice: $($Global:AvailableVoices[0].Name)" -Level "Info"
                $voiceSelected = $true
            } catch {
                Write-AlarmLog -Message "Failed to select any voice" -Level "Error"
            }
        }
        
        if (-not $voiceSelected) {
            Write-AlarmLog -Message "No voices could be selected" -Level "Error"
            $speaker.Dispose()
            return $false
        }
        
        # Speak the text - this should play through speakers
        Write-AlarmLog -Message "Speaking now..." -Level "Info"
        $speaker.Speak($Text)
        $speaker.Dispose()
        
        Write-AlarmLog -Message "Speech completed" -Level "Info"
        return $true
        
    } catch {
        Write-AlarmLog -Message "Speech failed: $_" -Level "Error"
        return $false
    }
}

function Play-Sound {
    param(
        [string]$SoundName,
        [string]$CustomSoundPath,
        [int]$Duration = 3,
        [int]$Volume = 100
    )
    
    if ($Global:SoundCancellation -and -not $Global:TestMode) { return }
    
    # Store original volume to restore later
    $originalVolume = $null
    try {
        $wmp = New-Object -ComObject "WMPlayer.OCX"
        $originalVolume = $wmp.settings.volume
        $wmp.settings.volume = $Volume
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wmp) | Out-Null
    } catch {
        # Continue without volume control
    }
    
    if ([System.Threading.Monitor]::TryEnter($Global:SoundPlayerLock, 500)) {
        try {
            $repeatCount = [Math]::Max(1, $Duration)
            
            for ($i = 0; $i -lt $repeatCount; $i++) {
                if ($Global:SoundCancellation -and -not $Global:TestMode) { break }
                
                if ($SoundName -eq "Custom Sound" -and $CustomSoundPath -and (Test-Path $CustomSoundPath)) {
                    try {
                        $Global:SoundPlayer = New-Object System.Media.SoundPlayer
                        $Global:SoundPlayer.SoundLocation = $CustomSoundPath
                        $Global:SoundPlayer.Load()
                        $Global:SoundPlayer.PlaySync()
                    } catch {
                        [System.Media.SystemSounds]::Asterisk.Play()
                        Start-Sleep -Milliseconds 300
                    } finally {
                        if ($Global:SoundPlayer) { $Global:SoundPlayer.Dispose() }
                    }
                } else {
                    switch ($SoundName) {
                        "Windows Default" { [System.Media.SystemSounds]::Asterisk.Play() }
                        "Windows Beep" { [System.Media.SystemSounds]::Exclamation.Play() }
                        "Windows Critical" { [System.Media.SystemSounds]::Hand.Play() }
                        "Windows Notification" { 
                            try {
                                $path = "$env:SystemRoot\Media\Windows Notify System Generic.wav"
                                if (Test-Path $path) {
                                    $Global:SoundPlayer = New-Object System.Media.SoundPlayer
                                    $Global:SoundPlayer.SoundLocation = $path
                                    $Global:SoundPlayer.Load()
                                    $Global:SoundPlayer.PlaySync()
                                } else {
                                    [System.Media.SystemSounds]::Asterisk.Play()
                                }
                            } catch {
                                [System.Media.SystemSounds]::Asterisk.Play()
                            } finally {
                                if ($Global:SoundPlayer) { $Global:SoundPlayer.Dispose() }
                            }
                        }
                        "Classic Bell" { 
                            for ($j = 0; $j -lt 2; $j++) {
                                [System.Console]::Beep(800, 300)  # Fixed frequency, volume handled by system
                            }
                        }
                        "Continuous Ring" {
                            for ($j = 0; $j -lt 3; $j++) {
                                [System.Console]::Beep(800, 200)
                                [System.Console]::Beep(1000, 200)
                            }
                        }
                        "Alert Tone" {
                            [System.Console]::Beep(900, 500)
                            Start-Sleep -Milliseconds 200
                            [System.Console]::Beep(900, 500)
                        }
                        "Soft Chime" {
                            [System.Console]::Beep(520, 500)
                            Start-Sleep -Milliseconds 100
                            [System.Console]::Beep(660, 500)
                        }
                        "No Sound" { break }
                        default { [System.Media.SystemSounds]::Asterisk.Play() }
                    }
                    
                    if ($SoundName -ne "No Sound" -and $i -lt $repeatCount - 1) {
                        Start-Sleep -Milliseconds 400
                    }
                }
            }
        }
        catch {
            Write-AlarmLog -Message "Error playing sound: $_" -Level "Error"
        }
        finally {
            [System.Threading.Monitor]::Exit($Global:SoundPlayerLock)
        }
    }
    
    # Restore original volume (optional)
    if ($originalVolume) {
        try {
            $wmp = New-Object -ComObject "WMPlayer.OCX"
            $wmp.settings.volume = $originalVolume
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wmp) | Out-Null
        } catch {}
    }
}

function Set-NextOccurrence {
    param([PSCustomObject]$Alarm)
    
    # Use original base time, not modified time (which may have been snoozed)
    $baseTime = $Alarm.BaseTimeOfDay
    $originalDate = $Alarm.OriginalDate
    $currentDateTime = Get-Date  # Use current time, not alarm time
    
    Write-AlarmLog -Message "Setting next occurrence for: $($Alarm.Name), recurrence: $($Alarm.Recurrence)" -Level "Info"
    
    switch ($Alarm.Recurrence) {
        "Daily" { 
            $Alarm.DateTime = (Get-Date).Date.AddDays(1).Add($baseTime)
        }
        "Weekly" { 
            $Alarm.DateTime = (Get-Date).Date.AddDays(7).Add($baseTime)
        }
        "Weekdays (Mon-Fri)" {
            $checkDate = (Get-Date).Date.AddDays(1)
            while ($checkDate.DayOfWeek -eq [System.DayOfWeek]::Saturday -or 
                   $checkDate.DayOfWeek -eq [System.DayOfWeek]::Sunday) {
                $checkDate = $checkDate.AddDays(1)
            }
            $Alarm.DateTime = $checkDate.Add($baseTime)
        }
        "Weekends" {
            $checkDate = (Get-Date).Date.AddDays(1)
            while ($checkDate.DayOfWeek -ne [System.DayOfWeek]::Saturday -and 
                   $checkDate.DayOfWeek -ne [System.DayOfWeek]::Sunday) {
                $checkDate = $checkDate.AddDays(1)
            }
            $Alarm.DateTime = $checkDate.Add($baseTime)
        }
        "Monthly" { 
            $targetDay = $originalDate.Day
            $nextMonth = (Get-Date).AddMonths(1)
            $daysInMonth = [DateTime]::DaysInMonth($nextMonth.Year, $nextMonth.Month)
            $adjustedDay = [Math]::Min($targetDay, $daysInMonth)
            $Alarm.DateTime = Get-Date -Year $nextMonth.Year -Month $nextMonth.Month -Day $adjustedDay -Hour $baseTime.Hours -Minute $baseTime.Minutes -Second 0
        }
        "Yearly" { 
            $nextYear = (Get-Date).AddYears(1)
            $targetMonth = $originalDate.Month
            $targetDay = $originalDate.Day
            $daysInMonth = [DateTime]::DaysInMonth($nextYear.Year, $targetMonth)
            $adjustedDay = [Math]::Min($targetDay, $daysInMonth)
            $Alarm.DateTime = Get-Date -Year $nextYear.Year -Month $targetMonth -Day $adjustedDay -Hour $baseTime.Hours -Minute $baseTime.Minutes -Second 0
        }
        "Custom Days" {
            if ($Alarm.SelectedDays.Count -gt 0) {
                $found = $false
                $checkDate = (Get-Date).Date.AddDays(1)
                $maxDays = 14
                $daysChecked = 0
                
                $dayMapping = @{
                    "Mon" = [System.DayOfWeek]::Monday
                    "Tue" = [System.DayOfWeek]::Tuesday
                    "Wed" = [System.DayOfWeek]::Wednesday
                    "Thu" = [System.DayOfWeek]::Thursday
                    "Fri" = [System.DayOfWeek]::Friday
                    "Sat" = [System.DayOfWeek]::Saturday
                    "Sun" = [System.DayOfWeek]::Sunday
                }
                
                while (-not $found -and $daysChecked -lt $maxDays) {
                    $dayOfWeek = $checkDate.DayOfWeek
                    foreach ($selectedDay in $Alarm.SelectedDays) {
                        if ($dayMapping[$selectedDay] -eq $dayOfWeek) {
                            $Alarm.DateTime = $checkDate.Add($baseTime)
                            $found = $true
                            break
                        }
                    }
                    if (-not $found) { $checkDate = $checkDate.AddDays(1); $daysChecked++ }
                }
                if (-not $found) { $Alarm.DateTime = (Get-Date).Date.AddDays(7).Add($baseTime) }
            } else {
                $Alarm.DateTime = (Get-Date).Date.AddDays(7).Add($baseTime)
            }
        }
    }
    
    $Alarm.Alerted = $false
    $Alarm.AdvanceAlerted = $false
    
    Write-AlarmLog -Message "Next occurrence set to: $($Alarm.DateTime)" -Level "Info"
}

function Update-Status {
    param([string]$Message)
    
    $form = $Global:MainForm
    if ($form -and !$form.IsDisposed) {
        $lblStatus = $form.Controls.Find("lblStatus", $true)[0]
        if ($lblStatus) { $lblStatus.Text = "$Message" }
    }
}

function Update-NextAlarmStatus {
    $nextAlarm = $Global:Alarms | Where-Object { -not $_.Alerted -and $_.Enabled } | Sort-Object DateTime | Select-Object -First 1
    
    $form = $Global:MainForm
    if ($form -and !$form.IsDisposed) {
        $lblStatus = $form.Controls.Find("lblStatus", $true)[0]
        if ($lblStatus -and $nextAlarm) {
            $timeUntil = ($nextAlarm.DateTime - (Get-Date)).TotalMinutes
            $timeStr = if ($timeUntil -lt 60) { "$([math]::Round($timeUntil)) minutes" } 
                      else { "$([math]::Round($timeUntil/60,1)) hours" }
            $lblStatus.Text = "Next: '$($nextAlarm.Name)' at $($nextAlarm.DateTime.ToString('hh:mm tt')) (in $timeStr)"
        } elseif ($lblStatus) {
            $lblStatus.Text = "No upcoming alarms"
        }
    }
}

function Save-Alarms {
    $savePath = Join-Path $Global:AppDataPath "alarms.xml"
    
    try {
        $exportAlarms = $Global:Alarms | ForEach-Object {
            [PSCustomObject]@{
                Id = $_.Id
                Name = $_.Name
                DateTime = $_.DateTime
                Recurrence = $_.Recurrence
                Voice = $_.Voice
                Sound = $_.Sound
                CustomSoundPath = $_.CustomSoundPath
                RingDuration = $_.RingDuration
                AdvanceNotification = $_.AdvanceNotification
                PopupNotification = $_.PopupNotification
                EmailNotification = $_.EmailNotification
                SelectedDays = $_.SelectedDays
                Alerted = $_.Alerted
                AdvanceAlerted = $_.AdvanceAlerted
                Enabled = $_.Enabled
                CreatedDate = $_.CreatedDate
                SnoozeCount = $_.SnoozeCount
                OriginalRecurrence = $_.OriginalRecurrence
                OriginalDateTime = $_.OriginalDateTime
            }
        }
        $exportAlarms | Export-Clixml -Path $savePath -Force
        Write-AlarmLog -Message "Saved $($Global:Alarms.Count) alarms"
    } catch {
        Write-AlarmLog -Message "Failed to save alarms: $_" -Level "Error"
    }
}

function Load-SavedAlarms {
    $savePath = Join-Path $Global:AppDataPath "alarms.xml"
    
    if (Test-Path $savePath) {
        try {
            $loadedAlarms = Import-Clixml -Path $savePath
            if ($loadedAlarms -is [System.Collections.IList]) {
                $Global:Alarms = $loadedAlarms
            } elseif ($loadedAlarms -ne $null) {
                $Global:Alarms = @($loadedAlarms)
            } else {
                $Global:Alarms = @()
            }
            
            # Add Category property to any alarms that don't have it
            foreach ($alarm in $Global:Alarms) {
                # Check if Category property exists
                $hasCategory = $false
                foreach ($prop in $alarm.PSObject.Properties) {
                    if ($prop.Name -eq "Category") {
                        $hasCategory = $true
                        break
                    }
                }
                
                # Add Category if it doesn't exist
                if (-not $hasCategory) {
                    Add-Member -InputObject $alarm -MemberType NoteProperty -Name "Category" -Value "General" -Force
                }
            }
            
            Write-AlarmLog -Message "Loaded $($Global:Alarms.Count) alarms"
        } catch {
            $Global:Alarms = @()
            Write-AlarmLog -Message "Failed to load alarms: $_" -Level "Error"
        }
    } else {
        $Global:Alarms = @()
    }
}

function Export-Alarms {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|XML files (*.xml)|*.xml"
    $saveFileDialog.FileName = "Alarms_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        try {
            if ($saveFileDialog.FilterIndex -eq 1) {
                $Global:Alarms | Select-Object Name, DateTime, Recurrence, Sound, Voice, AdvanceNotification, Enabled, SnoozeCount | 
                    Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            } else {
                $Global:Alarms | Export-Clixml -Path $saveFileDialog.FileName
            }
            [System.Windows.Forms.MessageBox]::Show("Alarms exported successfully!", "Export Complete", "OK", "Information")
            Write-AlarmLog -Message "Exported alarms to $($saveFileDialog.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Export failed: $_", "Export Error", "OK", "Error")
            Write-AlarmLog -Message "Export failed: $_" -Level "Error"
        }
    }
}
#endregion

function Import-Alarms {
    <#
    .SYNOPSIS
        Import alarms from CSV or XML file
    #>
    
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "All Supported Files (*.xml;*.csv)|*.xml;*.csv|XML files (*.xml)|*.xml|CSV files (*.csv)|*.csv"
    $openFileDialog.Title = "Import Alarms"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($openFileDialog.ShowDialog() -eq "OK") {
        try {
            $importCount = 0
            $filePath = $openFileDialog.FileName
            
            if ($filePath -like "*.csv") {
                $imported = Import-Csv $filePath
                foreach ($item in $imported) {
                    # Create new alarm from CSV data
                    $alarm = [PSCustomObject]@{
                        Id = [System.Guid]::NewGuid().ToString()
                        Name = if ($item.Name) { $item.Name } else { "Imported Alarm" }
                        DateTime = if ($item.DateTime) { [datetime]$item.DateTime } else { Get-Date }
                        Recurrence = if ($item.Recurrence) { $item.Recurrence } else { "One Time" }
                        Voice = if ($item.Voice) { $item.Voice } else { "No Voice" }
                        Sound = if ($item.Sound) { $item.Sound } else { "Windows Default" }
                        CustomSoundPath = $null
                        RingDuration = if ($item.RingDuration) { [int]$item.RingDuration } else { 0 }
                        AdvanceNotification = if ($item.AdvanceNotification) { [bool]$item.AdvanceNotification } else { $true }
                        PopupNotification = if ($item.PopupNotification) { [bool]$item.PopupNotification } else { $true }
                        EmailNotification = if ($item.EmailNotification) { [bool]$item.EmailNotification } else { $false }
                        SelectedDays = @()
                        Alerted = $false
                        AdvanceAlerted = $false
                        Enabled = $true
                        CreatedDate = Get-Date
                        SnoozeCount = 0
                        OriginalRecurrence = if ($item.Recurrence) { $item.Recurrence } else { "One Time" }
                        OriginalDateTime = if ($item.DateTime) { [datetime]$item.DateTime } else { Get-Date }
                    }
                    $Global:Alarms += $alarm
                    $importCount++
                }
            } else {
                # XML import
                $imported = Import-Clixml $filePath
                foreach ($item in $imported) {
                    # Generate new IDs to avoid conflicts
                    $item.Id = [System.Guid]::NewGuid().ToString()
                    $item.CreatedDate = Get-Date
                    $Global:Alarms += $item
                    $importCount++
                }
            }
            
            $Global:AlarmsModified++
			Refresh-AlarmList
			Save-Alarms
			[System.Windows.Forms.MessageBox]::Show("Successfully imported $importCount alarm(s)!", "Import Successful", "OK", "Information")
            Write-AlarmLog -Message "Imported $importCount alarms from $filePath"
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Import failed: $_", "Import Error", "OK", "Error")
            Write-AlarmLog -Message "Import failed: $_" -Level "Error"
        }
    }
}

#region Main Execution
try {
    Write-AlarmLog -Message "=" * 50 -Level "Info"
    Write-AlarmLog -Message "Application starting..." -Level "Info"
    Write-AlarmLog -Message "=" * 50 -Level "Info"
    
    Show-MainForm | Out-Null
    
    $Global:ApplicationRunning = $true
    
    while ($Global:ApplicationRunning -and -not $Global:ManualClose) {
        [System.Windows.Forms.Application]::DoEvents() | Out-Null
        Start-Sleep -Milliseconds 100 | Out-Null
        
        if ($Global:MainForm -and $Global:MainForm.IsDisposed) {
            $Global:ApplicationRunning = $false
        }
    }
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)`n`nPlease contact support at https://igrf.co.in/en/software/", "Windows Alarm Pro", "OK", "Error")
    Write-AlarmLog -Message "Fatal error: $_" -Level "Error"
}
finally {
    Cleanup-Resources
    Write-AlarmLog -Message "=" * 50 -Level "Info"
    Write-AlarmLog -Message "Application closed" -Level "Info"
    Write-AlarmLog -Message "=" * 50 -Level "Info"
    
    # Check if this was a normal exit
    if (-not $Global:ManualClose) {
        # Restart if crashed? Optional
        # Start-Process powershell -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`""
    }
}
#endregion