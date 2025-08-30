

$Global:UseColors = $true
function Write-Info($msg){ if($UseColors){Write-Host $msg -ForegroundColor Green}else{Write-Host $msg}}
function Write-Warn($msg){ if($UseColors){Write-Host $msg -ForegroundColor Yellow}else{Write-Host $msg}}
function Write-Err($msg){ if($UseColors){Write-Host $msg -ForegroundColor Red}else{Write-Host $msg}}
function Write-Note($msg){ if($UseColors){Write-Host $msg -ForegroundColor Cyan}else{Write-Host $msg}}
function Line([int]$w=60){Write-Host ("-"*($w))}
function Banner(){
    Clear-Host
    Line
    Write-Host "By @IBRHUB TikTok 4K Glitch Exporter"
    Write-Host ("PowerShell {0}  |  {1}" -f $PSVersionTable.PSVersion.ToString(), (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
    Line
}

# Resolve/Install FFmpeg
function Resolve-Exe([string]$name) {
    $cmd = Get-Command $name -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}
function Ensure-FFmpeg {
    $global:FFMPEG  = Resolve-Exe 'ffmpeg'
    $global:FFPROBE = Resolve-Exe 'ffprobe'
    if ($FFMPEG -and $FFPROBE) { return $true }

    Write-Warn "FFmpeg not found in PATH. Trying to install..."

    if (Resolve-Exe 'choco') {
        Write-Note "Installing via Chocolatey..."
        try { choco install ffmpeg-full -y | Out-Null } catch {}
    } else { Write-Note "Chocolatey not found, skipping." }

    if (-not (Resolve-Exe 'ffmpeg')) {
        if (Resolve-Exe 'scoop') {
            Write-Note "Installing via Scoop..."
            try { scoop install ffmpeg | Out-Null } catch {}
        } else { Write-Note "Scoop not found, skipping." }
    }

    if (-not (Resolve-Exe 'ffmpeg')) {
        if (Resolve-Exe 'winget') {
            Write-Note "Installing via Winget..."
            try { winget install --id Gyan.FFmpeg.Full --source winget --accept-source-agreements --accept-package-agreements | Out-Null } catch {}
        } else { Write-Note "Winget not found, skipping." }
    }

    # Refresh PATH from Machine+User
    $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path','User')

    $global:FFMPEG  = Resolve-Exe 'ffmpeg'
    $global:FFPROBE = Resolve-Exe 'ffprobe'
    return ($FFMPEG -and $FFPROBE)
}
function Get-DurationSec([string]$p) {
    $raw = & $FFPROBE -v error -show_entries format=duration -of default=nokey=1:noprint_wrappers=1 -- "$p"
    if (-not $raw) { return 0.0 }
    try { return [double]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture) } catch { return 0.0 }
}
function Has-Audio([string]$p) {
    $a = & $FFPROBE -v error -select_streams a:0 -show_entries stream=index -of csv=p=0 -- "$p"
    return -not [string]::IsNullOrWhiteSpace($a)
}
function Get-AtempoChain([double]$factor) {
    $t = $factor
    $chain = @()
    while ($t -lt 0.5 -or $t -gt 2.0) {
        if     ($t -lt 0.5) { $chain += "atempo=0.5"; $t /= 0.5 }
        elseif ($t -gt 2.0) { $chain += "atempo=2.0"; $t /= 2.0 }
    }
    $chain += ("atempo={0}" -f ([math]::Round($t,3)))
    return ($chain -join ",")
}
function Show-Progress {
    param (
        [Parameter(Mandatory)][Single]$TotalValue,
        [Parameter(Mandatory)][Single]$CurrentValue,
        [Parameter(Mandatory)][string]$ProgressText,
        [Parameter()][int]$BarSize = 40,
        [Parameter()][switch]$Complete
    )
    $percent = if ($TotalValue -gt 0) { $CurrentValue / $TotalValue } else { 0 }
    if ($percent -gt 1) { $percent = 1 }
    if ($percent -lt 0) { $percent = 0 }
    $percentComplete = $percent * 100
    if ($psISE) {
        Write-Progress "$ProgressText" -Id 0 -PercentComplete $percentComplete
    } else {
        $filled = [int]([math]::Round($BarSize * $percent))
        if ($filled -lt 0) { $filled = 0 }
        if ($filled -gt $BarSize) { $filled = $BarSize }
        $bar = ''.PadRight($filled, [char]9608).PadRight($BarSize, [char]9617)
        Write-Host -NoNewLine ("`r{0} {1} {2,6:N2} %" -f $ProgressText, $bar, $percentComplete)
        if ($Complete) { Write-Host "" }
    }
}

# Start-up
Banner
if (-not (Ensure-FFmpeg)) { Write-Err "ffmpeg/ffprobe not available."; exit 1 }
Write-Info  ("Using ffmpeg : {0}" -f $FFMPEG)
Write-Info  ("Using ffprobe: {0}" -f $FFPROBE)
Line

# Pick input file 
Add-Type -AssemblyName System.Windows.Forms
$dlg = New-Object Windows.Forms.OpenFileDialog
$dlg.Filter = "Video Files|*.mp4;*.mov;*.mkv;*.avi;*.webm;*.flv;*.wmv|All Files|*.*"
$dlg.Multiselect = $false
if ($dlg.ShowDialog() -ne "OK") { Write-Warn "No file selected. Exiting."; exit 2 }
$InFile = $dlg.FileName
$InDir  = [System.IO.Path]::GetDirectoryName($InFile)

# Mode selection 
Write-Note "Choose mode:"
Write-Host "  1) GPU (NVENC) "
Write-Host "  2) CPU (x264)  "
Write-Host "  3) Legacy (Tiktok)"
$modeChoice = Read-Host "Enter 1, 2 or 3"
if ($modeChoice -notin @('1','2','3')) { Write-Err "Invalid choice."; exit 2 }

# Legacy Mode
if ($modeChoice -eq '3') {
    Write-Info "LEGACY mode: processing all videos in:"
    Write-Host "  $InDir"
    $exts = '*.mp4','*.mov','*.mkv','*.avi','*.webm','*.flv','*.wmv'
    $files = Get-ChildItem -Path (Join-Path $InDir '*') -File -Include $exts -ErrorAction SilentlyContinue |
             Where-Object { $_.Extension -match '^\.(mp4|mov|mkv|avi|webm|flv|wmv)$' }

    if (-not $files -or $files.Count -eq 0) {
        Write-Err "No matching video files found in that folder."
        Write-Note "Tip: Place your clips in that folder (same folder as the selected file) and run again."
        pause
        exit 3
    }

    $i = 0
    $n = $files.Count
    foreach ($f in $files) {
        $i++
        $base = [System.IO.Path]::GetFileNameWithoutExtension($f.FullName)
        $out  = Join-Path $InDir ($base + " (Glitched).mp4")

        Write-Host ""
        Write-Info ("[{0}/{1}] Input : {2}" -f $i,$n,$f.Name)
        Write-Info ("[{0}/{1}] Output: {2}" -f $i,$n,([System.IO.Path]::GetFileName($out)))

        $filters = '[0:v]setpts=2.0*PTS[v];[0:a]asetpts=2.0*PTS[a]'
        $inSec = Get-DurationSec $f.FullName
        $outTotalSec = [Math]::Max(0.01, $inSec * 2.0)

        $parts = @()
        $parts += '-hide_banner -y'
        $parts += ('-i "{0}"' -f $f.FullName)
        $parts += ('-filter_complex "{0}"' -f $filters)
        $parts += '-map [v] -map [a]'
        $parts += '-c:v libx264 -r 30 -pix_fmt yuv420p'
        $parts += '-c:a aac -b:a 320k'
        $parts += '-movflags +faststart'
        $parts += '-progress pipe:1 -nostats -v error'
        $parts += ('"{0}"' -f $out)
        $argString = ($parts -join ' ')

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $FFMPEG
        $psi.Arguments = $argString
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.UseShellExecute        = $false

        $p = [System.Diagnostics.Process]::Start($psi)
        if (-not $p) { Write-Err "Failed to start ffmpeg."; continue }

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $lastPct = -1.0
        while (-not $p.HasExited) {
            $line = $p.StandardOutput.ReadLine()
            if ($null -ne $line -and $line -match '^out_time_ms=(\d+)') {
                $sec = [double]$Matches[1] / 1000000.0
                $pct = [Math]::Min(100.0, [Math]::Max(0.0, ($sec / $outTotalSec) * 100.0))
                if ([Math]::Abs($pct - $lastPct) -gt 0.2) {
                    $etaSec = [Math]::Max(0.0, $outTotalSec - $sec)
                    $progressText = ("[{0}/{1}] Encoding  Elapsed {2:hh\:mm\:ss}  ETA {3:hh\:mm\:ss}" -f $i,$n,$sw.Elapsed,[TimeSpan]::FromSeconds($etaSec))
                    Show-Progress -TotalValue 100 -CurrentValue ([single]$pct) -ProgressText $progressText -BarSize 40
                    $lastPct = $pct
                }
            }
            Start-Sleep -Milliseconds 40
        }
        $p.WaitForExit()
        Show-Progress -TotalValue 100 -CurrentValue 100 -ProgressText ("[{0}/{1}] Encoding  Elapsed {2:hh\:mm\:ss}  ETA 00:00:00" -f $i,$n,$sw.Elapsed) -BarSize 40 -Complete
        Write-Host ""

        if (Test-Path $out) {
            Write-Info "Done."
        } else {
            Write-Err "Error exporting: $($f.Name)"
            $err = $p.StandardError.ReadToEnd()
            if ($err) { Write-Host $err }
        }
    }

    Write-Host ""
    Write-Info "LEGACY BATCH MODE finished."
    pause
    exit 0
}

# Enhanced Modes (1 & 2) 
# Query encoder support first
function Has-NVENC {
    try {
        $enc = & $FFMPEG -hide_banner -encoders 2>&1
        return ($enc -match 'h264_nvenc')
    } catch { return $false }
}
$nvencAvailable = Has-NVENC

# Profile/Options with prompts
function Ask($prompt, $default){
    $v = Read-Host ("{0} [{1}]" -f $prompt,$default)
    if([string]::IsNullOrWhiteSpace($v)){ return $default } else { return $v }
}

# Choose encoder (auto-fallback if user picks GPU without support)
Write-Note "Choose encoder:"
Write-Host "  1) GPU (NVIDIA NVENC) - faster"
Write-Host "  2) CPU (x264)         - maximum compatibility"
$encChoice = Read-Host "Enter 1 or 2"
$UseGPU = $false
if ($encChoice -eq '1') {
    if($nvencAvailable){ $UseGPU = $true } else { Write-Warn "NVENC not available. Falling back to CPU."; $UseGPU = $false }
} elseif ($encChoice -eq '2') {
    $UseGPU = $false
} else {
    Write-Warn "Unknown choice, defaulting to CPU (x264)."
    $UseGPU = $false
}

# Ask for profile and 4K
$profile = Ask "Profile (classic/smooth60)" "classic"
if($profile -notin @('classic','smooth60')){ Write-Warn "Unknown profile. Using classic."; $profile='classic' }
$UpscaleTo4K = (Ask "Upscale to 4K? (y/n)" "y") -match '^(y|Y)'

# Optional: adjust CPU CRF/Preset
$CRF    = [int](Ask "CPU CRF (lower=better, 16-24 typical)" "18")
$PRESET = Ask "CPU Preset" "slower"

# Derived
$SlowmoFactor = 2.0
$FPS = if ($profile -eq 'classic') { 30 } else { 60 }
$EnableDenoise = $true
$ColorEq = 'eq=contrast=1.15:brightness=0.02:saturation=1.35:gamma=1.1'
$Sharpen = 'unsharp=5:5:1.2:5:5:0.8'
$TargetW = 3840
$TargetH = 2160

# Output naming
$BaseName = [System.IO.Path]::GetFileNameWithoutExtension($InFile)
$OutDir   = $InDir
$SuffixParts = @()
if ($UseGPU) { $SuffixParts += 'GPU' } else { $SuffixParts += 'CPU' }

# PowerShell 5.1-safe replacements for ternary
$SuffixParts += $( if ($profile -eq 'classic') { 'Classic' } else { 'Smooth60' } )
$SuffixParts += $( if ($UpscaleTo4K) { '4K' } else { 'OrigRes' } )

$Suffix = " (Glitched - " + ($SuffixParts -join ', ') + ").mp4"
$OutFile  = Join-Path $OutDir ($BaseName + $Suffix)

Write-Host ""
Write-Info ("Input : {0}" -f $InFile)
Write-Info ("Output: {0}" -f $OutFile)
$EncLabel = if ($UseGPU) { 'GPU (NVENC)' } else { 'CPU (x264)' }
Write-Host ("Mode  : {0} | Slow x{1} | FPS {2} | 4K Upscale: {3} | Encoder: {4}" -f $profile, $SlowmoFactor, $FPS, $UpscaleTo4K, $EncLabel)
Line

# Confirm overwrite
if (Test-Path $OutFile) {
    $ow = Read-Host "Output exists. Overwrite? (y/n) [n]"
    if ($ow -notmatch '^(y|Y)$') { Write-Warn "Cancelled by user."; exit 0 }
}

# Timings & audio info
$inSec = Get-DurationSec $InFile
$outTotalSec = [Math]::Max(0.01, $inSec * $SlowmoFactor)
$estimate = [TimeSpan]::FromSeconds($outTotalSec)
Write-Host ("Estimated output duration (playback): {0:hh\:mm\:ss} (~{1:0.##} s)" -f $estimate, $outTotalSec)
$audioPresent = Has-Audio $InFile
Write-Host ("Audio stream detected: {0}" -f ($(if ($audioPresent) { 'Yes' } else { 'No' })))
Line

# Build video filters
$vfChainList = @()
if ($EnableDenoise) { $vfChainList += "atadenoise=0a=0.02:0b=0.04:1a=0.02:1b=0.04" }
$vfChainList += $ColorEq
$vfChainList += $Sharpen
$vfChainList += "setpts=$($SlowmoFactor)*PTS"
if ($profile -eq 'smooth60') {
    $vfChainList += "minterpolate=mi_mode=mci:mc_mode=aobmc:vsbmc=1:fps=$FPS"
}
if ($UpscaleTo4K) {
    $vfChainList += ('scale={0}:{1}:flags=lanczos+accurate_rnd+full_chroma_int' -f $TargetW, $TargetH)
}
$vf = ($vfChainList -join ",")

$filters = @("[0:v]$vf[v]")
$maps    = @('-map','[v]')

# Build audio filters
if ($audioPresent) {
    if ($profile -eq 'classic') {
        $filters += "[0:a]asetpts=$($SlowmoFactor)*PTS[a]"
    } else {
        $atempoChain = Get-AtempoChain -factor (1.0 / $SlowmoFactor)
        $filters += "[0:a]$atempoChain,aresample=async=1:min_comp=0.001:first_pts=0[a]"
    }
    $maps += @('-map','[a]')
}
$filterComplex = ($filters -join ';')

# Encoder args
$VideoCodecArgs = @()
if ($UseGPU) {
    $VideoCodecArgs += '-c:v h264_nvenc'
    $VideoCodecArgs += '-preset p5 -rc vbr -cq 19 -b:v 0 -profile high -level 5.2'
} else {
    $VideoCodecArgs += '-c:v libx264'
    $VideoCodecArgs += ("-preset {0}" -f $PRESET)
    $VideoCodecArgs += ("-crf {0}" -f $CRF)
    $VideoCodecArgs += '-profile:v high -level 5.2'
}

# Build args & run
$parts = @()
$parts += '-hide_banner -y'
$parts += ('-i "{0}"' -f $InFile)
$parts += ('-filter_complex "{0}"' -f $filterComplex)
$parts += ($maps -join ' ')
$parts += $VideoCodecArgs
$parts += ('-r {0}' -f $FPS)
$parts += '-pix_fmt yuv420p -colorspace bt709 -color_primaries bt709 -color_trc bt709'
if ($audioPresent) { $parts += '-c:a aac -b:a 256k -ar 48000' } else { $parts += '-an' }
$parts += '-movflags +faststart'
$parts += '-progress pipe:1 -nostats -v error'
$parts += ('"{0}"' -f $OutFile)
$argString = ($parts -join ' ')

Write-Note "ffmpeg command:"
Write-Host ($FFMPEG + " " + $argString)
Line

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $FFMPEG
$psi.Arguments = $argString
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError  = $true
$psi.UseShellExecute = $false

$p = [System.Diagnostics.Process]::Start($psi)
if (-not $p) { Write-Err "Failed to start ffmpeg."; exit 10 }

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$lastPct = -1.0
while (-not $p.HasExited) {
    $line = $p.StandardOutput.ReadLine()
    if ($null -ne $line -and $line -match '^out_time_ms=(\d+)') {
        $sec = [double]$Matches[1] / 1000000.0
        $pct = [Math]::Min(100.0, [Math]::Max(0.0, ($sec / $outTotalSec) * 100.0))
        if ([Math]::Abs($pct - $lastPct) -gt 0.2) {
            $etaSec = [Math]::Max(0.0, $outTotalSec - $sec)
            $progressText = ("Encoding  Elapsed {0:hh\:mm\:ss}  ETA {1:hh\:mm\:ss}" -f $sw.Elapsed, [TimeSpan]::FromSeconds($etaSec))
            Show-Progress -TotalValue 100 -CurrentValue ([single]$pct) -ProgressText $progressText -BarSize 40
            $lastPct = $pct
        }
    }
    Start-Sleep -Milliseconds 40
}
$p.WaitForExit()

Show-Progress -TotalValue 100 -CurrentValue 100 -ProgressText ("Encoding  Elapsed {0:hh\:mm\:ss}  ETA 00:00:00" -f $sw.Elapsed) -BarSize 40 -Complete
Write-Host ""

$stderrText = $p.StandardError.ReadToEnd()
if (Test-Path $OutFile) {
    $sw.Stop()
    Write-Info ("DONE. Export time: {0:hh\:mm\:ss}" -f $sw.Elapsed)
    Write-Host "Saved to: $OutFile"
    exit 0
} else {
    Write-Err "Export failed."
    if ($stderrText) {
        Line
        Write-Host "ffmpeg error output:"
        Line
        Write-Host $stderrText
        Line
    }
    exit 1
}
