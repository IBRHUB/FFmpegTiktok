# TikTok Shooter Color+4K Glitch Exporter (PS 5.1)

$FFMPEG  = "C:\ffmpeg\bin\ffmpeg.exe"
$FFPROBE = "C:\ffmpeg\bin\ffprobe.exe"

if (-not (Test-Path $FFMPEG))  { Write-Host "ERROR: ffmpeg not found at $FFMPEG"; exit 1 }
if (-not (Test-Path $FFPROBE)) { Write-Host "ERROR: ffprobe not found at $FFPROBE"; exit 1 }

Add-Type -AssemblyName System.Windows.Forms
$dlg = New-Object Windows.Forms.OpenFileDialog
$dlg.Filter = "Video Files|*.mp4;*.mov;*.mkv;*.avi;*.webm;*.flv;*.wmv|All Files|*.*"
$dlg.Multiselect = $false
if ($dlg.ShowDialog() -ne "OK") { Write-Host "No file selected. Exiting."; exit 2 }
$InFile = $dlg.FileName

# ===== Profiles =====
# classic  = تبطيء ×2 بالصوت والصورة عبر setpts (الصوت يطيح pitch) + FPS 30 (مطابق للسكربت القديم)
# smooth60 = تبطيء ×2 بالصورة + 60fps minterpolate، والصوت محافظ على البتش (atempo)
$GlitchProfile = 'classic'   # 'classic' or 'smooth60'

# ===== Encoding / Color Controls =====
$CRF     = 18
$PRESET  = 'slower'
$EnableDenoise = $true                  # فلتر تنعيم خفيف قبل كل شيء
# ضبط ألوان مناسب للشوتر:
$ColorEq      = 'eq=contrast=1.15:brightness=0.02:saturation=1.35:gamma=1.1'
$Sharpen      = 'unsharp=5:5:1.2:5:5:0.8'

# ===== Resolution Controls =====
$UpscaleTo4K = $true                    # حسب طلبك: ارفع إلى 4K
$KeepOriginalResolution = -not $UpscaleTo4K
$TargetW = 3840
$TargetH = 2160

# ===== Time / FPS Controls =====
$SlowmoFactor = 2.0                     # يُطابق السكربت القديم
$FPS = if ($GlitchProfile -eq 'classic') { 30 } else { 60 }

# ===== Output naming =====
$BaseName = [System.IO.Path]::GetFileNameWithoutExtension($InFile)
$OutDir   = [System.IO.Path]::GetDirectoryName($InFile)
$SuffixParts = @()
if ($GlitchProfile -eq 'classic') { $SuffixParts += 'Classic' } else { $SuffixParts += 'Smooth60' }
if ($UpscaleTo4K) { $SuffixParts += '4K' } else { $SuffixParts += 'OrigRes' }
$Suffix = " (Glitched - " + ($SuffixParts -join ', ') + ").mp4"
$OutFile  = Join-Path $OutDir ($BaseName + $Suffix)

Write-Host "Input : $InFile"
Write-Host "Output: $OutFile"
Write-Host "Mode  : $GlitchProfile | Slow x$SlowmoFactor | FPS $FPS | 4K Upscale: $UpscaleTo4K"
Write-Host ""

# ===== Probe helpers =====
function Get-DurationSec([string]$p) {
    $raw = & $FFPROBE -v error -show_entries format=duration -of default=nokey=1:noprint_wrappers=1 -- "$p"
    if (-not $raw) { return 0.0 }
    try { return [double]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture) } catch { return 0.0 }
}
function Has-Audio([string]$p) {
    $a = & $FFPROBE -v error -select_streams a:0 -show_entries stream=index -of csv=p=0 -- "$p"
    return -not [string]::IsNullOrWhiteSpace($a)
}
# atempo chain (لـ smooth60 فقط)
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

# ===== Progress bar =====
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
    }
    else {
        $filled = [int]([math]::Round($BarSize * $percent))
        if ($filled -lt 0) { $filled = 0 }
        if ($filled -gt $BarSize) { $filled = $BarSize }
        $bar = ''.PadRight($filled, [char]9608).PadRight($BarSize, [char]9617)
        Write-Host -NoNewLine ("`r{0} {1} {2,6:N2} % " -f $ProgressText, $bar, $percentComplete)
        if ($Complete) { Write-Host "" }
    }
}

# ===== Timings =====
$inSec = Get-DurationSec $InFile
$outTotalSec = [Math]::Max(0.01, $inSec * $SlowmoFactor)  # لأننا نبطّئ دائمًا ×SlowmoFactor في كلا البروفايلين
$estimate = [TimeSpan]::FromSeconds($outTotalSec)
Write-Host ("Estimated output duration: {0:hh\:mm\:ss} (~{1:0.##} s)" -f $estimate, $outTotalSec)

$audioPresent = Has-Audio $InFile
$audioText = if ($audioPresent) { "Yes" } else { "No" }
Write-Host ("Audio stream detected: {0}" -f $audioText)
Write-Host ""

# ===== Build filter graph (VIDEO) =====
$vfChainList = @()

if ($EnableDenoise) {
    $vfChainList += "atadenoise=0a=0.02:0b=0.04:1a=0.02:1b=0.04"
}

# تحسين ألوان الشوتر
$vfChainList += $ColorEq
$vfChainList += $Sharpen

# تبطيء الفيديو
$vfChainList += "setpts=$($SlowmoFactor)*PTS"

# minterpolate لـ Smooth60 فقط
if ($GlitchProfile -eq 'smooth60') {
    $vfChainList += "minterpolate=mi_mode=mci:mc_mode=aobmc:vsbmc=1:fps=$FPS"
}

# رفع الدقة إلى 4K (أو الإبقاء كما هو)
if ($UpscaleTo4K) {
   $vfChainList += ('scale={0}:{1}:flags=lanczos+accurate_rnd+full_chroma_int' -f $TargetW, $TargetH)
}

$vf = ($vfChainList -join ",")

$filters = @("[0:v]$vf[v]")
$maps    = @('-map','[v]')

# ===== Build audio chain =====
if ($audioPresent) {
    if ($GlitchProfile -eq 'classic') {
        # يطابق السكربت القديم: نزّل البتش (asetpts)
        $filters += "[0:a]asetpts=$($SlowmoFactor)*PTS[a]"
    } else {
        # smooth60: حافظ على البتش عبر atempo
        $atempoChain = Get-AtempoChain -factor (1.0 / $SlowmoFactor)
        $filters += "[0:a]$atempoChain,aresample=async=1:min_comp=0.001:first_pts=0[a]"
    }
    $maps += @('-map','[a]')
}
$filterComplex = ($filters -join ';')

# ===== Build args =====
$parts = @()
$parts += '-hide_banner -y'
$parts += ('-i "{0}"' -f $InFile)
$parts += ('-filter_complex "{0}"' -f $filterComplex)
$parts += ($maps -join ' ')
$parts += '-c:v libx264'
$parts += ('-preset {0}' -f $PRESET)
$parts += ('-crf {0}' -f $CRF)
$parts += '-profile:v high -level 5.2'  # 4K-compatible level
$parts += ('-r {0}' -f $FPS)
$parts += '-pix_fmt yuv420p -colorspace bt709 -color_primaries bt709 -color_trc bt709'
if ($audioPresent) { $parts += '-c:a aac -b:a 256k -ar 48000' } else { $parts += '-an' }
$parts += '-movflags +faststart'
$parts += '-progress pipe:1 -nostats -v error'
$parts += ('"{0}"' -f $OutFile)
$argString = ($parts -join ' ')

Write-Host "ffmpeg command:"
Write-Host ($FFMPEG + " " + $argString)
Write-Host ""

# ===== Start ffmpeg =====
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $FFMPEG
$psi.Arguments = $argString
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError  = $true
$psi.UseShellExecute = $false

$p = [System.Diagnostics.Process]::Start($psi)
if (-not $p) { Write-Host "ERROR: failed to start ffmpeg."; exit 10 }

# ===== Progress loop using Show-Progress =====
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

# finalize progress line
Show-Progress -TotalValue 100 -CurrentValue 100 -ProgressText ("Encoding  Elapsed {0:hh\:mm\:ss}  ETA 00:00:00" -f $sw.Elapsed) -BarSize 40 -Complete
Write-Host ""

$stderrText = $p.StandardError.ReadToEnd()

if (Test-Path $OutFile) {
    $sw.Stop()
    Write-Host ("DONE. Export time: {0:hh\:mm\:ss}" -f $sw.Elapsed)
    Write-Host "Saved to: $OutFile"
} else {
    Write-Host "ERROR: Export failed."
    if ($stderrText) {
        Write-Host ""
        Write-Host "ffmpeg error output:"
        Write-Host "---------------------"
        Write-Host $stderrText
        Write-Host "---------------------"
    }
    exit 1
}
