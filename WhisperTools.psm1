# WhisperTools.psm1  single-file module

# --- CONSTANTS (persist in user profile) ---
$script:ToolsHome     = Join-Path $env:USERPROFILE ".whisper-tools"
$script:VenvPath      = Join-Path $script:ToolsHome "whisper-cpu"
$script:WhisperExe    = Join-Path $script:VenvPath "Scripts\whisper.exe"
$script:ModelsCache   = Join-Path $env:LOCALAPPDATA "whisper"
$script:WingetCommon  = "--silent --accept-source-agreements --accept-package-agreements --scope user"

# Ensure home folder exists
if (-not (Test-Path $script:ToolsHome)) { New-Item -ItemType Directory -Path $script:ToolsHome -Force | Out-Null }

function Test-CommandAvailable {
    param([Parameter(Mandatory)][string]$Name)
    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

function Resolve-OutputFormatFromPath {
    param([Parameter(Mandatory)][string]$Path)
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant().TrimStart('''.''')
    switch ($ext) {
        "txt" { "txt" }
        "srt" { "srt" }
        "vtt" { "vtt" }
        default { "srt" }
    }
}

function Install-VideoToAudio {
    [CmdletBinding()]
    param()

    if (-not (Test-CommandAvailable -Name "winget")) {
        throw "winget is not available. Install App Installer from Microsoft Store, then re-run."
    }

    Write-Host "Installing FFmpeg (Gyan.FFmpeg) via winget for current user..."
    $id = "Gyan.FFmpeg"
    $args = @("install","-e","--id=$id") + $script:WingetCommon.Split(" ")
    $process = Start-Process -FilePath "winget" -ArgumentList $args -NoNewWindow -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "winget failed installing FFmpeg (exit $($process.ExitCode))."
    }

    # Locate ffmpeg installation directory
    $ffmpegDir = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe\ffmpeg-*-full_build\bin"
    $resolvedPaths = Get-ChildItem -Path $ffmpegDir -ErrorAction SilentlyContinue
    if ($resolvedPaths) {
        $ffmpegPath = $resolvedPaths.FullName
        Write-Host "Adding FFmpeg to PATH: $ffmpegPath"

        # Add to PATH for current session
        $env:Path += ";$ffmpegPath"

        # Persist PATH update for future sessions
        [Environment]::SetEnvironmentVariable("Path", $env:Path, [EnvironmentVariableTarget]::User)

        Write-Host "FFmpeg is ready and added to PATH." -ForegroundColor Green
    } else {
        Write-Warning "FFmpeg installed, but could not locate installation directory. Restart PowerShell or log off/on."
    }
}

function Install-AudioToTxt {
    [CmdletBinding()]
    param(
        [string]$PythonId = "Python.Python.3.13"
    )

    if (-not (Test-CommandAvailable -Name "winget")) {
        throw "winget is not available. Install App Installer from Microsoft Store, then re-run."
    }

    # 1) Install Python (user-scope)
    Write-Host "Installing Python via winget (user scope): $PythonId ..."
    $args = @("install","-e","--id=$PythonId") + $script:WingetCommon.Split(" ")
    $process = Start-Process -FilePath "winget" -ArgumentList $args -NoNewWindow -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        Write-Warning "winget returned code $($process.ExitCode). Continuing if Python is already present in PATH..."
    }

    # Prefer "py" launcher if present, otherwise "python"
    $pythonLaunchers = @("py","python")
    $py = $null
    foreach ($cand in $pythonLaunchers) {
        if (Test-CommandAvailable -Name $cand) { $py = $cand; break }
    }
    if (-not $py) { throw "Python launcher not found in PATH after install. Restart PowerShell and try again." }

    # 2) Create venv in user profile
    if (-not (Test-Path $script:VenvPath)) {
        Write-Host "Creating venv in $script:VenvPath ..."
        & $py -m venv $script:VenvPath
        if ($LASTEXITCODE -ne 0) { throw "Failed to create venv at $script:VenvPath" }
    } else {
        Write-Host "Venv already exists at $script:VenvPath"
    }

    $pip = Join-Path $script:VenvPath "Scripts\pip.exe"
    $python = Join-Path $script:VenvPath "Scripts\python.exe"
    if (-not (Test-Path $pip)) { throw "pip not found in venv. Venv may be corrupted." }

    # 3) Upgrade pip/wheel/setuptools
    & $python -m pip install --upgrade pip wheel setuptools
    if ($LASTEXITCODE -ne 0) { throw "Failed to upgrade pip/wheel/setuptools." }

    # 4) Install CPU Torch and Whisper
    Write-Host "Installing Torch (CPU) ..."
    & $pip install torch --index-url https://download.pytorch.org/whl/cpu
    if ($LASTEXITCODE -ne 0) { throw "Failed to install torch (cpu)." }

    Write-Host "Installing openai-whisper ..."
    & $pip install -U openai-whisper
    if ($LASTEXITCODE -ne 0) { throw "Failed to install openai-whisper." }

    # 5) Create models cache folder (optional)
    if (-not (Test-Path $script:ModelsCache)) { New-Item -ItemType Directory -Path $script:ModelsCache -Force | Out-Null }

    Write-Host "AudioText environment is ready." -ForegroundColor Green
}

function Convert-VideoToAudio {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InputFile,
        [string]$Output
    )

    function Find-Ffmpeg {
        # Try PATH first
        if (Test-CommandAvailable -Name "ffmpeg") { return "ffmpeg" }

        # Expect ffmpeg in the WinGet path
        $wingetPath = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe\ffmpeg-*-full_build\bin\ffmpeg.exe"
        $resolvedPaths = Get-ChildItem -Path $wingetPath -ErrorAction SilentlyContinue
        if ($resolvedPaths) { return $resolvedPaths.FullName }

        throw "ffmpeg is not available. Run Install-VideoToAudio or add ffmpeg to PATH."
    }

    $ffmpegPath = Find-Ffmpeg

    if (-not (Test-Path $InputFile)) { throw "Input not found: $InputFile" }

    if (-not $Output) {
        $inputDir = [System.IO.Path]::GetDirectoryName($InputFile)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
        $Output = Join-Path $inputDir "$base.wav"
    }

    Write-Host "Converting $InputFile -> $Output (16 kHz mono) ..."
    & $ffmpegPath -y -i "$InputFile" -ar 16000 -ac 1 "$Output"
    if ($LASTEXITCODE -ne 0) { throw "ffmpeg failed to convert." }
    Write-Host "Audio written: $Output" -ForegroundColor Green
}

function Convert-AudioToTxt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InputFile,
        [string]$Output,
        [ValidateSet("tiny","base","small","medium","large-v3")]
        [string]$Model = "medium",
        [string]$Language = "ru",
        [int]$Threads,
        [ValidateSet("txt", "srt", "vtt")]
        [string]$Format = "txt"
    )

    # Ensure venv + whisper exist
    if (-not (Test-Path $script:WhisperExe)) {
        throw "Whisper CLI not found in venv. Run Install-AudioToTxt first."
    }
    if (-not (Test-Path $InputFile)) { throw "Input not found: $InputFile" }

    # Determine output file
    if (-not $Output) {
        $inputDir = [System.IO.Path]::GetDirectoryName($InputFile)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
        $Output = Join-Path $inputDir "$base.$Format"
    }

    # Whisper always names output by input base; we will rename to user-specified file after.
    $inBaseName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    $workDir = $inputDir

    # Build arguments
    $args = @(
        $InputFile,
        "--model", $Model,
        "--language", $Language,
        "--task", "transcribe",
        "--output_dir", $workDir,
        "--output_format", $Format,
        "--word_timestamps", "True",
        "--fp16", "False"
    )
    if ($Threads -gt 0) { $args += @("--threads", $Threads) }

    Write-Host "Transcribing (CPU-only) using model $Model, language $Language, format $Format ..."
    & $script:WhisperExe @args
    if ($LASTEXITCODE -ne 0) { throw "Whisper transcribe failed." }

    Write-Host "Transcript written: $Output" -ForegroundColor Green
}

function Convert-VideoToTxt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InputFile,
        [string]$Output,
        [ValidateSet("tiny","base","small","medium","large-v3")]
        [string]$Model = "medium",
        [string]$Language = "ru",
        [int]$Threads,
        [ValidateSet("txt", "srt", "vtt")]
        [string]$Format = "txt"
    )

    # Step 1: Convert video to audio
    $tempAudioFile = [System.IO.Path]::GetTempFileName() + ".wav"
    try {
        Convert-VideoToAudio -InputFile $InputFile -Output $tempAudioFile

        # Step 2: Transcribe audio to text
        Convert-AudioToTxt -InputFile $tempAudioFile -Output $Output -Model $Model -Language $Language -Threads $Threads -Format $Format
    } finally {
        # Clean up temporary audio file
        if (Test-Path $tempAudioFile) {
            Remove-Item -Path $tempAudioFile -Force
        }
    }
}

Export-ModuleMember -Function Install-VideoToAudio, Install-AudioToTxt, Convert-VideoToAudio, Convert-AudioToTxt, Convert-VideoToTxt
