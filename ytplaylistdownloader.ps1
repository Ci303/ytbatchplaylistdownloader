# Check if the script is run with administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')

if (-not $isAdmin) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Check if Chocolatey is installed, install if not
$chocoPath = Get-Command choco.exe -ErrorAction SilentlyContinue

if ($chocoPath -eq $null) {
    Set-ExecutionPolicy Bypass -Scope Process -Force
    iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    RefreshEnv
} else {
    Write-Host "Chocolatey is already installed."
}

# Check and update yt-dlp using Chocolatey
$ytDlpPath = Get-Command yt-dlp.exe -ErrorAction SilentlyContinue
$ytDlpVersion = $null  # Initialize $ytDlpVersion

if ($ytDlpPath -eq $null) {
    Write-Host "Installing yt-dlp..."
    choco install yt-dlp -y
} else {
    Write-Host "yt-dlp is already installed."
    # Set $ytDlpVersion if yt-dlp is already installed
    $ytDlpVersion = & $ytDlpPath --version
}

# Check and update yt-dlp using Chocolatey
$latestVersion = ((choco list yt-dlp --local-only).Split('|')[1]).Trim()

if ($ytDlpPath -ne $null -and $ytDlpVersion -ne $latestVersion) {
    Write-Host "Updating yt-dlp..."
    choco upgrade yt-dlp -y
    # Re-check version after upgrade
    $ytDlpVersion = & $ytDlpPath --version
}

# Check and install ffmpeg using Chocolatey
$ffmpegPath = Get-Command ffmpeg.exe -ErrorAction SilentlyContinue

if ($ffmpegPath -eq $null) {
    choco install ffmpeg -y
} else {
    Write-Host "ffmpeg is already installed."
}

# Check and update ffmpeg if outdated using Chocolatey
$ffmpegVersionOutput = & $ffmpegPath -version

# Extracting version from output
$ffmpegVersion = $ffmpegVersionOutput | Select-String -Pattern 'version\s\d+\.\d+\.\d+' | ForEach-Object { $_.Matches[0].Value -replace 'version ' }

$latestFfmpegVersion = ((choco list --local-only ffmpeg).Split('|')[1]).Trim()

if ($ffmpegVersion -ne $latestFfmpegVersion) {
    Write-Host "Updating ffmpeg..."
    choco upgrade ffmpeg -y
    # Re-check version after upgrade
    $ffmpegVersionOutput = & $ffmpegPath -version
    $ffmpegVersion = $ffmpegVersionOutput | Select-String -Pattern 'version\s\d+\.\d+\.\d+' | ForEach-Object { $_.Matches[0].Value -replace 'version ' }
}

$ffmpegPath = "C:\ProgramData\chocolatey\bin\ffmpeg.exe"
$ytdlpPath = "C:\ProgramData\chocolatey\bin\yt-dlp.exe"
$downloadPath = "$env:userprofile\Desktop\YouTube"

# Function to sanitize a string for file names
function Sanitize-FileName($name) {
    $illegalChars = [IO.Path]::GetInvalidFileNameChars()
    $sanitized = $name -replace "[$([RegEx]::Escape($illegalChars))]", "_"
    $sanitized = $sanitized -replace "\.+", "."  # Remove consecutive dots
    return $sanitized
}

# Function to validate URL
function Validate-URL($url) {
    $urlRegex = "^(http|https)://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?$|^https://www\.youtube\.com/watch\?v=.*$|^https://youtu\.be/.*$"
    return $url -match $urlRegex
}

# Function to display a text-based progress bar
function Show-ProgressBar($completed, $total) {
    # Check if $total is zero to avoid division by zero
    if ($total -eq 0) {
        Write-Host -NoNewline "`r[No videos to download] 0% Complete"
        return
    }

    $percentComplete = 0
    try {
        $percentComplete = [math]::Round(($completed / $total) * 100, 2)
        if ($percentComplete -lt 0) { $percentComplete = 0 }
        if ($percentComplete -gt 100) { $percentComplete = 100 }
    } catch {
        Write-Host "Error in calculating percentage: $_"
    }

    $progressBar = ("=" * [math]::Round($percentComplete))
    $spaces = " " * (100 - $progressBar.Length)
    Write-Host -NoNewline "`r[$progressBar$spaces] $percentComplete% Complete"
}

# Continuous loop to download playlists
while ($true) {
    # Prompt for playlist URL until a valid URL is provided
    do {
        $playlistUrl = Read-Host "Enter the playlist or Video URL (or type 'exit' to end)"
        if ($playlistUrl -eq 'exit') {
            Write-Host "Goodbye!"
            exit
        }

        if (-not (Validate-URL $playlistUrl)) {
            Write-Host "Invalid URL. Please enter a valid URL or type 'exit' to end."
        }
    } until (Validate-URL $playlistUrl)

    # Download and process videos
    $playlistInfo = & $ytdlpPath --dump-json --flat-playlist $playlistUrl
    $playlist = $playlistInfo | ConvertFrom-Json

    # Debug output
    Write-Host "Total Videos: $($playlist.Count)"

    $totalVideos = $playlist.Count
    $processedVideos = 0

    foreach ($video in $playlist) {
        $processedVideos++
        Show-ProgressBar $processedVideos $totalVideos
        
        $videoTitle = $video.title
        $sanitizedTitle = Sanitize-FileName $videoTitle
        $outputFileName = "$downloadPath\$sanitizedTitle.%(ext)s"

        & $ytdlpPath --yes-playlist --sponsorblock-remove all --windows-filenames --remux-video mkv --audio-quality 0 --ffmpeg-location $ffmpegPath $video.url -o "$downloadPath/%(title)s.%(ext)s"
        # --write-thumbnail --embed-thumbnail  --add-metadata --merge-output-format mkv --audio-format mp3 --no-check-certificate --sponsorblock-mark all --no-write-comment --format bestvideo+bestaudio/best
    }

    Write-Host  # Move to a new line after the progress bar
    Write-Host "The script has finished. The files have been downloaded to: $downloadPath"
}
