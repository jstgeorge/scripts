# Description: This script downloads a portion of a YouTube video using yt-dlp and ffmpeg based on the start and end timestamps provided by the user. It supports the following extensions: mp4, mkv, flv, ogg, webm, avi, and mp3.
# Author: jstgeorge
# Date: 2024-04-29
# Version: 1.0
# Dependencies: yt-dlp, ffmpeg
# Installation: Download yt-dlp from https://github.com/yt-dlp/yt-dlp
# Installation: Download ffmpeg from https://ffmpeg.org/download.html
# Installation: You must configure your path to use yt-dlp and ffmpeg.
# Usage: Run the script and follow the prompts to enter the YouTube URL, start time, end time, and desired file extension.
# If you want to automate the process, you can set the variables at the beginning of the script to negate the Read-Host prompts.

#$youtubeUrl = "https://www.youtube.com/watch?v=3zD7rBcgduw"
#$startTime = "00:17:04.00"
#$endTime = "00:17:13.00"
#$fileExtension = "mp3"

# Check if yt-dlp.exe and ffmpeg.exe are present in the current directory
Write-Output "Checking dependencies..."

try {
    $checkYtpDlp = Get-Command yt-dlp -ErrorAction Stop
    $checkFfmpeg = Get-Command ffmpeg -ErrorAction Stop
} catch {
    Write-Output "Dependency check failed. Please ensure yt-dlp and ffmpeg are installed and accessible from the current directory."
    exit
}

Write-Output "Dependencies are met."

# Get the YouTube URL, start time, end time and fileExtension from the user
# Only prompt the user if the variables are not already set
if (!$youtubeUrl) {
    $youtubeUrl = Read-Host -Prompt "Enter the YouTube URL"
}
if (!$startTime) {
    $startTime = Read-Host -Prompt "Enter the start time (HH:MM:SS.MS)"
}
if (!$endTime) {
    $endTime = Read-Host -Prompt "Enter the end time (HH:MM:SS.MS)"
}
if (!$fileExtension) {
    $fileExtension = Read-Host -Prompt "Enter the desired file extension (mp4|mkv|flv|ogg|webm|avi|mp3)"
}
Write-Output "youtubeUrl: $youtubeUrl"
Write-Output "startTime: $startTime"
Write-Output "endTime: $endTime"
Write-Output "fileExtension: $fileExtension"

# Validate the YouTube URL
$youtubeUrlRegex = "^(https?://)?(www\.)?(youtube\.com|youtu\.?be)/.+v=([^&]+)"
if ($youtubeUrl -notmatch $youtubeUrlRegex) {
    Write-Output "Invalid YouTube URL. Please enter a valid YouTube URL."
    exit
}

# Validate the timestamp format
$timestampRegex = "^\d{2}:\d{2}:\d{2}.\d{2}$"
$timestampStrings = @($startTime, $endTime)
foreach ($timestamp in $timestampStrings) {
    if ($timestamp -notmatch $timestampRegex) {
        Write-Output "Invalid timestamp format. Please use the HH:MM:SS.MS format."
        exit
    }
}

# Validate the file extension
if ($fileExtension -notmatch "^(mp4|mkv|flv|ogg|webm|avi|mp3)$") {
    Write-Output "Invalid file extension. Please enter a valid file extension (mp4|mkv|flv|ogg|webm|avi|mp3)."
    exit
}

# Get the current date and time
$currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"

# Specify the output file name with the video title, the current date and time, and the desired file extension
$fileName = "%(title)s_$currentDateTime.$fileExtension"

# Construct the yt-dlp command
if ($fileExtension -eq "mp3") {
    $ytDlpCommand = "yt-dlp -f `"(bestvideo+bestaudio/best)[protocol!*=dash]`" --extract-audio --audio-format mp3 --force-keyframes-at-cuts --external-downloader ffmpeg --external-downloader-args `"ffmpeg_i:-ss $startTime -to $endTime`" -o `"$fileName`" `"$youtubeUrl`""
} else {
    $ytDlpCommand = "yt-dlp -f `"(bestvideo+bestaudio/best)[protocol!*=dash]`" --merge-output-format $fileExtension --force-keyframes-at-cuts --external-downloader ffmpeg --external-downloader-args `"ffmpeg_i:-ss $startTime -to $endTime`" -o `"$fileName`" `"$youtubeUrl`""
}

# Execute the yt-dlp command
Invoke-Expression $ytDlpCommand


if ($LASTEXITCODE -ne 0) {
    Write-Output "Error occured. Please check the error message above."
    exit
}else{
    Write-Output "Filename is above."
    Write-Output "Download completed successfully."
}