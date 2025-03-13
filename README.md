# psMediaConvert

This powershell program will look at your media and shrink the filesize, it provides log files, an html status page showing how much data has been saved.

You will need to download HandBrakeCLI and ffmpeg.

## Variables you can edit

### V2 Variables

- $processingFile = "D:\MEDIA\processing.txt"
  - This is a file that keeps track of what is being processed, basically a short term log.
- $processingIgnore = @("-chanche", "processing-temp.mkv")
  - This should at a minimum include your processing tag and the temp filename that is used to process the video
  - I should really update to just use the $processingTag, this tag tell the program to not process the file

### Paths

- $handbrakeCliExe = "C:\HandBrakeCLI\HandBrakeCLI.exe"
  - Point this to your HandBrakeCLI location
- $ffprobeExe = "C:\ffmpeg\bin\ffprobe.exe"
  - Point this to your ffprobe location
- $mediaDir = "D:\MEDIA"
  - This is where your media lives that you want to shink
- $logFile = "D:\MEDIA\psMC.log"
  - This will track everything that happens, a long term log file with rotation and compression
- $logArchivePath = "D:\MEDIA\logArchive"
  - This folder will store archived log files
- $htmlLogFile = "D:\MEDIA\psMC.html"
  - This is a self refreshing webpage that will give you status on current jobs and how much data was saved
- $savedSpaceLogFile = "D:\MEDIA\psMC-saved-space.log"
  - This log tracks how much data was saved by reprocessing media
- $mediaDrive = "D"
  - Drive letter where your media lives

### Bitrates - If FFMPEG can do it its available

Define the audio and video bitrates (in kbps), this will be used to estimate the final file size to determine if we should process the file

- $videoBitrate = 2800
- $audioBitrate = 128
- $videoEncoder = "nvenc_h265"
- $audioEncoder = "mp3"
- $videoVerticalResolution = 1080

### File Tagging

- $processedTag = "-chanche"
  - This is a custom tag added to files to tell the program the media has been modified

### Source File Extensions

- $mediaExtensions = @("*.mkv", "*.mp4", "*.avi", "*.mov", "*.flv", "*.wmv", "*.mpg", "*.mpeg", "*.m4v", "*.ts", "*.mts", "*.m2ts")
  - This is a list of files the program will process, if FFMPEG supports it you can add/remove as needed.

### Log file size in MB before archiving happens

- $maxLogFileSize = 1
  - Size in MB before the log is rotated

### Sleep time in seconds

- $sleepTime = (.5*3600)
  - This is the sleep time between runs, if you are unsure, set this to be once per day.
