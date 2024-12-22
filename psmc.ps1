<#
To do list:

#>

# Variables

# V2 Variables
$processingFile = "D:\MEDIA\processing.txt"
$processingIgnore = @("-chanche", "processing-temp.mkv")

# Paths
$handbrakeCliExe = "C:\HandBrakeCLI\HandBrakeCLI.exe"
$ffprobeExe = "C:\ffmpeg\bin\ffprobe.exe"
$mediaDir = "D:\MEDIA"
$logFile = "D:\MEDIA\psMC.log"
$logArchivePath = "D:\MEDIA\logArchive"
$htmlLogFile = "D:\MEDIA\psMC.html"
$savedSpaceLogFile = "D:\MEDIA\psMC-saved-space.log"
$mediaDrive = "D"

# Define the audio and video bitrates (in kbps), this will be used to estimate the final file size to determine if we should process the file
$videoBitrate = 2800
$audioBitrate = 128
$videoEncoder = "nvenc_h265"
$audioEncoder = "mp3"
$videoVerticalResolution = 1080

# File Tagging
$processedTag = "-chanche"

# Source File Extensions
$mediaExtensions = @("*.mkv", "*.mp4", "*.avi", "*.mov", "*.flv", "*.wmv", "*.mpg", "*.mpeg", "*.m4v", "*.ts", "*.mts", "*.m2ts")

# Log file size in MB before archiving happens
$maxLogFileSize = 1

# Sleep time in seconds
$sleepTime = (.5*3600)


# Functions

# Remove existing processing file
function removeProcessingFile($processingFile) {
    if (Test-Path $processingFile) {
        Remove-Item -Path $processingFile
        Write-Host "Removed existing processing file: $processingFile"
        logMessage "Removed existing processing file: $processingFile"

    }
}

# Create files if they don't exist
function createFilesIfNotExist($processingFile, $logFile, $htmlLogFile, $savedSpaceLogFile) {
    $files = @($processingFile, $logFile, $htmlLogFile, $savedSpaceLogFile)

    foreach ($file in $files) {
        if (!(Test-Path $file)) {
            New-Item -Path $file -ItemType File
            Write-Host "Created file: $file"
            logMessage "Created file: $file"
        }
    }
}

# Create archive directory if it doesn't exist
function ensureLogArchivePathExists($logArchivePath) {
    if (!(Test-Path $logArchivePath)) {
        New-Item -Path $logArchivePath -ItemType Directory
        Write-Host "Created log archive directory: $logArchivePath"
        logMessage "Created log archive directory: $logArchivePath"
    }
}

# Check if executables exist
function ensureExecutablesExist($handbrakeCliExe, $ffprobeExe) {
    if (!(Test-Path $handbrakeCliExe)) {
        Write-Host "HandBrakeCLI executable not found at $handbrakeCliExe"
        logMessage "HandBrakeCLI executable not found at $handbrakeCliExe"
    }
    if (!(Test-Path $ffprobeExe)) {
        Write-Host "ffprobe executable not found at $ffprobeExe"
        logMessage "ffprobe executable not found at $ffprobeExe"
    }
    Write-Host "Executables found"
    logMessage "Executables found"
}

# Initialize saved space log file
function initializeSavedSpaceLogFile($savedSpaceLogFile) {
    if (!(Test-Path $savedSpaceLogFile) -or !(Get-Content $savedSpaceLogFile)) {
        Set-Content -Path $savedSpaceLogFile -Value 0
        Write-Host "Initialized saved space log file: $savedSpaceLogFile"
        logMessage "Initialized saved space log file: $savedSpaceLogFile"
    }
}

# Loop through all files in the media directory and its subdirectories and gather files
function loadProcessingFile($mediaDir, $processingFile) {
    Get-ChildItem -Path $mediaDir -Recurse -File | ForEach-Object {
        Add-Content -Path $processingFile -Value $_.FullName
        Write-Host "Added file to processing file: $_"
        #logMessage "Added file to processing file: $_"
    }
}

# Clean processing file of non-media files
function cleanProcessingFileExt($processingFile, $mediaExtensions) {
    $content = Get-Content -Path $processingFile
    $cleanContent = $content | Where-Object {
        foreach ($extension in $mediaExtensions) {
            if ($_ -like "*$extension") {
                return $true
                Write-Host "Added file to processing file: $_"
                #logMessage "Added file to processing file: $_"
            }
        }
        return $false
    }
    Set-Content -Path $processingFile -Value $cleanContent
}

# Remove lines from processing file that contain the ignore patterns
function removeIgnoredLines($processingFile, $processingIgnore) {
    $content = Get-Content -Path $processingFile
    $filteredContent = $content | Where-Object {
        $line = $_
        $ignore = $false
        foreach ($ignorePattern in $processingIgnore) {
            if ($line -like "*$ignorePattern*") {
                $ignore = $true
                Write-Host "Removed file from processing file: $line"
                #logMessage "Removed file from processing file: $line"
                break
            }
        }
        return -not $ignore
    }
    Set-Content -Path $processingFile -Value $filteredContent
}



# Function to find out which process is locking a filefunction Get-LockingProcess {
    function Get-LockingProcess {
        param($filePath)
    
        # Path to handle.exe
        $handleExePath = "C:\Handle\handle.exe"
    
        # Create a new ProcessStartInfo
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $handleExePath
        $startInfo.Arguments = "-accepteula -nobanner -p *"
        $startInfo.RedirectStandardOutput = $true
        $startInfo.UseShellExecute = $false
    
        # Start the process
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        $process.Start() | Out-Null
    
        # Read the output
        $output = $process.StandardOutput.ReadToEnd()
    
        # Wait for the process to exit
        $process.WaitForExit()
    
        # Print the output to the console
        Write-Output $output
    }

# Test to see if a file is locked
function testFileLock {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    try {
        $fileStream = [System.IO.File]::Open($Path, 'Open', 'Read', 'None')
        if ($null -ne $fileStream) {
            $fileStream.Close()
        }
        return $false
    } catch {
        Write-Host "File is locked: $Path"
        logMessage "File is locked: $Path"

        # Call Get-LockingProcess to find out which process is locking the file
        Get-LockingProcess -filePath $Path

        return $true
    }
}




function getEstimatedFileSize {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [Parameter(Mandatory=$true)]
        [long]$OriginalFileSize
    )

    $duration = & $ffprobeExe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 $Path

    
    $estimatedFileSize = ((($videoBitrate + $audioBitrate) * $duration) / 8) * 1024
    Write-Host "Original File Size: $OriginalFileSize"
    Write-Host "Estimated File Size: $estimatedFileSize"
    logMessage "Original File Size: $OriginalFileSize"
    logMessage "Estimated File Size: $estimatedFileSize"

    if ($estimatedFileSize -lt $OriginalFileSize) {
        return $true
    } else {
        # Rename the original file with the processed tag at the end so it is no longer considered for processing
        $renamedFileName = [IO.Path]::GetFileNameWithoutExtension($Path) + $processedTag + [IO.Path]::GetExtension($Path)
        Rename-Item $Path -NewName $renamedFileName
        Write-Host "Renamed original file to: $renamedFileName"
        logMessage "Renamed original file to: $renamedFileName"
        return $false
    }
}

function cleanUpFiles {
    param (
        [Parameter(Mandatory=$true)]
        [string]$MediaFolderPath,
        [Parameter(Mandatory=$true)]
        [string[]]$MediaExtensions
    )

    $mediaFiles = Get-ChildItem $MediaFolderPath -Include $MediaExtensions -Recurse -File

    foreach ($mediaFile in $mediaFiles) {
        $tempOutputFileName = [IO.Path]::GetDirectoryName($mediaFile.FullName) + '\' + "processing-temp.mkv"
        $finalOutputFileName = [IO.Path]::GetDirectoryName($mediaFile.FullName) + '\' + [IO.Path]::GetFileNameWithoutExtension($mediaFile.FullName) + $processedTag + ".mkv"

        if (Test-Path $tempOutputFileName) {
            Remove-Item $tempOutputFileName
            Write-Host "Removed temporary file: $tempOutputFileName"
            logMessage "Removed temporary file: $tempOutputFileName"
        }

        if (Test-Path $finalOutputFileName) {
            Remove-Item $mediaFile.FullName
            Write-Host "Removed original file: $mediaFile"
            logMessage "Removed original file: $mediaFile"
        }
    }
}

function logMessage {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message
    )

    $currentDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"

    Add-Content -Path $logFile -Value ("[" + $currentDateTime + "] " + $Message)
}

function ArchiveLogFile {
    param (
        [string]$logFile,
        [string]$logArchivePath
    )

    $logFileSize = (Get-Item $logFile).Length / 1MB

    if ($logFileSize -gt $maxLogFileSize) {
        $archiveFileName = "psMC-" + (Get-Date).ToString('yyyy-MM-dd-HH-mm') + ".zip"
        $archiveFilePath = Join-Path -Path $logArchivePath -ChildPath $archiveFileName

        Compress-Archive -Path $logFile -DestinationPath $archiveFilePath

        Remove-Item -Path $logFile

        Write-Host "Archived log file: $logFile"
    }
}

# Function to generate HTML file
function GenerateHTMLReport {
    param (
        [string]$currentJob,
        [string]$mediaDrive,
        [string]$savedSpaceLogFile,
        [string]$htmlLogFile
    )

    $driveInfo = Get-PSDrive $mediaDrive
    $totalCapacity = ($driveInfo.Used + $driveInfo.Free) / 1GB
    $remainingSpace = $driveInfo.Free / 1GB
    $totalUsedSpace = ($driveInfo.Used / 1GB)
    $totalSavedSpace = [double]::Parse((Get-Content $savedSpaceLogFile))
    $percentageSaved = ($totalSavedSpace / ($totalUsedSpace + $totalSavedSpace)) * 100

    # Read the contents of the processing.txt file
    $processingInfo = Get-Content -Path $processingFile

    # Create a variable that contains all the lines of $processingInfo but adds "br" tags to the end of each line
    $processingInfo = $processingInfo | ForEach-Object {$_ + "<br>"}
    

    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv='refresh' content='60'>
    <style>
        .light-mode {
            background-color: white;
            color: black;
        }
        .dark-mode {
            background-color: black;
            color: lightgray;
        }
    </style>
</head>
<body class="dark-mode">
    <button onclick="toggleMode()">Toggle Mode</button>
    <script>
        function toggleMode() {
            var body = document.body;
            if (body.classList.contains('light-mode')) {
                body.classList.remove('light-mode');
                body.classList.add('dark-mode');
            } else {
                body.classList.remove('dark-mode');
                body.classList.add('light-mode');
            }
        }
    </script>
    <h1>Total capacity: $totalCapacity GB</h1>
    <h1>Remaining space: $remainingSpace GB</h1>
    <h1>Total used space: $totalUsedSpace GB</h1>
    <h1>Total saved space: $totalSavedSpace GB</h1>
    <h1>Percentage saved: $percentageSaved%</h1>
    <h2>Completed job: $currentJob</h2>
    <p>$processingInfo</p>
</body>
</html>
"@

$htmlContent | Out-File $htmlLogFile

Write-Host "Writing to HTML file..."
logMessage "Writing to HTML file..."
}

# Function to update saved space log file
function UpdateSavedSpace {
    param(
        [double]$originalFileSize,
        [double]$newFileSize,
        [string]$savedSpaceLogFile,
        [string]$logFile,
        [string]$mediaFile,
        [string]$renamedFileName
    )

    $savedSpace = ($originalFileSize - $newFileSize) / 1GB
    $totalSavedSpace = [double]::Parse((Get-Content $savedSpaceLogFile))
    $totalSavedSpace += $savedSpace

    $totalSavedSpace | Out-File $savedSpaceLogFile

    Write-Host "Writing savedspace to file: $savedSpaceLogFile"
    logMessage "Writing savedspace to file: $savedSpaceLogFile"

    Add-Content -Path $logFile -Value ("[" + (Get-Date).ToString() + "] Successfully converted $mediaFile to $renamedFileName. Saved space: $savedSpace GB. Total saved space: $totalSavedSpace GB.")
}

function processMedia {
    $mediaFiles = Get-Content $processingFile

    if ($mediaFiles.Count -eq 0) {
        Write-Host "No files to be processed."
        logMessage "No files to be processed."
    } else {
        foreach ($mediaFile in $mediaFiles) {
            cleanUpFiles -MediaFolderPath $mediaDir -MediaExtensions $mediaExtensions

            Write-Host "Processing file: $mediaFile"
            logMessage "Processing file: $mediaFile"

            $fileLocked = testFileLock $mediaFile

            if ($fileLocked -eq $false) {
                Write-Host "File is unlocked, processing..."
                logMessage "File is unlocked, processing..."

                $shouldProcess = getEstimatedFileSize $mediaFile (Get-Item $mediaFile).Length

                if ($shouldProcess) {
                    $originalFileSize = (Get-Item $mediaFile).Length
                    $tempOutputFileName = [IO.Path]::GetDirectoryName($mediaFile) + "\processing-temp.mkv"

                    Write-Host "Processing file...We can reduce file size"
                    logMessage "Processing file...We can reduce file size"

                    Write-Host "Processing file...Starting"
                    logMessage "Processing file...Starting"

                    Write-Host "Temp file name: $tempOutputFileName"
                    logMessage "Temp file name: $tempOutputFileName"

                    & $handbrakeCliExe -i $mediaFile -o $tempOutputFileName -e $videoEncoder -b:v $videoBitrate -E $audioEncoder -s scan -Y $videoVerticalResolution

                    $newFileSize = (Get-Item $tempOutputFileName).Length

                    if (Test-Path $tempOutputFileName) {
                        $renamedFileName = [IO.Path]::GetFileNameWithoutExtension($mediaFile) + $processedTag + ".mkv"
                        Rename-Item $tempOutputFileName -NewName $renamedFileName
                        Write-Host "Renamed temp file to: $([IO.Path]::GetFileNameWithoutExtension($mediaFile) + $processedTag + ".mkv")"
                        logMessage "Renamed temp file to: $([IO.Path]::GetFileNameWithoutExtension($mediaFile) + $processedTag + ".mkv")"
                    }

                    # Update saved space log file
                    UpdateSavedSpace -originalFileSize $originalFileSize -newFileSize $newFileSize -savedSpaceLogFile $savedSpaceLogFile -logFile $logFile -mediaFile $mediaFile -renamedFileName $renamedFileName
                    
                    # Generate HTML report
                    GenerateHTMLReport -currentJob "Converting $mediaFile" -mediaDrive $mediaDrive -savedSpaceLogFile $savedSpaceLogFile -htmlLogFile $htmlLogFile


                    if (Test-Path $renamedFileName) {
                        Remove-Item $mediaFile
                        Write-Host "Removed original file: $mediaFile"
                        logMessage "Removed original file: $mediaFile"
                    }

                    cleanUpFiles -MediaFolderPath $mediaDir -MediaExtensions $mediaExtensions
                    Write-Host "Running Cleanup Process"
                    logMessage "Running Cleanup Process"
                } else {
                    Write-Host "Skipping file...Estimated file size is larger than original file size"
                    logMessage "Skipping file...Estimated file size is larger than original file size"
                }
            } else {
                Write-Host "Skipping file...File is locked"
                logMessage "Skipping file...File is locked"
            }
        }
    }
}

# Program

while ($true) {

    # Start the processing file removal, creation, loading, and sanitization process
    removeProcessingFile $processingFile

    createFilesIfNotExist $processingFile $logFile $htmlLogFile $savedSpaceLogFile

    ensureLogArchivePathExists $logArchivePath

    ensureExecutablesExist $handbrakeCliExe $ffprobeExe

    initializeSavedSpaceLogFile $savedSpaceLogFile

    loadProcessingFile $mediaDir $processingFile

    cleanProcessingFileExt $processingFile $mediaExtensions

    removeIgnoredLines $processingFile $processingIgnore

    # Start the media processing process
    processMedia

    # Archive the log file if it is larger than the max log file size
    ArchiveLogFile -logFile $logFile -logArchivePath $logArchivePath

    # Log to console the program has completed
    Write-Host "Completed processing files going to sleep"

    # Log to console that the program is sleeping for X hours
    Write-Host "Sleeping for" ($sleepTime / 3600) "hours..."
    Write-Host "Next start time:" (Get-Date).AddSeconds($sleepTime)

    # Sleep for X hours
    Start-Sleep -Seconds $sleepTime

}

