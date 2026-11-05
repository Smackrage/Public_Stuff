<#
.SYNOPSIS
    Renames ShareX screenshots with AI-generated descriptions, rescues unsaved clipboard
    screenshots, and registers the scheduled tasks that automate both.

.DESCRIPTION
    RENAME (default)
        Scans the screenshot folder (including ShareX year-month subfolders), asks the
        Anthropic Claude API what each image shows and renames it to
        "Description_YYYY-MM-DD.ext". If the API is unsure, the file is left untouched and
        listed in ShareXRename_Uncertain_<timestamp>.log. Files already listed in an earlier
        uncertain log are skipped on normal runs so they are not re-sent every time.

    RETRY (-RetryUncertain)
        Re-submits files listed in the uncertain logs (all of them, or a single log chosen
        with -UncertainLogPath).

    CLIPBOARD WATCHER (-WatchClipboard)
        Saves every image that lands on the clipboard (Snipping Tool, Win+Shift+S and so on)
        as <ScreenshotFolder>\yyyy-MM\Snip_<date>_<time>.png, so taking another snip can no
        longer overwrite the first one. The next rename run then describes and renames them.
        Requires an STA thread (powershell.exe is STA by default).

    SCHEDULED TASKS (elevated session required)
        -SetupScheduledTask  Runs the rename on the days and hours you choose.
        -SetupWatcherTask    Starts the clipboard watcher hidden at every logon (and now).

    Logs (CMTrace format, UTF-8 without BOM) go to the root of the screenshot folder:
        ShareXRename_<timestamp>.log            - every action of a rename or retry run
        ShareXRename_Uncertain_<timestamp>.log  - files the AI could not identify (not renamed)
        ShareXClipboardWatcher_<yyyyMMdd>.log   - every clipboard image saved or ignored

    The API key is read from -ApiKey, then the ANTHROPIC_API_KEY environment variable
    (Process, User, then Machine scope). It is never written into a scheduled task.

    STANDARD ACCOUNT PLUS SEPARATE ADMIN ACCOUNT
        Run the -Setup* switches from the elevated admin session. The tasks are registered to
        run as the standard (console) user, not the admin, so they use the standard user's
        profile, clipboard and environment. -ScreenshotFolder is only baked into a task when
        you pass it explicitly; otherwise each task works out the standard user's own
        default folder when it runs. The API key must be visible to the standard user: set it
        at User scope while logged on as that user (non-elevated), or at Machine scope.

.PARAMETER ScreenshotFolder
    ShareX screenshot root folder. Default: "$env:USERPROFILE\Documents\ShareX\Screenshots".

.PARAMETER Unattended
    Suppresses the PowerShell version prompt. Used by the scheduled tasks.

.PARAMETER ApiKey
    Anthropic API key. Optional if ANTHROPIC_API_KEY is set.

.PARAMETER Model
    Claude model used to describe screenshots. Default: claude-sonnet-5-5.

.PARAMETER MaxTokens
    Maximum tokens in each API response. Default: 100.

.PARAMETER RetryUncertain
    Re-submit files from the uncertain logs.

.PARAMETER UncertainLogPath
    Retry only the files listed in this one uncertain log.

.PARAMETER WatchClipboard
    Run the clipboard watcher in the foreground (Ctrl+C to stop).

.PARAMETER PollMilliseconds
    How often the watcher checks for a clipboard change. Default: 500.

.PARAMETER SetupWatcherTask
    Register a hidden logon task that runs the clipboard watcher, and start it now.

.PARAMETER SetupScheduledTask
    Register (or update) the scheduled rename task.

.PARAMETER ScheduleDays
    Days the rename task runs. Default: every day.

.PARAMETER ScheduleStartHour
    Hour (0-23) of the first daily run. Default: 9.

.PARAMETER ScheduleTimesPerDay
    Evenly spaced runs per day from ScheduleStartHour. Default: 1, maximum 24.

.PARAMETER TaskName
    Name of the scheduled rename task. Default: ShareXRename.

.PARAMETER TaskUser
    Account the scheduled tasks run as (DOMAIN\username). Default: the interactively
    logged-on console user, falling back to the current Windows identity. Environment
    variables are deliberately not used. SYSTEM and machine accounts are refused, because
    the tasks need your own session and clipboard.

.EXAMPLE
    .\Invoke-ShareXRename.ps1
    Rename run against the default ShareX folder.

.EXAMPLE
    .\Invoke-ShareXRename.ps1 -RetryUncertain
    Re-submit every file still listed in the uncertain logs.

.EXAMPLE
    powershell.exe -STA -File .\Invoke-ShareXRename.ps1 -WatchClipboard
    Run the clipboard watcher in the foreground to try it out.

.EXAMPLE
    .\Invoke-ShareXRename.ps1 -SetupWatcherTask
    Make the clipboard watcher start hidden at every logon (elevated prompt).

.EXAMPLE
    .\Invoke-ShareXRename.ps1 -SetupScheduledTask -ScheduleDays Monday,Wednesday,Friday -ScheduleStartHour 8 -ScheduleTimesPerDay 3
    Rename at 08:00, 13:00 and 18:00 on Mon/Wed/Fri.

.NOTES
    Author:  Marty Smith
    Date:    2026-11-05
    Version: 3.2
    Requires: Windows PowerShell 5.1, internet access to api.anthropic.com.
    The script is pure ASCII. Save it as a .ps1 file before using the task options.
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Rename')]
Param (
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $ScreenshotFolder = "$env:USERPROFILE\Documents\ShareX\Screenshots",

    [Parameter()]
    [switch] $Unattended,

    [Parameter(ParameterSetName = 'Rename')]
    [Parameter(ParameterSetName = 'Retry')]
    [string] $ApiKey,

    [Parameter(ParameterSetName = 'Rename')]
    [Parameter(ParameterSetName = 'Retry')]
    [Parameter(ParameterSetName = 'Schedule')]
    [ValidateNotNullOrEmpty()]
    [string] $Model = 'claude-sonnet-5-5',

    [Parameter(ParameterSetName = 'Rename')]
    [Parameter(ParameterSetName = 'Retry')]
    [ValidateRange(50, 500)]
    [int] $MaxTokens = 100,

    [Parameter(ParameterSetName = 'Retry', Mandatory = $True)]
    [switch] $RetryUncertain,

    [Parameter(ParameterSetName = 'Retry')]
    [string] $UncertainLogPath,

    [Parameter(ParameterSetName = 'Watch', Mandatory = $True)]
    [switch] $WatchClipboard,

    [Parameter(ParameterSetName = 'Watch')]
    [ValidateRange(100, 5000)]
    [int] $PollMilliseconds = 500,

    [Parameter(ParameterSetName = 'WatcherTask', Mandatory = $True)]
    [switch] $SetupWatcherTask,

    [Parameter(ParameterSetName = 'WatcherTask')]
    [Parameter(ParameterSetName = 'Schedule')]
    [string] $TaskUser,

    [Parameter(ParameterSetName = 'Schedule', Mandatory = $True)]
    [switch] $SetupScheduledTask,

    [Parameter(ParameterSetName = 'Schedule')]
    [ValidateSet('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')]
    [string[]] $ScheduleDays = @('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'),

    [Parameter(ParameterSetName = 'Schedule')]
    [ValidateRange(0, 23)]
    [int] $ScheduleStartHour = 9,

    [Parameter(ParameterSetName = 'Schedule')]
    [ValidateRange(1, 24)]
    [int] $ScheduleTimesPerDay = 1,

    [Parameter(ParameterSetName = 'Schedule')]
    [ValidateNotNullOrEmpty()]
    [string] $TaskName = 'ShareXRename'
)

$ErrorActionPreference  = 'Stop'
$ScriptStartTime        = Get-Date
$Script:ScriptPath      = $MyInvocation.MyCommand.Path
$Script:LogFile         = $Null
$Script:UncertainLog    = $Null
$Script:Utf8NoBom       = New-Object System.Text.UTF8Encoding($False)
$Script:ImageExtensions = @('.png', '.jpg', '.jpeg', '.gif', '.webp')
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# ===========================================================================
# FUNCTIONS
# ===========================================================================

<#
.SYNOPSIS
    Writes a CMTrace-compatible entry to the log file ($Script:LogFile) and to the console.

.PARAMETER Message
    The message to log.

.PARAMETER Severity
    1 = Informational (default), 2 = Warning, 3 = Error.

.PARAMETER Component
    Label for the CMTrace component column. Defaults to the calling function name.
#>
Function Write-Log {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $Message,

        [Parameter()]
        [ValidateSet(1, 2, 3)]
        [int] $Severity = 1,

        [Parameter()]
        [string] $Component
    )

    If (-Not $Component) {
        $Caller    = (Get-PSCallStack)[1].FunctionName
        $Component = If ($Caller -And $Caller -ne '<ScriptBlock>') { $Caller } Else { 'Main' }
    }

    $Now  = Get-Date
        # CMTrace wants an explicit sign: "+600" or "-600". A negative bias already carries its minus.
    $BiasText = If ($Bias -ge 0) { "+$Bias" } Else { "$Bias" }

    $LogLine = '<![LOG[{0}]LOG]!><time="{1}{2}" date="{3}" component="{4}" context="" type="{5}" thread="{6}" file="">' -f
        $Message, $Now.ToString('HH:mm:ss.fff'), $BiasText, $Now.ToString('MM-dd-yyyy'), $Component, $Severity, $PID

        
    $Colour = Switch ($Severity) { 1 { 'Gray' } 2 { 'Yellow' } 3 { 'Red' } }
    Write-Host ('[{0}] {1}' -f $Now.ToString('HH:mm:ss'), $Message) -ForegroundColor $Colour

    If ($Script:LogFile) {
        $LogDir = Split-Path -Path $Script:LogFile -Parent
        If (-Not (Test-Path -LiteralPath $LogDir)) {
            New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
        }
        [System.IO.File]::AppendAllText($Script:LogFile, $LogLine + [Environment]::NewLine, $Script:Utf8NoBom)
    }
}

<#
.SYNOPSIS
    Checks the script is running on Windows PowerShell 5.1 and prompts if it is not.

.PARAMETER NoPrompt
    When true, a version mismatch is logged as a warning and the script carries on.
#>
Function Test-PowerShellVersion {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        [bool] $NoPrompt
    )

    $Version = $PSVersionTable.PSVersion
    If ($Version.Major -eq 5 -And $Version.Minor -eq 1) {
        Return
    }

    Write-Log -Message "This script is written for Windows PowerShell 5.1. Detected version: $Version." -Severity 2
    If ($NoPrompt) {
        Write-Log -Message 'Unattended mode: continuing on the detected version.' -Severity 2
        Return
    }

    $Answer = Read-Host 'This script targets Windows PowerShell 5.1. Continue on this version anyway? (Y/N)'
    If ($Answer -notmatch '^[Yy]') {
        Throw "Cancelled by user: unsupported PowerShell version $Version."
    }
}

<#
.SYNOPSIS
    Throws unless the session is elevated (needed to register scheduled tasks).
#>
Function Assert-Elevated {
    [CmdletBinding()]
    Param ()

    $Principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    If (-Not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Throw 'Registering a scheduled task needs an elevated (Run as Administrator) PowerShell session.'
    }
}

<#
.SYNOPSIS
    Returns the Anthropic API key from -ApiKey or the ANTHROPIC_API_KEY environment variable.

.PARAMETER ApiKey
    Explicit key. Wins over the environment.

.OUTPUTS
    [string] The API key.
#>
Function Resolve-ApiKey {
    [CmdletBinding()]
    Param (
        [Parameter()]
        [string] $ApiKey
    )

    If ($ApiKey) {
        Return $ApiKey
    }
    ForEach ($Scope in 'Process', 'User', 'Machine') {
        $Value = [System.Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY', $Scope)
        If ($Value) {
            Return $Value
        }
    }
    Throw 'No API key found. Use -ApiKey or set the ANTHROPIC_API_KEY environment variable.'
}

<#
.SYNOPSIS
    Sends an image to the Anthropic Claude API and returns a short filename-safe description.

.PARAMETER ImagePath
    Full path to the image file.

.PARAMETER ApiKey
    Anthropic API key.

.PARAMETER Model
    Claude model identifier.

.PARAMETER MaxTokens
    Maximum tokens for the response.

.OUTPUTS
    [PSCustomObject] with Description, Uncertain (bool) and RawResponse.
#>
Function Get-ImageDescription {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $ImagePath,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $ApiKey,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $Model,

        [Parameter()]
        [int] $MaxTokens = 100
    )

    $MimeMap = @{
        '.png'  = 'image/png'
        '.jpg'  = 'image/jpeg'
        '.jpeg' = 'image/jpeg'
        '.gif'  = 'image/gif'
        '.webp' = 'image/webp'
    }
    $Extension = [System.IO.Path]::GetExtension($ImagePath).ToLower()
    If (-Not $MimeMap.ContainsKey($Extension)) {
        Return [PSCustomObject]@{
            Description = $Null
            Uncertain   = $True
            RawResponse = "Unsupported file extension: $Extension"
        }
    }

    $Base64Image  = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($ImagePath))
    $SystemPrompt = (
        'You are a screenshot description assistant. Produce a concise, filename-safe description of what a screenshot shows.',
        'Rules:',
        '- Reply with ONLY the description: no preamble and no trailing punctuation.',
        '- Use Title Case with spaces between words, for example: Visual Studio Code Python Error',
        '- Maximum 8 words.',
        '- If you cannot tell what the screenshot depicts with reasonable confidence, reply with exactly: UNCERTAIN'
    ) -join ' '

    $Body = @{
        model      = $Model
        max_tokens = $MaxTokens
        system     = $SystemPrompt
        messages   = @(
            @{
                role    = 'user'
                content = @(
                    @{
                        type   = 'image'
                        source = @{
                            type       = 'base64'
                            media_type = $MimeMap[$Extension]
                            data       = $Base64Image
                        }
                    },
                    @{
                        type = 'text'
                        text = 'Describe this screenshot.'
                    }
                )
            }
        )
    } | ConvertTo-Json -Depth 10

    $BodyBytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $Headers   = @{
        'x-api-key'         = $ApiKey
        'anthropic-version' = '2023-06-01'
    }
    $RequestParams = @{
        Uri         = 'https://api.anthropic.com/v1/messages'
        Method      = 'Post'
        Headers     = $Headers
        Body        = $BodyBytes
        ContentType = 'application/json; charset=utf-8'
        TimeoutSec  = 120
    }

    # Retry transient failures (rate limit or server side) up to three attempts
    $Attempt = 0
    While ($True) {
        $Attempt++
        Try {
            $Response = Invoke-RestMethod @RequestParams
            Break
        }
        Catch {
            $Status = 0
            If ($_.Exception.Response) {
                $Status = [int]$_.Exception.Response.StatusCode
            }
            If ($Attempt -lt 3 -And ($Status -eq 429 -Or $Status -ge 500)) {
                Write-Log -Message "API returned HTTP $Status (attempt $Attempt of 3). Waiting before retry." -Severity 2
                Start-Sleep -Seconds (5 * $Attempt)
            }
            Else {
                Throw
            }
        }
    }

    $TextBlock = $Response.content | Where-Object { $_.type -eq 'text' } | Select-Object -First 1
    $RawText   = ''
    If ($TextBlock) {
        $RawText = ([string]$TextBlock.text).Trim()
    }

    If ([string]::IsNullOrWhiteSpace($RawText) -Or $RawText -match '^\s*UNCERTAIN\b') {
        Return [PSCustomObject]@{ Description = $Null; Uncertain = $True; RawResponse = $RawText }
    }

    # Make it safe for a Windows filename
    $Clean = $RawText -replace '[<>:"/\\|?*\[\]\r\n]', ' '
    $Clean = ($Clean -replace '\s+', ' ').Trim()
    If ($Clean.Length -gt 100) {
        $Clean = $Clean.Substring(0, 100).Trim()
    }
    $Clean = $Clean.TrimEnd('.', ' ')

    If ([string]::IsNullOrWhiteSpace($Clean)) {
        Return [PSCustomObject]@{ Description = $Null; Uncertain = $True; RawResponse = $RawText }
    }

    Return [PSCustomObject]@{ Description = $Clean; Uncertain = $False; RawResponse = $RawText }
}

<#
.SYNOPSIS
    Returns the capture date: from the file name when it contains YYYY-MM-DD, else CreationTime.

.PARAMETER FilePath
    Full path to the image file.

.OUTPUTS
    [datetime]
#>
Function Get-CaptureDateFromFile {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $FilePath
    )

    $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    If ($BaseName -match '(\d{4})-(\d{2})-(\d{2})') {
        Try {
            Return [datetime]::new([int]$Matches[1], [int]$Matches[2], [int]$Matches[3])
        }
        Catch {
            Write-Log -Message "File name date in '$BaseName' is not a valid date. Using CreationTime." -Severity 2
        }
    }
    Return (Get-Item -LiteralPath $FilePath).CreationTime
}

<#
.SYNOPSIS
    Returns a path that does not exist yet, appending _N when the name is taken.

.PARAMETER Directory
    Target directory.

.PARAMETER BaseName
    Wanted file name without extension.

.PARAMETER Extension
    Extension including the leading dot.

.OUTPUTS
    [string]
#>
Function ConvertTo-UniqueFilePath {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        [string] $Directory,

        [Parameter(Mandatory = $True)]
        [string] $BaseName,

        [Parameter(Mandatory = $True)]
        [string] $Extension
    )

    $Candidate = Join-Path -Path $Directory -ChildPath "$BaseName$Extension"
    $Counter   = 1
    While (Test-Path -LiteralPath $Candidate) {
        $Candidate = Join-Path -Path $Directory -ChildPath "${BaseName}_$Counter$Extension"
        $Counter++
    }
    Return $Candidate
}

<#
.SYNOPSIS
    Adds a file path to this run's uncertain log, creating the log (with header) on first use.

.PARAMETER FilePath
    Full path of the file that was not renamed.
#>
Function Add-UncertainEntry {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $FilePath
    )

    If (-Not (Test-Path -LiteralPath $Script:UncertainLog)) {
        $Header = (
            'ShareX Rename - Uncertain Files',
            "Run    : $($ScriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))",
            "Folder : $ScreenshotFolder",
            '---',
            'Files below could not be identified by the AI and have NOT been renamed.',
            'Use -RetryUncertain to submit them again.'
        ) -join [Environment]::NewLine
        [System.IO.File]::WriteAllText($Script:UncertainLog, $Header + [Environment]::NewLine, $Script:Utf8NoBom)
    }
    [System.IO.File]::AppendAllText($Script:UncertainLog, $FilePath + [Environment]::NewLine, $Script:Utf8NoBom)
}

<#
.SYNOPSIS
    Collects every file path listed in any earlier uncertain log in the screenshot folder.

.PARAMETER Folder
    Screenshot root folder.

.OUTPUTS
    [System.Collections.Generic.HashSet[string]] (case-insensitive)
#>
Function Get-KnownUncertainPath {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        [string] $Folder
    )

    $Known = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $Logs  = Get-ChildItem -LiteralPath $Folder -Filter 'ShareXRename_Uncertain_*.log' -File
    ForEach ($Log in $Logs) {
        ForEach ($Line in (Get-Content -LiteralPath $Log.FullName)) {
            If ($Line -match '^([A-Za-z]:\\|\\\\)') {
                [void]$Known.Add($Line.Trim())
            }
        }
    }
    Return , $Known
}

<#
.SYNOPSIS
    Describes each image with the Claude API and renames it, or logs it as uncertain.

.PARAMETER ImageFiles
    Files to process.

.PARAMETER ApiKey
    Anthropic API key.

.PARAMETER Model
    Claude model identifier.

.PARAMETER MaxTokens
    Maximum tokens per API response.

.OUTPUTS
    [PSCustomObject] with Renamed, Skipped, Uncertain and Failed counts.
#>
Function Invoke-RenameImages {
    [CmdletBinding(SupportsShouldProcess)]
    Param (
        [Parameter(Mandatory = $True)]
        [System.IO.FileInfo[]] $ImageFiles,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $ApiKey,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $Model,

        [Parameter(Mandatory = $True)]
        [int] $MaxTokens
    )

    $Total     = $ImageFiles.Count
    $Current   = 0
    $Renamed   = 0
    $Skipped   = 0
    $Uncertain = 0
    $Failed    = 0
    $StartTime = Get-Date

    Write-Log -Message "Processing $Total image(s) with model $Model."

    ForEach ($ImageFile in $ImageFiles) {
        $Current++
        $Percent = [Math]::Round(($Current / $Total) * 100, 1)
        $Elapsed = (Get-Date) - $StartTime

        Write-Progress -Activity 'Renaming ShareX screenshots' `
            -Status "[$Current/$Total] $($ImageFile.Name)" -PercentComplete $Percent
        Write-Log -Message "[$Current/$Total] ($Percent%) $($ImageFile.FullName) | elapsed $($Elapsed.ToString('hh\:mm\:ss'))"

        Try {
            If ($ImageFile.BaseName -match '^.+_\d{4}-\d{2}-\d{2}(_\d+)?$') {
                Write-Log -Message "SKIP - already in the target name format: $($ImageFile.Name)"
                $Skipped++
                Continue
            }
            If (-Not (Test-Path -LiteralPath $ImageFile.FullName)) {
                Write-Log -Message "SKIP - file no longer exists: $($ImageFile.FullName)" -Severity 2
                $Skipped++
                Continue
            }

            $DateSuffix = (Get-CaptureDateFromFile -FilePath $ImageFile.FullName).ToString('yyyy-MM-dd')
            $AiResult   = Get-ImageDescription -ImagePath $ImageFile.FullName -ApiKey $ApiKey `
                -Model $Model -MaxTokens $MaxTokens

            If ($AiResult.Uncertain) {
                Write-Log -Message "UNCERTAIN - not renamed. API said: '$($AiResult.RawResponse)'" -Severity 2
                Add-UncertainEntry -FilePath $ImageFile.FullName
                $Uncertain++
                Continue
            }

            $NewFilePath = ConvertTo-UniqueFilePath -Directory $ImageFile.DirectoryName `
                -BaseName "$($AiResult.Description)_$DateSuffix" -Extension $ImageFile.Extension
            $NewName     = Split-Path -Path $NewFilePath -Leaf

            Write-Log -Message "RENAME: '$($ImageFile.Name)' -> '$NewName'"
            If ($PSCmdlet.ShouldProcess($ImageFile.FullName, "Rename to $NewName")) {
                Rename-Item -LiteralPath $ImageFile.FullName -NewName $NewName
                $Renamed++
            }
        }
        Catch {
            Write-Log -Message "ERROR on '$($ImageFile.Name)' (line $($_.InvocationInfo.ScriptLineNumber)): $($_.Exception.Message)" -Severity 3
            $Failed++
        }
    }

    Write-Progress -Activity 'Renaming ShareX screenshots' -Completed

    Return [PSCustomObject]@{
        Renamed   = $Renamed
        Skipped   = $Skipped
        Uncertain = $Uncertain
        Failed    = $Failed
    }
}

<#
.SYNOPSIS
    Reads the image currently on the clipboard and returns it as PNG bytes.

.DESCRIPTION
    Retries for about 1.5 seconds if another program has the clipboard locked.

.OUTPUTS
    [byte[]] PNG data, or $null when the clipboard holds no image.
#>
Function Get-ClipboardImageBytes {
    [CmdletBinding()]
    Param ()

    $MaxAttempts = 10
    For ($Attempt = 1; $Attempt -le $MaxAttempts; $Attempt++) {
        Try {
            If (-Not [System.Windows.Forms.Clipboard]::ContainsImage()) {
                Return $Null
            }
            $Image = [System.Windows.Forms.Clipboard]::GetImage()
            If ($Null -eq $Image) {
                Return $Null
            }
            $Stream = New-Object System.IO.MemoryStream
            Try {
                $Image.Save($Stream, [System.Drawing.Imaging.ImageFormat]::Png)
                Return , $Stream.ToArray()
            }
            Finally {
                $Stream.Dispose()
                $Image.Dispose()
            }
        }
        Catch [System.Runtime.InteropServices.ExternalException] {
            Write-Log -Message "Clipboard busy (attempt $Attempt of $MaxAttempts)." -Severity 2
            Start-Sleep -Milliseconds 150
        }
    }
    Throw "Clipboard stayed locked for $MaxAttempts attempts."
}

<#
.SYNOPSIS
    Runs until stopped, saving every new clipboard image to the screenshot folder as a PNG.

.PARAMETER ScreenshotFolder
    Root folder. Images go to <root>\yyyy-MM\Snip_yyyy-MM-dd_HH.mm.ss.png.

.PARAMETER PollMilliseconds
    How often the clipboard change counter is checked.
#>
Function Start-ClipboardWatcher {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $ScreenshotFolder,

        [Parameter()]
        [ValidateRange(100, 5000)]
        [int] $PollMilliseconds = 500
    )

    If ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
        Throw 'The clipboard watcher needs an STA thread. Start it with: powershell.exe -STA -File <script> -WatchClipboard'
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    If (-Not ('Native.Clipboard' -as [type])) {
        Add-Type -Namespace Native -Name Clipboard -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetClipboardSequenceNumber();'
    }

    If (-Not (Test-Path -LiteralPath $ScreenshotFolder -PathType Container)) {
        New-Item -ItemType Directory -Path $ScreenshotFolder -Force | Out-Null
    }

    $Sha256       = [System.Security.Cryptography.SHA256]::Create()
    $SeenHashes   = New-Object 'System.Collections.Generic.HashSet[string]'
    $LastSequence = [Native.Clipboard]::GetClipboardSequenceNumber()
    $SavedCount   = 0

    $Script:LogFile = Join-Path -Path $ScreenshotFolder -ChildPath "ShareXClipboardWatcher_$(Get-Date -Format 'yyyyMMdd').log"
    Write-Log -Message '===== Clipboard watcher started ====='
    Write-Log -Message "User: $env:USERNAME | Computer: $env:COMPUTERNAME | PowerShell: $($PSVersionTable.PSVersion)"
    Write-Log -Message "Saving clipboard images under: $ScreenshotFolder (poll every $PollMilliseconds ms)"

    Try {
        While ($True) {
            Start-Sleep -Milliseconds $PollMilliseconds

            $Sequence = [Native.Clipboard]::GetClipboardSequenceNumber()
            If ($Sequence -eq $LastSequence) {
                Continue
            }
            $LastSequence   = $Sequence
            $Script:LogFile = Join-Path -Path $ScreenshotFolder -ChildPath "ShareXClipboardWatcher_$(Get-Date -Format 'yyyyMMdd').log"

            Try {
                $Bytes = Get-ClipboardImageBytes
                If ($Null -eq $Bytes) {
                    Continue
                }

                $Hash = [System.BitConverter]::ToString($Sha256.ComputeHash($Bytes)).Replace('-', '')
                If (-Not $SeenHashes.Add($Hash)) {
                    Write-Log -Message "Duplicate image ignored (hash $($Hash.Substring(0, 12)))."
                    Continue
                }

                $Now       = Get-Date
                $TargetDir = Join-Path -Path $ScreenshotFolder -ChildPath $Now.ToString('yyyy-MM')
                If (-Not (Test-Path -LiteralPath $TargetDir)) {
                    New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
                }

                $TargetPath = ConvertTo-UniqueFilePath -Directory $TargetDir `
                    -BaseName "Snip_$($Now.ToString('yyyy-MM-dd_HH.mm.ss'))" -Extension '.png'
                [System.IO.File]::WriteAllBytes($TargetPath, $Bytes)
                $SavedCount++
                Write-Log -Message "SAVED clipboard image ($([Math]::Round($Bytes.Length / 1KB, 1)) KB, #$SavedCount this session): $TargetPath"
            }
            Catch {
                Write-Log -Message "Could not save clipboard image (line $($_.InvocationInfo.ScriptLineNumber)): $($_.Exception.Message)" -Severity 3
            }
        }
    }
    Finally {
        $Sha256.Dispose()
        Write-Log -Message "===== Clipboard watcher stopped. Images saved this session: $SavedCount ====="
    }
}

<#
.SYNOPSIS
    Works out which account the scheduled tasks should run as.

.DESCRIPTION
    Order of preference: -TaskUser, the interactively logged-on console user, then the
    current Windows identity. Environment variables are not used because they can be wrong
    in elevated sessions. SYSTEM, service and machine accounts are refused: the tasks need
    your own session (and clipboard).

.PARAMETER TaskUser
    Optional DOMAIN\username supplied by the caller.

.OUTPUTS
    [string] The account name.
#>
Function Resolve-TaskUser {
    [CmdletBinding()]
    Param (
        [Parameter()]
        [string] $TaskUser
    )

    $Candidate = $TaskUser
    $Source    = 'the -TaskUser parameter'

    If (-Not $Candidate) {
        $Candidate = [string]((Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).UserName)
        $Source    = 'the logged-on console user'
    }
    If (-Not $Candidate) {
        $Candidate = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $Source    = 'the current Windows identity'
    }
    If ($Candidate -match '(^|\\)(SYSTEM|LOCAL SERVICE|NETWORK SERVICE)$' -Or $Candidate -match '\$$') {
        Throw "Resolved task account '$Candidate' (from $Source) is a service or machine account. Pass -TaskUser DOMAIN\username with the standard account the tasks should run as."
    }

    Write-Log -Message "Scheduled tasks will run as: $Candidate (taken from $Source)."
    Return $Candidate
}

<#
.SYNOPSIS
    Finds where the task account will get ANTHROPIC_API_KEY from, without using the
    elevated (admin) session's own environment.

.DESCRIPTION
    Checks the Machine environment first, then the target user's User environment through
    their registry hive (HKEY_USERS\<SID>\Environment). The hive is only loaded while that
    user is logged on, which is the case for an interactive task anyway.

.PARAMETER UserId
    DOMAIN\username the task will run as.

.OUTPUTS
    [string] A description of where the key was found, or $Null if it was not found.
#>
Function Get-ApiKeySource {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $UserId
    )

    If ([System.Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY', 'Machine')) {
        Return 'the Machine environment'
    }

    Try {
        $Sid  = (New-Object System.Security.Principal.NTAccount($UserId)).Translate([System.Security.Principal.SecurityIdentifier]).Value
        $Item = Get-ItemProperty -LiteralPath "Registry::HKEY_USERS\$Sid\Environment" -Name 'ANTHROPIC_API_KEY' -ErrorAction Stop
        If ($Item.ANTHROPIC_API_KEY) {
            Return "the User environment of $UserId"
        }
    }
    Catch {
        Write-Log -Message "Could not find the key in the User environment of ${UserId}: $($_.Exception.Message)" -Severity 2
    }

    Return $Null
}

<#
.SYNOPSIS
    Registers (or overwrites) a scheduled task under \ShareXRename\ for the given user and checks it landed.

.PARAMETER TaskName
    Task name.

.PARAMETER Arguments
    Argument string passed to powershell.exe.

.PARAMETER Triggers
    One or more scheduled task triggers.

.PARAMETER Settings
    A scheduled task settings object.

.PARAMETER Description
    Task description.

.PARAMETER UserId
    DOMAIN\username the task runs as (see Resolve-TaskUser).
#>
Function Register-ShareXTask {
    [CmdletBinding(SupportsShouldProcess)]
    Param (
        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $TaskName,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $Arguments,

        [Parameter(Mandatory = $True)]
        [object[]] $Triggers,

        [Parameter(Mandatory = $True)]
        [object] $Settings,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $Description,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $UserId
    )

    Assert-Elevated
    If (-Not $Script:ScriptPath -Or -Not (Test-Path -LiteralPath $Script:ScriptPath -PathType Leaf)) {
        Throw 'Cannot find this script on disk. Save it as a .ps1 file and run it from there.'
    }

    $Action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $Arguments -ErrorAction Stop
    $Principal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive -RunLevel Limited -ErrorAction Stop

    $RegisterParams = @{
        TaskName    = $TaskName
        TaskPath    = '\ShareXRename\'
        Action      = $Action
        Trigger     = $Triggers
        Settings    = $Settings
        Principal   = $Principal
        Description = $Description
        Force       = $True
        ErrorAction = 'Stop'
    }

    If ($PSCmdlet.ShouldProcess("\ShareXRename\$TaskName", 'Register scheduled task')) {
        Register-ScheduledTask @RegisterParams | Out-Null

        # Confirm it really exists: the ScheduledTasks cmdlets can report an error without stopping the script
        $Registered = Get-ScheduledTask -TaskPath '\ShareXRename\' -TaskName $TaskName -ErrorAction Stop
        Write-Log -Message "Scheduled task registered and verified: \ShareXRename\$TaskName (runs as $UserId, state $($Registered.State))"
    }
}

# ===========================================================================
# MAIN
# ===========================================================================

Try {
    Test-PowerShellVersion -NoPrompt $Unattended.IsPresent
    $Mode = $PSCmdlet.ParameterSetName

    # Only bake a folder or model into a task if it was supplied explicitly. Otherwise the task
    # (running as the standard user) works out its own default. The elevated admin session's
    # profile folder must never leak into the task.
    $FolderArgument = ''
    If ($PSBoundParameters.ContainsKey('ScreenshotFolder')) {
        $FolderArgument = " -ScreenshotFolder `"$ScreenshotFolder`""
    }
    $ModelArgument = ''
    If ($PSBoundParameters.ContainsKey('Model')) {
        $ModelArgument = " -Model `"$Model`""
    }

    # ---- Mode: register the scheduled rename task ----------------------------
    If ($Mode -eq 'Schedule') {
        $RunAs     = Resolve-TaskUser -TaskUser $TaskUser
        $KeySource = Get-ApiKeySource -UserId $RunAs
        If (-Not $KeySource) {
            Throw "ANTHROPIC_API_KEY is not visible to $RunAs. Log on as that account and, in a NON-elevated window, run: [Environment]::SetEnvironmentVariable('ANTHROPIC_API_KEY', '<your key>', 'User')  (or set it at Machine scope from this session), then run this again."
        }
        Write-Log -Message "The task will use the API key from $KeySource."

        # Spread the runs evenly from the start hour towards midnight
        $IntervalHours = 0
        If ($ScheduleTimesPerDay -gt 1) {
            $IntervalHours = [Math]::Max(1, [Math]::Floor((24 - $ScheduleStartHour) / $ScheduleTimesPerDay))
        }
        $RunHours = New-Object 'System.Collections.Generic.List[int]'
        For ($Index = 0; $Index -lt $ScheduleTimesPerDay; $Index++) {
            $Hour = [int]($ScheduleStartHour + ($Index * $IntervalHours))
            If ($Hour -le 23) {
                $RunHours.Add($Hour)
            }
            Else {
                Write-Log -Message "Run $($Index + 1) would fall at hour $Hour (past 23:00) and was skipped." -Severity 2
            }
        }
        If ($RunHours.Count -eq 0) {
            Throw 'No valid run times. Adjust -ScheduleStartHour or -ScheduleTimesPerDay.'
        }

        $Triggers = @(
            ForEach ($Hour in $RunHours) {
                New-ScheduledTaskTrigger -Weekly -DaysOfWeek $ScheduleDays -At (Get-Date -Hour $Hour -Minute 0 -Second 0) -ErrorAction Stop
            }
        )
        $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -StartWhenAvailable -RunOnlyIfNetworkAvailable -MultipleInstances IgnoreNew `
            -ExecutionTimeLimit (New-TimeSpan -Hours 2) -ErrorAction Stop
        $Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$Script:ScriptPath`" -Unattended" +
            $FolderArgument + $ModelArgument

        Register-ShareXTask -TaskName $TaskName -Arguments $Arguments -Triggers $Triggers -Settings $Settings `
            -UserId $RunAs -Description 'Renames ShareX screenshots using Claude. Managed by Invoke-ShareXRename.ps1.'

        Write-Log -Message "Runs on: $($ScheduleDays -join ', ')"
        Write-Log -Message "Run times: $(($RunHours | ForEach-Object { '{0:D2}:00' -f $_ }) -join ', ')"
        Write-Log -Message "Remove with: Unregister-ScheduledTask -TaskName '$TaskName' -TaskPath '\ShareXRename\' -Confirm:`$false"
        Return
    }

    # ---- Mode: register the clipboard watcher logon task ---------------------
    If ($Mode -eq 'WatcherTask') {
        $WatcherTaskName = 'ShareXClipboardWatcher'
        $RunAs     = Resolve-TaskUser -TaskUser $TaskUser
        $Triggers  = @(New-ScheduledTaskTrigger -AtLogOn -User $RunAs -ErrorAction Stop)
        $Settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero) `
            -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -ErrorAction Stop
        $Arguments = "-NoProfile -STA -WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass " +
            "-File `"$Script:ScriptPath`" -WatchClipboard -Unattended" +
            $FolderArgument

        Register-ShareXTask -TaskName $WatcherTaskName -Arguments $Arguments -Triggers $Triggers -Settings $Settings `
            -UserId $RunAs -Description 'Saves every clipboard image to the ShareX folder so no snip is lost. Managed by Invoke-ShareXRename.ps1.'

        If (-Not $WhatIfPreference) {
            Start-ScheduledTask -TaskPath '\ShareXRename\' -TaskName $WatcherTaskName -ErrorAction Stop
            Write-Log -Message 'Clipboard watcher task started. It will also start at every logon.'
        }
        Return
    }

    # ---- Mode: clipboard watcher ---------------------------------------------
    If ($Mode -eq 'Watch') {
        Start-ClipboardWatcher -ScreenshotFolder $ScreenshotFolder -PollMilliseconds $PollMilliseconds
        Return
    }

    # ---- Modes: Rename and Retry ---------------------------------------------
    $ApiKey = Resolve-ApiKey -ApiKey $ApiKey
    If (-Not (Test-Path -LiteralPath $ScreenshotFolder -PathType Container)) {
        Throw "Screenshot folder not found: $ScreenshotFolder"
    }

    $TimeStamp           = $ScriptStartTime.ToString('yyyyMMdd_HHmmss')
    $Script:LogFile      = Join-Path -Path $ScreenshotFolder -ChildPath "ShareXRename_$TimeStamp.log"
    $Script:UncertainLog = Join-Path -Path $ScreenshotFolder -ChildPath "ShareXRename_Uncertain_$TimeStamp.log"

    Write-Log -Message '===== Invoke-ShareXRename started ====='
    Write-Log -Message "Mode: $Mode | User: $env:USERNAME | Computer: $env:COMPUTERNAME | PowerShell: $($PSVersionTable.PSVersion)"
    Write-Log -Message "ScreenshotFolder: $ScreenshotFolder | Model: $Model | MaxTokens: $MaxTokens | WhatIf: $WhatIfPreference"

    $KnownUncertain = Get-KnownUncertainPath -Folder $ScreenshotFolder

    If ($Mode -eq 'Retry') {
        If ($UncertainLogPath) {
            If (-Not (Test-Path -LiteralPath $UncertainLogPath -PathType Leaf)) {
                Throw "Uncertain log not found: $UncertainLogPath"
            }
            Write-Log -Message "Retrying files listed in: $UncertainLogPath"
            $Candidates = @(
                Get-Content -LiteralPath $UncertainLogPath |
                    Where-Object { $_ -match '^([A-Za-z]:\\|\\\\)' } |
                    ForEach-Object { $_.Trim() }
            )
        }
        Else {
            Write-Log -Message 'Retrying every file listed in the uncertain logs.'
            $Candidates = @($KnownUncertain)
        }

        $Existing = @($Candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf })
        If ($Existing.Count -lt $Candidates.Count) {
            Write-Log -Message "$($Candidates.Count - $Existing.Count) listed file(s) no longer exist (renamed or moved) and were dropped." -Severity 2
        }
        $ImageFiles = @(
            $Existing |
                ForEach-Object { Get-Item -LiteralPath $_ } |
                Where-Object { $Script:ImageExtensions -contains $_.Extension.ToLower() }
        )
    }
    Else {
        $AllImages  = @(
            Get-ChildItem -LiteralPath $ScreenshotFolder -Recurse -File |
                Where-Object { $Script:ImageExtensions -contains $_.Extension.ToLower() }
        )
        $ImageFiles = @($AllImages | Where-Object { -Not $KnownUncertain.Contains($_.FullName) })
        If ($AllImages.Count -gt $ImageFiles.Count) {
            Write-Log -Message "Skipping $($AllImages.Count - $ImageFiles.Count) file(s) already listed as uncertain. Use -RetryUncertain to submit them again."
        }
    }

    If ($ImageFiles.Count -eq 0) {
        Write-Log -Message 'No images to process.'
        Return
    }

    $Results = Invoke-RenameImages -ImageFiles $ImageFiles -ApiKey $ApiKey -Model $Model -MaxTokens $MaxTokens

    $TotalElapsed = (Get-Date) - $ScriptStartTime
    Write-Log -Message "Summary: processed $($ImageFiles.Count) | renamed $($Results.Renamed) | skipped $($Results.Skipped) | uncertain $($Results.Uncertain) | errors $($Results.Failed)"
    If ($Results.Uncertain -gt 0) {
        Write-Log -Message "Uncertain files listed in: $Script:UncertainLog" -Severity 2
    }
    Write-Log -Message "===== Invoke-ShareXRename finished. Elapsed: $($TotalElapsed.ToString('hh\:mm\:ss')) ====="
}
Catch {
    Write-Log -Message "FATAL (line $($_.InvocationInfo.ScriptLineNumber)): $($_.Exception.Message)" -Severity 3
    Throw
}
Finally {
    Write-Progress -Activity 'Renaming ShareX screenshots' -Completed
}
