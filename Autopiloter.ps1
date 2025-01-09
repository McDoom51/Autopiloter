

[CmdletBinding(DefaultParameterSetName = 'Default')]
param (
    [Parameter(Mandatory = $true, ParameterSetName = 'Online')] 
    [switch] $Online = $false,

    [Parameter(Mandatory = $false, ParameterSetName = 'Online')]
    [string] $tenant = "",

    [Parameter(Mandatory = $false, ParameterSetName = 'Online')]
    [string] $clientId = "",

    [Parameter(Mandatory = $false, ParameterSetName = 'Online')]
    [string] $clientSecret = "",

    [Parameter]
    [switch] $ownISO = $false,

    [Parameter(Mandatory = $false, ParameterSetName = 'ownISO')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [string] $isoPath = "",

    [Parameter(Mandatory = $false, ParameterSetName = 'ownISO', HelpMessage = "Used with the download link from Microsoft 365 Admin Center")]
    [string] $enterpriseUri = "",

    [Parameter(Mandatory = $true)]
    [ValidateSet(Own, Example)]
    [string] $unattendFile = "Example"

)

# Collection of Functions

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Position = 2, Mandatory = $true)]
        [string] $Message,

        [Parameter(Position = 3, Mandatory = $false)]
        [array] $Arguments,

        [Parameter(Position = 4, Mandatory = $false)]
        [object] $Body = $null,

        [Parameter(Position = 5, Mandatory = $false)]
        [System.Management.Automation.ErrorRecord] $ExceptionInfo = $null
    )

    DynamicParam {
        New-LoggingDynamicParam -Level -Mandatory $false -Name "Level"
        $PSBoundParameters["Level"] = "INFO"
    }

    End {
        $levelNumber = Get-LevelNumber -Level $PSBoundParameters.Level
        $invocationInfo = (Get-PSCallStack)[$Script:Logging.CallerScope]

        # Split-Path throws an exception if called with a -Path that is null or empty.
        [string] $fileName = [string]::Empty
        if (-not [string]::IsNullOrEmpty($invocationInfo.ScriptName)) {
            $fileName = Split-Path -Path $invocationInfo.ScriptName -Leaf
        }

        $logMessage = [hashtable] @{
            timestamp    = [datetime]::now
            timestamputc = [datetime]::UtcNow
            level        = Get-LevelName -Level $levelNumber
            levelno      = $levelNumber
            lineno       = $invocationInfo.ScriptLineNumber
            pathname     = $invocationInfo.ScriptName
            filename     = $fileName
            caller       = $invocationInfo.Command
            message      = [string] $Message
            rawmessage   = [string] $Message
            body         = $Body
            execinfo     = $ExceptionInfo
            pid          = $PID
        }

        if ($PSBoundParameters.ContainsKey('Arguments')) {
            $logMessage["message"] = [string] $Message -f $Arguments
            $logMessage["args"] = $Arguments
        }

        #This variable is initiated via Start-LoggingManager
        $Script:LoggingEventQueue.Add($logMessage)
    }
}

function Get-Modules {
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [ValidateScript({ Test-Path -Path $_ -PathType Container})]
        [string] $path
    )

    $modules = @(
        [PSCustomObject]@{
            Name = 'Microsoft.Graph.Authentication';
            MinVersion = '2.0.0'
        },
        [PSCustomObject]@{
            Name = 'Microsoft.Graph.Groups';
            MinVersion = '5.0.0'
        },
        [PSCustomObject]@{
            Name = 'Microsoft.Graph.Identity.DirectoryManagement';
            MinVersion = '2.2.0'
        }
    )

    foreach ($module in $modules) {
        try {
            Save-Module -Name $module.Name -Path "$path\modules" -Repository PSGallery -IncludeDependencies -MinimumVersion $module.MinVersion -Force -AllowClobber
            Write-Log -Level 'INFO' -Message "Module $($module.Name) downloaded."
        }
        catch {
            Write-Log -Level 'ERROR' -Message "Failed to download module $($module.Name)." -ExceptionInfo $_
        }
    }
}

function Get-EnterpriseISO {
    param (
        [Parameter]
        [ValidateNotNullOrEmpty()]
        [string] $uri
    )

    Invoke-DownloadISO -Uri $uri

}

function Get-FidoISO {
    param (
        [Parameter(Mandatory = $true)]
        [array] $versions
    )

    $isofilename = "$path\microsoftwindows.iso"
    Write-Host "Selecting OS"
    Write-Host "Finding latest supported versions"

    $options = @()

    foreach ($foundversion in $versions) {
        $options += $foundversion.Name
    }

    $object = foreach ($option in $options) {
        New-Object psobject -Property @{'Pick your option' = $option}
    }

    $osinput = $object | Out-GridView -Title "Windows Selection" -PassThru

    $selectedname = $osinput.'Pick your option'

    $selectedos = $versions | Where-Object Name -eq "$selectedname"

    Write-Host "Selected OS: $($selectedos.Name)"

    # Prompt for language
    Write-OutPut "Select the language for the ISO"
    $url = "https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/available-language-packs-for-windows?view=windows-11"
    $content = (Invoke-WebRequest -Uri $url -UseBasicParsing).content

    # Use regex to extract the first table from the HTML content
    $tableRegex = '<table.*?>(.*?)</table>'
    $tableMatches = [regex]::Matches($content, $tableRegex, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $firstTable = $tableMatches[0].Value
    $rowRegex = '<tr.*?>\s*<td.*?>.*?</td>\s*<td.*?>(.*?)</td>'
    $rowMatches = [regex]::Matches($firstTable, $rowRegex, [System.Text.RegularExpressions.RegexOptions]::Singleline)

    $rowgroups = $rowMatches.Groups
    $languages = @()
    foreach ($row in $rowgroups) {
        $secondColumnContent = [regex]::Match($row.Value, '<td.*?>(.*?)</td>\s*<td.*?>(.*?)</td>').Groups[2].Value
        if ($secondColumnContent) {
            if ($secondColumnContent -notlike "*<p>*") {
        $languages += $secondColumnContent
            }
        }
    }

    $selectedlanguage = $languages | Out-GridView -Title "Select a Language" -PassThru

    Write-Host "Selected Language: $selectedlanguage"

    # Convert to text
    switch ($selectedlanguage) {
        "ar-SA" { $Locale = "Arabic" }
        "pt-BR" { $Locale = "Brazilian Portuguese" }
        "bg-BG" { $Locale = "Bulgarian" }
        "zh-CN" { $Locale = "Chinese (Simplified)" }
        "zh-TW" { $Locale = "Chinese (Traditional)" }
        "hr-HR" { $Locale = "Croatian" }
        "cs-CZ" { $Locale = "Czech" }
        "da-DK" { $Locale = "Danish" }
        "nl-NL" { $Locale = "Dutch" }
        "en-US" { $Locale = "English" }
        "en-GB" { $Locale = "English International" }
        "et-EE" { $Locale = "Estonian" }
        "fi-FI" { $Locale = "Finnish" }
        "fr-FR" { $Locale = "French" }
        "fr-CA" { $Locale = "French Canadian" }
        "de-DE" { $Locale = "German" }
        "el-GR" { $Locale = "Greek" }
        "he-IL" { $Locale = "Hebrew" }
        "hu-HU" { $Locale = "Hungarian" }
        "it-IT" { $Locale = "Italian" }
        "ja-JP" { $Locale = "Japanese" }
        "ko-KR" { $Locale = "Korean" }
        "lv-LV" { $Locale = "Latvian" }
        "lt-LT" { $Locale = "Lithuanian" }
        "nb-NO" { $Locale = "Norwegian" }
        "pl-PL" { $Locale = "Polish" }
        "pt-PT" { $Locale = "Portuguese" }
        "ro-RO" { $Locale = "Romanian" }
        "ru-RU" { $Locale = "Russian" }
        "sr-Latn-RS" { $Locale = "Serbian Latin" }
        "sk-SK" { $Locale = "Slovak" }
        "sl-SI" { $Locale = "Slovenian" }
        "es-ES" { $Locale = "Spanish" }
        "es-MX" { $Locale = "Spanish (Mexico)" }
        "sv-SE" { $Locale = "Swedish" }
        "th-TH" { $Locale = "Thai" }
        "tr-TR" { $Locale = "Turkish" }
        "uk-UA" { $Locale = "Ukrainian" }
        default { $Locale = $selectedlanguage }
    }

    Write-Host "Selected Locale: $Locale"

    # Download Fido
    $fidourl = "https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1"
    $fidopath = $path + "\fido.ps1"
    Write-Log -Level 'INFO' -Message "Downloading Fido"

    try {
        Invoke-WebRequest -Uri $fidourl -OutFile $fidopath -UseBasicParsing
        Write-Log -Level 'INFO' -Message "Fido downloaded."
    }
    catch {
        Write-Log -Level 'ERROR' -Message "Failed to download Fido." -ExceptionInfo $_
    }

    # Set the parameters for Fido
    $Win = $selectedos.Major
    $Rel = $selectedos.Minor
    $Ed = "Pro"
    $GetUrl = $true

    # Build the command to run Fido
    $Command =  "$fidopath -Lang '$Locale' -Win $Win -Rel $Rel -Ed $Ed -GetUrl $GetUrl"

    try {
        Write-Log -Level 'INFO' -Message "Running Fido"
        $process = Start-Process -FilePath PowerShell.exe -ArgumentList "-Command & {$Command}" -NoNewWindow -PassThru -Wait -RedirectStandardOutput "$path\output.txt"
        Write-Log -Level 'INFO' -Message "Fido completed."
    }
    catch {
        Write-Log -Level 'ERROR' -Message "Failed to run Fido." -ExceptionInfo $_
    }

    $windowsuri = Get-Content "$path\output.txt"

    return $windowsuri
}


function Check-EnterpriseISO {
    param (
        [Parameter]
        [ValidateNotNullOrEmpty()]
        [string] $path
    )

    $allbuilds = Get-WindowsReleases -all
    $latestbuilds = Get-WindowsReleases -latest

    # Mount the ISO to check the version
    $mount = Mount-DiskImage -ImagePath $path -PassThru
    
    $driveLetter = ($mount | Get-Volume).DriveLetter

    # Path to the install.wim or boot.wim
    $wimPath = "$driveLetter`:sources\install.wim"
    if (-Not (Test-Path $wimPath)) {
        $wimPath = "$driveLetter`:sources\install.esd"
    }

    # Extract build information
    $isoBuildInfo = (dism /Get-WimInfo /WimFile:$wimPath | Select-String "Version").Line
    $isoBuildVersion = ($isoBuildInfo -split ":")[1].Trim()

    Dismount-DiskImage -ImagePath $path

    if ($isoBuildVersion -like "10.0.1*") {
        $osType = "Windows 10"
    }
    elseif ($isoBuildVersion -like "10.0.2*") {
        $osType = "Windows 11"
    }
    else {
        $osType = "Unknown"
    }

    # Compare the ISO build version with the latest available builds
    $matchedBuild = $allBuilds | Where-Object { $_.Build -eq $isoBuildVersion }
    $latestBuild = $latestBuilds | Where-Object { $_.Major -eq $osType.Split(" ")[1] }

    if ($matchedBuild) {
        if ($isoBuildVersion -eq $latestBuild.Build) {
            Write-Host "The ISO contains the latest build: $($latestBuild.Version) ($($latestBuild.Version))."
        }
        else {
            Write-Host "The ISO is not the latest."
            Write-Host "The latest build for $osType is: $($latestBuild.Name) ($($latestBuild.Version))."
        }
    }
    else {
        Write-Host "The ISO Build Version does not match any known Windows 10 or Windows 11 releases."
    }
}

function Invoke-DownloadISO {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $uri
    )

    try {
        Write-Log -Level 'INFO' -Message "Downloading ISO from $uri"
        $download = Start-BitsTransfer -Source $uri -Destination $script:isofilename -Asynchronous
        while ($download.JobState -ne "Transferred") {
            [int] $dlProgress = ($download.BytesTransferred / $download.BytesTotal) * 100;
        }
        Complete-BitsTransfer -BitsJob $download.JobId;
        Write-Log -Level 'INFO' -Message "ISO downloaded."
    }
    catch {
        Write-Log -Level 'ERROR' -Message "Failed to download ISO." -ExceptionInfo $_
    }
}

function Get-WindowsReleases {
    param (
        [switch] $Latest,
        [switch] $All
    )

    $urlList = @(
        [PSCustomObject]@{ OS = "Windows 10"; URL = "https://learn.microsoft.com/en-us/windows/release-health/release-information" },
        [PSCustomObject]@{ OS = "Windows 11"; URL = "https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information" }
    )

    $allVersions = @()

    foreach ($entry in $urlList) {
        $osType = $entry.OS
        $url = $entry.URL

        # Fetch content from the URL
        $content = (Invoke-WebRequest -Uri $url -UseBasicParsing).Content
        [regex]$regex = "(?s)<tr class=.*?</tr>"
        $tables = $regex.Matches($content).Groups.Value

        foreach ($table in $tables) {
            $cleanedTable = $table.Replace("<td>", "").Replace("</td>", "").Replace('<td align="left">', "").Replace('<tr class="highlight">', "").Replace("</tr>", "")
            $fields = $cleanedTable -split "`n" | Where-Object { $_.Trim() -ne "" }

            if ($fields.Length -ge 5) {
                $version = $fields[0].Trim()
                $build = $fields[4].Trim()

                $allVersions += [PSCustomObject]@{
                    Name = "$osType"
                    Version = $version
                    Build = $build
                }
            }
        }
    }

    if ($Latest) {
        return $allVersions | Sort-Object -Property Build -Descending | Select-Object -First 2
    } 
    elseif ($All) {
        return $allVersions
    }
}

function Set-GroupTag {

}

function New-AutopilotScript {
    param (
        [Parameter(Mandatory = $true)]
        [string] $tenant,

        [Parameter(Mandatory = $true)]
        [string] $clientId,

        [Parameter(Mandatory = $true)]
        [string] $clientSecret,

        [Parameter(Mandatory = $false)]
        [string] $GroupTag,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string] $path
    )

}

function New-AutopilotMenuScript {
    param (
        [Parameter(Mandatory = $true)]
        [string] $tenant,

        [Parameter(Mandatory = $true)]
        [string] $clientId,

        [Parameter(Mandatory = $true)]
        [string] $clientSecret,

        [Parameter(Mandatory = $true)]
        [array] $GroupTags,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string] $path
    )

}

##############################################################################################################
#                                                   Execution
##############################################################################################################


# Create path for files
$DirectoryToCreate = "c:\temp"
if (-not (Test-Path -LiteralPath $DirectoryToCreate)) {
    try {
        New-Item -Path $DirectoryToCreate -ItemType Directory -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Log -Level 'ERROR' -Message "Unable to create directory '$DirectoryToCreate'. Error was: $_" -ExceptionInfo $_
    }
}
else {
    "Directory already existed"
}

$random = Get-Random -Maximum 1000
$random = $random.ToString()
$date =get-date -format yyMMddmmss
$date = $date.ToString()
$path2 = $random + "-"  + $date
$path = "c:\temp\" + $path2

New-Item -ItemType Directory -Path $path

# Set filename and filepath
$isocontents    = "$path\iso\"
$wimname        = "$isocontents\sources\install.wim"
$wimnametemp    = "$path\installtemp.wim"

# Check if ISO path has been passed
$isocheck = $PSBoundParameters.ContainsKey('isopath')

if ($isocheck -eq $true) {
    $isofilename = $isopath
}

# Download Modules


# Download WindowsAutopilotInfoCommunity Script
$apcommunityurl = "https://raw.githubusercontent.com/andrew-s-taylor/WindowsAutopilotInfo/refs/heads/main/Community%20Version/get-windowsautopilotinfocommunity.ps1"
$apcommunitypath = $path + "scripts\get-windowsautopilotinfocommunity.ps1"
Write-Log -Level 'INFO' -Message "Downloading WindowsAutopilotInfoCommunity"
try {
    Invoke-WebRequest -Uri $apcommunityurl -OutFile $apcommunitypath -UseBasicParsing
    Write-Log -Level 'INFO' -Message "WindowsAutopilotInfoCommunity downloaded."
}
catch {
    Write-Log -Level 'ERROR' -Message "Failed to download WindowsAutopilotInfoCommunity." -ExceptionInfo $_
}