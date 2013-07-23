<#
 
.SYNOPSIS
ExRelease helps preparing softsubbed fansub releases for comparison. Extracts subtitles and fonts, loads fonts, creates indexes for ffms2 and AviSynth scripts from template.
 
.DESCRIPTION
The script works on a specified directory with any number of mkv files and does the following operations:
(1a) Looks for first subtitle stream in every mkv and extracts it to filename.[ass|srt|xxx].
(1b) Creates an index of every mkv for use with FFVideoSource().
(1c) Extracts all attachments of every to the \Fonts subdirectory. Duplicates will be overwritten.
(1d) Creates an AviSynth script for each mkv file from a template script, by default .\template.avs
(2) Looks for typesetting inside extracted .ass subtiles and exports chapters text files, which can be imported as bookmarks in AvsPmod
(3) Loads extracted fonts from the \Fonts subdirectory and makes them available to Windows applications for the duration of the session or until they are manually unloaded.
--------
(4) Unloads fonts present in the \Fonts subdirectory or whatever was specified in -fontdir

(5a) Gets a list of installed fonts and the names (not filenames) of TrueType & OpenType fonts in the \Fonts subdirectory
(5b) Checks which of the Fonts in \Fonts are not already installed and writes them to fonts.txt
(6) Installs the fonts listed in fonts.txt. The fonts will NOT be copied to the Windows fonts directory but registered for immediate use from the directory they currently reside in.
(7a) Uninstalls the fonts in fonts.txt to restore the original state
(7b) deletes fonts.txt

In it's default mode, the script will do (1), (2), and prompt you on whether or not to do (3)
The switches -extract, -loadfonts, -unloadfonts, -findts -fontlist, -install and -uninstall will let you run (1),(2),(3),(4),(5) separately or in any combination. No sanity checks.


Template variables:

{{?VIDEO_FILE}} : Inserts full path to video file
{{?SUBTITLE_FILE}} : Inserts full path to subtitle file
{{?GROUP_NAME}} : Inserts group tag, if present and encapsulated in square brackets at the beginning of the file name
{{?TARGET_RES_X}} : Inserts highest horizontal resolution found in the releases beind processed
{{?TARGET_RES_Y}} : Inserts highest vertical resolution found in the release being processed
{{?IS_UPSCALED}} : Inserts true for releases upscaled for the comparison, false otherwise

.EXAMPLE
.\ExRelease.ps1 F:\Some\Directory
Runs the complete script. You will be offered the asked whether or not to install fonts.

.EXAMPLE
.\ExRelease.ps1 -l -fontdir X:\Fonts\ShowX
Loads all Fonts from "X:\Fonts\ShowX".

.EXAMPLE
.\ExRelease.ps1 -u -d F:\Some\Directory
Unloads all Fonts from "F:\Some\Directory\Fonts".

.EXAMPLE
.\ExRelease.ps1 F:\Some\Directory -uninstall
Uninstalls fonts previously installed by the script.

.EXAMPLE
.\ExRelease.ps1 F:\Some\Directory -fontlist -install
Checks the \Fonts subdirectory for already installed fonts and installs the fonts not yet available on the host system.

.EXAMPLE
.\ExRelease.ps1 . -ts "F:\Some\Directory\subtitle.ass"
Looks for typesetting inside subtitles.ass and writes a bookmarks file to F:\Some\Directory\subtitle.TSChapters.txt

.NOTES
Requires PowerShell 3, MKVToolNix, FFMS2 and SIL FontUtils in PATH. 
Currently only works with mkv files.
Only the first subtitle stream is extracted.

ATTENTION:
Currently the script can only check for installed font families but not for different font styles and weights of fonts with the same name. As a result the script may not install a font even if the required style is not yet available on the host system.
tl;dr use of the -install option is not recommended

#>

function ExRelease
{
[CmdletBinding()]
param
(
[Parameter(Position=0, Mandatory=$true, HelpMessage='Specify the directory your mkv files are in. Examples: "F:\Some\Directory", "."')]
[alias("d")]
[string]$dir,
[Parameter(Mandatory=$false, HelpMessage='Specify the path to the template avs you want to generate scripts from Examples: "F:\Some\Directory\template.avs", ".\template.avs"')]
[alias("avs","template")]
[string]$avsTemplate = (Join-Path (Split-Path -parent $PSCommandPath) "template.avs"),
[Parameter(Mandatory=$false, HelpMessage='Specify the directory fonts will be extracted to or loaded/installed from.')]
[string]$fontDir = (Join-Path $dir "Fonts"),
[Parameter(Mandatory=$false, HelpMessage='Loads the extracted fonts and makes them available for the duration of the session.')]
[alias("l")]
[switch]$loadfonts = $false,
[Parameter(Mandatory=$false, HelpMessage='Unloads the extracted fonts.')]
[alias("u")]
[switch]$unloadfonts = $false,
[Parameter(Mandatory=$false, HelpMessage='Extracts the first subtitle tracks and fonts from all mkv files in the working directory and creates indexes for FFVideoSource.')]
[alias("e")]
[switch]$extract = $false,
[Parameter(Mandatory=$false, HelpMessage='Reads names from extracted TrueType and OpenType fonts and writes a list fonts not yet available on the host system .')]
[alias("f")]
[switch]$fontlist = $false,
[Parameter(Mandatory=$false, HelpMessage='Installs the extracted fonts marked for installation. Requires fonts.txt unless run with -fontlist.')]
[alias("i")]
[switch]$install = $false,
[Parameter(Mandatory=$false, HelpMessage='Uninstalls previously installed fonts as specified in fonts.txt.')]
[alias("r")]
[switch]$uninstall = $false,
[Parameter(Mandatory=$false, HelpMessage='Looks for typesetting inside supplied .ass subtitle and writes a timecode list to a .txt with the same base name')]
[alias("ts")]
[string]$findts
)

    $scriptVersion = 8
    Write-Host "ExRelease r$scriptVersion ($((Get-Item $PSCommandPath).LastWriteTime.toString("yyyy-MM-dd")))`n`n" -ForegroundColor Gray 
    If($PSVersionTable.PSVersion.Major -lt 3) {Throw "Powershell Version 3 required."}

    # some basic error checking

    try 
    {
        $dirContents = get-childitem $dir -ErrorAction Stop
    }
    catch
    { 
        Write-Host "Fatal Error: Failed processing directory $($_.Exception.ItemName): " -ForegroundColor Red -NoNewline 
        if($_.Exception.GetType().Name -eq "ItemNotFoundException")
        {
            Write-Host "Directory not found" -ForegroundColor Red
        }
        else { Write-Host "`nError Message: $($_.Exception.Message)" -ForegroundColor Red } 
        break
    }
    if (!(Test-Path $dir -pathType container))
    {
        Write-Host "Fatal Error: $dir is not a directory. ExRelease only works on folders." -ForegroundColor Red    
        break
    }

    $fontListPath = Join-Path $dir "fonts.txt"


    if(($fontlist -or $install -or $uninstall -or $fontlist -or $loadfonts -or $unloadfonts -or $findts)-eq $false)
    {
        $extractData = Extract -dir $dir -fontdir $fontDir 
    
        WriteAVS -extractData $extractData -avsTemplate $avsTemplate
    
        foreach ($subFile in ($extractData | ?{$_.subFile -match ".ass$"} | Select -Expand subFile))
        {
            WriteTSBookmarks -assName $subFile 
        }

        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No",""
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
        $caption = "Ready to load fonts in $($fontDir)."
        $message = "Do you want to load fonts now?"
        $result = $Host.UI.PromptForChoice($caption,$message,$choices,0)
        switch($result)
        {
	        0 { $loadFontsStatus = LoadFonts -fontDir $fontDir }
	        1 { Write-Host "To load extracted fonts at a later time , run this script with -l. To install extracted fonts, run this script with -f." }
	        default { throw "bad choice" }
        }
    }
    else
    {
        if($uninstall) { UninstallFonts -fontListPath $fontListPath }
        if($extract) { Extract -dir $dir -fontdir $fontDir }
        if($loadfonts) { $loadFontsStatus = LoadFonts -fontDir $fontDir }
        if($unloadfonts) { $unLoadFontsStatus = LoadFonts -fontDir $fontDir -unload $true}
        if($fontlist) { WriteFontList -dir $dir -fontDir $fontDir -fontListPath $fontListPath }
        if($install) { InstallFonts -fontListPath $fontListPath }
        if($findts) { WriteTSBookmarks -assName $findts }
    }
}

function CheckMissingCommands([string[]]$commands)
{
    $commandDetails = @{
                'ffmsindex.exe' = 'FFMS2'
                'mkvextract.exe' = 'MKVToolNix'
                'mkvinfo.exe' = 'MKVToolNix'
                'dumpfont.exe' = 'SIL FontUtils'
                'addfont.exe' = 'SIL FontUtils'
                }
    
    $commandDetails.GetEnumerator() | ForEach-Object {$cmdNotFound = $false} {
        if (($commands -contains $_.Key) -and !(Get-Command $_.Key -ErrorAction SilentlyContinue))
        {
            $cmdNotFound = $true
            Write-Host "Fatal Error: missing $($_.Key). Make sure $($_.Value) is installed and in your PATH." -ForegroundColor Red              
        }
    } 
    if($cmdNotFound) { break }
}

function GetMkvData([string]$file)
{
    $mkvData = &mkvinfo  --ui-language en $file

    $regex = '(?:.*?\+ A track[\s\n\|]+\+ Track number: [0-9]+ \(track ID for mkvmerge & mkvextract: )([0-9]+)(?:\)[\s\n\|]+\+ Track UID: [0-9]+[\s\n\|]+\+ Track type: video[\s\n\|]+.*?\+ Video track[\s\n\|]+'`
           + '\+ Pixel width: (?<pwidth>[0-9]+)[\s\n\|]+\+ Pixel height: (?<pheight>[0-9]+)[\s\n\|]+(?:\+ Interlaced: (?<interlaced>[0-9])[\s\n\|]+)?\+ Display width: (?<dwidth>[0-9]+)[\s\n\|]+\+ Display height: (?<dheight>[0-9]+)[\s\n\|]+)'`
           + '(?:.*?\+ A track[\s\n\|]+\+ Track number: [0-9]+ \(track ID for mkvmerge & mkvextract: (?<subtracknum>[0-9]+)\)[\s\n\|]+\+ Track UID: [0-9]+[\s\n\|]+\+ Track type: subtitles[\s\n\|]+.*?\+ Codec ID: (?<subtype>.*?))?'`
           + '(?:(?:[\s\n\|]+.*?Attachments)(?:[\s\n\|]+\+ Attached[\s\n\|]+\+ File name: (?<attachment>.*?)[\s\n]+\|.*?\+ File UID: [0-9]+)+(?:[\s\n\|]+)|[\s\n\|]+)'

    $matches = select-string -InputObject [string]$mkvData  -pattern $regex  | select -expand Matches
	
    $subType = $matches.groups["subtype"].value
	$subTrackId = $matches.groups["subtracknum"].value
	$attachments = $matches.groups["attachment"].captures
    $dResX = [int]($matches.groups["dwidth"].value)
    $dResY = [int]($matches.groups["dheight"].value)

    $mkvData = New-Object PsObject 
    if ($subType) { Add-Member -InputObject $mkvData -Name subType -Value $subType -MemberType NoteProperty }
    if ($subTrackId) { Add-Member -InputObject $mkvData -Name subTrackId -Value $subTrackId -MemberType NoteProperty }
    if ($attachments) { Add-Member -InputObject $mkvData -Name attachments -Value $attachments -MemberType NoteProperty }
    if ($dResX) { Add-Member -InputObject $mkvData -Name dResX -Value $dResX -MemberType NoteProperty }
    if ($dResY) { Add-Member -InputObject $mkvData -Name dResY -Value $dResY -MemberType NoteProperty }

    return $mkvData
}

function Extract([string]$dir, [string]$fontDir)
{
    CheckMissingCommands -commands "ffmsindex.exe", "mkvinfo.exe"

    $mkvFiles= Join-Path $dir "*" | get-childitem -include ('*.mkv')

    [PSObject[]]$extractData = @()

    foreach($file in $mkvFiles)
    {
	    Write-Host "Indexing $($file.Name) ..." -foreground yellow
	    &ffmsindex $file `        | Tee-Object -Variable ffmsOutput | %{$_.Split("`n")} `        | Select-String -pattern "(?:Indexing, please wait... )([0-9]{1,3})(?:%)" -AllMatches `        | %{$last=-1}{if ($_.Matches.groups[1].value -ne $last -and $_.Matches.groups[1].value % 5 -eq 0) `                        { Write-Host "$($_.Matches.groups[1].value)% " -ForegroundColor Gray -NoNewline;                            $last=$_.Matches.groups[1].value 
                        } 
           }
        
        # TODO: add more error checking
        if ($ffmsOutput -match "index file already exists")
             { Write-Host "Index file already exists.`n" -ForegroundColor Gray }
        else { Write-Host "Done.`n" -ForegroundColor Green } 
             
	    $mkvData = GetMkvData -file $file
        Write-Host "Display Resolution: $($mkvData.dResX)x$($mkvData.dResY)`n"     

        Write-Host "Extracting subtitles..." -foreground yellow
        if ($mkvData.subTrackId)
        {
            $subExt=switch($mkvData.subType)
	        {
		        "S_TEXT/ASS" { "ass" }
		        "S_TEXT/UTF8" { "srt" }
                "S_VOBSUB" { "sub" }
		        default { "unknown" }
	        }
            
            $subFile = (Join-Path $file.Directory $file.BaseName) + ".$subExt"

  	        # properly filters and outputs mkvextract progress without spamming the shell             # TODO: parse $mkvexEOutput for potential errors            &mkvextract tracks $file "$($mkvData.subTrackId):$subFile" `            | Tee-Object -Variable mkvexEOutput | %{$_.Split("`n")} `            | Select-String -pattern "(?:Progress: )([0-9]{1,3})(?:%)" -AllMatches `            | %{$last=-1}{if ($_.Matches.groups[1].value -ne $last -and $_.Matches.groups[1].value % 5 -eq 0)                           { Write-Host "$($_.Matches.groups[1].value)% " -ForegroundColor Gray -NoNewline;                             $last=$_.Matches.groups[1].value 
                           }
               }{Write-Host "Done.`n" -ForegroundColor Green }
        }
        else 
        { Write-Host "No subtitles found.`n" -foreground gray }
	
        
        Write-Host "Extracting fonts..." -foreground yellow

	    if($mkvData.attachments.count -gt 0)
        {
            [string[]] $attachmentArgs=@()
	        $i=1
	        foreach($attachment in $mkvData.attachments)
	        {
                $fontDirUnescaped = [System.Management.Automation.WildcardPattern]::Unescape($fontDir)
		        $attachmentArgs += "$($i):" + (Join-Path $fontDirUnescaped $attachment.value)
		        $i++
	        }
	
	        &mkvextract attachments $file $attachmentArgs `            | Tee-Object -Variable mkvexAOutput | %{$_.Split("`n")} `            | Select-String -pattern "(#[0-9]+)(?:.*?, is written to ')(.*?)(?:'.)" -AllMatches `            | %{ Write-Host "$($_.Matches.groups[1].value): $($_.Matches.groups[2].value)" -ForegroundColor Gray } -end {Write-Host "Done.`n" -ForegroundColor Green}
        }
        else { Write-Host "No fonts attached to $($file.Name).`n" -ForegroundColor Gray }

        $extractData += (New-Object PsObject -Property @{videoFile=$file; dResX=$mkvData.dResX; dResY=$mkvData.dResY} | ? { (!$subFile) -or (Add-Member -InputObject $_ -MemberType NoteProperty -Name subFile -Value $subFile -PassThru) })
        Remove-Variable subFile -ErrorAction SilentlyContinue
        
    }
    return $extractData
}

function WriteAVS([PSObject[]]$extractData, [string]$avsTemplate)
{
    if(Test-Path $avsTemplate)
    {
        [System.IO.FileSystemInfo]$avsTemplateFile = Get-Childitem $avsTemplate
        Write-Host "Generating scripts from $($avsTemplateFile.Name)" -foreground yellow
        $avsTemplateContent = Get-Content -Path $avsTemplateFile -Raw
        $targetResX = $extractData | %{ $max=0 } { if($_.dResX -gt $max){$max=$_.dResX} } {$max}
        $targetResY = $extractData | %{ $max=0 } { if($_.dResY -gt $max){$max=$_.dResY} } {$max}

        Write-Host "Target Resolution: $($targetResX)x$targetResY"
        foreach ($release in $extractData)
        {
            
            $releaseInfo=select-string -InputObject $release.videoFile.Name  -pattern '^(\[(?<group>.*?)\])?(.*?)([\[\(](?<crc>[0-9A-Fa-f]{8})[\]\)])?.mkv'  | select -expand Matches

            # Replace template variables using lookup table
            $templateVars = @{
                '{{\?VIDEO_FILE}}' = [string]$release.videoFile.FullName
                '{{\?SUBTITLE_FILE}}' = if($release.subFile) { [string]$release.subFile } else { "" }
                '{{\?GROUP_NAME}}' = [string]$releaseInfo.groups["group"].value 
                '{{\?TARGET_RES_X}}' = if($targetResX -gt 0){[string]$targetResX} else {"w"}
                '{{\?TARGET_RES_Y}}' = if($targetResY -gt 0){[string]$targetResY} else {"h"}
                '{{\?IS_UPSCALED}}' = if(($targetResX -gt $release.dResX) -or ($targetResY -gt $release.dResY)) {"true"} else {"false"}
            } #escaped for regex

            $templateVars.GetEnumerator() | ForEach-Object {$avsScriptContent=$avsTemplateContent} {
                if ($avsTemplateContent -match $_.Key)
                {
                    $avsScriptContent = $avsScriptContent -replace $_.Key, $_.Value
                }
            }

            Out-File -LiteralPath "$(Join-Path $release.videoFile.Directory $release.videoFile.BaseName).avs" -InputObject $avsScriptContent -Encoding oem
         }
    }

    else { Write-Host "AVS template file  $($avsTemplateFile.BaseName) not found, skipping script generation." }
}

function WriteTsBookmarks([string]$assName)
{
    $overrideTags = @{
        '\\blur' = 1
        '\\fscx' = 1
        '\\fscy' = 1
        '\\bord' = 2
        '\\be[0-9]' = 1
        '\\fn' = 2
        '\\fs(?!p[0-9]|cy[0-9]|cx[0-9])' = 2
        '\\fsp' = 1
        '\\frx' = 3
        '\\fry' = 3
        '\\frz' = 2
        '\\fax' = 3
        '\\fay' = 3
        '\\[1-4]?c&H' = 2
        '\\[1-4]?a&H' = 3
        '\\an[1-9]' = 1
        '\\pos' = 1
        '\\move' = 4
        '\\org' = 4
        '\\fade?' = 1
        '\\t\(.*?\)' = 4
        '\\i?clip' = 4
        '\\p1' = 4


    } #escaped for regex
    
    $tsScoreTresh = 5
    $timeDiffTresh = 2 # number of seconds TS lines need to be apart to be logged
    $lineDiffTresh = 2 # number of lines between the current and last TS required for the current TS to be logged


    [System.IO.FileSystemInfo]$assFile = Get-Childitem -LiteralPath ([System.Management.Automation.WildcardPattern]::Unescape($assName))
    
    
    Write-Host "Looking for typesetting in $($assFile.Name)..." -foreground yellow

    $assContent = Get-Content -Raw -LiteralPath $assFile

    $assContentSections = $assContent -split "\[Script Info\]|\[V4\+ Styles\]|\[Events\]",0, "multiline"

    [int]$lineNum = 0
    [int]$lineNumLastTS = -$lineDiffTresh
    $lineStartTimeLastTS = Get-Date -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    [string[]] $bookmarks=@()

    $linesSorted = $assContentSections[3] -split [environment]::NewLine`
                   | Sort-Object -Property @{Expression=`                   {$_ -replace "(.*?: [0-9]*,)([0-9]+:[0-9]{2}:[0-9]{2}.[0-9]{2})(.*)", "`$2" | Get-Date}`                   }
    
    foreach ($line in $linesSorted)
    {
        $regex = '(.*?): [0-9]*,([0-9]+:[0-9]{2}:[0-9]{2}.[0-9]{2})(?:,.*?){8}(.*)'
        $lineMatches = select-string -InputObject $line -pattern $regex  | select -expand Matches
        if($lineMatches)
        {
            $lineType =  $lineMatches.groups[1].value
            $lineStartTime = ([DateTime]$lineMatches.groups[2].value).AddMilliseconds(100)
            $timeDiff = New-TimeSpan -Start $lineStartTimeLastTS -End $lineStartTime
            $lineText = $lineMatches.groups[3].value
              

            $overrideTags.GetEnumerator() | ForEach-Object {[int]$tsScore = 0} {
            $matches = select-string -InputObject $lineText -pattern $_.Key -AllMatches
            $tsScore = $tsScore + ([int]$matches.Matches.Count * [int]$_.Value)
            }

            if ($tsScore -ge $tsScoreTresh -and $lineType -eq "Dialogue")
            {
                if($timeDiff -ge $timeDiffTresh -and ($lineNum - $lineNumLastTS) -ge $lineDiffTresh)
                {     
                    $bookmarks += "$lineNum=$($lineStartTime.toString("HH:mm:ss.fff"))"
                    $lineStartTimeLastTS = $lineStartTime
                }
                $lineNumLastTS = $lineNum
            }
            $lineNum++ 
        }  
    }
    Write-Host "Found $($bookmarks.Count) signs." -ForegroundColor Green
    Out-File -LiteralPath "$(Join-Path $assFile.Directory $assFile.BaseName).TSChapters.txt" -InputObject $bookmarks -Encoding oem
}

function GetStringAtLocale($fontMatches,[string]$groupLocale,[string]$groupString)
{
     [string]$locale=[string]1033 # hardcoded en-us for now
     $idx=[array]::IndexOf($fontMatches.groups[$groupLocale].Captures.Value, $locale)
     if ($idx -eq -1) { $idx=0 } 
     return $fontMatches.groups[$groupString].Captures[$idx].Value
}

function PerlToUnicode([string]$string)
{
    $matchEval = { 
        param($m)
        $charCode = $m.Groups[1].Value
        [char][int] "0x$charCode"
    }
    return [regex]::Replace($string, '\\x\{([0-9a-fA-F]{1,4})\}', $matchEval)
}

function GetFontData([string]$fontPath)
{
    CheckMissingCommands -commands "dumfont.exe"

    $fontData = dumpfont -t name $fontPath
    # Yeah, this is fucking terrible, also broken for some fonts. Replace with something sane ASAP
    $fontMatches = select-string -InputObject [string]$fontdata  -pattern '^(?:.*?\[.[\s\n]+#0.[\s\n]+)(?:\[(?:[\n\s]+#[0-9][\n\s]+(?:(?:\[.*?\],?)|(?:undef,)))+(?:[\n\s]+\])|undef),(?:[\s\n]+#1[\s\n]+\[)(?:[\n\s]+#[0-2][\n\s]+(?:(?:\[.*?\],?)|(?:undef,)))+(?:[\s\n]+#3[\s\n]+\[.*?(?:#0[\s\n]+undef,[\s\n]+#1|#0)[\s\n]+\{)(?:[\s\n]+(?<floc>[0-9]+)[\s]+=>[\s]+"(?<fname>.*?)",?)+(?:[\s\n]+\}[\s\n]+\][\s\n]+\],[\s\n]+#2[\s\n]+\[)(?:[\n\s]+#[0][\n\s]+(?:(?:\[.*?\],?)|(?:undef,))[\s\n]+)(#1[\s\n]+(?:(?:\[.*?#0[\s\n]+\{)(?:[\s\n]+[0-9]+[\s]+=>[\s]+"(?<fstyle>.*?)")(?:[\s\n]+\}(?:(?:,[\s\n]+#[1-9][\s\n]+\{[\s\n]+[0-9]+[\s]+=>[\s]+".*?"[\s\n]+\})+)*[\s\n]+\])|(?:undef)),[\s\n]+)(?:[\n\s]+#2[\n\s]+(?:(?:\[.*?\],?)|(undef,)))[\n\s]+#3[\n\s]+(?:(?:\[[\n\s]+(?:#0[\n\s]+undef,[\n\s]+)?#[0-1][\s\n]+\{(?:[\s\n]+(?<fweightloc>[0-9]+)[\s]+=>[\s]+"(?<fweightname>.*?)",?)+)(?:[\s\n]+\}[\s\n]+\],?))' | select -expand Matches
       
    $fontName = PerlToUnicode -string (GetStringAtLocale -fontMatches $fontMatches -groupLocale 'floc' -groupString 'fname')
    $fontStyle=$fontMatches.groups["fstyle"].value
    $fontWeight = GetStringAtLocale -fontMatches $fontMatches -groupLocale 'fweightloc' -groupString 'fweightname'

    return New-Object PsObject -Property @{Name=$fontName ; Style=$fontStyle; Weight=$fontWeight}   
}

function WriteFontList([string]$dir, [string]$fontDir, [string]$fontListPath)
{
    Write-Host "Checking for already installed fonts..." -foreground yellow
    $fonts = Join-Path $fontDir "*" | get-childitem -include ('*.ttf', '*.ttc', '*.otf')
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    $objFontsInstalled = New-Object System.Drawing.Text.InstalledFontCollection
    [string[]] $fontsToInstall=@()

    foreach($font in $fonts) 
    {
	    $fontData = GetFontData -fontPath $font

	    $isFontInstalled = $objFontsInstalled.Families -contains $fontData.Name
	    write-host  "$($fontData.Name) ($($font.Name)): " -nonewline
	    if($isFontInstalled -eq 1)
	    {
		    Write-Host "already installed" -foreground gray
	    }
	    else 
	    {
		    $fontsToInstall += $font
		    Write-Host "marked for installation" -foreground yellow		
	    }

    }
    if(!$fontlist -and (Test-Path $fontListPath))
    {
        throw "$fontListPath already exists. To avoid losing track of fonts installed by the script it will not be overwritten by default. If you want to do so, run the script with -fontlist."
    }
    else 
    { 
        $fontsToInstall > $fontListPath 
        Write-Host "$fontListPath written." -foreground green
    }
    return $fontsToInstall
}

function InstallFonts([string[]]$fontsToInstall, [string]$fontListPath)
{
    CheckMissingCommands -commands "addfont.exe"

    if(-not $fontsToInstall)
    {
        $fontsToInstall=GetFontListFromFile($fontListPath)
    }

	Write-Host "Installing Fonts..." -foreground yellow
	for($i=0; $i -le $fontsToInstall.count -1; $i++)
	{

		echo $fontsToInstall[$i]
        $arg = $fontsToInstall[$i]
		# Work around addfont.exe bug
		if(select-string -InputObject $arg -pattern "\s")
		{
			&addfont `'$arg`'
		}
		else { &addfont $arg }
	}
		
	Write-Host "Done installing fonts." -foreground green

}

function LoadFonts([string]$fontDir, [boolean]$unload)
{

    if (!$unload) { 
        $text = "Loading fonts..." 
        $signature = @’
        [DllImport("gdi32.dll")]
        public static extern int AddFontResourceEx(
        string lpszFilename, 
        uint fl, 
        IntPtr pdv);
‘@

        $dotnet = Add-Type -MemberDefinition $signature `
         -Name GDI32 -Namespace AddFontResourceEx `
         -Using System.Text  -PassThru
    } 

    else { 
        $text = "Unloading fonts..." 

        $signature = @’
        [DllImport("gdi32.dll")]
        public static extern bool RemoveFontResourceEx(
        string lpFileName, 
        uint fl,
        IntPtr pdv);
   
‘@

    $dotnet = Add-Type -MemberDefinition $signature `
        -Name GDI32 -Namespace RemoveFontResourceEx `
        -Using System.Text  -PassThru 
    }
    
    Write-Host $text -foreground yellow
    try 
    {
        $fonts = Join-Path $fontDir "*" | get-childitem -include ('*.ttf', '*.ttc', '*.otf') -ErrorAction Stop
    }
    catch
    { 
        Write-Host "Failed loading fonts from $($_.Exception.ItemName): " -ForegroundColor Red -NoNewline 
        if(($_.Exception.GetType().Name -eq "ItemNotFoundException") -or ($_.Exception.GetType().Name -eq "DriveNotFoundException"))
        {
            Write-Host "Directory not found" -ForegroundColor Red
        }
        else { Write-Host "`nError Message: $($_.Exception.Message)" -ForegroundColor Red } 
        return $false
    }

    $isAllSuccessful = $true
    foreach($font in $fonts) 
    {
	    $fontData = GetFontData -fontPath $font
        
	    write-host  "$($fontData.Name) ($($font.Name)): " -nonewline
        
        if (!$unload) { $isFontLoaded = $dotnet::AddFontResourceEx([string]$font,0,0) }
        else { $isFontLoaded = $dotnet::RemoveFontResourceEx([string]$font,0,0)}

	    if($isFontLoaded -eq 1)
	    {
		    Write-Host "successful" -foreground green
	    }
	    else 
	    {
		    Write-Host "failed" -foreground red
            $isAllSuccessful = $false		
	    }
    }
    
    if(!$isAllSuccessful) { Write-Host "`nOne or more fonts failed to $(if ($unload) {"un"})load. This may or may not be a problem." -ForegroundColor Red } 
    if (!$unload) { Write-Host "Fonts will be available for the duration of the current session or until manually unloaded." -foreground gray }
    
    return $isAllSuccessful   
}

function GetFontListFromFile([string]$fontListPath)
{
    if (Test-Path $fontListPath)
    {
        return Get-Content $fontListPath 
    }
    else {throw "Font list not found."}
}

function UninstallFonts([string]$fontListPath)
{
    CheckMissingCommands -commands "addfont.exe"

    $fontsToUninstall=GetFontListFromFile($fontListPath)

    Write-Host "Uninstalling previously installed fonts..." -foreground yellow
    foreach($font in $fontsToUninstall)
    {
        if(select-string -InputObject $font -pattern "\s")
        {
	        &addfont -r `'$font`'
        }
        else { &addfont -r $font }
    }
    Remove-Item $fontListPath
    Write-Host "Done uninstalling fonts." -foreground green
}

Export-ModuleMember ExRelease