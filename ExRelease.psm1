<#
 
.SYNOPSIS
ExRelease helps preparing softsubbed fansub releases for comparison. Extracts subtitles and fonts, loads fonts, creates indexes for ffms2 and AviSynth scripts from template.
 
.DESCRIPTION
The script works on a specified directory with any number of mkv files and does the following operations:
(1) Creates an index of every mkv for use with FFVideoSource().
(2a) Looks for subtitle streams in every mkv and extracts it to filename.trackNum.[ass|srt|xxx].
(2b) Extracts all attachments of every to the \Fonts subdirectory. Duplicates will be overwritten.
(2c) Creates an AviSynth script for each mkv file from a template script, by default .\template.avs
(3) Looks for typesetting inside extracted .ass subtiles and exports chapters text files, which can be imported as bookmarks in AvsPmod
(4) Loads extracted fonts from the \Fonts subdirectory and makes them available to Windows applications for the duration of the session or until they are manually unloaded.
--------
(5) Unloads fonts present in the \Fonts subdirectory or whatever was specified in -fontdir

(6a) Gets a list of installed fonts and the names (not filenames) of TrueType & OpenType fonts in the \Fonts subdirectory
(6b) Checks which of the Fonts in \Fonts are not already installed and writes them to fonts.txt
(7) Installs the fonts listed in fonts.txt. The fonts will NOT be copied to the Windows fonts directory but registered for immediate use from the directory they currently reside in.
(8a) Uninstalls the fonts in fonts.txt to restore the original state
(8b) deletes fonts.txt

In it's default mode, the script will do (1), (2), and prompt you on whether or not to do (3)
The switches -index, -extract, -findts, -loadfonts, -unloadfonts, -fontlist, -install and -uninstall will let you run (1)-(8) separately or in any combination. No sanity checks.


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
#requires -version 3

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
[string]$findts,
[Parameter(Mandatory=$false, HelpMessage='Creates FFMS2 indexes for all files.')]
[string]$index,
[Parameter(Mandatory=$false, HelpMessage='Returns script version Number')]
[switch]$version
)

    $scriptVersion = [version]"0.9.5"
    Write-Host "ExRelease r$scriptVersion ($((Get-Item $PSCommandPath).LastWriteTime.toString("yyyy-MM-dd")))`n`n" -ForegroundColor Gray 
    If($PSVersionTable.PSVersion.Major -lt 3) {Throw "Powershell Version 3 required."}

    if(!($dir -match '`')) { $dir = [System.Management.Automation.WildcardPattern]::Escape($dir) }

    # some basic error checking
    try
    {
        $dirContents = get-childitem -LiteralPath ([System.Management.Automation.WildcardPattern]::Unescape($dir)) -ErrorAction Stop
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


    if(($fontlist -or $install -or $uninstall -or $fontlist -or $loadfonts -or $unloadfonts -or $findts -or $version -or $index -or $extract)-eq $false)
    {
        $extractData = Extract -dir $dir -fontdir $fontDir
        Index -dir $dir 

        WriteAVS -extractData $extractData -avsTemplate $avsTemplate
    
        foreach ($subFile in ($extractData | %{ $_.subFiles} | ?{$_ -match ".ass$"}))
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
        if($version) { return $scriptVersion; break }
        if($uninstall) { UninstallFonts -fontListPath $fontListPath }
        if($index) { Index -dir $dir }
        if($extract) { $extractData = Extract -dir $dir -fontdir $fontDir }
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

function OptAdd([object]$obj,$props)
{
    $props | ?{ $_.Val }| %{ 
        $val = if(!$_.Type) {$_.Val} else {$_.Val -as $_.Type}    
        Add-Member -InputObject $obj -Name $_.Prop -Value $val -MemberType NoteProperty             
    }
    return $obj
}

function Get-MkvInfo([string]$file)
{
    $mkvInfo = (&mkvinfo  --ui-language en $file) | ? {$_.trim() -ne "" }

    $i=0
    $mkvInfoFilt = @()
    foreach ($line in $mkvInfo)
    {
        $regex = '^([| ]*)\+ (.*?)(?:(?:\: (.*?))|(?:, size [0-9]*))?(?:$)'
        $matches = select-string -InputObject $line -pattern $regex  | Select -ExpandProperty Matches

        [int]$depth = $matches.Groups[1].Length #get tree depth from indentation
        $prop = $matches.Groups[2].Value
        $val = $matches.Groups[3].Value

        $mkvInfoFilt += ,@($i, $depth, $prop, $val)
        $i++
    }

    $maxDepth = $mkvInfoFilt | %{$_[1]} | Measure-Object -Maximum
    [array]::reverse($mkvInfoFilt)

    for([int]$i=$maxDepth.Maximum; $i -ge 0; $i--)
    {
        $lines = @($mkvInfoFilt | ?{$_[1] -eq $i}) + ,@(-1,-1,-1) # additional dummy line required to parse last line

        Remove-Variable lastLine -ErrorAction SilentlyContinue
        foreach($line in $lines)
        {
            if(!$lastLine) { $arr = @() }

            elseif(($lastLine[0]-1) -ne $line[0]) #non-consecutive line indexes mean we've found our parent node
            {
                $idx = [array]::IndexOf(($mkvInfoFilt | %{$_[0]}),$lastLine[0]-1)
                [array]::reverse($arr)
                $mkvInfoFilt[$idx][3] = $arr
                $arr = @()
            }

            if($line[0] -ne -1) #skip dummy entries
            {
                if (!$arr.($line[2])) #if a sibling node with the same name doesn't exist...
                { 
                    $hash = @{$line[2] = ,@($line[3])} # create new hashtable and add array with value of the current line as first element
                    $arr += $hash

                }
                else
                {
                    $arr[-1].Set_Item($line[2],(,@($line[3]))+$arr[-1].($line[2])) # else add value as new array item to the existing hashtable
                }
                $lastLine = $line
            }

        }
        $mkvInfoFilt = $mkvInfoFilt | ?{$_[1] -eq 0 -or $_[1] -ne $i}    # remove processed lines except the parent nodes 
        0..($mkvInfoFilt.Length-1) |%{$mkvInfoFilt[$_][0]=$mkvInfoFilt.Length-1-$_} # make line indexes continuous again
    }

    $segments = $mkvInfoFilt | ?{ $_[2] -eq "Segment"} | %{$_[3]}
    $tracks = $segments.("Segment tracks")[0]["A track"]

    $segmentInfo = [PSCustomObject]@{
        UID = [byte[]]$(,@($segments.("Segment information").("Segment UID")[0] -split "\s" | %{[byte]$_}))
        Duration = [TimeSpan]($segments.("Segment information").("Duration")[0] -creplace ".*?s \(([0-9]*:[0-9]{2}:[0-9]{2}.[0-9]{3})\)","`$1")
        TrackCount = $tracks.Length
        }

    if($segments.("Segment information").("Title")) 
        { Add-Member -InputObject $segmentInfo -Name Title -Value $segments.("Segment information").("Title")[0] -MemberType NoteProperty }


    $tracksInfo = @()
    foreach ($track in $tracks)
    {
        $trackId = [int]($track.("Track number")[0] -creplace "[0-9]+ \(track ID for mkvmerge \& mkvextract\: ([0-9]*)\)","`$1")
        $trackType = $track.("Track type")[0]

        $trackInfo = [PSCustomObject]@{
            ID = $trackId
            Type = $trackType
            Codec = $track.("Codec ID")[0]
        }

        $trackInfo = OptAdd -obj $trackInfo -props @(@{Prop="Name"; Val=$track.("Name")},
                                                     @{Prop="Lang"; Val=$track.("Language")}
                                                     @{Prop="Enabled"; Val=$track.("Enabled"); Type=[type]"bool"})

        if($trackType -eq "video")
        {
            $trackInfoVideo = @{
                Framerate = [float]($track.("Default duration")[0] -creplace "[0-9]*.[0-9]*ms \(([0-9]*.[0-9]*) frames.*?\)","`$1")
                dResX = [int]$track.("Video track").("Display width")[0]
                dResY = [int]$track.("Video track").("Display height")[0]
                pResX = [int]$track.("Video track").("Pixel width")[0]
                pResY = [int]$track.("Video track").("Pixel height")[0]
            }
            $trackInfoVideo = OptAdd -obj $trackInfoVideo -props @(@{Prop="Interlaced"; Val=$track.("Video track").("Interlaced"); Type=[type]"bool"})
            Add-Member -InputObject $trackInfo -NotePropertyMembers $trackInfoVideo
        }

        if($trackType -eq "audio")
        {
            $trackInfoAudio = @{
                SampleRate = [int]$track.("Audio track").("Sampling Frequency")[0]
                ChannelCount = [int]$track.("Audio track").("Channels")[0]
            }
            $trackInfoAudio = OptAdd -obj $trackInfoAudio -props @(@{Prop="BitDepth"; Val=$track.("Audio track").("Bit Depth"); Type=[type]"int"})
            Add-Member -InputObject $trackInfo -NotePropertyMembers $trackInfoAudio
        }

        $tracksInfo += $trackInfo
    }
    $segmentInfo = $segmentInfo | Add-Member -MemberType NoteProperty -Name Tracks -Value $tracksInfo -PassThru

    
    if($segments.("Attachments")) 
    { 
        $attsInfo = @()
        #$atts = $segments.("Attachments").("Attached")
        $atts = $segments.("Attachments")[0]["Attached"] # because the above line apparently resolves array too far when there's only one element
   
        foreach ($att in $atts)
        {
            $attInfo = [PSCustomObject]@{
                UID = [uint64]$att.("File UID")[0]
                MimeType = $att.("Mime type")[0]
                Name = $att.("File Name")[0]
                }
            $attsInfo += $attInfo
        }
        $segmentInfo = $segmentInfo | Add-Member -MemberType NoteProperty -Name Attachments -Value $attsInfo -PassThru
    }


    $segmentInfo = $segmentInfo | Add-Member -MemberType ScriptMethod -Value `
    { param([Parameter(Mandatory=$true)][string]$type) 
        return $this.Tracks | ?{ $_.Type -eq $type }
    } -Name GetTracksByType -PassThru

    $segmentInfo = $segmentInfo | Add-Member -MemberType ScriptMethod -Value `
    { param([Parameter(Mandatory=$true)][int]$id) 
        return $this.Tracks | ?{ $_.ID -eq $id }
    } -Name GetTrackById -PassThru

    return $segmentInfo
}
function Index([string]$dir)
{
    CheckMissingCommands -commands "ffmsindex.exe"

    $mkvFiles= Join-Path $dir "*" | get-childitem -include ('*.mkv')
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
    }
}

function Extract([string]$dir, [string]$fontDir)
{
    CheckMissingCommands -commands "mkvinfo.exe"

    $mkvFiles= Join-Path $dir "*" | get-childitem -include ('*.mkv')
    [PSObject[]]$extractData = @()

    foreach($file in $mkvFiles)
    {
        $mkvInfo = Get-MkvInfo -file $file
        $vInfo = $mkvInfo.GetTracksByType("video")[0]
        $subs = $mkvInfo.GetTracksByType("subtitles")
       
        Write-Host "$($file.Name)$(if($mkvInfo.Title){": " + $mkvInfo.Title}) ($($mkvInfo.TrackCount) tracks)"
        Write-Host "Display Resolution: $($vInfo.dResX)x$($vInfo.dResY)`n"     
        Write-Host "Extracting subtitles..." -foreground yellow
        if ($subs)
        {
            $subFiles = @()
            foreach ($sub in $subs) 
            {
                $subExt=switch($sub.Codec)
	            {
		            "S_TEXT/ASS" { "ass" }
		            "S_TEXT/UTF8" { "srt" }
                    "S_VOBSUB" { "sub" }
		            default { "unknown" }
	            }
                Write-Host "#$($sub.ID): $(if($sub.Name) {$sub.Name} else {("unnamed")}) ($subExt)" -ForegroundColor Gray

                $subFiles += ((Join-Path $file.Directory $file.BaseName) + ".$($sub.ID).$subExt")

                if(!(Test-Path -LiteralPath $subFiles[-1] -PathType Leaf))                {
  	                # properly filters and outputs mkvextract progress without spamming the shell                     # TODO: parse $mkvexEOutput for potential errors                    &mkvextract tracks $file "$($sub.ID):$($subFiles[-1])" `                    | Tee-Object -Variable mkvexEOutput | %{$_.Split("`n")} `                    | Select-String -pattern "(?:Progress: )([0-9]{1,3})(?:%)" -AllMatches `                    | %{$last=-1}{if ($_.Matches.groups[1].value -ne $last -and $_.Matches.groups[1].value % 5 -eq 0)                                   { Write-Host "$($_.Matches.groups[1].value)% " -ForegroundColor Gray -NoNewline;                                     $last=$_.Matches.groups[1].value 
                                   }
                       }{Write-Host "Done.`n" -ForegroundColor Green }
                } else { Write-Host "$($subFiles[-1]) already exists, skipping." -ForegroundColor Gray }
            }

        }
        else 
        { Write-Host "No subtitles found.`n" -foreground gray }
	
        
        Write-Host "Extracting fonts..." -foreground yellow

        $atts = $mkvInfo.Attachments
	    if($atts.count -gt 0)
        {
            [string[]] $attachmentArgs=@()
	        $i=1
	        foreach($att in $atts)
	        {
                $fontDirUnescaped = [System.Management.Automation.WildcardPattern]::Unescape($fontDir)
		        $attachmentArgs += "$($i):" + (Join-Path $fontDirUnescaped $att.Name)
		        $i++
	        }
	
	        &mkvextract attachments $file $attachmentArgs `            | Tee-Object -Variable mkvexAOutput | %{$_.Split("`n")} `            | Select-String -pattern "(#[0-9]+)(?:.*?, is written to ')(.*?)(?:'.)" -AllMatches `            | %{ Write-Host "$($_.Matches.groups[1].value): $($_.Matches.groups[2].value)" -ForegroundColor Gray } -end {Write-Host "Done.`n" -ForegroundColor Green}
        }
        else { Write-Host "No fonts attached to $($file.Name).`n" -ForegroundColor Gray }

        $extractData += (New-Object PsObject -Property @{videoFile=$file; dResX=$vInfo.dResX; dResY=$vInfo.dResY} | ? { (!$subFiles) -or (Add-Member -InputObject $_ -MemberType NoteProperty -Name subFiles -Value $subFiles -PassThru) })
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
                '{{\?SUBTITLE_FILE}}' = if($release.subFiles) { [string]$release.subFiles[0] } else { "" }
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
        '\\(?:fsc[xy])|(?:be[0-9])|(?:fsp)|(?:an[1-9])|(?:fade?)' = 1
        '\\(?:blur)|(?:bord)|(?:fn)|(?:fs(?!p[0-9]|cy[0-9]|cx[0-9]))|(?:[1-4]?c&H)|(?:pos\(.*?\))' = 2
        '\\(?:fr[xyz])|(?:fa[xy])|(?:[1-4]?a(?:lpha)?&H)' = 3
        '\\(?:move)|(?:org)|(?:t\(?:.*?\))|(?:i?clip)|(?:p1)' = 4
        '\\[kK][fo]?[0-9]+' = 5
    } #escaped for regex
    
    $tsScoreTresh = 5
    $timeDiffTresh = 2 # number of seconds TS lines need to be apart to be logged
    $lineDiffTresh = 4 # number of lines between the current and last TS required for the current TS to be logged


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
        $regex = '(Dialogue): [0-9]*,([0-9]+:[0-9]{2}:[0-9]{2}.[0-9]{2})(?:,.*?){7}(.*?),(.*)'
        $lineMatches = select-string -InputObject $line -pattern $regex  | select -expand Matches
        if($lineMatches)
        {
            $lineType =  $lineMatches.groups[1].value
            $lineStartTime = ([DateTime]$lineMatches.groups[2].value)
            $timeDiff = New-TimeSpan -Start $lineStartTimeLastTS -End $lineStartTime
            $lineText = $lineMatches.groups[4].value
            $lineEffect = $lineMatches.groups[3].value

            $overrideTags.GetEnumerator() | ForEach-Object {[int]$tsScore = 0} {
            $matches = select-string -InputObject $lineText -pattern $_.Key -AllMatches
            $tsScore = $tsScore + ([int]$matches.Matches.Count * [int]$_.Value)
            }

            if ($tsScore -ge $tsScoreTresh -or $lineEffect)
            {  
                if($timeDiff -ge $timeDiffTresh -and ($lineNum - $lineNumLastTS) -ge $lineDiffTresh)
                {   
                    Write-Host "Found typesetting at line $lineNum ($($lineStartTime.toString("HH:mm:ss")))" -ForegroundColor Gray  
                    $bookmarks += "$lineNum=$($lineStartTime.AddMilliseconds(100).toString("HH:mm:ss.fff"))"
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