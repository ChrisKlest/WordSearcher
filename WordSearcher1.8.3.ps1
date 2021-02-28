#STRINGSEARCHER

# USED TO PERFORM RECURSIVE LOOKUP OF KEY WORDS IN SPECIFIED FOLDER DIRECTORY AND FILE OR SEARCH BY FILETYPE

# WARNING: RECURSIVE LOOKUP WILL LOOK IN EACH FOLDER WITHIN THE FOLDER YOU SPECIFY, THIS CAN TAKE CONSIDERABLE AMOUNTS OF TIME DEPENDING ON YOUR STARTING LOCATION.

# Once you have configured the script to your liking, remove the leading # and press play.

#TO UPDATE: Line 137 Select-String. Create and/or options
#TO UPDATE: Create Form for search parameter entry


cls


#region starttimer

   $starttimer = Get-date

#endregion starttimer


#region config
    
    #Search This Folder
    $dirs = "F:\Documents +\Contracts\Jose"
        # Full Folder path in quotes. This can be a single folder, or multiple folders separated by commas.
        # e.g. "F:\TextmessagesBackup\CombinedAllTextMessages", "C:\UserName\Desktop" 

    #Search through subfolders? aka Recursive Search
    $r = $true
        # $true or $false. Note, setting to $true can take a long time if searching through all subfolders.

    #Filter for these File Properties
    $year            = ""
    $DayofWeek       = ""
    $ext             = "txt"
    $FileNameNotLike = ""
    $FileNameLike    = ""
        # $year           = Last Write Time.           Double Quotes if no filter (zero length string) "", else 4 digit year separated by comma. e.g. 2012, 2013, 2014, 2015, 2016, 2017
        # $DayofWeek      = Last Write Time.           Double Quotes if no filter (zero length string) "", else full day name in quotes and separated by comma. e.g. "Monday","Tuesday"
        # $ext            = File Extension.            Double Quotes if no filter (zero length string) "", else file extension in quotes and separated by comma. e.g. "xlsx","csv","docx","xlsm","pdf"
        # FileNameNotLike = File Name not Like string. Double Quotes if no filter (zero length string) "", else strings in quotes separated by comma. e.g. "nameNotLike","oranges","bananas"
        # FileNameLike    = File Name like string.     Double Quotes if no filter (zero length string) "", else strings in quotes separated by comma. e.g. "nameLike","mangos","fruit"

    #File Content Must Contain
    $string = ""
        # Double Quotes if no filter (zero length string) "", else String or Array of Strings in quotes and separated by commas.
        # e.g. "on my way", "be there in", "should be about", "omw", "work?"

    #Open results file when done?
    $openfile = 2
        # 2 Always open results file.
        # 1 Ask to open.
        # 0 Never open.
    
    #Results File Properties  
    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    $filename = "WordSearcher"
    $fileext = ".txt"
    $datetime = Get-Date -f "M-dd-yyyy HHmmss"
    $exportfile = "$DesktopPath\$filename $datetime$fileext" #dynamic desktop
    #$exportfile = "C:\Users\cjohn\Desktop\WordSearcher.csv" #hardcoded

    $ErrorActionPreference='Continue'
        <#
        Break: Enter the debugger when an error occurs or when an exception is raised.
        Continue: (Default) Displays the error message and continues executing.
        Ignore: Suppresses the error message and continues to execute the command. The Ignore value is intended for per-command use, not for use as saved preference. Ignore isn't a valid value for the $ErrorActionPreference variable.
        Inquire: Displays the error message and asks you whether you want to continue.
        SilentlyContinue: No effect. The error message isn't displayed and execution continues without interruption.
        Stop: Displays the error message and stops executing.
        Suspend: is only available for workflows which aren't supported in PowerShell 6 and beyond.
        #>

#endregion config


#region defaultvalues

    $getEverything = 0
    $Weekdays = "Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"
    if($Year -eq ""){$Year = 1980..2060}elseif($Year -notmatch '\b\d{4}\b'){write-host "Year needs to be 4 digits";exit}
    if($DayofWeek -eq ""){$DayofWeek = $Weekdays}elseif($($DayofWeek | Where-Object -FilterScript {$_ -in $Weekdays}).Count -ne $DayofWeek.count){write-host "Check `$DayofWeek spelling."; exit}
    $ext = if($ext -eq ""){$ext = "*.*"}else{$($ext | %{$_.insert(0,"*.")})}
    $FileNameNotLikeSB = if(!$FileNameNotLike){[scriptblock]{$_.Name -NotLike ""}}else{[scriptblock]{$_.Name -NotMatch $($FileNameNotLike -join "|")}}
    $FileNameLike = if(!$FileNameLike){".*"}else{$FileNameLike -join "|"}
    if($string -eq ""){$getEverything = 1}
    $dirs = $dirs | %{$_ + "\*"}

#endregion defaultvalues


#region FilterByFileProperties
    
    write-host "Script started at $starttimer"
    
    #Checks directory for files using your specified properties
    ForEach($dir in $dirs){
        dir $dir -Include $ext -Recurse:$r |
        Select-Object FullName,LastWriteTime,{$_.LastWriteTime.Year},{$_.LastWriteTime.DayOfWeek},DirectoryName,Name |
        Where-Object {$_.LastWriteTime.Year} -In $year |
        Where-Object {$_.LastWriteTime.DayOfWeek} -In $DayofWeek |
        Where-Object {$_.Name -match $FileNameLike} |
        Where-Object $FileNameNotLikeSB -OutVariable files |
        %{if($_.DirectoryName -eq $files[-2].DirectoryName){}else{write-host "Searching $($_.DirectoryName) | Total Items = $($(Get-ChildItem $_.DirectoryName | Measure-Object ).Count)"}}
    }

    write-host "`n$($files.Count) file(s) match file property criteria."
    
    #Exit if no files match properties
    If($files.Count -eq 0)
    {
       $endtimer = Get-date
       $time = (New-TimeSpan -Start $starttimer -End $endtimer).TotalSeconds
       write-host "`nScript took $time seconds to run.`n"
       exit
    }

#endregion FilterByFileProperties


#region parsefiles #Checks each file for string

    $outfile = @()
    $n = 0
    $i = 0
    $y = 100 / $files.Count
    write-host "`nParsing..."
    ForEach($file in $files){

        $i = $i + $y
        $n = $n + 1
        
        #write-host "$i of $($files.count) files parsed"
        Write-Progress -Activity "$n of $($files.count) files parsed" -Status "$([Math]::Round($i))% Complete" -PercentComplete $i

        if($getEverything -ne 1){
            $select = Select-String $string $file.FullName -SimpleMatch
        
            $temp = New-Object System.Object

            Add-Member -InputObject $temp -MemberType NoteProperty -Name Path -Value "" -Force
            Add-Member -InputObject $temp -MemberType NoteProperty -Name LineNumber -Value "" -Force
            Add-Member -InputObject $temp -MemberType NoteProperty -Name Line -Value "" -Force

            $temp.Path = ($($select | Select-Object -ExpandProperty Path) | Out-String).Trim()
            $temp.LineNumber = ($($select | Select-Object -ExpandProperty LineNumber) | Out-String).Trim()
            $temp.Line = ($($select | Select-Object -ExpandProperty Line) | Out-String).Trim()
        
        
            If($temp.Path -ne $null -and $temp.Path -ne "`r" -and $temp.Path -ne "`n" -and $temp.Path -ne ""){
                $outfile += $temp
                $outfile += "`n"
                $outfile += "`n"
                $outfile += "`n"
            }
        }else{
            $getContent = get-content $file.FullName
        
            $temp = New-Object System.Object

            Add-Member -InputObject $temp -MemberType NoteProperty -Name Path -Value "" -Force
            Add-Member -InputObject $temp -MemberType NoteProperty -Name LineNumber -Value "" -Force
            Add-Member -InputObject $temp -MemberType NoteProperty -Name Line -Value "" -Force

            $temp.Path = $file.FullName
            $temp.LineNumber = ""
            $temp.Line = "`n"+($getContent | Out-String).Trim()
        
        
            If($temp.Path -ne $null -and $temp.Path -ne "`r" -and $temp.Path -ne "`n" -and $temp.Path -ne ""){
                $outfile += $temp
                $outfile += "`n"
                $outfile += "`n"
                $outfile += "`n"
            }

        }
    }


#endregion parsefiles


#region CountMatches

    #Count the number of matches
    if($outfile.count -ne 0){
        if($outfile.count % 4 -eq 0){
             $outfileCnt = $($outfile.Count / 4)
        }else{
            $outfileCnt = $outfile.count
        }
    }

    #Display the number of matches
    if($getEverything -ne 1){
        write-host "`n$outfileCnt match(es) found."
    }else{
        write-host "$outfileCnt match(es) found."
        write-host "String was blank, so all file content was selected."
    }

#endregion CountMatches


#region AppendSearchCriteria

    $parameters =  [PSCustomObject]@{
        Title = $MyInvocation.MyCommand
        DateRan = $(Get-Date)
        FileNameNotLike = $FileNameNotLike
        FileNameLike = $FileNameLike
        DayofWeek = $DayofWeek
        year = $year
        ext = $ext
        string = $string
    }
    $parametersString = ($parameters | FL | Out-string).Trim()

#endregion AppendSearchCriteria

#region outfile
    
    #Prepend Title, Date and Search Criteria to File
    $outPut = @()
    $outPut += $parametersString
    $outPut += "`n"
    $outPut += $outfile | ConvertTo-Csv -NoTypeInformation -Delimiter ","
    
    #OutFile
    $outPut | Out-File $exportfile
    Write-host "$outfileCnt match(es) exported to $exportfile"

#endregion outfile


#region endtimer

   $endtimer = Get-date
   $time = (New-TimeSpan -Start $starttimer -End $endtimer).TotalSeconds
   write-host "`nScript took $time seconds to run.`n"

#endregion endtimer


#region openfile
    
    If($openfile -eq 1){
        If($($outfile.Count) -gt 0){
            $open = Read-Host "`nWould you like to open the results file? 1 for yes, 0 for no."
            If($open -eq 1){
                Invoke-Item $exportfile
            }elseif($open -eq 0){

            }else{write-host "`nYou pushed $open. I was expecting 1 or 0.`n";
                $open = Read-Host "Please confirm. Would you like to open the results file? 1 for yes, 0 for no."
                If($open -eq 1){
                    Invoke-Item $exportfile
                }else{}
            }    
        }
    }elseif($openfile -eq 2){
        Invoke-Item $exportfile
    }else{}

#endregion openfile