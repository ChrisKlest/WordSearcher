#STRINGSEARCHER

# USED TO PERFORM RECURSIVE LOOKUP OF KEY WORDS IN SPECIFIED FOLDER DIRECTORY AND FILE OR SEARCH BY FILETYPE

# WARNING: RECURSIVE LOOKUP WILL LOOK IN EACH FOLDER WITHIN THE FOLDER YOU SPECIFY, THIS CAN TAKE CONSIDERABLE AMOUNTS OF TIME DEPENDING ON YOUR STARTING LOCATION.

# Once you have configured the script to your liking, remove the leading # and press play.

#TO UPDATE: Line 137 Select-String. Create and/or options
#TO UPDATE: Create Form for search parameter entry


cls


#region config
    
    #Look In. #Note, this can be 1 element or multiple elements. Just be sure the directory is in quotes and has a \* at the end of it. If multiple elements, separate by comma.
    $dirs = "F:\TextmessagesBackup\CombinedAllTextMessages\*" #, "2nd directory"

    #Recursive #If you have many subfolders, but you only want to search the single parent folder, than set to $false
    $r = $true #set $true or $false

    #File Properties Filter #If these are left as 0 length strings "", it will not filter by file property.
    $year = "" #2012, 2013, 2014, 2015, 2016, 2017
    $DayofWeek = "" #"Monday", "Tuesday"
    $ext = "" #"*.xlsx", "*.csv" #, "*.docx", "*.xlsx", "*.xlsm", "*.pdf"
    $FileNameNotLike = ""
    $FileNameLike = "2016-09", "2016-10" #, "2016-08", "2016-10"

    #Content like # Can be a single string in quotes, or multiple strings in quotes, separated by commas
    $string = "on my way", "be there in", "should be about", "omw", "work?"

    #Open results file when done?
    $openfile = 2
        # 2 Always open results file.
        # 1 Ask to open.
        # 0 Never open.
    
    #Results file; save name, type and location  
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
    if($ext -eq ""){$ext = "*.*"}
    if($FileNameNotLike -eq ""){$FileNameNotLike = ""}elseif($FileNameNotLike -ne ""){$FileNameNotLike = "*$FileNameNotLike*"}
    if($FileNameLike -eq ""){$FileNameLike = "*"}elseif($FileNameLike -ne ""){$FileNameLike = $FileNameLike | %{ $_ -replace $_,"*$_*"}} #{$FileNameLike = "*$FileNameLike*"}
    if($string -eq ""){$getEverything = 1}

#endregion defaultvalues


$parameters =  [PSCustomObject]@{
    FileNameNotLike = $FileNameNotLike
    FileNameLike = $FileNameLike
    DayofWeek = $DayofWeek
    year = $year
    ext = $ext
    string = $string
}
$parametersString = ($parameters | FL | Out-string).Trim()


#region starttimer

   $starttimer = Get-date

#endregion starttimer


#region getfiles #Checks directory for files using your specified properties
    
    write-host "Script started at $starttimer"
    ForEach($dir in $dirs){
        #write-host "`nSearching $dir | Total Items = $($(Get-ChildItem $dir | Measure-Object ).Count)"
        dir $dir -Include $ext -Recurse:$r |
        Select-Object FullName,LastWriteTime,{$_.LastWriteTime.Year},{$_.LastWriteTime.DayOfWeek}, DirectoryName |
        Where-Object {$_.LastWriteTime.Year} -In $year |
        Where-Object {$_.LastWriteTime.DayOfWeek} -In $DayofWeek |
        Where-Object FullName -like $FileNameLike |
        Where-Object FullName -notlike $FileNameNotLike -OutVariable files |
        %{if($_.DirectoryName -eq $files[-2].DirectoryName){}else{write-host "Searching $($_.DirectoryName) | Total Items = $($(Get-ChildItem $_.DirectoryName | Measure-Object ).Count)"}}
    }

#endregion getfiles


write-host "`n$($files.Count) file(s) match file property criteria."
If($files.Count -eq 0){write-host "Exiting.";exit}


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

#region outfile

    $outfile | ConvertTo-Csv -NoTypeInformation -Delimiter "," | Out-File $exportfile
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