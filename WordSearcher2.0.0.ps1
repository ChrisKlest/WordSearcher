#WordSearcher

#SEARCH FOR FILES ON YOUR PC OR EXTERNAL HD USING FILE PROPERTIES AND KEYWORD SEARCH FROM WITHIN THE FILE

#SCRIPT IS CONFIGURED IN A FORM DURING RUNTIME

cls 

function wrapText( $text, $delimiter, $width )
{
    $words = $text -split "\s+"
    $col = 0
    $NewText = @()
    foreach ( $word in $words )
    {
        $col += $word.Length + 1

        #newline using delimiter
        if($word -eq "|"){
            $word = "`n`n"
            $col = $word.Length + 1
            $newLine = 1
        }
        
        #newline using width
        if ( $col -gt $width )
        {
            #Write-Host ""
            $word = "`n$word"
            $col = $word.Length + 1
        }

        #add proper word
        if($newLine -ne 1)
        {
            $NewText = "$newText$word "
        }else{
            $NewText = "$newText $word"
        }
        $newLine=0

    }
    Return $NewText
}


#region formText

    $description="DESCRIPTION: Allows you to retrieve file content by filtering file content by search strings and by file property filters.
    This is a step up from Windows's file explorer search bar in that it will search and retrieve the content of a file. This 
    script also gives additional search functionality, such as filtering for last modified-by Day Name and by allowing you to
    search in multiple non-nested folders."

    $warning="WARNING: This script can run for considerable amounts of time, depending on your
    search criteria. For example, setting your starting directory as C:\ would search
    for all files in your C drive."

    $disclaimer="DISCLAIMER: The author of this script is unaware of any malicious code in this program. 
    The author of this script is not responsible for any damage it may cause to your computer 
    or any time or money wasted to recover from said damages."
    
    $formText="$description | $warning | $disclaimer" | Out-string

    $instructionsText="INSTRUCTIONS: Directories to Search and Save File are mandatory fields.
    The rest are optional. If the optional fields are left blank, default values
    will be provided. Do not put quotes around strings or numbers. Use
    commas to delimit multiple criteria in the same field. Directories to Search
    allows multiple directories. Save File Directory does not allow multiple paths.
    Be sure to verify the Search Subfolders checkbox.
    "

#endregion formText


Function LoadFormDescription{

    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms

#GET MONITOR SIZE
    $monitors = [System.Windows.Forms.Screen]::AllScreens
    $i=0
    ForEach($item in $monitors){
        Try{Add-Member -InputObject $monitors[$i] -MemberType NoteProperty -Name "ScreenArea" -Value $($($item.WorkingArea.size.Height) * $($item.WorkingArea.size.Width)) -EA Stop}Catch{}
        $i++
    }
    $max = ($monitors | measure -Property ScreenArea -Maximum).Maximum
    $formWidth = ($monitors | ? {$_.ScreenArea -eq $max}).WorkingArea.Width
    $formHeight = ($monitors | ? {$_.ScreenArea -eq $max}).WorkingArea.Height
    $formHeightWidthratio = [math]::Round($formHeight/$formWidth,1)
    $formAwayFromEdgeWidth = $formWidth/2
    $formAwayFromEdgeHeight = $formAwayFromEdgeWidth*$formHeightWidthratio

#CREATE FORM
    $UserForm = New-Object System.Windows.Forms.Form
    $UserForm.Text = "$($PSCommandPath)"
    $UserForm.Height = $formHeight - $formAwayFromEdgeHeight
    $UserForm.Width = $formWidth - $formAwayFromEdgeWidth
    $UserForm.StartPosition = "CenterScreen"
    $UserForm.FormBorderStyle = 4
    #$UserForm.BackgroundImage = [System.Drawing.Image]::FromFile("$PSScriptRoot\Image\imgButtonFuzzy.png") #Camo or button theme
    #$UserForm.BackgroundImageLayout = "Stretch" # None, Tile, Center, Stretch, Zoom
    $UserForm.Name = "UserForm"

#LABEL FORM
    $CreateFormLabel = New-Object System.Windows.Forms.Label
    $CreateFormLabel.AutoSize = $True
    $CreateFormLabel.Left = 40
    $CreateFormLabel.Top = 30
    $CreateFormLabel.Text = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))"
    $CreateFormLabel.Name = "CreateFormLabel"
    $CreateFormLabel.Font = New-Object System.Drawing.Font("Arial",20,[System.Drawing.FontStyle]::Bold)
    #$CreateFormLabel.backcolor = "White"
    $CreateFormLabel.forecolor = "Black"
    $CreateFormLabel.BackColor = "Transparent"
    $UserForm.Controls.Add($CreateFormLabel)

#TEXT FORM
    $CreateFormText = New-Object System.Windows.Forms.Label
    $CreateFormText.AutoSize = $True
    $CreateFormText.Left = 40
    $CreateFormText.Top = $CreateFormLabel.Top + 40
    $CreateFormText.Text = $(wrapText $formText "|" 100)
    $CreateFormText.Name = "CreateFormText"
    $CreateFormText.Font = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Bold)
    #$CreateFormText.backcolor = "White"
    $CreateFormText.forecolor = "Black"
    $CreateFormText.BackColor = "Transparent"
    $UserForm.Controls.Add($CreateFormText)

#PROCEED BUTTON
    $script:proceed=0
    $ProceedButton = New-Object System.Windows.Forms.Button
    $ProceedButton.Size = New-Object System.Drawing.Size(200,60) #W,H
    $ProceedButton.left = $UserForm.Width - ($UserForm.Width / 3)
    $ProceedButton.Top = $UserForm.Height - 150
    #$ProceedButton.Right = $($($UserForm.Width - $QuickStartbutton.Width)/2 - 8)
    #$ProceedButton.Bottom = $($AnchorTopLeft.Bottom + 39)
    $ProceedButton.Font = New-Object System.Drawing.Font("Arial",18,[System.Drawing.FontStyle]::Bold)
    $ProceedButton.Text = $("PROCEED")
    $ProceedButton.Name = "ProceedButton"
    $UserForm.Controls.Add($ProceedButton)
    $ProceedButton.Add_Click({
        #$script:user = $UserListBox.SelectedItem
        $script:Proceed = 1;
        $UserForm.Dispose()
    })

#EXIT BUTTON
    $script:exit=0
    $exitbutton = New-Object System.Windows.Forms.Button
    $exitbutton.Size = New-Object System.Drawing.Size(100,50) #W,H
    $exitbutton.Left = $ProceedButton.left - ($ProceedButton.Width - $exitbutton.Width + 20)
    $exitbutton.Top = $ProceedButton.Top + (($ProceedButton.Height - $exitbutton.Height)/2)
    $exitbutton.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
    $exitbutton.Text = "Exit Program"
    $exitbutton.Name = "exitbutton"
    [void]$UserForm.Controls.Add($exitbutton)
    [void]$exitbutton.Add_Click({
        $script:exit = 1;
        $UserForm.Dispose()
    })


#SHOW DIALOG
    $UserForm.Activate();
    $UserForm.ShowDialog() | Out-Null

#RETURN VALUES
    $attributes = @($script:exit,$script:proceed)
    Return $attributes
}


Function LoadFormConfig{

    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms

#GET MONITOR SIZE
    $monitors = [System.Windows.Forms.Screen]::AllScreens
    $i=0
    ForEach($item in $monitors){
        Try{Add-Member -InputObject $monitors[$i] -MemberType NoteProperty -Name "ScreenArea" -Value $($($item.WorkingArea.size.Height) * $($item.WorkingArea.size.Width)) -EA Stop}Catch{}
        $i++
    }
    $max = ($monitors | measure -Property ScreenArea -Maximum).Maximum
    $formWidth = ($monitors | ? {$_.ScreenArea -eq $max}).WorkingArea.Width
    $formHeight = ($monitors | ? {$_.ScreenArea -eq $max}).WorkingArea.Height
    $formHeightWidthratio = [math]::Round($formHeight/$formWidth,1)
    $formAwayFromEdgeWidth = $formWidth/2
    $formAwayFromEdgeHeight = $formAwayFromEdgeWidth*$formHeightWidthratio

#CREATE FORM
    $UserForm = New-Object System.Windows.Forms.Form
    $UserForm.Text = "$($PSCommandPath)"
    $UserForm.Height = $formHeight - $formAwayFromEdgeHeight
    $UserForm.Width = $formWidth - $formAwayFromEdgeWidth
    $UserForm.StartPosition = "CenterScreen"
    $UserForm.FormBorderStyle = 4
    #$UserForm.BackgroundImage = [System.Drawing.Image]::FromFile("$PSScriptRoot\Image\imgButtonFuzzy.png") #Camo or button theme
    #$UserForm.BackgroundImageLayout = "Stretch" # None, Tile, Center, Stretch, Zoom
    $UserForm.Name = "UserForm"

#LABEL FORM
    $CreateFormLabel = New-Object System.Windows.Forms.Label
    $CreateFormLabel.AutoSize = $True
    $CreateFormLabel.Left = 40
    $CreateFormLabel.Top = 30
    $CreateFormLabel.Text = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))"
    $CreateFormLabel.Name = "CreateFormLabel"
    $CreateFormLabel.Font = New-Object System.Drawing.Font("Arial",20,[System.Drawing.FontStyle]::Bold)
    #$CreateFormLabel.backcolor = "White"
    $CreateFormLabel.forecolor = "Black"
    $CreateFormLabel.BackColor = "Transparent"
    $UserForm.Controls.Add($CreateFormLabel)

#CREATE DIRECTORY LABEL
    $CreateDirectoryLabel = New-Object System.Windows.Forms.Label
    $CreateDirectoryLabel.AutoSize = $True
    $CreateDirectoryLabel.Left = $($CreateFormLabel.Left)
    $CreateDirectoryLabel.Top = $($CreateFormLabel.Top + $CreateFormLabel.Height + 10)
    $CreateDirectoryLabel.Text = "Directories to Search. e.g. C:\Users\cjohn\Desktop"
    $CreateDirectoryLabel.Name = "CreateDirectoryLabel"
    $CreateDirectoryLabel.Font = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Bold)
    #$CreateDirectoryLabel.backcolor = "White"
    $CreateDirectoryLabel.forecolor = "Black"
    $CreateDirectoryLabel.BackColor = "Transparent"
    $UserForm.Controls.Add($CreateDirectoryLabel)

#CREATE DIRECTORY TEXTBOX
    $CreateDirectoryTextBox = New-Object System.Windows.Forms.TextBox
    $CreateDirectoryTextBox.Size = New-Object System.Drawing.Size($($UserForm.Width/2),20)#W,H
    $CreateDirectoryTextBox.Top = $CreateDirectoryLabel.Bottom
    $CreateDirectoryTextBox.Left = $($CreateDirectoryLabel.Left + 3)
    $CreateDirectoryTextBox.Text = "$([Environment]::GetFolderPath("Desktop"))"
    $CreateDirectoryTextBox.Name = "CreateDirectoryTextBox"
    $CreateDirectoryTextBox.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
    $CreateDirectoryTextBox.Add_Gotfocus( { $this.SelectAll(); $this.Focus() })
    #$CreateDirectoryTextBox.Add_Click( { $this.SelectAll(); $this.Focus() })
    $UserForm.Controls.Add($CreateDirectoryTextBox)

#CREATE RECURSIVE CHECKBOX
    $RecursiveCheckbox = New-Object System.Windows.Forms.Checkbox 
    $RecursiveCheckbox.AutoSize = $false
    $RecursiveCheckbox.Size = New-Object System.Drawing.Size(160,50)#W,H
    $RecursiveCheckbox.Top = $($CreateDirectoryTextBox.Top -20)
    $RecursiveCheckbox.Left = $($CreateDirectoryTextBox.Right + 20)
    $RecursiveCheckbox.Text = "Search Subfolders?"
    $RecursiveCheckbox.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
    $RecursiveCheckbox.Name = "RecursiveCheckbox"
    $RecursiveCheckbox.forecolor = "Black"
    $RecursiveCheckbox.TabIndex = 1
    $UserForm.Controls.Add($RecursiveCheckbox)
    $RecursiveCheckbox.Checked = $false

#CREATE YEAR LABEL
    $CreateYearLabel = New-Object System.Windows.Forms.Label
    $CreateYearLabel.AutoSize = $True
    $CreateYearLabel.Left = $($CreateDirectoryLabel.Left)
    $CreateYearLabel.Top = $($CreateDirectoryTextBox.Bottom + 10)
    $CreateYearLabel.Text = "Last modified Year. e.g. 2008,2009"
    $CreateYearLabel.Name = "CreateYearLabel"
    $CreateYearLabel.Font = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Bold)
    #$CreateYearLabel.backcolor = "White"
    $CreateYearLabel.forecolor = "Black"
    $CreateYearLabel.BackColor = "Transparent"
    $UserForm.Controls.Add($CreateYearLabel)

#CREATE YEAR TEXTBOX
    $CreateYearTextBox = New-Object System.Windows.Forms.TextBox
    $CreateYearTextBox.Size = New-Object System.Drawing.Size($($UserForm.Width/2),20)#W,H
    $CreateYearTextBox.Top = $CreateYearLabel.Bottom
    $CreateYearTextBox.Left = $($CreateYearLabel.Left + 3)
    $CreateYearTextBox.Text = ""
    $CreateYearTextBox.Name = "CreateYearTextBox"
    $CreateYearTextBox.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
    $CreateYearTextBox.Add_Gotfocus( { $this.SelectAll(); $this.Focus() })
    #$CreateYearTextBox.Add_Click( { $this.SelectAll(); $this.Focus() })
    $UserForm.Controls.Add($CreateYearTextBox)

#CREATE DAY OF WEEK LABEL
    $CreateDayOfWeekLabel = New-Object System.Windows.Forms.Label
    $CreateDayOfWeekLabel.AutoSize = $True
    $CreateDayOfWeekLabel.Left = $($CreateDirectoryLabel.Left)
    $CreateDayOfWeekLabel.Top = $($CreateYearTextBox.Bottom + 10)
    $CreateDayOfWeekLabel.Text = "Last modified DayName. e.g. Monday,Tuesday"
    $CreateDayOfWeekLabel.Name = "CreateDayOfWeekLabel"
    $CreateDayOfWeekLabel.Font = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Bold)
    #$CreateDayOfWeekLabel.backcolor = "White"
    $CreateDayOfWeekLabel.forecolor = "Black"
    $CreateDayOfWeekLabel.BackColor = "Transparent"
    $UserForm.Controls.Add($CreateDayOfWeekLabel)

#CREATE DAY OF WEEK TEXTBOX
    $CreateDayOfWeekTextBox = New-Object System.Windows.Forms.TextBox
    $CreateDayOfWeekTextBox.Size = New-Object System.Drawing.Size($($UserForm.Width/2),20)#W,H
    $CreateDayOfWeekTextBox.Top = $CreateDayOfWeekLabel.Bottom
    $CreateDayOfWeekTextBox.Left = $($CreateYearLabel.Left + 3)
    $CreateDayOfWeekTextBox.Text = ""
    $CreateDayOfWeekTextBox.Name = "CreateDayOfWeekTextBox"
    $CreateDayOfWeekTextBox.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
    $CreateDayOfWeekTextBox.Add_Gotfocus( { $this.SelectAll(); $this.Focus() })
    #$CreateDayOfWeekTextBox.Add_Click( { $this.SelectAll(); $this.Focus() })
    $UserForm.Controls.Add($CreateDayOfWeekTextBox)

#CREATE EXT LABEL
    $CreateExtLabel = New-Object System.Windows.Forms.Label
    $CreateExtLabel.AutoSize = $True
    $CreateExtLabel.Left = $($CreateDirectoryLabel.Left)
    $CreateExtLabel.Top = $($CreateDayOfWeekTextBox.Bottom + 10)
    $CreateExtLabel.Text = "File extensions. e.g. txt,csv"
    $CreateExtLabel.Name = "CreateExtLabel"
    $CreateExtLabel.Font = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Bold)
    #$CreateExtLabel.backcolor = "White"
    $CreateExtLabel.forecolor = "Black"
    $CreateExtLabel.BackColor = "Transparent"
    $UserForm.Controls.Add($CreateExtLabel)

#CREATE EXT TEXTBOX
    $CreateExtTextBox = New-Object System.Windows.Forms.TextBox
    $CreateExtTextBox.Size = New-Object System.Drawing.Size($($UserForm.Width/2),20)#W,H
    $CreateExtTextBox.Top = $CreateExtLabel.Bottom
    $CreateExtTextBox.Left = $($CreateExtLabel.Left + 3)
    $CreateExtTextBox.Text = ""
    $CreateExtTextBox.Name = "CreateDayOfWeekTextBox"
    $CreateExtTextBox.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
    $CreateExtTextBox.Add_Gotfocus( { $this.SelectAll(); $this.Focus() })
    #$CreateExtTextBox.Add_Click( { $this.SelectAll(); $this.Focus() })
    $UserForm.Controls.Add($CreateExtTextBox)

#CREATE FILE NAME NOT LIKE LABEL
    $CreateFileNameNotLikeLabel = New-Object System.Windows.Forms.Label
    $CreateFileNameNotLikeLabel.AutoSize = $True
    $CreateFileNameNotLikeLabel.Left = $($CreateDirectoryLabel.Left)
    $CreateFileNameNotLikeLabel.Top = $($CreateExtTextBox.Bottom + 10)
    $CreateFileNameNotLikeLabel.Text = "Filename does not contain. e.g. apples,oranges"
    $CreateFileNameNotLikeLabel.Name = "CreateFileNameNotLikeLabel"
    $CreateFileNameNotLikeLabel.Font = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Bold)
    #$CreateFileNameNotLikeLabel.backcolor = "White"
    $CreateFileNameNotLikeLabel.forecolor = "Black"
    $CreateFileNameNotLikeLabel.BackColor = "Transparent"
    $UserForm.Controls.Add($CreateFileNameNotLikeLabel)

#CREATE FILE NAME NOT LIKE TEXTBOX
    $CreateFileNameNotLikeTextBox = New-Object System.Windows.Forms.TextBox
    $CreateFileNameNotLikeTextBox.Size = New-Object System.Drawing.Size($($UserForm.Width/2),20)#W,H
    $CreateFileNameNotLikeTextBox.Top = $CreateFileNameNotLikeLabel.Bottom
    $CreateFileNameNotLikeTextBox.Left = $($CreateFileNameNotLikeLabel.Left + 3)
    $CreateFileNameNotLikeTextBox.Text = ""
    $CreateFileNameNotLikeTextBox.Name = "CreateFileNameNotLikeTextBox"
    $CreateFileNameNotLikeTextBox.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
    $CreateFileNameNotLikeTextBox.Add_Gotfocus( { $this.SelectAll(); $this.Focus() })
    #$CreateFileNameNotLikeTextBox.Add_Click( { $this.SelectAll(); $this.Focus() })
    $UserForm.Controls.Add($CreateFileNameNotLikeTextBox)

#CREATE FILE NAME LIKE LABEL
    $CreateFileNameLikeLabel = New-Object System.Windows.Forms.Label
    $CreateFileNameLikeLabel.AutoSize = $True
    $CreateFileNameLikeLabel.Left = $($CreateDirectoryLabel.Left)
    $CreateFileNameLikeLabel.Top = $($CreateFileNameNotLikeTextBox.Bottom + 10)
    $CreateFileNameLikeLabel.Text = "Filename contains. e.g. bananas,turnips"
    $CreateFileNameLikeLabel.Name = "CreateFileNameLikeLabel"
    $CreateFileNameLikeLabel.Font = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Bold)
    #$CreateFileNameLikeLabel.backcolor = "White"
    $CreateFileNameLikeLabel.forecolor = "Black"
    $CreateFileNameLikeLabel.BackColor = "Transparent"
    $UserForm.Controls.Add($CreateFileNameLikeLabel)

#CREATE FILE NAME LIKE TEXTBOX
    $CreateFileNameLikeTextBox = New-Object System.Windows.Forms.TextBox
    $CreateFileNameLikeTextBox.Size = New-Object System.Drawing.Size($($UserForm.Width/2),20)#W,H
    $CreateFileNameLikeTextBox.Top = $CreateFileNameLikeLabel.Bottom
    $CreateFileNameLikeTextBox.Left = $($CreateFileNameLikeLabel.Left + 3)
    $CreateFileNameLikeTextBox.Text = ""
    $CreateFileNameLikeTextBox.Name = "CreateFileNameLikeTextBox"
    $CreateFileNameLikeTextBox.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
    $CreateFileNameLikeTextBox.Add_Gotfocus( { $this.SelectAll(); $this.Focus() })
    #$CreateFileNameLikeTextBox.Add_Click( { $this.SelectAll(); $this.Focus() })
    $UserForm.Controls.Add($CreateFileNameLikeTextBox)

#CREATE CONTENT SEARCH LABEL
    $CreateContentSearchLabel = New-Object System.Windows.Forms.Label
    $CreateContentSearchLabel.AutoSize = $True
    $CreateContentSearchLabel.Left = $($CreateDirectoryLabel.Left)
    $CreateContentSearchLabel.Top = $($CreateFileNameLikeTextBox.Bottom + 10)
    $CreateContentSearchLabel.Text = "Keywords in Content of File. e.g. bundles,baskets"
    $CreateContentSearchLabel.Name = "CreateContentSearchLabel"
    $CreateContentSearchLabel.Font = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Bold)
    #$CreateContentSearchLabel.backcolor = "White"
    $CreateContentSearchLabel.forecolor = "Black"
    $CreateContentSearchLabel.BackColor = "Transparent"
    $UserForm.Controls.Add($CreateContentSearchLabel)

#CREATE CONTENT SEARCH TEXTBOX
    $CreateContentSearchTextBox = New-Object System.Windows.Forms.TextBox
    $CreateContentSearchTextBox.Size = New-Object System.Drawing.Size($($UserForm.Width/2),20)#W,H
    $CreateContentSearchTextBox.Top = $CreateContentSearchLabel.Bottom
    $CreateContentSearchTextBox.Left = $($CreateContentSearchLabel.Left + 3)
    $CreateContentSearchTextBox.Text = ""
    $CreateContentSearchTextBox.Name = "CreateContentSearchTextBox"
    $CreateContentSearchTextBox.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
    $CreateContentSearchTextBox.Add_Gotfocus( { $this.SelectAll(); $this.Focus() })
    #$CreateContentSearchTextBox.Add_Click( { $this.SelectAll(); $this.Focus() })
    $UserForm.Controls.Add($CreateContentSearchTextBox)

#CREATE DESTINATION File LABEL
    $CreateDestinationFileLabel = New-Object System.Windows.Forms.Label
    $CreateDestinationFileLabel.AutoSize = $True
    $CreateDestinationFileLabel.Left = $($CreateDirectoryLabel.Left)
    $CreateDestinationFileLabel.Top = $($CreateContentSearchTextBox.Bottom + 10)
    $CreateDestinationFileLabel.Text = "Save File Directory. If blank, will save to desktop."
    $CreateDestinationFileLabel.Name = "CreateDestinationFileLabel"
    $CreateDestinationFileLabel.Font = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Bold)
    #$CreateDestinationFileLabel.backcolor = "White"
    $CreateDestinationFileLabel.forecolor = "Black"
    $CreateDestinationFileLabel.BackColor = "Transparent"
    $UserForm.Controls.Add($CreateDestinationFileLabel)

#CREATE DESTINATION File TEXTBOX
    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    $filename = "WordSearcher"
    $fileext = ".txt"
    $datetime = Get-Date -f "M-dd-yyyy HHmmss"
    $exportfile = "$DesktopPath\$filename $datetime$fileext" #dynamic desktop
    $CreateDestinationFileTextBox = New-Object System.Windows.Forms.TextBox
    $CreateDestinationFileTextBox.Size = New-Object System.Drawing.Size($($UserForm.Width/2),20)#W,H
    $CreateDestinationFileTextBox.Top = $CreateDestinationFileLabel.Bottom
    $CreateDestinationFileTextBox.Left = $($CreateDestinationFileLabel.Left + 3)
    $CreateDestinationFileTextBox.Text = "$exportfile"
    $CreateDestinationFileTextBox.Name = "CreateDestinationFileTextBox"
    $CreateDestinationFileTextBox.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
    $CreateDestinationFileTextBox.Add_Gotfocus( { $this.SelectAll(); $this.Focus() })
    #$CreateDestinationFileTextBox.Add_Click( { $this.SelectAll(); $this.Focus() })
    $UserForm.Controls.Add($CreateDestinationFileTextBox)

#CREATE INSTRUCTIONS LABEL
    $CreateInstructionsLabel = New-Object System.Windows.Forms.Label
    $CreateInstructionsLabel.AutoSize = $True
    $CreateInstructionsLabel.Left = $($CreateDestinationFileTextBox.Right + 20)
    $CreateInstructionsLabel.Top = $($CreateYearTextBox.Top)
    $CreateInstructionsLabel.Text = $(wrapText $instructionsText $null 40)
    $CreateInstructionsLabel.Name = "CreateInstructionsLabel"
    $CreateInstructionsLabel.Font = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Bold)
    #$CreateInstructionsLabel.backcolor = "White"
    $CreateInstructionsLabel.forecolor = "Black"
    $CreateInstructionsLabel.BackColor = "Transparent"
    $UserForm.Controls.Add($CreateInstructionsLabel)

#PROCEED BUTTON
    $script:proceed=0
    $ProceedButton = New-Object System.Windows.Forms.Button
    $ProceedButton.Size = New-Object System.Drawing.Size(200,60) #W,H
    $ProceedButton.left = $UserForm.Width - ($UserForm.Width / 1.5)
    $ProceedButton.Top = $CreateDestinationFileTextBox.Bottom + 20
    #$ProceedButton.Right = $($($UserForm.Width - $QuickStartbutton.Width)/2 - 8)
    #$ProceedButton.Bottom = $($AnchorTopLeft.Bottom + 39)
    $ProceedButton.Font = New-Object System.Drawing.Font("Arial",18,[System.Drawing.FontStyle]::Bold)
    $ProceedButton.Text = $("PROCEED")
    $ProceedButton.Name = "ProceedButton"
    $UserForm.Controls.Add($ProceedButton)
    $ProceedButton.Add_Click({
        $script:directories     = $CreateDirectoryTextBox.Text
        $script:years           = $CreateYearTextBox.Text
        $script:dayOfWeek       = $CreateDayOfWeekTextBox.Text
        $script:ext             = $CreateExtTextBox.Text
        $script:fileNameNotLike = $CreateFileNameNotLikeTextBox.Text
        $script:fileNameLike    = $CreateFileNameLikeTextBox.Text
        $script:contentSearch   = $CreateContentSearchTextBox.Text
        $script:saveFile        = $CreateDestinationFileTextBox.Text
        $script:recursiveSearch = $RecursiveCheckbox.Checked
        $script:proceed = 1;
        $UserForm.Dispose()
    })

#EXIT BUTTON
    $script:exit=0
    $exitbutton = New-Object System.Windows.Forms.Button
    $exitbutton.Size = New-Object System.Drawing.Size(100,50) #W,H
    $exitbutton.Left = $ProceedButton.left - ($ProceedButton.Width - $exitbutton.Width + 20)
    $exitbutton.Top = $ProceedButton.Top + (($ProceedButton.Height - $exitbutton.Height)/2)
    $exitbutton.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
    $exitbutton.Text = "Exit Program"
    $exitbutton.Name = "exitbutton"
    [void]$UserForm.Controls.Add($exitbutton)
    [void]$exitbutton.Add_Click({
        $script:exit = 1;
        $UserForm.Dispose()
    })

#MINIMUM SIZE FORM
    $minHeight = 
    $CreateFormLabel.Height +
    $CreateDirectoryLabel.Height +
    $CreateDirectoryTextBox.Height +
    $RecursiveCheckbox.Height +
    $CreateYearLabel.Height +
    $CreateYearTextBox.Height +
    $CreateDayOfWeekLabel.Height +
    $CreateDayOfWeekTextBox.Height +
    $CreateExtLabel.Height +
    $CreateExtTextBox.Height +
    $CreateFileNameNotLikeLabel.Height +
    $CreateFileNameNotLikeTextBox.Height +
    $CreateFileNameLikeLabel.Height +
    $CreateFileNameLikeTextBox.Height +
    $CreateContentSearchLabel.Height +
    $CreateContentSearchTextBox.Height +
    $CreateDestinationFileLabel.Height +
    $CreateDestinationFileTextBox.Height +
    $ProceedButton.Height +
    $exitbutton.Height +
    110

    $minWidth = $CreateContentSearchLabel.Width

    $UserForm.minimumSize = New-Object System.Drawing.Size($($minWidth,$minHeight)) #W,H

#SHOW DIALOG
    $UserForm.Activate();
    $UserForm.ShowDialog() | Out-Null

#RETURN VALUES
    $attributes = @(
        $script:exit,
        $script:proceed,
        $script:directories,
        $script:years,
        $script:dayOfWeek,
        $script:ext,
        $script:fileNameNotLike,
        $script:fileNameLike,
        $script:contentSearch
        $script:saveFile
        $script:recursiveSearch

    )
    Return $attributes
}


#region callForm
    #CALL FORM
    $null = $outDescription = LoadFormDescription
    If($outDescription[0] -eq 1){write-host "User Clicked Exit";[System.GC]::Collect();exit}
    $null = $outConfig = LoadFormConfig
    If($outConfig[0] -eq 1){"User Clicked Exit";[System.GC]::Collect();exit}
#endregion callForm


#region starttimer

   $starttimer = Get-date

#endregion starttimer


#region config
    
    #Search This Folder
    $dirs = $outConfig[2].Split(",")
        # Full Folder path in quotes. This can be a single folder, or multiple folders separated by commas.
        # e.g. "F:\TextmessagesBackup\CombinedAllTextMessages", "C:\UserName\Desktop" 

    #Search through subfolders? aka Recursive Search
    $r = $outConfig[10]
        # $true or $false. Note, setting to $true can take a long time if searching through all subfolders.

    #Filter for these File Properties
    $year            = $outConfig[3].Split(",")
    $DayofWeek       = $outConfig[4].Split(",")
    $ext             = $outConfig[5].Split(",")
    $FileNameNotLike = $outConfig[6].Split(",")
    $FileNameLike    = $outConfig[7].Split(",")
        # $year           = Last Write Time.           Double Quotes if no filter (zero length string) "", else 4 digit year separated by comma. e.g. 2012, 2013, 2014, 2015, 2016, 2017
        # $DayofWeek      = Last Write Time.           Double Quotes if no filter (zero length string) "", else full day name in quotes and separated by comma. e.g. "Monday","Tuesday"
        # $ext            = File Extension.            Double Quotes if no filter (zero length string) "", else file extension in quotes and separated by comma. e.g. "xlsx","csv","docx","xlsm","pdf"
        # FileNameNotLike = File Name not Like string. Double Quotes if no filter (zero length string) "", else strings in quotes separated by comma. e.g. "nameNotLike","oranges","bananas"
        # FileNameLike    = File Name like string.     Double Quotes if no filter (zero length string) "", else strings in quotes separated by comma. e.g. "nameLike","mangos","fruit"

    #File Content Must Contain
    $string = $outConfig[8].Split(",")
        # Double Quotes if no filter (zero length string) "", else String or Array of Strings in quotes and separated by commas.
        # e.g. "on my way", "be there in", "should be about", "omw", "work?"

    #Open results file when done?
    $openfile = 2
        # 2 Always open results file.
        # 1 Ask to open.
        # 0 Never open.
    
    #Results File Properties  
    if($($outConfig[9]) -eq "" -or $([string]::IsNullOrEmpty($outConfig[9]))){
        $DesktopPath = [Environment]::GetFolderPath("Desktop")
        $filename = "WordSearcher"
        $fileext = ".txt"
        $datetime = Get-Date -f "M-dd-yyyy HHmmss"
        $exportfile = "$DesktopPath\$filename $datetime$fileext" #dynamic desktop
    }else{
        $exportfile = $outConfig[9]
    }

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
    if($Year -eq "" -or [string]::IsNullOrEmpty($Year)){$Year = 1980..2060}elseif($Year -notmatch '\b\d{4}\b'){write-host "Year needs to be 4 digits";[System.GC]::Collect();exit}
    if($DayofWeek -eq "" -or [string]::IsNullOrEmpty($DayofWeek)){$DayofWeek = $Weekdays}elseif($($DayofWeek | Where-Object -FilterScript {$_ -in $Weekdays}).Count -ne $DayofWeek.count){write-host "Check `$DayofWeek spelling.";[System.GC]::Collect();exit}
    $ext = if(!$ext){"*.*"}else{$($ext | %{$_.insert(0,"*.")})}
    $FileNameNotLikeSB = if(!$FileNameNotLike){[scriptblock]{$_.Name -NotLike ""}}else{[scriptblock]{$_.Name -NotMatch $($FileNameNotLike -join "|")}}
    $FileNameLike = if(!$FileNameLike){".*"}else{$FileNameLike -join "|"}
    if($string -eq "" -or [string]::IsNullOrEmpty($string)){$getEverything = 1}
    $dirs = $dirs | %{$_ + "\*"}

#endregion defaultvalues


#region FilterByFileProperties
    
    write-host "Script started at $starttimer"
    
    #Checks directory for files using your specified properties
    $Collection=@()
    ForEach($dir in $dirs){
        [System.GC]::Collect()
        dir $dir -Include $ext -Recurse:$r |
        Select-Object FullName,LastWriteTime,{$_.LastWriteTime.Year},{$_.LastWriteTime.DayOfWeek},DirectoryName,Name |
        Where-Object {$_.LastWriteTime.Year} -In $year |
        Where-Object {$_.LastWriteTime.DayOfWeek} -In $DayofWeek |
        Where-Object {$_.Name -match $FileNameLike} |
        Where-Object $FileNameNotLikeSB -OutVariable files |
        %{if($_.DirectoryName -eq $files[-2].DirectoryName){}else{write-host "Searching $($_.DirectoryName) | Total Items = $($(Get-ChildItem $_.DirectoryName | Measure-Object ).Count)"}}
        $collection += $files
    }
    $files=$collection

    write-host "`n$($files.Count) file(s) match file property criteria."
    
    #Exit if no files match properties
    If($files.Count -eq 0)
    {
       $endtimer = Get-date
       $time = (New-TimeSpan -Start $starttimer -End $endtimer).TotalSeconds
       write-host "`nScript took $time seconds to run.`n"
       write-host "Exiting since 0 files were matched."
       [System.GC]::Collect()
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
        
        Write-Progress -Activity "$n of $($files.count) files parsed" -Status "$([Math]::Round($i))% Complete" -PercentComplete $i

        if($getEverything -ne 1){
            $skip=0
            try{$select = Select-String $string $file.FullName -SimpleMatch -EA Stop}catch{$skip=1}
            
            #if($skip -eq 1){
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
            #}
        }else{
            $skip=0
            Try{$getContent = get-content $file.FullName -EA Stop}Catch{$skip=1}
            #if($skip -eq 1){
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
            #}
        }
    }


#endregion parsefiles


#region CountMatches
    
    #Count the number of matches
    $outfilecnt=0
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
        FileCount = $($files.Count)
        Matches = $outfileCnt
        Directories = $dirs
        Recurse = $r
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

    #Exit script if 0 matches found
    if($outfile.count -eq 0){
        write-host "`n"
        write-host $parametersString
        write-host "`nNot exporting to file due to 0 matches found."
        $endtimer = Get-date
        $time = (New-TimeSpan -Start $starttimer -End $endtimer).TotalSeconds
        write-host "`nScript took $time seconds to run."
        write-host "`nExiting"
        exit
    }
    
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

[System.GC]::Collect()