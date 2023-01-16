
$programs = @{
    "Adobe Acrobat"            = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    "Excel"                    = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
    "Google Chrome"            = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    "OneNote"                  = "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE"
    "Outlook"                  = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
    "PowerPoint"               = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
    "Project"                  = "C:\Program Files\Microsoft Office\root\Office16\WINPROJ.EXE"
    "Visio"                    = "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE"
    "Word"                     = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.exe"
}

 

#Check for shortcuts in Start Menu, if program is available and the shortcut isn't... Then recreate the shortcut
$programs.GetEnumerator() | ForEach-Object {
    if (Test-Path -Path $_.Value) {
        if (-not (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\$($_.Key).lnk")) {
            write-host ("Shortcut for {0} not found in {1}, creating it now..." -f $_.Key, $_.Value)
            $shortcut = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\$($_.Key).lnk"
            $target = $_.Value
            $description = $_.Key
            $workingdirectory = (Get-ChildItem $target).DirectoryName
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($shortcut)
            $Shortcut.TargetPath = $target
            $Shortcut.Description = $description
            $shortcut.WorkingDirectory = $workingdirectory
            $Shortcut.Save()
        }
    }
}