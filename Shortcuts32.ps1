$programs = @{
    "Access"                   = "C:\C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE"
    "Adobe Acrobat"            = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    "Adobe Reader DC"          = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    "Excel"                    = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    "Firefox Private Browsing" = "C:\Program Files\Mozilla Firefox\private_browsing.exe"
    "Firefox"                  = "C:\Program Files\Mozilla Firefox\firefox.exe"
    "Google Chrome"            = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    "Microsoft Edge"           = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    "Notepad++"                = "C:\Program Files\Notepad++\notepad++.exe"
    "OneNote"                  = "C:\Program Files (x86)\Microsoft Office\root\Office16\ONENOTE.EXE"
    "Outlook"                  = "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
    "PowerPoint"               = "C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"
    "Project"                  = "C:\Program Files (x86)\Microsoft Office\root\Office16\WINPROJ.EXE"
    "Publisher"                = "C:\Program Files (x86)\Microsoft Office\root\Office16\MSPUB.EXE"
    "Remote Desktop"           = "C:\Program Files\Remote Desktop\msrdcw.exe"
    "TeamViewer"               = "C:\Program Files\TeamViewer\TeamViewer.exe"
    "Visio"                    = "C:\Program Files (x86)\Microsoft Office\root\Office16\VISIO.EXE"
    "Word"                     = "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.exe"
    "Cisco AnyConnect"         = "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"
    "Cisco Jabber"             = "C:\Program Files (x86)\Cisco Systems\Cisco Jabber\CiscoJabber.exe"
    "Citrix Workspace"         = "C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\SelfService.exe"

    
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