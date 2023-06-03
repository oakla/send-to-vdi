#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

#NoEnv

EnvGet, hdrive, Homedrive
EnvGet, hpath, Homepath
repoTransmissionFolder := hdrive . hpath . "\RepoTransmission"
bareRepoFolder := hdrive . hpath . "\code\bare-repos"

; NOTE: The s-drive path didn't work because it has spaces
; repoTransmissionFolder := "S:\Operational Excellence\Improvement\AutoBots Initiative\99._RemoteSync"

; paste to the transmission folder
^!r::
    OpenInFileExplorer(repoTransmissionFolder)
    Sleep 2000 ; Wait for the folder to load
    SendInput, ^v ; Simulate pressing Ctrl+V

    ; Sleep, 1000

    ; ;wait for a paste operation to complete
    ; while DllCall("GetOpenClipboardWindow", Ptr)
    ;     Sleep 50
    ; if WinActive("ahk_class XLMAIN")
    ;     while (A_Cursor = "Wait")
    ;         Sleep 50
    ; MsgBox, % "Done!"

    ; Wait for the window to become active again (i.e. wait for paste action to finish)
    WinWaitActive, ahk_class CabinetWClass

    Sleep 5000 ; Wait for the folder to load



    ; Get path of newly pasted zip file
    mostRecentlyModifiedFile := getMostRecent(repoTransmissionFolder)

    ; Unzip file 
    ; MsgBox, mostRecentlyModifiedFile is %mostRecentlyModifiedFile%
    unzip(mostRecentlyModifiedFile, bareRepoFolder)
    ; MsgBox, repoTransmissionFolder is %repoTransmissionFolder%
    ; Close active window
    WinClose, ahk_class CabinetWClass

    return

OpenInFileExplorer(targetFolder)
{
    ; Run the File Explorer command
    Run, explorer.exe
    Sleep 500 ; Wait for File Explorer window to open
    NavigateToFolder(targetFolder)
    return
}

NavigateToFolder(folderPath)
{
    ; ; Set the active window to File Explorer
    ; WinActivate, ahk_class CabinetWClass

    ; ; Wait for the window to become active
    ; WinWaitActive, ahk_class CabinetWClass

    Sleep 500
    SendInput, !d
    
    Sleep 500
    ; Send the folder path
    SendInput, %folderPath%{Enter}
    return
}

unzip(zipFile, targetFolder)
{
    ; MsgBox, Unzipping to %targetFolder%
    RunWait PowerShell.exe -Command Expand-Archive -Force -LiteralPath '%zipFile%' -DestinationPath %targetFolder%,, ; hide
    return
}

getMostRecent(targetFolder) {
    ; Retrieve files in a certain directory sorted by modification date:
    FileList :=  "" ; Initialize to be blank
    ; Create a list of those files consisting of the time the file was modified and the file path separated by tab
    ; Loop, %A_MyDocuments%\*.*
    ; MsgBox, A_UserName is %A_UserName%
    ; MsgBox, % "A_UserName is " A_UserName

    Loop, %targetFolder%\*.*
    FileList .= A_LoopFileTimeModified . "`t" . A_LoopFileLongPath . "`n"
    Sort, FileList, R  ;   ; Sort by time modified in reverse order
    Loop, Parse, FileList, `n
    {
        If (A_LoopField = "") ; omit the last linefeed (blank item) at the end of the list.
            Continue
        StringSplit, FileItem, A_LoopField, %A_Tab%  ; Split into two parts at the tab char
        ; FileItem1 is FileTimeModified und FileItem2 is FileName

        ; MsgBox, 36, Last modified file, %FileItem1% - %FileItem2%`n`nDo you want to continue?
        ; IfMsgBox, Yes
        ;     return FileItem2
        ; }    

        return FileItem2
    }
    return
}