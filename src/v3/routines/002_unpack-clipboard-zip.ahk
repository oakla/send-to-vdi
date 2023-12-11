#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%



unpack_clipboard_zip(to_parent_folder) {
    ; homePath := A_Home ; alternative homepath access method
    ; MsgBox, Target parent for unzipped file is %to_parent_folder%
    EnvGet, hdrive, Homedrive
    EnvGet, hpath, Homepath
    ; Specify the location to send newly created zip files before they get sent to the VDI
    intermediateReceivingDirectory := hdrive . hpath . "\ZipReceivingFolder"

    paste_to_folder(intermediateReceivingDirectory)

    ; MsgBox, Clipboard file was pasted to %intermediateReceivingDirectory%

    mostRecentlyModifiedFile := getMostRecent(intermediateReceivingDirectory)
    ; MsgBox, Most recently modified file is %mostRecentlyModifiedFile%

    ; MsgBox, Unzipping %mostRecentlyModifiedFile% to %to_parent_folder%
    unzip(mostRecentlyModifiedFile, to_parent_folder)
    ; MsgBox, repoTransmissionFolder is %repoTransmissionFolder%
    ; ; Close active window
    ; WinClose, ahk_class CabinetWClass
}


paste_to_folder(receiving_folder){

    OpenFolderExplorer(receiving_folder)
    ; Wait for explorer window to open
    Sleep, 1000
    SendInput, ^v
    MsgBox, Confirm paste has finished
}


OpenFolderExplorer(targetFolder)
{
    ; Run the File Explorer command
    Run, explorer.exe
    Sleep, 1000
    ; Wait for the File Explorer window to open
    ; WinWait, ahk_class CabinetWClass
    NavigateToFolder(targetFolder)
    return
}


NavigateToFolder(folderPath)
{
    ; ; Set the active window to File Explorer
    ; WinActivate, ahk_class CabinetWClass

    ; ; Wait for the window to become active
    ; WinWaitActive, ahk_class CabinetWClass

    Sleep 1000
    SendInput, !d
    
    Sleep 1000
    ; Send the folder path
    SendInput, %folderPath%{Enter}

    ; ensure navigation is complete
    Sleep 1000
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


unzip(zipFile, targetFolder) {

    ; RunWait PowerShell.exe -Command Expand-Archive -Force -LiteralPath '%zipFile%' -DestinationPath %targetFolder%,, ; hide
    RunWait PowerShell.exe -Command "try { Expand-Archive -Force -LiteralPath '%zipFile%' -DestinationPath %targetFolder% } catch { Write-Host $_.Exception.Message; pause }",, ; Hide
    
    return
}


unzip_no_exit(zipFile, targetFolder) {

    RunWait PowerShell.exe -NoExit -Command "try { Expand-Archive -Force -LiteralPath '%zipFile%' -DestinationPath %targetFolder% } catch { Write-Host $_.Exception.Message; pause }",, ; Hide
    ; Wait for PowerShell to finish
    ; Process, WaitClose, PowerShell.exe
    Return
}