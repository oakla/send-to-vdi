#Persistent
#SingleInstance, force
SetWorkingDir, %A_ScriptDir%

#Include WatchFolder.ahk

EnvGet, hdrive, Homedrive
EnvGet, hpath, Homepath
targetFolder := hdrive . hpath . "\Bare-Repos\"

WatchFolder(targetFolder, "ChangeResponse", true)

; MsgBox, Staring script
ChangeResponse(EffectedPath, Changes) {
    For Each, Change in Changes
        {
            Name := Change.Name
            Action := Change.Action
            IsDir := Change.IsDir
            MsgBox, % "Name: " . Name . "`nAction: " . Action . "`nIsDir: " . IsDir
        }
    MsgBox, % "Folder: " . EffectedPath . "`nChangeInfo: " . ChangeInfo
}

