#Persistent
#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

; This feels like a bodge. I took a function that I knew found a folder's immediate
; parent, and used that in a recursive funtion (ScanUp) that finds the first folder
; that is a child of a specified folder. I'm sure there's a better way to do this.

changedFile := "C:\Users\Alexander.Oakley\Bare-Repos\small-sample-folder\some-sub\another-1\somedoc.txt"
; changedFile := "C:\Users\Alexander.Oakley\Bare-Repos\small-sample-folder"

firstRootChild := GetFirstRootChild(changedFile, "C:\Users\Alexander.Oakley")
MsgBox, firstRootChild is %firstRootChild%

getParent(folderPath) {
    ; Get the immediate parent of the folder specified by folderPath
    ; Note: I copied this for the internet, and I have no idea why it needs 
    ; to use a loop (source: https://www.autohotkey.com/boards/viewtopic.php?t=70027)

    loop, Files, % folderPath "\..", D
    {
        parentPath := A_LoopFileLongPath
        ; I would have thought removing 'break' would allow this loop to continue
        ; moving up the folder tree. But it doesn't. It just stops at the first
        ; parent folder. I don't know why.
        ; break
        MsgBox, % parentPath
    }
    return parentPath
}

GetFirstRootChild(LeafFolder, Root) {
    firstRootChild := ScanUp(LeafFolder, Root)
    return firstRootChild
}

ScanUp(current, tooHigh) {
    ; Scan up the folder tree until we find a folder that is a child of tooHigh
    parent := getParent(current)
    if(parent = tooHigh){
        return current
    }
    if not doesDirExist(parent) {
        MsgBox, %parent% does not exist
        return ""
    }
    return ScanUp(parent, tooHigh)

}

doesDirExist(DirPath){
    ; Make sure DirPath is actually a dir (or folder) and not a file.
    If( InStr( FileExist(DirPath), "D") )
        return true
    else
        return false
}