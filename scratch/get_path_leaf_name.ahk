#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%



^!t::
    rslt:=GetPathLeafName("C:\Users\Alexander.Oakley\OneDrive - Colonial First State\Code\Bare-Repos\AdConDeM.git")
    MsgBox, %rslt%


GetPathLeafName(folderPath) {
    return SubStr(folderPath, InStr(folderPath, "\", 0, -1) + 1)
}
