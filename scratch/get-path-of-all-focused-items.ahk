#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%


^!p::
{
    ; For each Windows Explorer window that is open, get the path of the currently focused item
    for window in ComObjCreate("Shell.Application").Windows
        if (window.Document.FocusedItem.Path)
            MsgBox, % window.Document.FocusedItem.Path
}