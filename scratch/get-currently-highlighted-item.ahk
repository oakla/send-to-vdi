#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%


; Get the path of the currently highlighted item in Windows Explorer
!^t::
    clipboard := ""
    SendInput, ^c
    ClipWait, 1
    path := clipboard

    MsgBox, The selected item's path is: %path%
