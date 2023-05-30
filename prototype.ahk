#NoEnv ; Prevents checking empty variables as environment variables
#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

SetTitleMatchMode, 2 ; Set the window title matching mode to match any part of the title

; Define a hotkey to trigger the action (you can customize the hotkey as per your preference)
^!f::
{
    ; Bring the Citrix VDI window into focus
    ; WinActivate, CVS Main Windows HVD – AU - CFS - Desktop Viewer
    ; WinActivate, Citrix Workspace - Google Chrome
    WinActivate, CVS Main Windows HVD
    Sleep, 500
    Send, ^!e

    return
}
