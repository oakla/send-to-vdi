#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

#Include, routines\001_zip_current_explorer_window_to_clipboard.ahk

send_folder_contents_in_explorer_window_to_vdi() {
    zip_current_explorer_window_to_clipboard()
    FocusVDI()
    ConfirmVDIFocus()
    ActivateVDISideScript()
}


FocusVDI() {
    ; Assumes the VDI is open on the current desktop environment
    ; Bring the Citrix VDI window into focus

    ; WinActivate, CVS Main Windows HVD – AU - CFS - Desktop Viewer
    WinActivate, CVS Main Windows HVD
    return
}

ConfirmVDIFocus() {
    FocusVDI()
    ; Check if the VDI window is in focus
    if (WinActive("CVS Main Windows HVD"))
    {
        return
    }
    else
    {
        MsgBox, The VDI window is not in focus.
    }
}

ActivateVDISideScript() {
    ; Hotkey must distinct from any MMD-side hotkey otherwise collision will 
    ; cause MMD hotkey to activate and VDI-side will be ignored
    Send, ^!r
}
