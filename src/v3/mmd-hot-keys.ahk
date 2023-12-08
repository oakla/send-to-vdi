

#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

#Include, routines\01_send_folder_contents_in_explorer_window_to_vdi.ahk
#Include, routines\001_zip_current_explorer_window_to_clipboard.ahk
#Include, routines\002_unpack-clipboard-zip.ahk

EnvGet, hdrive, Homedrive
EnvGet, hpath, Homepath


!^z::zip_current_explorer_window_to_clipboard()

!^s::send_folder_contents_in_explorer_window_to_vdi()

; Specify the location to send newly created zip files before they get sent to the VDI
; ...todo
; Specify destination of unpacked files
bareRepoFolder := hdrive . hpath . "\code\bare-repos"
!^v::unpack_clipboard_zip(bareRepoFolder) 