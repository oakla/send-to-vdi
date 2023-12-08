#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%




; clipboard := ""  ; Clear the clipboard variable
; Send, ^c  ; Copy the selected text to the clipboard
; ClipWait  ; Wait for the clipboard to contain data

MsgBox, % "Clipboard contents: " . clipboard  ; Print the clipboard contents
