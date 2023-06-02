#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

; source https://www.autohotkey.com/board/topic/98687-levelup-easy-way-to-refer-to-a-file-or-folder-that-is-higher-up-the-tree/
LevelUp(n){
	return SubStr(A_ScriptDir,1,!InStr(A_ScriptDir,"\",0,-1,n) ? InStr(A_ScriptDir,"\",0,1,1) : InStr(A_ScriptDir,"\",0,-1,n))
}


; source https://www.autohotkey.com/board/topic/32569-up-one-folder/
; 1 - Use split path
Folder := A_AppData
Loop {
  MsgBox, % Folder
  SplitPath, Folder,,Folder,,,Drv
  IfEqual,Folder,%Drv%, Break
}

; 2 - RegEx
; 2.1
NewFolder := RegExReplace(Folder,"\\[^\\]+\\$","\")
; 2.2 - apparently simpler
Folder := RegExReplace(address,"[^\\]+\\?$")