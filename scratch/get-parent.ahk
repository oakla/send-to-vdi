#Persistent
#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

changedFile := "C:\Users\Alexander.Oakley\Bare-Repos\small-sample-folder\some-sub\another-1\somedoc.txt"

loop, Files, % changedFile "\..", D
{
	parentPath := A_LoopFileLongPath
	; I would have thought removing 'break' would allow this loop to continue
	; moving up the folder tree. But it doesn't. It just stops at the first
	; parent folder. I don't know why.
	; break
    MsgBox, % parentPath
}