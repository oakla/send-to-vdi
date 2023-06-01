#Persistent
#SingleInstance, force
SetWorkingDir, %A_ScriptDir%

#include WatchFolder.ahk

EnvGet, hdrive, Homedrive
EnvGet, hpath, Homepath
targetFolder := hdrive . hpath . "\Downloads"

WatchFolder(targetFolder, "ReportFunction") ;

ReportFunction(Directory, Changes)
	{
	 For Each, Change In Changes
		{
         MsgBox, In ReportFunction
		 Action := Change.Action
		 Name := Change.Name
         MsgBox, %Name%
		;  ; -------------------------------------------------------------------------------------------------------------------------
		;  ; Action 1 (added) = File gets added in the watched folder
		;  If (Action = 1) 
		;  	{
		; 	 If RegExMatch(name,"i)report.*pdf$") 
		; 	 	{
		; 	 	 ; if excel file do something
		; 	 	 FileMove, % name, E:\Documents\Reports\report card.pdf
		; 		 ;MsgBox % name
		; 	 	}
		;  	}
		}
	}

