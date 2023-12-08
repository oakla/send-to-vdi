#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%


!^t::
    rtnpath := GetActiveExplorerPath()


    MsgBox, %rtnpath%


; Get target source
GetActiveExplorerPath()
{
	explorerHwnd := WinActive("ahk_class CabinetWClass")
	if (explorerHwnd)
	{
		for window in ComObjCreate("Shell.Application").Windows
		{
			if (window.hwnd==explorerHwnd)
			{
				return window.Document.Folder.Self.Path
			}
        }
    }
}