#NoEnv ; Prevents checking empty variables as environment variables
#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

SetTitleMatchMode, 2 ; Set the window title matching mode to match any part of the title

EnvGet, hdrive, Homedrive
EnvGet, hpath, Homepath
; Specify the location to send newly created zip files before they get sent to the VDI
intermediateDirectory := hdrive . hpath . "\RepoTransmission"

^!s::
{
    ; Get target source
    folderPath := GetActiveExplorerPath()
    ; Check if a folder path was obtained
    if folderPath = ""
    {
        MsgBox, No folder is currently selected in the File Explorer window.
        ExitApp
    }
    ; Generate a timestamp to append to the output ZIP file name
    formatTime, timeStamp, A_Now, yyyyMMdd_HHmmss

    ; Create the output ZIP file path using the folder name and timestamp
    ; outputZip := folderPath . "\" . GetFolderName(folderPath) . "_" . timeStamp . ".zip"

    outputZip := intermediateDirectory . "\" . GetFolderName(folderPath) . "_" . timeStamp . ".zip"

    ; Zip target source to a temp destination
    PowerShellZip(folderPath, outputZip)

    ; Copy newly created zip file to clip board
    ClipboardSetFiles(outputZip, "Copy")

    ; Make VDI window active & Activate VDI emissary script
    FocusVDI()
    Sleep, 1000
    ActivateVdiEndScript() 

    ; [Wait
    ; Delete zip file from MMD]
}


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

; Function to get the folder name from a given path
GetFolderName(path)
{
    SplitPath, path, , , folderName
    return folderName
}

; Zip target source to a temp destination
PowerShellZip(sourceFolder, outputZip){
    ; MsgBox, %sourceFolder%
    ; MsgBox, %outputZip%
    ; Zip
    RunWait PowerShell.exe -Command Compress-Archive -LiteralPath '%sourceFolder%' -CompressionLevel Optimal -DestinationPath '%outputZip%',, Hide
    ; MsgBox, Created: %outputZip% 
}

; Copy file to clip board
ClipboardSetFiles(FilesToSet, DropEffect := "Copy") {
   ; FilesToSet - list of fully qualified file pathes separated by "`n" or "`r`n"
   ; DropEffect - preferred drop effect, either "Copy", "Move" or "" (empty string)
   Static TCS := A_IsUnicode ? 2 : 1 ; size of a TCHAR
   Static PreferredDropEffect := DllCall("RegisterClipboardFormat", "Str", "Preferred DropEffect")
   Static DropEffects := {1: 1, 2: 2, Copy: 1, Move: 2}
   ; -------------------------------------------------------------------------------------------------------------------
   ; Count files and total string length
   TotalLength := 0
   FileArray := []
   Loop, Parse, FilesToSet, `n, `r
   {
      If (Length := StrLen(A_LoopField))
         FileArray.Push({Path: A_LoopField, Len: Length + 1})
      TotalLength += Length
   }
   FileCount := FileArray.Length()
   If !(FileCount && TotalLength)
      Return False
   ; -------------------------------------------------------------------------------------------------------------------
   ; Add files to the clipboard
   If DllCall("OpenClipboard", "Ptr", A_ScriptHwnd) && DllCall("EmptyClipboard") {
      ; HDROP format ---------------------------------------------------------------------------------------------------
      ; 0x42 = GMEM_MOVEABLE (0x02) | GMEM_ZEROINIT (0x40)
      hDrop := DllCall("GlobalAlloc", "UInt", 0x42, "UInt", 20 + (TotalLength + FileCount + 1) * TCS, "UPtr")
      pDrop := DllCall("GlobalLock", "Ptr" , hDrop)
      Offset := 20
      NumPut(Offset, pDrop + 0, "UInt")         ; DROPFILES.pFiles = offset of file list
      NumPut(!!A_IsUnicode, pDrop + 16, "UInt") ; DROPFILES.fWide = 0 --> ANSI, fWide = 1 --> Unicode
      For Each, File In FileArray
         Offset += StrPut(File.Path, pDrop + Offset, File.Len) * TCS
      DllCall("GlobalUnlock", "Ptr", hDrop)
      DllCall("SetClipboardData","UInt", 0x0F, "UPtr", hDrop) ; 0x0F = CF_HDROP
      ; Preferred DropEffect format ------------------------------------------------------------------------------------
      If (DropEffect := DropEffects[DropEffect]) {
         ; Write Preferred DropEffect structure to clipboard to switch between copy/cut operations
         ; 0x42 = GMEM_MOVEABLE (0x02) | GMEM_ZEROINIT (0x40)
         hMem := DllCall("GlobalAlloc", "UInt", 0x42, "UInt", 4, "UPtr")
         pMem := DllCall("GlobalLock", "Ptr", hMem)
         NumPut(DropEffect, pMem + 0, "UChar")
         DllCall("GlobalUnlock", "Ptr", hMem)
         DllCall("SetClipboardData", "UInt", PreferredDropEffect, "Ptr", hMem)
      }
      DllCall("CloseClipboard")
      Return True
   }
   Return False
}
; Make VDI window active
; Activate VDI emissary script


; Define a hotkey to trigger the action (you can customize the hotkey as per your preference)
FocusVDI() {
    ; Assumes the VDI is open on the current desktop environment
    ; Bring the Citrix VDI window into focus
    WinActivate, CVS Main Windows HVD
    return
}

ActivateVdiEndScript() {
    Send, ^!r
}

