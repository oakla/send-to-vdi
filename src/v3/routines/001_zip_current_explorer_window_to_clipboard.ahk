#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%


; Define the hotkey
zip_current_explorer_window_to_clipboard() {

   ; homePath := A_Home ; alternative homepath access method
   EnvGet, hdrive, Homedrive
   EnvGet, hpath, Homepath
   ; Specify the location to send newly created zip files before they get sent to the VDI
   intermediateZipDirectory := hdrive . hpath . "\ZipTransmission"


    ; Get target source
    folderPath := GetActiveExplorerPath()
    ; Check if a folder path was obtained
    if folderPath = ""
    {
        MsgBox, No folder is currently selected in the File Explorer window.
        ExitApp
    }

   ;  MsgBox, The selected item's folderPath is: %folderPath%

   ;  MsgBox, intermediateZipDirectory is: %intermediateZipDirectory%
    ; Zip the folder to an intermediate folder using PowerShell

    ; Generate a timestamp to append to the output ZIP file name
    formatTime, timeStamp, A_Now, yyyyMMdd_HHmmss

    zipPath := intermediateZipDirectory . "\" . GetPathLeafName(folderPath) . "_" . timeStamp . ".zip"
    MsgBox, Zip path is: %zipPath%

    ;  PowerShellZipNoExit(folderPath, zipPath)
     PowerShellZip(folderPath, zipPath)
 
    ; Copy newly created zip file to clip board
    ClipboardSetFiles(zipPath, "Copy")
}

; Zip target source to a temp destination
PowerShellZip(sourceFolder, outputZip){
    ; MsgBox, %sourceFolder%
    ; MsgBox, %outputZip%
    ; Zip
    RunWait PowerShell.exe -Command "Compress-Archive -LiteralPath '%sourceFolder%' -CompressionLevel Optimal -DestinationPath '%outputZip%'",, ; Hide
    ; Wait for PowerShell to finish
    Process, WaitClose, PowerShell.exe
}
    
PowerShellZipNoExit(sourceFolder, outputZip){
    ; MsgBox, %sourceFolder%
    ; MsgBox, %outputZip%
    ; Zip
    RunWait PowerShell.exe -NoExit -Command "try { Compress-Archive -LiteralPath '%sourceFolder%' -CompressionLevel Optimal -DestinationPath '%outputZip%' } catch { Write-Host $_.Exception.Message; pause }",, ; Hide
    ; Wait for PowerShell to finish
    Process, WaitClose, PowerShell.exe
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


GetPathLeafName(folderPath) {
    return SubStr(folderPath, InStr(folderPath, "\", 0, -1) + 1)
}


; Copy file to clip board
ClipboardSetFiles(FilesToSet, DropEffect := "Copy") {
   if not (FileExist(FilesToSet))
   {
      MsgBox, The path %path% is invalid.
      ExitApp
   }
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
      MsgBox, % "Files copied to clipboard"
      Return True
   }
   MsgBox, % "Files copied FAILED to be clipboard"
   Return False
}