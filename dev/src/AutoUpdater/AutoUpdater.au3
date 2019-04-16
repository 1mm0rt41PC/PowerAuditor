; PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
;
; Filename: AutoUpdater.au3
;
; This program is free software; you can redistribute it and/or modify
; it under the terms of the GNU General Public License as published by
; the Free Software Foundation; either version 2 of the License, or
; (at your option) any later version.
;
; This program is distributed in the hope that it will be useful,
; but WITHOUT ANY WARRANTY; without even the implied warranty of
; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
; GNU General Public License for more details.
;
; You should have received a copy of the GNU General Public License
; along with this program; see the file COPYING. If not, write to the
; Free Software Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
#NoTrayIcon
#AutoIt3Wrapper_Res_HiDpi=1
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/nsdp
#Au3Stripper_On
#pragma compile(Out, ../../../bin/AutoUpdater.exe)
#pragma compile(UPX, True)
#pragma compile(Icon, ./iconfinder_superhero_deadpool_hero_comic_1380751.ico)
#pragma compile(AutoItExecuteAllowed, False)
#pragma compile(Console, False)
#pragma compile(x64, True)
#pragma compile(FileDescription, 'AutoUpdater for PowerAuditor')
#pragma compile(ProductName, AutoUpdater)
#pragma compile(ProductVersion, 1.0)
#pragma compile(FileVersion, 1.0) ; The last parameter is optional.
#pragma compile(LegalCopyright, © 1mm0rt41PC)
#pragma compile(LegalTrademarks, 'See https://github.com/ImmortalPC. Icon from https://www.iconfinder.com/icons/1380751/comic_deadpool_hero_superhero_icon by Aitor Picon')
#pragma compile(CompanyName, 'None')
Opt('TrayAutoPause', 0)
Opt('TrayIconDebug', 0)
Opt('TrayIconHide', 1)
Opt('GUICloseOnESC', False)

#include <File.au3>
#include <WindowsConstants.au3>
#include <WinAPISys.au3>
#include <WinAPIvkeysConstants.au3>
#include <GUIConstantsEx.au3>

Global Const $sPIDFile = @TempDir & '\AutoUpdater-PowerAuditor.pid'


DllCall('User32.dll', 'bool', 'SetProcessDPIAware') ; Support du DPI

; We avoid to boot multiple time
If FileExists($sPIDFile) And ProcessExists(FileRead($sPIDFile)) Then Exit

; If the binary is not in the temp folder, we 'fork' this process to allow update
If Not StringInStr(FileGetLongName(@ScriptDir), FileGetLongName(@TempDir)) Then
	Local $sExeFile = _TempFile(@TempDir, '~', '.exe')
	FileCopy(@ScriptFullPath, $sExeFile)
	FileChangeDir(@ScriptDir & '\..\')
	Run($sExeFile, @WorkingDir)
	Exit
EndIf

; Lock the binary Singleton
FileWrite($sPIDFile, @AutoItPID)

Local $iLastTimeExeUpdated = FileGetTime(@WorkingDir & '\bin\AutoUpdater.exe', $FT_MODIFIED, $FT_STRING)
Local $iLoop = 0
While 1
	If GUIGetMsg() == $GUI_EVENT_CLOSE Then ExitLoop
	Sleep(5 * 1000)
	If $iLoop > 10 Then
		$iLoop = -1
		git('')
		git('vulndb')
		git('template')
		If Not FileExists($sPIDFile) Or $iLastTimeExeUpdated <> FileGetTime(@WorkingDir & '\bin\AutoUpdater.exe', $FT_MODIFIED, $FT_STRING) Then
			FileDelete($sPIDFile)
			Local $sTmpBat = _TempFile(@TempDir, '~', '.bat')
			FileWrite($sTmpBat, 'ping -n 5 127.0.0.1' & @CRLF & 'del /F /Q "' & @ScriptFullPath & '" "' & $sTmpBat & '"')
			Run($sTmpBat, @WorkingDir, @SW_HIDE)
			Exit
		EndIf
	EndIf
	$iLoop += 1
WEnd


Func git($sRepo)
	Local $iPID = Run('git pull', @WorkingDir & '\' & $sRepo, @SW_HIDE, $STDOUT_CHILD)
	ProcessWaitClose($iPID)
	$retCode = @extended
	Local $sOutput = StdoutRead($iPID)
	If $retCode <> 0 Then
		If $sRepo == '' Then
			$sRepo = 'main'
		EndIf
		MsgBox(0, 'AutoUpdater for PowerAuditor', 'There is an error when pulling the repo >' & $sRepo & '<' & @CRLF & $sOutput, 3)
	EndIf
EndFunc   ;==>git
