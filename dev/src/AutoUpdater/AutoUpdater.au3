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
#include <TrayConstants.au3> ; Required for the $TRAY_ICONSTATE_SHOW constant.
#AutoIt3Wrapper_Res_HiDpi=1
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/nsdp
#Au3Stripper_On
#pragma compile(Out, ../../../bin/AutoUpdater.exe)
#pragma compile(UPX, True)
#pragma compile(Icon, ../PowerAuditor.ico)
#pragma compile(AutoItExecuteAllowed, False)
#pragma compile(Console, False)
#pragma compile(x64, True)
#pragma compile(FileDescription, 'AutoUpdater for PowerAuditor')
#pragma compile(ProductName, AutoUpdater)
#pragma compile(ProductVersion, 1.0)
#pragma compile(FileVersion, 1.0) ; The last parameter is optional.
#pragma compile(LegalCopyright, © 1mm0rt41PC)
#pragma compile(LegalTrademarks, 'See https://github.com/1mm0rt41PC. Icon from https://www.iconfinder.com/')
#pragma compile(CompanyName, 'None')
Opt('TrayAutoPause', 0)
Opt('TrayIconDebug', 0)
;~ Opt('TrayIconHide', 1)
Opt('TrayMenuMode', 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Opt('GUICloseOnESC', False)

#include <File.au3>
#include <WindowsConstants.au3>
#include <WinAPISys.au3>
#include <WinAPIvkeysConstants.au3>
#include <GUIConstantsEx.au3>

Global Const $sPIDFile = @TempDir & '\AutoUpdater-PowerAuditor.pid'
Global $iCounterLastError = 0

DllCall('User32.dll', 'bool', 'SetProcessDPIAware') ; Support du DPI
TraySetState($TRAY_ICONSTATE_SHOW) ; Show the tray menu.
TraySetToolTip('AutoUpdate for PowerAuditor')

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
FileDelete($sPIDFile)
FileWrite($sPIDFile, @AutoItPID)

Local $iLastTimeExeUpdated = FileGetTime(@WorkingDir & '\bin\AutoUpdater.exe', $FT_MODIFIED, $FT_STRING)
Local $iLoop = 0
Global $idForceUpdate = TrayCreateItem('Force update')
Local $idExit = TrayCreateItem('Exit')
Global $tray
While 1
	$tray = TrayGetMsg()
	If GUIGetMsg() == $GUI_EVENT_CLOSE Or $tray == $idExit Then ExitLoop
	Sleep(100)
	If $iLoop > 10 * 60 * 60 Or $tray == $idForceUpdate Then
		If $tray == $idForceUpdate Then TrayTip('PowerAuditor', 'Updating all repositories', 5, $TIP_ICONASTERISK)
		$iLoop = -1
		git('')
		git('vulndb')
		git('template')
		UpdateVulnDBFolder()
		If Not FileExists($sPIDFile) Or $iLastTimeExeUpdated <> FileGetTime(@WorkingDir & '\bin\AutoUpdater.exe', $FT_MODIFIED, $FT_STRING) Then
			FileDelete($sPIDFile)
			Local $sTmpBat = _TempFile(@TempDir, '~', '.bat')
			FileWrite($sTmpBat, 'ping -n 5 127.0.0.1' & @CRLF & 'del /F /Q "' & @ScriptFullPath & '" "' & $sTmpBat & '"' & @CRLF & '"' & @WorkingDir & '\bin\AutoUpdater.exe"')
			Run($sTmpBat, @WorkingDir, @SW_HIDE)
			Exit
		EndIf
	EndIf
	$iLoop += 1
WEnd


Func myMsgBox($msg)
	If $iCounterLastError == @MDAY And $tray <> $idForceUpdate Then Return Null
	TrayTip('PowerAuditor', $msg, 5, $TIP_ICONEXCLAMATION)
	$iCounterLastError = @MDAY
EndFunc   ;==>myMsgBox


Func git($sRepo)
	Local $iPID = Run('git pull', @WorkingDir & '\' & $sRepo, @SW_HIDE, BitOR($STDERR_CHILD, $STDOUT_CHILD))
	ProcessWaitClose($iPID)
	$retCode = @extended
	Local $sOutput = StdoutRead($iPID)
	$sOutput &= StderrRead($iPID)
	If $retCode <> 0 Then
		If StringInStr($sOutput, "Could not resolve host") Then
			Return Null
		EndIf
		If $sRepo == '' Then
			$sRepo = 'main'
		EndIf
		myMsgBox('There is an error when pulling the repo >' & $sRepo & '<' & @CRLF & $sOutput)
	EndIf
EndFunc   ;==>git


Func UpdateVulnDBFolder()
	Local $sPath = @WorkingDir & '\vulndb'
	Local $hSearch = FileFindFirstFile($sPath & '\*')
	If $hSearch = -1 Then Return Null
	Local $sFileName = ""

	While 1
		$sFileName = FileFindNextFile($hSearch)
		If @error Then ExitLoop

		If StringLeft($sFileName, 1) <> '.' Then
			FileSetAttrib($sPath & '\' & $sFileName, '+R')
			FileSetAttrib($sPath & '\' & $sFileName & '\desktop.ini', '+ASH')
		EndIf
	WEnd

	FileClose($hSearch)
EndFunc   ;==>UpdateVulnDBFolder


