; PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
;
; Filename: CVSSEditor.au3
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
#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/striponly
#pragma compile(Out, ../../../bin/CVSSEditor.exe)
#pragma compile(UPX, True)
#pragma compile(Icon, ../PowerAuditor.ico)
#pragma compile(AutoItExecuteAllowed, False)
#pragma compile(Console, True)
#pragma compile(x64, True)
#pragma compile(FileDescription, "CVSS Editor for PowerAuditor")
#pragma compile(ProductName, CVSSEditor)
#pragma compile(ProductVersion, 1.0)
#pragma compile(FileVersion, 1.0) ; The last parameter is optional.
#pragma compile(LegalCopyright, © 1mm0rt41PC)
#pragma compile(LegalTrademarks, 'See https://github.com/1mm0rt41PC. Icon from https://www.iconfinder.com/')
#pragma compile(CompanyName, 'None')
Opt('TrayAutoPause', 0)
Opt('TrayIconDebug', 0)
Opt('TrayIconHide', 1)
Opt('GUICloseOnESC', True)

#include <IE.au3>
#include <File.au3>
#include <WindowsConstants.au3>
#include <WinAPISys.au3>
#include <WinAPIvkeysConstants.au3>
#include <GUIConstantsEx.au3>


If $CmdLine[0] == 1 And ($CmdLine[1] == '/help' Or $CmdLine[1] == '/h' Or $CmdLine[1] == '/?' Or $CmdLine[1] == '-h' Or $CmdLine[1] == '-help' Or $CmdLine[1] == '--help') Then
	ConsoleWriteError('Usage !' & @CRLF)
	ConsoleWriteError(@AutoItExe & ' : Run a new CVSS From scratch' & @CRLF)
	ConsoleWriteError(@AutoItExe & ' CVSS:3.0/AV:P/AC:H/PR:H/UI:R/S:C/C:H/I:H/A:H : Run a CVSS editor initialized with the this CVSS value' & @CRLF)
	Exit
EndIf

$cvssEditor = _TempFile(@TempDir, "~", ".html")
FileInstall('cvss.htm', $cvssEditor)

DllCall("User32.dll", "bool", "SetProcessDPIAware") ; Support du DPI
Local $oIE = _IECreateEmbedded()
Local $width = @DesktopWidth / 1.5
Local $height = @DesktopHeight / 1.5
Local $hGui = GUICreate("PowerAuditor - CVSS Editor", $width, $height, -1, -1, $WS_OVERLAPPEDWINDOW + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
Local $hIE = GUICtrlCreateObj($oIE, 10, 10, $width - 20, $height - 20)
If $CmdLine[0] == 0 Then
	_IENavigate($oIE, $cvssEditor)
Else
	_IENavigate($oIE, $cvssEditor & '#' & $CmdLine[1])
EndIf
GUISetState(@SW_SHOW) ;Show GUI

Local Const $iBitMask = 0x8000 ; a bit mask to strip the high word bits from the return of the function.
Local $tag
While 1
	If BitAND(_WinAPI_GetAsyncKeyState($VK_ESCAPE), $iBitMask) <> 0 And WinActive($hGui) Then
		$tag = StringSplit(_IEPropertyGet($oIE, 'locationurl'), '#')
		ExitLoop
	EndIf
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			$tag = StringSplit(_IEPropertyGet($oIE, 'locationurl'), '#')
			ExitLoop
		Case $GUI_EVENT_RESTORE, $GUI_EVENT_MAXIMIZE, $GUI_EVENT_RESIZED, $GUI_FOCUS
			$size = WinGetClientSize($hGui)
			GUICtrlSetPos($hIE, 10, 10, $size[0] - 20, $size[1] - 20)
	EndSwitch
WEnd

If $tag[0] == 2 Then
	ConsoleWrite(StringReplace($tag[2], ';', @CRLF))
EndIf
FileDelete($cvssEditor)
