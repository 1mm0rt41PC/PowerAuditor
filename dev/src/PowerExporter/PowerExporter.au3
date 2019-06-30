; PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
;
; Filename: PowerExporter.au3
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
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_parameters=/STRIPONLY
#Tidy_Parameters=/nsdp
#pragma compile(Out, ../../../bin/PowerExporter.exe)
#pragma compile(UPX, True)
#pragma compile(Icon, ../PowerAuditor.ico)
#pragma compile(AutoItExecuteAllowed, False)
#pragma compile(Console, True)
#pragma compile(x64, True)
#pragma compile(FileDescription, "Export Editor for PowerAuditor")
#pragma compile(ProductName, PowerExporter)
#pragma compile(ProductVersion, 1.0)
#pragma compile(FileVersion, 1.0) ; The last parameter is optional.
#pragma compile(LegalCopyright, © 1mm0rt41PC)
#pragma compile(LegalTrademarks, 'See https://github.com/1mm0rt41PC. Icon from https://www.iconfinder.com/')
#pragma compile(CompanyName, 'None')
Opt('TrayAutoPause', 0)
Opt('TrayIconDebug', 0)
Opt('TrayIconHide', 1)
;~ Opt('GUICloseOnESC', True)

#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <Array.au3>
#include <GuiListView.au3>
#include <EditConstants.au3>
#include <IE.au3>
#include <WindowsConstants.au3>
#include <WinAPISys.au3>
#include <WinAPIvkeysConstants.au3>
Global Const $SC_CLOSE = 0xF060 ;

If $CmdLine[0] <> 1 Or ($CmdLine[0] == 1 And ($CmdLine[1] == '/help' Or $CmdLine[1] == '/h' Or $CmdLine[1] == '/?' Or $CmdLine[1] == '-h' Or $CmdLine[1] == '-help' Or $CmdLine[1] == '--help')) Then
	ConsoleWriteError('Invalid usage !' & @CRLF)
	ConsoleWriteError(@AutoItExe & ' <file>: Read vulnerability to export from the file <file> with one vuln per row (crlf)' & @CRLF)
	Exit
EndIf

Global $sFileSwap = $CmdLine[1]
Local $aVuln = StringStripWS(FileRead($sFileSwap), 3)
FileDelete($sFileSwap)
If $aVuln == "" Then
	ConsoleWriteError('Invalid usage !' & @CRLF)
	ConsoleWriteError(@AutoItExe & ' <file>: Read vulnerability to export from the file <file> with one vuln per row (crlf)' & @CRLF)
	Exit
EndIf
$aVuln = StringSplit($aVuln, @CRLF, 1)

Global $editor = _TempFile(@TempDir, "~", ".html")
FileInstall('PowerExporter.html', $editor)

Local $sHTML = ''
Local $aSplit
If $aVuln <> 0 Then
	For $i = 1 To $aVuln[0]
		$aSplit = StringSplit($aVuln[$i], @TAB, 1)
		$sHTML &= '<tr>'
		$sHTML &= '	<td onclick="toogleCheckbox(this);"><input type="checkbox" value="' & $aSplit[1] & '" class="checkSelect" disabled="disabled" /></td>'
		$sHTML &= '	<td class="vulnname" onclick="toogleCheckbox(this);">' & $aSplit[2] & '</td>'
		$sHTML &= '	<td class="exploit"><input type="checkbox" value="e" class="checkExploit" onclick="updateURL();" /></td>'
		$sHTML &= '	<td class="status"><select onchange="updateURL();"><option value="d" selected="selected" class="status_Draft">Draft</option><option value="v" class="status_Validated">Validated</option></select></td>'
		$sHTML &= '</tr>'
	Next
EndIf


DllCall("User32.dll", "bool", "SetProcessDPIAware") ; Support du DPI
Global $oIE = _IECreateEmbedded()
Local $width = @DesktopWidth / 1.5
Local $height = @DesktopHeight / 1.5
Local $hGui = GUICreate('PowerAuditor - Export vulnerability', $width, $height, -1, -1, $WS_OVERLAPPEDWINDOW + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
Local $hIE = GUICtrlCreateObj($oIE, 10, 10, $width - 20, $height - 20)
_IENavigate($oIE, $editor)
GUIRegisterMsg($WM_SYSCOMMAND, "On_Exit")
GUISetState(@SW_SHOW) ;Show GUI



Local $sFileHTML = FileRead($editor)
_IEBodyWriteHTML($oIE, StringReplace($sFileHTML, '%INSERTDATA%', $sHTML))

Local Const $iBitMask = 0x8000 ; a bit mask to strip the high word bits from the return of the function.
While 1
	If BitAND(_WinAPI_GetAsyncKeyState($VK_ESCAPE), $iBitMask) <> 0 And WinActive($hGui) Then
		ExitLoop
	EndIf
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			MsgBox(0, 0, 1)
			ExitLoop
		Case $GUI_EVENT_RESTORE, $GUI_EVENT_MAXIMIZE, $GUI_EVENT_RESIZED, $GUI_FOCUS
			$size = WinGetClientSize($hGui)
			GUICtrlSetPos($hIE, 10, 10, $size[0] - 20, $size[1] - 20)
	EndSwitch
WEnd

On_Exit()


Func On_Exit($hWnd = Null, $Msg = Null, $wParam = $SC_CLOSE, $lParam = Null)
	If $wParam <> $SC_CLOSE Then Return $GUI_RUNDEFMSG
	Local $tag = StringSplit(_IEPropertyGet($oIE, 'locationurl'), '#')
	If $tag[0] == 2 Then
		FileWrite($sFileSwap, DecodeUrl(StringReplace(DecodeUrl($tag[2]), '%0A', @CRLF)))
	Else
		FileWrite($sFileSwap, '')
	EndIf
	FileDelete($editor)
	Exit
EndFunc   ;==>On_Exit


Func DecodeUrl($src)
	Local $i
	Local $ch
	Local $buff

	;Init Counter
	$i = 1

	While ($i <= StringLen($src))
		$ch = StringMid($src, $i, 1)
		;Correct spaces
		If ($ch = "+") Then
			$ch = " "
		EndIf
		;Decode any hex values
		If ($ch = "%") Then
			$ch = Chr(Dec(StringMid($src, $i + 1, 2)))
			$i += 2
		EndIf
		;Build buffer
		$buff &= $ch
		;Inc Counter
		$i += 1
	WEnd

	Return $buff
EndFunc   ;==>DecodeUrl
