:: PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
::
:: Filename: setup.bat - System autoinstall
::
:: This program is free software; you can redistribute it and/or modify
:: it under the terms of the GNU General Public License as published by
:: the Free Software Foundation; either version 2 of the License, or
:: (at your option) any later version.
::
:: This program is distributed in the hope that it will be useful,
:: but WITHOUT ANY WARRANTY; without even the implied warranty of
:: MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
:: GNU General Public License for more details.
::
:: You should have received a copy of the GNU General Public License
:: along with this program; see the file COPYING. If not, write to the
:: Free Software Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
set scriptpath=%~dp0
set scriptpath=%scriptpath:~0,-1%
set DST=%userprofile%\PowerAuditor
echo [.ShellClassInfo]     > %DST%\desktop.txt
echo ConfirmFileOp=0      >> %DST%\desktop.txt
echo NoSharing=1          >> %DST%\desktop.txt
echo IconFile=%DST%\install\icon.ico       >> %DST%\desktop.txt
echo IconIndex=0          >> %DST%\desktop.txt
echo IconResource=%DST%\install\icon.ico,0 >> %DST%\desktop.txt
echo InfoTip=PowerAuditor >> %DST%\desktop.txt
echo [ViewState]          >> %DST%\desktop.txt
echo Mode=                >> %DST%\desktop.txt
echo Vid=                 >> %DST%\desktop.txt
echo FolderType=Generic   >> %DST%\desktop.txt
chcp 1252 >NUL
attrib -S -H -R %DST%\desktop.ini 2>NUL
::cmd.exe /D /A /C (SET/P=ÿþ)<NUL > desktop.ini 2>NUL
cmd.exe /D /U /C type %DST%\desktop.txt > %DST%\desktop.ini
del /q %DST%\desktop.txt
attrib +S +H +R %DST%\desktop.ini

powershell -exec bypass -nop -File "%scriptpath%\wrapper.ps1"
