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
powershell -exec bypass -nop -File "%scriptpath%\wrapper.ps1"
