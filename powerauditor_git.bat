@echo off
:: PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
::
:: Filename: powerauditor_git.bat - Script for auto update
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

set mypath=%~dp0
set mypath=%mypath:~0,-1%
pushd %mypath%
git %2 %3 %4 %5 %6 %7 > %1.log 2>&1
echo %ERRORLEVEL% > %1.ret 2>&1
popd