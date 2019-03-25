# PowerAuditor - A simple script to help report writing by https://github.com/1mm0rt41PC
#
# Filename: wrapper.ps1 - System autoinstall
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 2 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; see the file COPYING. If not, write to the
# Free Software Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.

if( [System.IO.File]::Exists('C:\Program Files\Git\bin\git.exe') ){
	Write-Host "Git allready installed"
	exit
}

# Elevate to Admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
	Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs;
	exit
}

# Install Choco
if( -not [System.IO.File]::Exists('C:\ProgramData\chocolatey\bin\choco.exe') ){
	Set-ExecutionPolicy Bypass -Scope Process -Force;
	iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
}

# Install git
choco install git.install --force -y
# Requis pour ActiveWorkbook.VBProject.VBComponents
reg ADD HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security /v AccessVBOM /t REG_DWORD /d 1 /f

$email = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your address email", "GIT configuration")
$pseudo = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your name (ie: NUEL Guillaume)", "GIT configuration")
git config --global user.name $pseudo
git config --global user.email $email

echo "FriendlyName=$pseudo" | Out-File -FilePath $env:USERPROFILE\PowerAuditor\config.ini -Encoding ascii
echo "EmailAddress=$email" | Out-File -Append -FilePath $env:USERPROFILE\PowerAuditor\config.ini -Encoding ascii

