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
[void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")


$DesktopIni = @"
[.ShellClassInfo]
ConfirmFileOp=1
NoSharing=1
IconFile=$env:USERPROFILE\PowerAuditor\install\icon.ico
IconIndex=0
IconResource=$env:USERPROFILE\PowerAuditor\install\icon.ico,0
InfoTip=PowerAuditor
[ViewState]
Mode=
Vid=
FolderType=Generic
[DeleteOnCopy]
Personalized=5
"@
rm -Force -Recurse "$($env:USERPROFILE)\PowerAuditor\desktop.ini"
#Create/Add content to the desktop.ini file
Add-Content "$($env:USERPROFILE)\PowerAuditor\desktop.ini" -Value $DesktopIni
#Set the attributes for $DesktopIni
(Get-Item "$($env:USERPROFILE)\PowerAuditor\desktop.ini" -Force).Attributes = 'Hidden, System, Archive'
#Finally, set the folder's attributes
(Get-Item "$($env:USERPROFILE)\PowerAuditor" -Force).Attributes = 'ReadOnly, Directory'



# We ask for the user identity
if( (Get-ChildItem $env:USERPROFILE\PowerAuditor\config.ini -ErrorAction SilentlyContinue).Count -eq 0 ){
	$email = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your address email for reports (ie: fake@domain.com)", "Office356 configuration")
	$pseudo = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your name for reports (ie: NUEL Guillaume)", "Office356 configuration")
	echo "FriendlyName=$pseudo" | Out-File -Encoding ascii $env:USERPROFILE\PowerAuditor\config.ini
	echo "EmailAddress=$email" | Out-File -Encoding ascii -Append $env:USERPROFILE\PowerAuditor\config.ini
}


echo '
{
	"monaco": {
		"editorOptions": {
			"minimap": {
				"enabled": false
			},
			"wordWrap": "bounded"
		}
	},
	"sorting": {
		"by": "title",
		"type": "ascending"
	},
	"tutorial": true,
	"cwd": "%USERPROFILE%\\PowerAuditor\\vulndb\\.notable"
}'.Replace('%USERPROFILE%',($env:USERPROFILE).Replace("\", "\\")) | Out-File -Encoding ascii $env:USERPROFILE\.notable.json


# White list git host to avoid error "unknown host key"
mkdir $env:USERPROFILE\.ssh -ErrorAction SilentlyContinue
Get-ChildItem $env:USERPROFILE\PowerAuditor -Recurse -Force | where {  $_.FullName.Contains(".git\config") } | foreach {
	$tmp=cat $_.Fullname;
	$rx = [regex]::Match($tmp, "url = [a-z]+@([^:\r\n]+)");
	if( $rx.Success -eq $false ){
		$rx = [regex]::Match($tmp, "url = https?://([^/:\r\n]+)");
	}
	$rx = $rx.Captures.Groups[1].Value
	& $env:USERPROFILE\PowerAuditor\install\ssh-keyscan.exe $rx | Out-File -Encoding ascii -Append $env:USERPROFILE\.ssh\known_hosts
}

#if( [System.IO.File]::Exists('C:\Program Files\Git\bin\git.exe') ){
#	Write-Host "Git allready installed"
#	# We ask for the user identity
#	if( (Get-ChildItem $env:USERPROFILE\.gitconfig -ErrorAction SilentlyContinue).Count -eq 0 ){
#		$email = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your address email for GIT", "GIT configuration")
#		$pseudo = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your name for GIT (ie: NUEL Guillaume)", "GIT configuration")
#		git config --global user.name $pseudo
#		git config --global user.email $email
#	}
#	exit
#}

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

# Install git and notable
choco install git.install -y
choco install notable -y
New-Alias git "C:\Program Files\Git\bin\git.exe"
mkdir $env:USERPROFILE\PowerAuditor\vulndb\.notable\ -ErrorAction SilentlyContinue
mkdir $env:USERPROFILE\PowerAuditor\vulndb\.notable\notes -ErrorAction SilentlyContinue
mkdir $env:USERPROFILE\PowerAuditor\vulndb\.notable\attachments -ErrorAction SilentlyContinue
# Required for ActiveWorkbook.VBProject.VBComponents
reg ADD HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security /v AccessVBOM /t REG_DWORD /d 1 /f

# We ask for the user identity
if( (Get-ChildItem $env:USERPROFILE\.gitconfig -ErrorAction SilentlyContinue).Count -eq 0 ){
	$email = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your address email for GIT (ie: fake@domain.com)", "GIT configuration")
	$pseudo = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your name for GIT (ie: NUEL Guillaume)", "GIT configuration")
	git config --global user.name $pseudo
	git config --global user.email $email
}

cmd /c mklink $env:USERPROFILE\Desktop\PowerAuditor.xlsm $env:USERPROFILE\PowerAuditor\PowerAuditor_last.xlsm
New-Item -ItemType SymbolicLink -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup" -Name "PowerAuditor_AutoUpdater.lnk" -Value "$env:USERPROFILE\PowerAuditor\bin\AutoUpdater.exe"

$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\PowerAuditor_AutoUpdater.lnk")
$Shortcut.TargetPath = "$env:USERPROFILE\PowerAuditor\bin\AutoUpdater.exe"
$Shortcut.Save()

# Create a lnk to the vulndb
$Shortcut = $WshShell.CreateShortcut("$($env:USERPROFILE)\Desktop\PowerAuditor.lnk")
$Shortcut.TargetPath = "$($env:USERPROFILE)\PowerAuditor\vulndb\"
$shortcut.IconLocation = "$($env:USERPROFILE)\PowerAuditor\install\icon.ico,0"
$Shortcut.Save()
