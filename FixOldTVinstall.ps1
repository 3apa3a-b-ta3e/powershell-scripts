#
# Uninstall TeamViewer completely on Windows x64. Used as a quick`n`dirty fix for corrupted installations (required re-install after apply).
#

$WindowsRegUninstallPath="HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
$TVRegUninstallPath="$WindowsRegUninstallPath\TeamViewer"
$TVPFDir="ProgramFiles(x86)\TeamViewer"
$TVRegPath="HKLM:\SOFTWARE\Wow6432Node\TeamViewer"

$TVMSICODE=((Get-ChildItem $WindowsRegUninstallPath | foreach { Get-ItemProperty $_.PSPath } | Where { $_ -match "TeamViewer" }).ModifyPath -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X","").Trim()

start-process "msiexec.exe" -arg "/X $TVMSICODE /qn" -Wait

if (Test-Path $TVPFDir) 
{
    Remove-Item $TVPFDir
}

if (Test-Path $TVRegPath) 
{
    Remove-Item $TVRegPath
}

if (Test-Path $TVRegUninstallPath) 
{
    Remove-Item $TVRegUninstallPath
}

$RE = 'TeamViewer*'
$Key = 'HKLM:\SOFTWARE\Classes\Installer\Products'
Get-ChildItem $Key -Rec -EA SilentlyContinue | ForEach-Object {
	$CurrentKey = (Get-ItemProperty -Path $_.PsPath)
		If ($CurrentKey -match $RE){
			$CurrentKey|Remove-Item -Force -Recurse -Confirm:$false | Out-Null
			}
}

if (Test-Path $TVRegPath) 
 {
    Return
	} else {
    New-Item $TVRegPath
 }
