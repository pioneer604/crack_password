# Determine script location for PowerShell
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
cd $ScriptDir

# pyinstaller  -y --upx-dir . main.spec

py -3.10 -m PyInstaller  -y main.spec
