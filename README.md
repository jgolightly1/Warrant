# Warrant
-------------------------------------------------------------------------------
INSTALL INSTRUCTIONS
-------------------------------------------------------------------------------
(Menu > Type: cmd > right-click Command Prompt > Run as Administrator > Enter the following:)

@"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command " [System.Net.ServicePointManager]::SecurityProtocol = 307t System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"2; iex ((New-Objec

choco install python3 --version 3.12.0-a2 --pre -y
choco install git --version 2.39.2 -y

winget install -e --id Microsoft.VisualStudio.2022.BuildTools --override "--passive --wait --add Microsoft.VisualStudio.Workload.VCTools;includeRecommended"

refreshenv

set PYTHONPATH=list;of;paths

cd %USERPROFILE%\Documents
git clone https://github.com/jgolightly1/Warrant.git

pip install -r requirements.txt
