# Warrant
-------------------------------------------------------------------------------
INSTALL INSTRUCTIONS
-------------------------------------------------------------------------------
(Menu > Type: cmd > right-click Command Prompt > Run as Administrator > Enter the following:)

@"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command " [System.Net.ServicePointManager]::SecurityProtocol = 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"

choco install git --version 2.39.2 -y

refreshenv

cd %USERPROFILE%\Documents
git clone https://github.com/jgolightly1/Warrant.git

pip install -r requirements.txt
