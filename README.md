# Warrant
-------------------------------------------------------------------------------
INSTALL INSTRUCTIONS
-------------------------------------------------------------------------------
---(Menu > Type: cmd > right-click Command Prompt > Run as Administrator > Copy the Following Commands:)

@"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command " [System.Net.ServicePointManager]::SecurityProtocol = 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"

choco install git --version 2.39.2 -y

refreshenv

cd %USERPROFILE%\Documents

git clone https://github.com/jgolightly1/Warrant.git

---Leave the command prompt open, we'll come back to it in a minute...

-------------------------------------------------------------------------------
> Open File Explorer and go to your Documents folder. There should be a new folder called "Warrant"
> Open the folder called "PreRequirements" and click to install Python. The windows should look like this:
![win_installer](https://user-images.githubusercontent.com/32847002/221447750-ab62647c-9b4f-479c-a9bd-583df71d2adb.png)

>Make sure "Add Python to Path" at the bottom is CHECKED then click "Install Now".
>Click through the prompts to install.
>When the installation is finished, click the option to "Disable Path Limit"

-------------------------------------------------------------------------------
---Go back to the command prompt:

refreshenv

cd %USERPROFILE%\Documents\Warrant

pip install -r requirements.txt

mkdir "%USERPROFILE%\Desktop\Search Warrant Tool"

cd Shortcuts

copy  "CDR Warrants.lnk" "%USERPROFILE%\Desktop\Search Warrant Tool\"

copy  "General Warrants.lnk" "%USERPROFILE%\Desktop\Search Warrant Tool\"

copy  "Inventory Warrants.lnk" "%USERPROFILE%\Desktop\Search Warrant Tool\"

copy  "Social Warrants.lnk" "%USERPROFILE%\Desktop\Search Warrant Tool\"

-------------------------------------------------------------------------------

--- Completed warrants are in the "Documents\Warrant\Complete" folder
