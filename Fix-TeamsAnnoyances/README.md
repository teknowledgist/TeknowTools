### Introduction

Users complain enough that first logins are slow, but Microsoft decided that they don't care and designed the machine/deployable install of Teams such that it auto-installs for every user the first time they log in.  This is an attempt to "fix" Teams to work the way it should:

* The Teams installer only runs when the user wants to use Teams for the first time.
* The user sees a "Microsoft Teams" icon in the Start Menu.
* Clicking the icon the first time will install Teams.
* Clicking subsequent times will start Teams.
* Although the installer and the application are different, the Start Menu icon (pinned or not) does not change/break like it would if attempting to modify it through some other scripted process.

### Files

* `README.txt` - This documentation.
* `Fix-MSTeamsAnnoyances.ps1` - Script to run during deployment or as an admin *after* the machine-wide install of Teams is complete.  
* `Start-MSTeams.ps1` - Script that starts Teams and/or installs Teams (for just one user) if not installed.
* `Teams.ico` - a handy Teams icon you can use (made via [BeCyIconGrabber](http://www.becyhome.de/download/BeCyIGrab230Eng.zip)).

### How to use

1. Download the [PS2EXE utility](https://github.com/MScholtes/TechNet-Gallery/tree/master/PS2EXE-GUI) (script or GUI) and create an .exe file from the `Start-MSTeams.ps1` file.  
   * Use the icon included here or make your own.
   * Name it something meaningful (like "Start-Teams.exe").  
2. Put the new .exe and `Fix-MSTeamsAnnoyances.ps1` in the same directory where they can be run locally or via a centralized utility (like SCCM).
3. Run the `Fix-MSTeamsAnnoyances.ps1` script as an admin or SYSTEM.

### What *should* happen

1. The registry entry to start Teams on login is removed. 
2. The "real" Teams installer is prevented from auto-starting Teams after install.
3. The new .exe file you created gets copied to: `$env:ProgramData\Organization\Scripts\`
4. A shortcut of that copy gets created at: `C:\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk`

### What `Start-MSTeams.ps1` does

1. Checks if the (machine-wide) Teams installer is installed.
2. Checks if Teams is installed for this user.
3. Locks the Start Menu shortcut for the custom .exe.  (This is necessary because when Teams installs, it pushes an icon into the Start Menu that overwrites the one we want there.)
4. Installs Teams for this user (if not already).
5. Starts Teams via the "normal" Teams Updater process (to ensure it is up-to-date).

