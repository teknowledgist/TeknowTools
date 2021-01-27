Auto-reboot Process
=========

This page is intended to explain the Auto-reboot process I created for the college in which I work and to provide some troubleshooting starting points and technical information.  "Auto-reboot" is probably not the best name for it, because it is mostly a reminder and reboot deadline process, but that's the name it has.

#### Origin
SCCM does a great job pushing updates, but when I wanted to configure the SCCM client to auto-reboot when needed, I discovered that the notification time-window has a hard-coded maximum of 24 hours.  Also, once the pending reboot is recognized, it is somewhat non-trivial (and definitely non-immediate -- as is *everything* with SCCM) to stop it in case of an emergency.  

Many users in my workplace are "forgetful professor" types.  They are so focused on the teaching and research they love to do, that they don't consider details like computer (or even solar) schedules.  They have desktop computers that are always on with multiple, partially finished email/documents open and/or they have a multi-day computation running.  Their schedules may have them in the office only one day a week.  Forcing a reboot with a maximum, 24-hour window of notice would simply not be acceptable.  Disseminating information on when it is "safe" to start computations or configuring different SCCM client settings schedules for different systems would never work. 

The process I have created is insistent, yet forgiving and is not dependent (once deployed) on GPO (i.e. off campus  laptop computers still go through the process).  A local administrator (i.e. without IT approval) can configure a machine to never reboot, and an imminent forced reboot can be delayed within seconds by any logged in user in case of emergencies.  

#### Reasoning
It is my firm belief that users do not want to be vulnerable by not rebooting; They are just busy and need time to both accept the eventuality and prepare for the interruption.  They *will* come around to do it themselves with some firm, clear reminders of the need.  I also believe that providing IT Services with a more friendly, *client* relationship will generate more respect than providing them with a *consumer* relationship like what an inflexible mandate establishes.

#### Summary of the process
1. A PowerShell script is run at startup (or with elevated rights) that embeds itself into a VBS script and establishes several scheduled tasks:
    1. Logon task:  Runs one hour after any user logs in.
    2. Daily task:  Runs once daily as the SYSTEM account.
    3. Wake/Startup task:  Runs on boot and when the machines wakes from sleep.
    4. NIC task: Triggers on specific events in the event log and enables or disables network interfaces.
2. The first three tasks each run the same VBS script which checks for several indicators of a pending reboot.
    1. **If a reboot is _NOT_ pending** the script run via logon-task sets/modifies another scheduled task to re-run itself in 4 hours (within the user login session).  The script running via the other tasks does nothing.
    2. **If a reboot _is_ pending**, [an initial window](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/InitialWindow.png) opens giving the user a choice to reboot immediately, or "snooze" for 4 hours.  If the Windows is not the front-most window, a [balloon (Win7)](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/BalloonNotice.png) or [Toast (Win10)](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/Win10FirstToast.png) notice pops up informing the user of the need to reboot.
3. If the user chooses to "snooze", **they have now acknowledged** that they know the computer needs a reboot **and accepted** the reboot time shown in the window.
4. The acknowledgement initiates a deadline for automatically rebooting the system.  Whether the Friday is the current week or next depends on when the acknowledgement occurred (see [Feature #4](#Features)).
5. After 4 hours, [a second window](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/FurtherWindows.png) will appear with a countdown indicating the remaining time until auto-reboot.  Again, if the window is not in front, a [balloon (Win7)](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/BalloonNotice.png) or [Toast (Win10)](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/Win10FirstToast.png) notice will appear.
    1. "Snoozing" will set the window to appear again.
    2. The user selecting "Hold Reminders" is accepting responsibility for rebooting.  If they forget, the auto-reboot will still occur, but they will only get warnings of that in the last 4 hours.
6. By default, a machine with a pending reboot that has not been acknowledged when Microsoft's "Patch Tuesday" comes around will have it's network interfaces disabled.  They will automatically be re-enabled by the script on reboot.

###### First Window
![screenshot](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/InitialWindow.png)

###### Further Windows
![screenshot](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/FurtherWindows.png)

###### Windows 10 Balloon Notice
![screenshot](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/Win10FirstToast.png)

###### Windows 7 Balloon Notice
![screenshot](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/Win7NextBalloon.png)

#### Features
1. Runs with PowerShell v2+ and tested to work on Windows XP through Windows 10.
2. Many options can be set via command-line switches (including installing, uninstalling and testing when no reboot is pending), and all strings are in one section (starting at around line 600) to allow for easy localization. 
3. Every action of the script is tracked in the System Application log under its own log source (i.e. easy to filter).
4. Users will always have between 54 hours and 9 days of warning after they acknowledge the need for a reboot before an automatic reboot will occur.  
5. If the window is not in front, a balloon/toast notice will appear every 24 minutes.
6. If the initial notice is not acknowledged for approximately 52 hours, all windows will be minimized to force the notice window (which cannot be minimized) into focus.  Window minimization will also occur when less than 4 hours remain before auto-reboot.
7. Regular users cannot modify the script, but they can delay the reboot in urgent situations by deleting a pending-reboot log file.
8. Network interfaces will (by default) be automatically disabled if a reboot is still pending for a previous update when "Patch Tuesday" arrives.  

#### Troubleshooting

- A scheduled reboot (i.e. the user has acknowledged) can be reset before the next iteration to a pending reboot (potentially extending the reboot by a week) by deleting the "*C:\\ProgramData\\Institution\Reboot\\Shutdown.xml*" file.  This is the best, **occasional solution** for someone who needs to delay the reboot.
- The entire process can be permanently stopped at the next iteration by an administrator of the system placing a "NoReboot" file at the root of the C: drive. 
    - The file can be empty, and should not have an extension. 
    - The very first thing the script does is check for the file and abort if found.  No further scheduled tasks are set.  The impact on such a system doing number-crunching and/or data-collection should be negligible and will only occur once per log-in session and once per day.
    - This "NoReboot" block should only be used for systems for which 54 hours is not enough lead time without causing a loss of data – i.e. very few.  Try to limit who knows this information to responsible and technically savvy users who understand the implications of not rebooting.  Users who are simply annoyed by having to reboot don't count!
- The notice window can be closed without an acknowledgement by killing the "mshta.exe" process.  The underlying PowerShell script should continue as if the notice had closed after a time-out.  I.E. the scheduled task will run the script again after a one-minute delay. 
- The script can be manually run at any time by double-clicking it.  Any choices made when run manually will update the scheduled tasks to run again as if the script had been called by the task itself.
- The XML-based, pending-reboot log file can be modified at any time (as long as the xml syntax and case is preserved).
- The XML, log file may remain even after a reboot if the reboot was not initiated by the script.  The first time the script runs, it should rename the file to "*reboot*&lt;timestamp&gt;*.log*".  In this case, the &lt;timestamp&gt; will reflect when the machine last booted.
- The scheduled tasks that calls the script (one triggered by login, and one with a time trigger) are inside an &lt;OrgName&gt; folder of the Task Scheduler
- All actions are logged in the *Application* log (under *Windows Logs*) in the event viewer and with a source of "*&lt;Organization&gt; reboot Check*".
    

Technical Info
---

#### Logs

- The log file is in xml format.  It tracks information in nodes:
    - &lt;RebootReason&gt; = Which of the 4 possible pending reboot triggers were found.
    - &lt;TriggerPoint&gt; = The datetime the script recognized the pending reboot.
    - &lt;Acknowledgements&gt; = A collection of &lt;Stamp&gt; subnodes, each of which is a time-stamp for when the user clicked a button in a notice window.
        - The earliest acknowledgement &lt;Stamp&gt; has a "Mark = 0" attribute.
    - &lt;Shutdown&gt; = Whether a user ticked the "Shutdown instead of reboot" checkbox.
    - &lt;Quiet&gt; = Whether a user used the "OK, I will shutdown later" button.
    - &lt;RebootPoint&gt; = The deadline for a reboot.
    - &lt;ActNow&gt; = Whether the user clicked the "Reboot now" option **OR** the script auto-rebooted the system.
- When the script reboots the computer (automatically or by user choice), the "Shutdown.xml" file is renamed to "Reboot&lt;time-stamp&gt;.log".  If the user shuts-down/reboots the system through other means, the script will create a log file with a time-stamp of the last boot time as stored in the system logs.

#### Script

This is an all-in-one PowerShell script that embeds itself into a [polyglot](https://en.wikipedia.org/wiki/Polyglot_%28computing%29) VBScript which then runs the PowerShell (v2) code which calls further nested HTA code (containing more VBScript and Javascript).  The reasons for the complexity are as follows:

- I wanted to keep the most-secure, default [PowerShell execution policy](http://windowsitpro.com/powershell/use-execution-policies-control-what-powershell-scripts-can-be-run) at the "Restricted" level.  This level means PowerShell scripts will not run unless initiated by Group Policy or SCCM or are explicitly asked to run from within a PowerShell console or a command window with a specific switch argument.
- There is apparently *no way to call a hidden PowerShell script from a scheduled task!*  A bug/feature prevents the usually-working `-windowstyle hidden` switch from working when used in the task scheduler.  A big, black console window every 4 hours (even if there is no pending reboot) would not be appreciated by users.
- While most (probably all) of the PowerShell functionality could be accomplished by VBScript alone, I found a [ready-made PowerShell function](http://blogs.technet.com/b/heyscriptingguy/archive/2013/06/11/determine-pending-reboot-status-powershell-style-part-2.aspx) that determined if a reboot was pending and the project grew from there.  Also, VBScript is very wordy, difficult to read and generally sucks!  :-)
- Graphical interfaces are possible in PowerShell, but they are long and bulky.  I happened upon an [HTA with a countdown timer](http://www.itsupportguides.com/windows-7/windows-7-shutdown-message-with-countdown-and-cancel/) for rebooting a machine and modified it to meet our needs.  Adjusting a GUI in HTML and a little VBscript and Javascript is much easier that custom building a PowerShell GUI.
- It could have been configured to run from multiple script files, but I didn't want to keep track of a mess of files (and the obfuscation of polyglot code makes non-IT tinkerers less likely to fiddle).  Also, passing variables from one script to another is a challenge that is easily overcome with nested scripts (by modifying them on the fly).  

 A bit more detail on how the script works:

###### PowerShell:

1. The script first checks for a "noreboot" file and exits if found.
2. After initializing all strings, the script checks if it should be enabling/disabling the network interface based on switches that are used when the NIC task is triggered.
3. Determine how the script was started (manually, via VBS and/or scheduled task).
4. Uninstall the script (if requested by admin via switch).
5. Check for the last time the system was booted and compare it to the last time a pending reboot was noticed.
6. If this is running as startup script (or the `-set Start` argument is used):
    1. Create an Application EventLog source
    2. Clean up previous reboot notices and enable any disabled NICs
    3. Create/update the polyglot, VBS script necessary to run without a console window
    4. Create/update the logon, daily, wake (or-boot), and NIC tasks
7. If not running at startup, check if users are logged in (to allow for reboot).
8. Check for a pending reboot.  The checks are:
    1. Registry for: Component Based Servicing (Windows Vista/2008+)
    2. Registry for: Windows Update / Auto Update (Windows 2003+)
    3. WMI for:  DetermineIfRebootPending method of SCCM 2012 Clients
    4. Registry for: PendingFileRenameOperations (Windows 2003+).  *This has been disabled as it tried the patience of users because it was so frequent.*
9. If there is a pending reboot and the machine has not rebooted in 24 hours:
    1. Establish a task service that triggers on logoff events to cause the system to reboot instead.
    2. Check the level of importance given to Patch Tuesday and if necessary, write a log entry that should trigger the NIC task.
    3. Check for an active, "shutdown.xml" log file and ensure it was created **after** the last boot.
    4. Determine if the user has asked to not be bothered or if time is running out.
    5. Calculate the potential reboot deadline (if the user acknowledges).
    6. Export a long here-string with calculated internal variables into a one-time-use, custom, HTA file in a temp directory
    7. Start an asycnronous process to cycle watching for the notice window to be in front and open balloon notices.
    8. Start the HTA and wait.
    9. When the HTA is done:
        1. Adjust the ACL on the "Shutdown.xml" file (if exists) in case another user needs to modify it.
        2. Recalculate the deadline for rebooting.
        3. Check (and execute) if "Reboot Now" was selected (&lt;ActNow&gt; node)
5. Calculate when the task should be run again and create or modify the scheduled task to make it happen.
    1. If there is no pending reboot, the check will run again in 4 hours.
    2. The default interval between checks/notices is 4 hours, but at every point at which the remaining time before the deadline is less than twice the current interval, the interval becomes half again.  Thus, 4-8 hours left -&gt; 2 hour "snooze", 2-4 hours left -&gt; 1 hour "snooze", etc.

- Whenever a reboot is occuring (or has occured), the script will clean up by:
    - Renaming the .xml file to *reboot*&lt;YYYY.MM.DD-HH.mm&gt;*.log*
    - deleting all but the last, temporary .hta file (for troubleshooting)
    - deleting any logoff-capture task.
    - deleting scheduled tasks previously set for the current user.

###### [HTA (Hyper-Text Application)](https://en.wikipedia.org/wiki/HTML_Application)

An [HTA](https://en.wikipedia.org/wiki/HTML_Application) is essentially an HTML page (with CSS and scripting) that executes as a "fully trusted" application.  It runs in a special compatibility mode version of Internet Explorer and can be made to have no "standard" window controls (taskbar button, close/min/max buttons, etc.).  A few things to note about this HTA:

- Both VBScript and Javascript are included. 
    - VBScript is needed to make the beep noise when the notice appears (and for every second of the last minute) and to create/make changes to the XML log file because javascript is blocked from external file access.
    - Javascript handles the countdown operation and opening a link in another browser.  This might be possible in VBScript, but this was already written by someone else and VBScript sucks.
- The reboot logo:
    - Is converted to a PNG graphics file from [here](https://upload.wikimedia.org/wikipedia/commons/f/f5/Reset_button.svg)
    - Has then been encoded (using [this site](https://www.base64-image.de/) into [Base64](https://en.wikipedia.org/wiki/Base64) (with ~23,050 characters) to allow it to exist within this single text file.
    - Can be replaced with any Base64 version of a PNG graphic.
    - Is at around line #672 of the script.  Just look for the **really** long line. :smile:

###### VBScript polyglot:

- The VBScript has the entire PowerShell script as comment lines. 
- The VBScript first checks for the "noreboot" file and exits (if necessary) before starting PowerShell. 
- It reads itself into a variable, strips the lines of VBScript out of the beginning and end, removes the ' (single-quote) comment marker, and finally calls PowerShell to run the resulting string/codeblock with the "-windowstyle hidden" switch.


#### Issues
This is not a perfect system.  
- It doesn't control the system like SCCM or other tools, so a knowledgeable  local administrator can bypass it. (Of course, an admin can always do what they want, but that's a different problem.)  I don't feel that most users are militantly against rebooting, they just need firm reminders and time to accept the inevitability.   
- It is home-grown, so configurations/possibilities I don't know about or didn't think of could cause an issue.  The developer is readily available though.  :grin:

