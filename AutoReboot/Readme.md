Auto-reboot Process
=========

This page is intended to explain the Auto-reboot process I created for the college in which I work and to provide some troubleshooting starting points and technical information.

#### Philosophy
Users don't want to totally avoid necessary reboots and remain vulnerable; They just need time to accept the eventuality and want to do it somewhat on their schedule.  With firm, clear reminders of the need, they will comply and feel better about their IT services than they would with an authoritarian mandate. 

#### Why
My workplace uses SCCM to push updates, but when I wanted to configure the SCCM client to auto-reboot based on need, I discovered that the notification time-window has a hard-coded maximum of 24 hours.  Also, once the pending reboot is recognized, it is somewhat non-trivial (and definitely non-immediate -- as is *everything* with SCCM) to stop it in case of an emergency.  

Many users in my workplace are "forgetful professor" types.  They are so focused on the cutting-edge or educational problems/research they are paid to do, they don't consider details like computer (or even solar) schedules.  They have desktop computers that are always on with multiple, partially finished email/documents open.  Their schedules may have them in the office only some days and/or they have a somewhat unexpected need to run multi-day computations.  Forcing a reboot with such a "small" window of notice would simply not be acceptable.  Setting up different SCCM client settings schedules and keeping track of (and modifying!) who should be in which and/or disseminating information on when it is "safe" to start computations would be a nightmare too.  And forget teaching them anything.  :wink:

The process I have created and outlined here is very forgiving and not dependent (once deployed) on GPO (i.e. off campus  laptop computers still go through the process).  A local administrator (i.e. without IT approval) can configure a machine to never reboot, and the process can be delayed/halted within seconds by any logged in user in case of emergencies.  


#### Summary of the process
1. A scheduled task is pushed at start-up through Group Policy (running a script) to run one hour after any user logs in.  If changes need to be made to the script without waiting for machines to reboot, it can also be pushed through SCCM.
2. The task runs a local script (that is "installed" by the former script) that checks for several indicators of a pending reboot.
    1. **If a reboot is _NOT_ pending**, the script sets/modifies another scheduled task to re-run itself in 4 hours.
    2. **If a reboot _is_ pending**, [an initial window](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/InitialWindow.png) pops up informing the user of the need to reboot and gives them a choice to reboot immediately, or "snooze" for 4 hours.
3. If the user chooses to "snooze", **they have now acknowledged** that they know the computer needs a reboot.
4. The acknowledgement initiates a deadline (defaults to Friday at 5pm) for automatically rebooting the system.  Whether the Friday is the current week or next depends on when the acknowledgement occurred (see [Note #2](#Notes)).
5. After 4 hours, [a second window](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/FurtherWindows.png) will appear with a countdown indicating the remaining time until auto-reboot.
    1. "Snoozing" will set the window to appear again.
    2. The user selecting "I'll reboot later" is accepting responsibility for rebooting.  If they forget, the auto-reboot will still occur, but they will only get warnings of that in the last 4 hours.

###### First Window
![screenshot](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/InitialWindow.png)

###### Further Windows
![screenshot](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/FurtherWindows.png)

###### Balloon Notice
![screenshot](https://cdn.rawgit.com/teknowledgist/TeknowTools/master/AutoReboot/BalloonNotice.png)

#### Notes
0. This does require PowerShell v2+, but otherwise, the code works (in my experience) on Windows XP through Windows 10.
1. As provided here, this script is "neutered" for testing/viewing purposes.  It will always think the machine has a pending reboot and will not reboot the machine if requested (or time runs out).   
    - The pending reboot check can be re-enabled by swapping the comment/uncomment marks around line 347-349.
    - Reboot/shutdown can be re-enabled by swapping the comment marks around lines 271-273.
2. Users will always have between 54 hours and 9 days of warning after they acknowledge the need for a reboot before an automatic reboot will occur.  Thus, an initial acknowledgement on a Thursday morning will set the deadline to the Friday, 8 days hence.
3. If the window is not in front, a balloon notice will appear every 24 minutes.
4. If the initial notice is not acknowledged for approximately 52 hours, all windows will be minimized to force the notice window (which cannot be minimized) into focus.  Window minimization will also occur when less than 4 hours remain before auto-reboot.
5. The script and log files default to "*C:\\ProgramData\\Institution\Reboot\\*". 
    - The script (*CheckAutoReboot.vbs*) can only be modified by admins and will be updated/recreated at next login by Group Policy if the *Initialize-AutoReboot.ps1* script is a startup script.
    - The active, log file (*Shutdown.xml*) can be modified by any user of the machine.

#### Troubleshooting

- A scheduled reboot (i.e. the user has acknowledged) can be reset before the next iteration to a pending reboot (potentially extending the reboot by a week) by deleting the "*C:\\ProgramData\\Institution\Reboot\\Shutdown.xml*" file.  This is the best, **occasional solution** for someone who needs to delay the reboot.
- The entire process can be permanently stopped at the next iteration by an administrator of the system placing a "NoReboot" file at the root of the C: drive. 
    - The file can be empty, and should not have an extension. 
    - The very first thing the script does is check for the file and abort if found.  No further scheduled tasks are set.  The impact on such a system doing number-crunching and/or data-collection should be negligible and will only occur once again per log-in session.
    - This "NoReboot" block should only be used for systems for which 54 hours is not enough lead time without causing a loss of data – i.e. very few.  Try to limit who knows this information to responsible and technically savvy users who understand the implications of not rebooting.  Users who are simply annoyed by having to reboot don't count!
- The notice window can be closed without an acknowledgement by killing the "mshta.exe" process.  The underlying PowerShell script should continue as if the notice had closed after a time-out.  I.E. the scheduled task will run the script again after a one-minute delay. 
- The script can be manually run at any time by double-clicking it.  Any choices made when run manually will update the scheduled tasks to run again as if the script had been called by the task itself.
- The XML, log file can be modified at any time (as long as the xml syntax and case is preserved).
- The XML, log file may remain even after a reboot if the reboot was not initiated by the script.  The first time the script runs (an hour after login), it will rename the file to "*reboot*&lt;timestamp&gt;*.log*".  In this case, the &lt;timestamp&gt; will reflect when the machine last booted.
- The scheduled tasks that calls the script (one triggered by login, and one with a time trigger) are inside an &lt;OrgName&gt; folder of the Task Scheduler and named:
    - Initialize auto-reboot checks
    - Check for Pending Reboot - &lt;username&gt;
    
    If they are deleted, they will be recreated the next time the machine boots (assuming *Initialize-AutoRebootTask.ps1* is a startup script).
    - Note: There could be scheduled tasks with a name that does not match the current user if the machine was rebooted outside of the script and a different user logged in.  The mismatched tasks shouldn't cause a problem because they will have a scheduled run time that has either passed, or is within 24 hours of a reboot (canceling any notice).  This may be resolved in a later version of the script.

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

This is a one-file, [polyglot](https://en.wikipedia.org/wiki/Polyglot_%28computing%29) script with VBScript calling nested PowerShell (v2) code calling further nested HTA code (containing more VBScript and Javascript).  The reasons are as follows:

- I wanted to keep the most-secure, default [PowerShell execution policy](http://windowsitpro.com/powershell/use-execution-policies-control-what-powershell-scripts-can-be-run) at the "Restricted" level.  This level means PowerShell scripts will not run unless initiated by Group Policy or SCCM or are explicitly asked to run from within a PowerShell console or a command window with a specific switch argument.
- There is apparently no way to call a PowerShell script from a scheduled task (using the cmd switch) and have the PoSh window be hidden.  There is a bug/feature that prevents the usually-working "-windowstyle hidden" switch from working when used in the task scheduler.  A big, black console window every 4 hours (even if there is no pending reboot) would not be appreciated by users.
- While most (probably all) of the PowerShell functionality could be accomplished by VBScript alone, I found a [ready-made PowerShell function](http://blogs.technet.com/b/heyscriptingguy/archive/2013/06/11/determine-pending-reboot-status-powershell-style-part-2.aspx) that determined if a reboot was pending and the project grew from there.  Also, VBScript is very wordy, difficult to read and generally sucks!  :-)
- Graphical interfaces are possible in PowerShell, but they are long and bulky.  I happened upon an [HTA with a countdown timer](http://www.itsupportguides.com/windows-7/windows-7-shutdown-message-with-countdown-and-cancel/) for rebooting a machine and modified it to meet our needs.  Adjusting a GUI in HTML and a little VBscript and Javascript is much easier that custom building a PowerShell GUI.
- It could have been configured to run from multiple script files, but I didn't want to keep track of a mess of files (and the obfuscation of polyglot code makes non-IT tinkerers less likely to fiddle).  Also, passing variables from one script to another is a challenge that is easily overcome with nested scripts (by modifying them on the fly).  

 A bit more detail on how the script works:

###### Outer VBScript:

- The VBScript, outer shell simply has the entire PowerShell script as comment lines. 
- The VBScript first checks for the "noreboot" file and exits if it finds it. 
- Then it reads itself into a variable, strips the lines of VBScript out of the beginning and end, removes the ' (single-quote) comment marker, and finally calls PowerShell to run the resulting string/codeblock with the "-windowstyle hidden" switch.

###### PowerShell:

1. The Powershell script again checks for a "noreboot" file (exiting if found).
2. A check for the last time the system was booted.
3. A check is made for a pending reboot.  The checks are:
    1. Registry for: Component Based Servicing (Windows Vista/2008+)
    2. Registry for: Windows Update / Auto Update (Windows 2003+)
    3. WMI for:  DetermineIfRebootPending method of SCCM 2012 Clients
    4. Registry for: PendingFileRenameOperations (Windows 2003+).  *This has been disabled as it tried the patience of users because it was so frequent.*
4.  If there is a pending reboot and the machine has not rebooted in 24 hours:
    1. A check for an active, "shutdown.xml" log file and ensure it was created **after** the last boot.
    2. Establish a task service that triggers on logoff events to cause the system to reboot instead.
    3. Determine if the user has asked to not be bothered or if time is running out.
    4. Export a long here-string with calculated internal variables into a one-time-use, custom, HTA file in a temp directory
    5. Start an asycnronous process to cycle watching for the notice window to be in front and open balloon notices.
    6. Start the HTA and wait.
    7. When the HTA is done:
        1. Adjust the ACL on the "Shutdown.xml" file (if exists) in case another user needs to modify it.
        2. Calculate when the deadline for rebooting is (this or next Friday?).
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
    - Can be replaced with any Base64 version of a PNG graphic.  (I use the official, University logo internally.)
    - Is at around line #656 of the .vbs.  Just turn off wrapping and look for the **really** long line. :smile:

Deployment
---
While I wanted to use a Group Policy Preference to create the initial scheduled task (that runs one hour after a user logs in), there are apparently bugs/quirks in GPP that prevented the specific settings in the task from being pushed out.  Specifically, defining the task as system-wide (to run for every user) to run as the logged in user, doesn't work.  GPP wants to run it as system/administrator.  To resolve this problem, a schedule task was manually made and exported to XML and a short PowerShell startup script (*Initialize-AutoRebootTask.ps1*) creates the initial scheduled task by directly importing the XML as a scheduled task. 

The script file itself is copied from a software deployment point (\\\\server\\sdp\$\\AutoReboot) whether deployed through GPO or SCCM.  Any changes to the script in that location should minimally start populating machines as they reboot.

#### Issues
This is not a perfect system.  
- It doesn't control the system like SCCM or other tools, so a knowledgeable  local administrator can bypass it. (Of course, an admin can always do what they want, but that's a different problem.)  I don't feel that most users are militantly against rebooting, they just need firm reminders and time to accept the inevitability.  
- It won't reboot idle systems.  Generally, an unused system is at significantly less risk.
- It is home-grown, so configurations/possibilities I don't know about or didn't think of could cause an issue.  The developer is readily available though.  :-)

