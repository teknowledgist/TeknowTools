Sometimes you want a Scheduled task that all users can run, but only Admins can modify.

Earlier versions of Windows (apparently) used file permissions on the task files contained in `C:\Windows\System32\Tasks` to manage security.  Although those file permissions still appear to control visibility of scheduled tasks, Windows 10 now controls the ability to run tasks using the registry value of *Security Descriptor* in the listed tasks under `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree`.    

This script will modify the registry value to grant the *Local Users* group read and execute permissions to the task. 

Because the registry key is protected, this is intended to be run as SYSTEM, 
but an admin user can successfully run it too. (It's just messier).

[Inspiration/initial code found here](https://social.technet.microsoft.com/Forums/windows/en-US/6b9b7ac3-41cd-419e-ac25-c15c45766c8e/scheduled-task-that-any-user-can-run)