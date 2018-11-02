# Usage Tracking

This script was created as a way to collect some basic Windows computer usage for the purpose of billing.  Specifically, a researcher had a piece of expensive and unusual scientific equipment and wanted to recover some costs by charging for its use.

It is intended to be run as a startup script (with system privileges), but it will also run under (elevated) admin privileges.  
   
Each time the script runs, there is a potential for 5 files to be created in a local archive plus an optional 5 duplicate files in a shared, network location.  

1. One file will hold *all* logon/off records found in the security log at the time the script is run.  
2. One file will simply have the time stamp of the last time the script was run.  
3. The logon/off records of the current month
4. The logon/off records of the last month
5. The logon/off records of the previous-to-last month.  

The last two will overwrite previous logs *only* for those months it is certain the current set contains the entire month in question.
   
This script does not attempt to determine the amount of time (or even if) a user has locked or slept the computer.  It simply calculates the difference between user logon and logoff (or other user-session-ending) events.  As such, this script assumes that only one user account is logged in at a time. Therefore, **Fast-User-Switching MUST BE TURNED OFF** for the best accuracy.

In my case, the network share is Read-Only for the Billing Managers (we don't care if they can each read the others billing logs), and only the Usage Tracking computers have write properties to the share.  