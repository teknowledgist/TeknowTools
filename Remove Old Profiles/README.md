This script is an attempt to delete old user profiles with options to keep some.  For many years [Delprof2 by Helge Klein](https://helgeklein.com/free-tools/delprof2-user-profile-deletion-tool/) did an excellent job, but it is no longer being maintained and has a few issues.  Specifically:

* DelProf2 relies on the modified date of the `NTUSER.dat` file in a user's profile.  At some point, Windows started updating the last-modified dates causing DelProf2 (and anything else that watches that date) to not recognize that a profile is old.  For a while, Microsoft's own "Delete user profiles older than a specified number of days on a system restart" GPO setting was affected by this for a while!
* As of 2018, Delprof2 has issues with UWP apps on Windows 10/11 that will not be addressed.

Some users have written scripts to re-adjust the `NTUSER.dat` modification date so they can continue to use DelProf2, but does does it make sense to write a script to fix a third-party tool that is no longer supported and has known issues?

I don't think so.  

*My* only reason to use Delprof2 rather than GPO was because I could exclude specific profiles from deletion. In particular, I needed to keep local account profiles because computers were set up by vendors who had to occasionally log into their local support account and those profiles often have tools/documentation that we don't want to lose.

If I have to write a script to identify outdated profiles, it may as well just do the whole job.  This script utilizes the same date stored in the registry that the GPO now uses to determine old profiles and deletes those older than the specified age using the built-in `Remove-CimInstance` PowerShell method, but it also allows for exclusions to be defined.  While Delprof2 allowed for wildcards (* and ?) in the exclusions list this script accepts (requires) much more powerfull regular expressions -- multiple regexes too.  It can also exclude all local accounts from deletion without needing to name them or match them to regexes.



