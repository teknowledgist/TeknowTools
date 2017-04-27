In May, 2014, Microsoft released [an update](https://technet.microsoft.com/library/security/ms14-025) that disabled the ability to set local passwords directly through Group Policy.  Microsoft (wisely) made this change because changing local account passwords this way was never secure although it was the suggested method in the past and many people used it.  

Best practice recommendations include (among others): 
- to not have any accounts with admin rights on all workstations (domain administrator excepted, of course)
- to have unique, complex, closely-guarded, local admin passwords for every machine
- to change local admin passwords regularly.  

Unfortunately, it’s a challenge to fit best practice into practical reality.  It is not uncommon for a technician to need a local admin password, but not be able to look up a local admin password stored in some kind of database without trecking all the way back to the office.  What follows is my best ideas on marrying the practice with the real (higher-ed) world.

1. Create a script that will set a (nearly) unique the password based on a human-computable rubric.  (See [MasterPassword](https://ssl.masterpasswordapp.com/) for a possibly even better option.)
2. Put the script into a shared location that only computer/system accounts (and domain admins) can access.
    1. Group Policy security filtering set to "Domain Computers" for the GPO holding the script will do this.
4. Call it as a startup script via GPO (or run it from SCCM).

#### The bad
- **It is discoverable.**  A user with (legit or not) local admin rights to a domain computer has the ability to impersonate that computer's AD account and read the script.  Reading the script could allow an attacker to determine the password if they know or iterate though all the (limited) options.
- If a machine has an unusual attribute (e.g. it won't return the "apparent" serial number to WMI queries - VMs are especially tricky) have to be identified and dealt with manually.

#### The good
- The password is re-enforced at each startup.
- The password can be easily changed -- manually or automatically (via date reference) -- for all systems.
- An attacker who knows that the rubric includes the serial number would still need to get at least brief physical access to each machine they seek to attack as the WMI service (hence, the serial number) is not remotely accessible (to my knowledge) by default.

#### External security enhancements
- Disable (or leave disabled) the built-in "Administrator" account and create another.  I like the idea of a local admin username similar to most other domain users so that it doesn't stand out.
- Use a GPO to disable the local administrators group's default right to remote login.  [Reference here.](https://security.berkeley.edu/node/94?destination=node/94) Thus a user who is a local admin and wants to RDP into their machine would not be able to simply allow remote desktop.  They would have to explicitly add themselves to the list of users.  This should require attackers who can know/guess the password to be physically present at each machine they attack -- at least long enough to add the local admin account to the "Remote Desktop Users" group.  (Of course, if they can do that much, they could easily do all kinds of other bad things.)
- A GPO to disable debugging for administrators.  [Reference here.](https://www.sans.org/reading-room/whitepapers/testing/passthehash-attacks-tools-and-mitigation-33283)
- If you use Remote Administration, only open the firewall to the machines from which you will be administering.
- Ditto for the WMI service.
- If you use Remote Management (WinRM), consider establishing subscriptions to security event logs customized to only logins of the local administrator account in question.  [Reference here.](http://technet.microsoft.com/en-us/library/cc748890.aspx)  (I had a sample, custom, event-log xml filter that I'll re-create and put up here someday.)
- Consider [tracking all explicitly added administrator accounts](https://github.com/teknowledgist/TeknowTools/tree/master/TrackLocalAdmins) (i.e. not set by GPO/Domain).  

##### Comment
Personally, I've never like the idea of on-site technical support using a domain account to log in -- especially one with rights beyond the local workstation.  You don't know what the normal user has done (intentionally or not) to compromise security and it leaves behind pieces that make it easier for someone to attack that account.  A unique, local admin password that has been discovered is still less dangerous than an admin-of-all-workstations, domain account with a memorable, occasionally-changed password used by rushed technicians.  

That said, having a domain account that is an administrator on all workstations (but has no other rights) is mighty handy on occasion.  Securing that is the reason for [something like this.](https://github.com/teknowledgist/TeknowTools/tree/master/Passwords)
 