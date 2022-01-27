## Summary
   Set and log a fairly readable/memorable but complex passphrase on an AD account.

## Description
The intended use of this script is to run it in a periodic scheduled task 
targeting a generic account that is dangerous if cracked and is shared 
among a group of people.  The password for these generic accounts does
need to be logged (for reference), but also transmitted through voice or 
handwritten notes.  Using a passphrase is easier to share and/or remember 
than a random string of characters.

The passphrase will combine random words from a list with one random letter
from each word capitalized and a random, non-alphabetic character used to 
separate the words.  If the dictionary to be used is too short, the phrase
will be further complicated by the replacement of characters by numerals.

This idea is loosely based on concepts demonstrated in two places:
    
    https://www.grc.com/haystack.htm
    http://xkcd.com/936/

**Note:** It is shocking how the mind finds meaning in randomly paired words,
      so it may be a good idea to remove words from the dictionary of
      potential impoliteness, violence, sexuality, morbidity or bias.

### Parameters
* **NetID**: The LogonName of the user whose password will be set.  For 
testing purposes, if the NetID is "password", an example passphrase will 
be generated and output to the console only but not logged.  
`Default: "password"`
* **Minimum**: The minimum length of the resulting password.  
`Default: 12`
* **LogPath**: Full path to the log of the passphrase.  For it to be 
useful, it probably should be on a network share with appropriately set 
ACL.  
`Default: <temp directory>\<NetID>.log`
* **Maximum**: The maximum length of the resulting password.  
`Default: 1.75*Minimum`
* **TimeStamp**: Include a timestamp for each passphrase in the log.   
`Default: False`
* **Overwrite**: The previous log file will be overwritten.  Without 
this switch, the created passphrase will be appended to the log.  
`Default: False`
* **Dictionary**: Path to the local text file containing a list of 
words to use for creating passphrases.  If this list is shorter than 
100 words, an internal, default list will be used instead.  Avoid 
dictionaries with words containing non-alphabetic characters (e.g. "-").  
`Default: <script directory>\Dictionary.txt`
* **Domain**: The FQDN for the domain in which the user resides.  
`Default: The domain of the computer`

### Examples
* `.\Set-MemorablePassphrase.ps1 -NetID tempvisitor`   
   This will set a 12+ character passphrase for the tempvisitor 
   account in the default domain and append it with a timestamp to 
   "tempvisitor.log" in the temp directory.
* `.\Set-MemorablePassphrase.ps1 -NetID tempvisitor -logpath \\FS1\Logs\VPass.txt -overwrite`   
   The same as above, but logged to a shared filepath containing 
   only the current passphrase.
* `.\Set-MemorablePassphrase.ps1 HDuser 20 \\FS1\Logs\Secret.txt -Dictionary .\LatinWords.txt -timestamp`   
   This will set the password for the HDuser account in the default 
   domain to a 20+ character passphrase derived from a custom dictionary 
   file and append the result (with a timestamp) to a text file on FS1.