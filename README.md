# Update-RibbonADCache.ps1
Run or schedule this script to force an update of the AD cache on your Ribbon SBC

Visitors browsing <a href="https://greiginsydney.com/category/sonus/" target="_blank"> the Ribbon/Sonus category</a> on my blog will quickly notice that I've written about the Sonus/Ribbon REST interface on and off for years now, as it's such a handy way of peeking and poking into the SBC without interacting with the browser. And  so it came to pass last week that when presented with a challenge, I've again resorted to REST and PowerShell to deliver the fix.
### The challenge - refresh the AD Cache
My customer is planning on migrating a couple of thousand users to Skype for Business quite gradually, and for that we've proposed the fairly standard upstream model, with the SBC doing AD lookups to determine if a user is enabled for SfB and thus  decide where to send the call. If you're not familiar with it I drew a picture and added an explanation at the top of <a href="https://greiginsydney.com/tweaking-sonus-message-translations/" target="_blank"> this post back in 2014</a>.
The catch is that the SBC's minimum refresh period is an hour, and the customer doesn't want to be potentially waiting that long for it to kick in after a user's been migrated.
Theyre also big users of automation, and so logging into the SBC to clear the cache by hand isn't really an option.
### The fix
Thankfully Sonus added the option to clear the cache to their REST interface, and it didn't take me a lot of effort to "save-as" a script I'm working on (more on that soon) to come up with "Update-RibbonADCache.ps1".
### Examples
If you just run the script on its own, it will prompt you for the FQDN and credentials, then go about its business:
```powershell
PS C:\> .\Update-RibbonADCache.ps1
About to login
SBC FQDN                          : mysbc.greigin.sydney
REST login name                   : REST
REST password                     : Pa$$w0rd
Login successful
Refresh of the Cache requested

DomainController : davros.greigin.sydney
BackupStatus     : Backup Not Applicable
CacheStatus      : Cache Not Applicable
ID               : 1
ADStatus         : AD Up

DomainController : davros.greigin.sydney
BackupStatus     : Backup Successful
CacheStatus      : Cache Active
ID               : 2
ADStatus         : AD Up

PS C:\>
```
In the above results you'll see Davros (yes, he's one of our DCs here) listed twice. That's because he exists in the SBC's config twice, once for Authentication and again for Call Routing - and that's why one instance reports a status of "Cache Not Applicable".
If you give it the lot in the one command it will do the deed, then query the SBC for an update:
```powershell
PS C:\> .\Update-RibbonADCache.ps1 -SkipUpdateCheck -SbcFQDN mySweLite.greigin.sydney -RestLogin REST -RestPassword Pa$$w0rd 
About to login
Login successful
Refresh of the Cache requested

DomainController : davros.greigin.sydney
BackupStatus     : Backup Successful
CacheStatus      : Cache Active
ID               : 1
ADStatus         : AD Up
```
PS C:\>
The script outputs a PowerShell Object, so you can use it in downstream tests, or pipe it to "ft" to customise the output display:
```powershell
DomainController       BackupStatus          CacheStatus          ID ADStatus
----------------       ------------          -----------          -- --------
davros.greigin.sydney  Backup Not Applicable Cache Not Applicable 1  AD Up
davros.greigin.sydney  Backup Successful     Cache Active         2  AD Up
```
If you add the "-QueryOnly" switch to the above it won't trash the cache, just query the status to check it's OK. If you have some kind of automated health checks, this might be a good one to add to your schedule!
### It logs too!
With his generous consent I've ~~stolen~~ included [Pat Richard's]("https://ucunleashed.com") logging function, so you'll find relatively detailed logs - sans password  though, naturally - in the /logs/ folder it creates where the script lives.
Running it with the extra -verbose or -debug switches I've included will spray more info to screen if you're needing some assistance debugging it, but hopefully you won't encounter too many problems beyond the obvious issues with typos in the FQDN or bad  REST credentials.
### Did it Work?
The SBC's Alarm/Event History is another way you can confirm the script reset the cache, and if you're automating it this will be an effective human-viewable way of keeping tabs on it.
<img title="SBC-AlarmEventHistory" src="https://user-images.githubusercontent.com/11004787/79121380-84a47600-7dd8-11ea-8fce-719d723d7488.png" border="0" alt="SBC-AlarmEventHistory" width="600" />

You can also have the SBC send these events as SNMP traps to your NMS. And no I don't know why my DC was uncontactable at midnight. Puzzling.
### Auto Update
I've added an update-checking component so it will let you know as updates become available. You can suppress the update check by running it with the "-SkipUpdateCheck" parameter, which you should remember to add if you're running the script  via a scheduled task or some other automated/unattended means.
### Revision History
1st August 2018. This is the initial release.  
 
<br>

\- G.

<br>

This script was originally published at [https://greiginsydney.com/update-ribbonadcache-ps1](https://greiginsydney.com/update-ribbonadcache-ps1).

