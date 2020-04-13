# Update-RibbonADCache.ps1
Run or schedule this script to force an update of the AD cache on your Ribbon SBC

<p>Visitors browsing <a href="https://greiginsydney.com/category/sonus/" target="_blank"> the Ribbon/Sonus category</a> on my blog will quickly notice that I&rsquo;ve written about the Sonus/Ribbon REST interface on and off for years now, as it&rsquo;s such a handy way of peeking and poking into the SBC without interacting with the browser. And  so it came to pass last week that when presented with a challenge, I&rsquo;ve again resorted to REST and PowerShell to deliver the fix.</p>
<h3>The challenge &ndash; refresh the AD Cache</h3>
<p>My customer is planning on migrating a couple of thousand users to Skype for Business quite gradually, and for that we&rsquo;ve proposed the fairly standard upstream model, with the SBC doing AD lookups to determine if a user is enabled for SfB and thus  decide where to send the call. If you&rsquo;re not familiar with it I drew a picture and added an explanation at the top of <a href="https://greiginsydney.com/tweaking-sonus-message-translations/" target="_blank"> this post back in 2014</a>.</p>
<p>The catch is that the SBC's minimum refresh period is an hour, and the customer doesn&rsquo;t want to be potentially waiting that long for it to kick in after a user&rsquo;s been migrated.</p>
<p>They&rsquo;re also big users of automation, and so logging into the SBC to clear the cache by hand isn&rsquo;t really an option.</p>
<h3>The fix</h3>
<p>Thankfully Sonus added the option to clear the cache to their REST interface, and it didn&rsquo;t take me a lot of effort to &ldquo;save-as&rdquo; a script I&rsquo;m working on (more on that soon) to come up with &ldquo;Update-RibbonADCache.ps1&rdquo;.</p>
<h3>Examples</h3>
<p>If you just run the script on its own, it will prompt you for the FQDN and credentials, then go about its business:</p>
<pre>PS C:\&gt; .\Update-RibbonADCache.ps1
About to login
SBC FQDN&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : mysbc.greigin.sydney
REST login name&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : REST
REST password&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : Pa$$w0rd
Login successful
Refresh of the Cache requested

DomainController : davros.greigin.sydney
BackupStatus&nbsp;&nbsp;&nbsp;&nbsp; : Backup Not Applicable
CacheStatus&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : Cache Not Applicable
ID&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : 1
ADStatus&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : AD Up

DomainController : davros.greigin.sydney
BackupStatus&nbsp;&nbsp;&nbsp;&nbsp; : Backup Successful
CacheStatus&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : Cache Active
ID&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : 2
ADStatus&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : AD Up

PS C:\&gt;</pre>
<p>In the above results you'll see Davros (yes, he's one of our DCs here) listed twice. That's because he exists in the SBC's config twice, once for Authentication and again for Call Routing - and that's why one instance reports a status of "Cache Not Applicable".</p>
<p>If you give it the lot in the one command it will do the deed, then query the SBC for an update:</p>
<pre>PS C:\&gt; .\Update-RibbonADCache.ps1 -SkipUpdateCheck -SbcFQDN mySweLite.greigin.sydney -RestLogin REST -RestPassword Pa$$w0rd 
About to login
Login successful
Refresh of the Cache requested

DomainController : davros.greigin.sydney
BackupStatus&nbsp;&nbsp;&nbsp;&nbsp; : Backup Successful
CacheStatus&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : Cache Active
ID&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : 1
ADStatus&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : AD Up
<br /></pre>
<pre>PS C:\&gt;</pre>
<p>The script outputs a PowerShell Object, so you can use it in downstream tests, or pipe it to "ft" to customise the output display:</p>
<pre>DomainController       BackupStatus          CacheStatus          ID ADStatus
----------------       ------------          -----------          -- --------
davros.greigin.sydney  Backup Not Applicable Cache Not Applicable 1  AD Up
davros.greigin.sydney  Backup Successful     Cache Active         2  AD Up</pre>
<p>If you add the &ldquo;-QueryOnly&rdquo; switch to the above it won&rsquo;t trash the cache, just query the status to check it&rsquo;s OK. If you have some kind of automated health checks, this might be a good one to add to your schedule!</p>
<h3>It logs too!</h3>
<p>With his generous consent I've <span style="text-decoration: line-through;">stolen</span> included&nbsp;<a rel="noopener" href="https://ucunleashed.com" target="_blank">Pat Richard's</a> logging function, so you'll find relatively detailed logs - sans password  though, naturally - in the /logs/ folder it creates where the script lives.</p>
<p>Running it with the extra -verbose or -debug switches I've included will spray more info to screen if you're needing some assistance debugging it, but hopefully you won't encounter too many problems beyond the obvious issues with typos in the FQDN or bad  REST credentials.</p>
<h3>Did it Work?</h3>
<p>The SBC&rsquo;s Alarm/Event History is another way you can confirm the script reset the cache, and if you&rsquo;re automating it this will be an effective human-viewable way of keeping tabs on it.</p>
<p><a href="https://greiginsydney.com/wp-content/uploads/2018/08/SBC-AlarmEventHistory.jpg"><img title="SBC-AlarmEventHistory" src="https://greiginsydney.com/wp-content/uploads/2018/08/SBC-AlarmEventHistory.jpg" border="0" alt="SBC-AlarmEventHistory" width="600" /></a></p>
<p>You can also have the SBC send these events as SNMP traps to your NMS. And no I don&rsquo;t know why my DC was uncontactable at midnight. Puzzling.</p>
<h3>Auto Update</h3>
<p>I've added an update-checking component so it will let you know as updates become available. You can suppress the update check by running it with the &ldquo;-SkipUpdateCheck&rdquo; parameter, which you should remember to add if you&rsquo;re running the script  via a scheduled task or some other automated/unattended means.</p>
<h3>Revision History</h3>
<p>1st August 2018. This is the initial release. &nbsp;</p>
<p>&nbsp;</p>
<p>- G.</p>
