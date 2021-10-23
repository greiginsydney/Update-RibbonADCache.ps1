<#
.SYNOPSIS
This script will login to the REST account in your SBC and refresh the AD cache.

.DESCRIPTION
This script will login to the REST account in your SBC and refresh the AD cache. Add the "-QueryOnly" switch for it to look but not touch.
Outputs an object per configured Domain Controller with the relevant settings/results.


.NOTES
	Version				: 1.1
	Date				: 23rd October 2021
	Author    			: Greig Sheridan

	Wishlist / TODO:
		#?

	Revision History:
			v1.1: 23rd October 2021
				Changed fn 'BasicHandler' to accept varying 'xml' responses from the SBC
				Updated the REST queries to all loop through BasicHandler

			v1.0: 1st August 2018
				With thanks to Pat Richard for the auto-update and logging modules.



.LINK
	https://greiginsydney.com/Update-RibbonADCache.ps1

.EXAMPLE
	.\Update-RibbonADCache.ps1

	Description
	-----------
	With no input parameters passed to it, the script will prompt you for an SBC FQDN & some REST credentials before refreshing the cache, then
	querying the status and reporting the output to screen.

.EXAMPLE
	.\Update-RibbonADCache.ps1 -SbcFQDN mySbc.greigin.sydney -RestLogin REST -RestPassword P@ssword1 -QueryOnly -SkipUpdateCheck

	Description
	-----------
	Running the script with the above combination of parameters will execute a non-invasive health check of your SBC's AD cache. Capture the returned
	object and add this to your daily automatic health checks!


.PARAMETER SbcFQDN
	String. The FQDN of your SBC.

.PARAMETER RestLogin
	String. The REST login name. (Set this in the SBC under Security / Users / Local User Management).

.PARAMETER RestPassword
	String. The REST password. Set with the above.

.PARAMETER QueryOnly
	Boolean. Set this and the script will only query and then report the status of the configured domain controllers.

.PARAMETER SkipUpdateCheck
	Boolean. Skips the automatic check for an Update. Courtesy of Pat: http://www.ucunleashed.com/3168

#>

[CmdletBinding(SupportsShouldProcess = $False)]
Param(

	[string]$SbcFQDN,
	[string]$RestLogin,
	[string]$RestPassword,
	[switch]$QueryOnly,

	[switch]$SkipUpdateCheck
)


#--------------------------------
# Setup hash tables--------------
#--------------------------------

$ADStatusLookup = @{'0' = 'AD Up'; '1' = 'AD Down'}
$ADCacheStatusLookup = @{'0' = 'Cache Disabled'; '1' = 'Cache Building'; '2' = 'Cache Updating'; '3' = 'Cache Active'; '4' = 'Cache Failed'; `
						'5' = 'Cache Backup'; '6' = 'Cache Truncated'; '7' = 'Cache Not Applicable'; '8' = 'Cache Incomplete'}
$ADBackupStatusLookup = @{'0' = 'Backup Successful'; '1' = 'Backup Failed'; '2' = 'Backup Disabled'; '3' = 'Backup Not Applicable'; '4' = 'Backup Truncated'; `
						'5' = 'Backup Updating'}


#--------------------------------
# START FUNCTIONS ---------------
#--------------------------------
#region functions

function Get-UpdateInfo
{
  <#
	  .SYNOPSIS
	  Queries an online XML source for version information to determine if a new version of the script is available.
	  *** This version customised by Greig Sheridan. @greiginsydney https://greiginsydney.com ***

	  .DESCRIPTION
	  Queries an online XML source for version information to determine if a new version of the script is available.

	  .NOTES
	  Version               : 1.2 - See changelog at https://ucunleashed.com/3168 for fixes & changes introduced with each version
	  Wish list             : Better error trapping
	  Rights Required       : N/A
	  Sched Task Required   : No
	  Lync/Skype4B Version  : N/A
	  Author/Copyright      : © Pat Richard, Office Servers and Services (Skype for Business) MVP - All Rights Reserved
	  Email/Blog/Twitter    : pat@innervation.com  https://ucunleashed.com  @patrichard
	  Donations             : https://www.paypal.me/PatRichard
	  Dedicated Post        : https://ucunleashed.com/3168
	  Disclaimer            : You running this script/function means you will not blame the author(s) if this breaks your stuff. This script/function
							is provided AS IS without warranty of any kind. Author(s) disclaim all implied warranties including, without limitation,
							any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use
							or performance of the sample scripts and documentation remains with you. In no event shall author(s) be held liable for
							any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss
							of business information, or other pecuniary loss) arising out of the use of or inability to use the script or
							documentation. Neither this script/function, nor any part of it other than those parts that are explicitly copied from
							others, may be republished without author(s) express written permission. Author(s) retain the right to alter this
							disclaimer at any time. For the most up to date version of the disclaimer, see https://ucunleashed.com/code-disclaimer.
	  Acknowledgements      : Reading XML files
							http://stackoverflow.com/questions/18509358/how-to-read-xml-in-powershell
							http://stackoverflow.com/questions/20433932/determine-xml-node-exists
	  Assumptions           : ExecutionPolicy of AllSigned (recommended), RemoteSigned, or Unrestricted (not recommended)
	  Limitations           :
	  Known issues          :

	  .EXAMPLE
	  Get-UpdateInfo -Title 'Update-RibbonADCache.ps1'

	  Description
	  -----------
	  Runs function to check for updates to script called 'Update-RibbonADCache.ps1'.

	  .INPUTS
	  None. You cannot pipe objects to this script.
  #>
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
	[string] $title
	)
	try
	{
		[bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
		if ($HasInternetAccess)
		{
			write-verbose -message 'Performing update check'
			# ------------------ TLS 1.2 fixup from https://github.com/chocolatey/choco/wiki/Installation#installing-with-restricted-tls
			$securityProtocolSettingsOriginal = [Net.ServicePointManager]::SecurityProtocol
			try {
			  # Set TLS 1.2 (3072). Use integers because the enumeration values for TLS 1.2 won't exist in .NET 4.0, even though they are
			  # addressable if .NET 4.5+ is installed (.NET 4.5 is an in-place upgrade).
			  [Net.ServicePointManager]::SecurityProtocol = 3072
			} catch {
			  write-verbose -message 'Unable to set PowerShell to use TLS 1.2 due to old .NET Framework installed.'
			}
			# ------------------ end TLS 1.2 fixup
			[xml] $xml = (New-Object -TypeName System.Net.WebClient).DownloadString('https://greiginsydney.com/wp-content/version.xml')
			[Net.ServicePointManager]::SecurityProtocol = $securityProtocolSettingsOriginal #Reinstate original SecurityProtocol settings
			$article  = select-XML -xml $xml -xpath ("//article[@title='{0}']" -f ($title))
			[string] $Ga = $article.node.version.trim()
			if ($article.node.changeLog)
			{
				[string] $changelog = 'This version includes: ' + $article.node.changeLog.trim() + "`n`n"
			}
			if ($Ga -gt $ScriptVersion)
			{
				$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
				$updatePrompt = $wshell.Popup(("Version {0} is available.`n`n{1}Would you like to download it?" -f ($ga), ($changelog)),0,'New version available',68)
				if ($updatePrompt -eq 6)
				{
					Start-Process -FilePath $article.node.downloadUrl
					write-warning -message "Script is exiting. Please run the new version of the script after you've downloaded it."
					exit
				}
				else
				{
					write-verbose -message ('Upgrade to version {0} was declined' -f ($ga))
				}
			}
			elseif ($Ga -eq $ScriptVersion)
			{
				write-verbose -message ('Script version {0} is the latest released version' -f ($Scriptversion))
			}
			else
			{
				write-verbose -message ('Script version {0} is newer than the latest released version {1}' -f ($Scriptversion), ($ga))
			}
		}
		else
		{
		}

	} # end function Get-UpdateInfo
	catch
	{
		write-verbose -message 'Caught error in Get-UpdateInfo'
		if ($Global:Debug)
		{
			$Global:error | Format-List -Property * -Force #This dumps to screen as white for the time being. I haven't been able to get it to dump in red
		}
	}
}


function Write-Log {
  <#
	  .SYNOPSIS
	  Extensive function to write data to either the console screen, a log file, and/or a Windows event log.

	  .DESCRIPTION
	  Extensive function to write data to either the console screen, a log file, and/or a Windows event log. Data can be written as info, warning, error, and includes indentation, time stamps, etc.

	  .NOTES
	  Version               : 3.2
	  Wish list             : Better error trapping
	  Rights Required       : Local administrator on server if writing to event log(s)
	  Sched Task Required   : No
	  Lync/Skype4B Version  : N/A
	  Author/Copyright      : © Pat Richard, Office Servers and Services (Skype for Business) MVP - All Rights Reserved
	  Email/Blog/Twitter		: pat@innervation.com 	https://www.ucunleashed.com @patrichard
	  Donations             : https://www.paypal.me/PatRichard
	  Dedicated Post        : http://poshcode.org/6894
	  Disclaimer            : You running this script/function means you will not blame the author(s) if this breaks your stuff. This script/function
						  is provided AS IS without warranty of any kind. Author(s) disclaim all implied warranties including, without limitation,
						  any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use
						  or performance of the sample scripts and documentation remains with you. In no event shall author(s) be held liable for
						  any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss
						  of business information, or other pecuniary loss) arising out of the use of or inability to use the script or
						  documentation. Neither this script/function, nor any part of it other than those parts that are explicitly copied from
						  others, may be republished without author(s) express written permission. Author(s) retain the right to alter this
						  disclaimer at any time. For the most up to date version of the disclaimer, see https://ucunleashed.com/code-disclaimer.
	  Acknowledgements      : Based on an original function by Any Arismendi, along with updates by others
						  http://poshcode.org/2566

						  Test for log names and sources
						  http://powershell.com/cs/blogs/tips/archive/2013/06/10/testing-event-log-names-and-sources.aspx

						  Writing to different event logs and sources registered to a single event log
						  http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/e172f039-ce88-4c9f-b19a-0dd6dc568fa0/
	  Assumptions           : ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
	  Limitations           : Writing to event logs requires admin rights
	  Known issues          :

	  .EXAMPLE
	  .\

	  Description
	  -----------


	  .INPUTS
	  System.String. You cannot pipe objects to this script.

	  .OUTPUTS
	  System.String
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
	# The type of message to be logged. Alias is 'type'.
	[Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
	[ValidateSet('Error', 'Warn', 'Info', 'Verbose')]
	[ValidateNotNullOrEmpty()]
	[string] $Level = 'Info',

	# The message to be logged.
	[Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true, HelpMessage = 'No message specified.')]
	[ValidateNotNullOrEmpty()]
	[string] $Message,

	# Specifies that $message should not the sent to the log file.
	[Parameter(ValueFromPipelineByPropertyName = $true)]
	[switch] $NoLog,

	# Specifies to not display the message to the console.
	[Parameter(ValueFromPipelineByPropertyName = $true)]
	[switch] $NoConsole,

	# The number of spaces to indent the message in the log file.
	[Parameter(ValueFromPipelineByPropertyName = $true)]
	[ValidateRange(1,30)]
	[ValidateNotNullOrEmpty()]
	[int] $Indent = 0,

	# Specifies what color the text should be be displayed on the console. Ignored when switch 'NoConsoleOut' is specified.
	[Parameter(ValueFromPipelineByPropertyName = $true)]
	[ValidateSet('Black', 'DarkMagenta', 'DarkRed', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkYellow', 'Red', 'Blue', 'Green', 'Cyan', 'Magenta', 'Yellow', 'DarkGray', 'Gray', 'White')]
	[ValidateNotNullOrEmpty()]
	[String] $ConsoleForeground = 'White',

	# Existing log file is deleted when this is specified. Alias is 'Overwrite'.
	[Parameter(ValueFromPipelineByPropertyName = $true)]
	[Switch] $Clobber,

	# The name of the system event log, e.g. 'Application'. The Skype for Business log is still called 'Lync Server'. Note that writing to the system event log requires elevated permissions.
	[Parameter(ValueFromPipelineByPropertyName = $true)]
	[ValidateSet('Application', 'System', 'Security', 'Lync Server', 'Microsoft Office Web Apps')]
	[ValidateNotNullOrEmpty()]
	[String] $EventLogName,

	# The name to appear as the source attribute for the system event log entry. This is ignored unless 'EventLogName' is specified.
	[Parameter(ValueFromPipelineByPropertyName = $true)]
	[ValidateNotNullOrEmpty()]
	[String] $EventSource = $([IO.FileInfo] $MyInvocation.ScriptName).Name,

	# The ID to appear as the event ID attribute for the system event log entry. This is ignored unless 'EventLogName' is specified.
	[Parameter(ValueFromPipelineByPropertyName = $true)]
	[ValidateRange(1,65535)]
	[ValidateNotNullOrEmpty()]
	[int] $EventID = 1,

	# The text encoding for the log file. Default is ASCII.
	[Parameter(ValueFromPipelineByPropertyName = $true)]
	[ValidateSet('Unicode','Byte','BigEndianUnicode','UTF8','UTF7','UTF32','ASCII','Default','OEM')]
	[ValidateNotNullOrEmpty()]
	[String] $LogEncoding = 'ASCII',

	#Divider line to be used to separate sections in the log file
	[Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Divider')]
	[ValidateNotNullOrEmpty()]
	[string] $LogDivider = '+------------------------------+'
  ) # end of param block
  BEGIN{

	[string]$TargetFolder = split-path ($MyInvocation.scriptname)
	#[string] $LogPath = "$TargetFolder\logs\update-RibbonADCache" + " {0:yyyy-MM-dd hh-mmtt}.log" -f (Get-Date)
	[string] $LogPath = "$TargetFolder\logs\update-RibbonADCache" + " {0:yyyy-MM-dd}.log" -f (Get-Date)
	[string]$LogFolder = Split-Path -Path $LogPath -Parent
	if (-not (Test-Path -Path $LogFolder)){
	  $null = New-Item -Path $LogFolder -ItemType Directory
	}
  } # end BEGIN
  PROCESS{
	try {
	  $Message = $($Message.trim())
	  $msg = '{0} : {1} : {2}{3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($Level.ToUpper()).PadRight(5," "), ('  ' * $Indent), $Message
	  if (-not ($NoConsole)){
		switch ($Level) {
		  'Error' {$Host.UI.WriteErrorLine("$Message")}
		  'Warn' {Write-Warning -Message $Message}
		  'Info' {Write-Host $Message -ForegroundColor $ConsoleForeground}
		  'Verbose' {Write-Verbose -Message $Message}
		}
	  }
	  if (-not ($NoLog)){
		if ($Clobber) {
		  $msg | Out-File -FilePath $LogPath -Encoding $LogEncoding -Force
		} else {
		  $msg | Out-File -FilePath $LogPath -Encoding $LogEncoding -Append
		}
	  }
	  if ($EventLogName) {
		if (-not $EventSource) {
		  [string] $EventSource = $([IO.FileInfo] $MyInvocation.ScriptName).Name
		}

		if(-not [Diagnostics.EventLog]::SourceExists($EventSource)) {
		  [Diagnostics.EventLog]::CreateEventSource($EventSource, $EventLogName)
		}

		switch ($Level) {
		  'Error' {$EntryType = 'Error'}
		  'Warn'  {$EntryType = 'Warning'}
		  'Info'  {$EntryType = 'Information'}
		  'Verbose' {$EntryType = 'Information'}
		  Default  {$EntryType = 'Information'}
		}
		Write-EventLog -LogName $EventLogName -Source $EventSource -EventId 1 -EntryType $EntryType -Message $Message
	  }
	  $msg = ''
	} # end try
	catch {
	  Throw "Failed to create log entry in: '$LogPath'. The error was: '$_'."
	} # end catch
  } # end PROCESS
  END{} # end END
} # end function Write-Log


function Read-UserInput
{
	param (
	[string] $prompt,
	[string] $default,
	[boolean] $displayOnly
	)

	#"Padright" done a little differently:
	while (($prompt.length + $default.length) -le 30)
	{
		$prompt = $prompt + " "
	}
	if ($default -ne "")
	{
		$prompt =  "{0} [{1}]" -f $prompt, $default
	}
	else
	{
		#Don't show the square brackets if there's no default value
		$prompt =  "{0}   " -f $prompt
	}

	if ($DisplayOnly)
	{
		Write-Host $prompt
	}
	else
	{
		if (($response = Read-Host -Prompt $prompt) -eq "")
		{
			$response = $default
		}
	}
	return $response
}


### Return the result of the request
Function BasicHandler
{
	Param($MyResult)

	if ($MyResult.GetType().Fullname -eq 'System.String')
	{
		[xml]$XmlResult = $MyResult.trimstart()
	}
	else
	{
		[xml]$XmlResult = $MyResult
	}

	if($XmlResult.root.status.http_code.contains("200"))
	{
		$info = @{
			"Success" = $True;
			"Result" = $XmlResult.root.status.http_code;
			"ErrorCode" = $null;
			"ErrorParam" = $null
		}
	}
	else
	{
		$info = @{
			"Success" = $False;
			"Result" = $XmlResult.root.status.http_code;
			"ErrorCode" = $XmlResult.root.status.app_status.app_status_entry.code;
			"ErrorParam" = $XmlResult.root.status.app_status.app_status_entry.params
		}
	}
	$resultInfo = New-Object -TypeName PSObject -Property $info
	return $resultInfo
}


function Login
{
	param (
	[string] $SbcFqdn,
	[string] $RestLogin,
	[string] $RestPassword
	)

add-type @"
	using System.Net;
	using System.Security.Cryptography.X509Certificates;

	public class IDontCarePolicy : ICertificatePolicy {
		public IDontCarePolicy() {}
		public bool CheckValidationResult(
			ServicePoint sPoint, X509Certificate cert,
			WebRequest wRequest, int certProb) {
			return true;
		}
	}
"@
	[System.Net.ServicePointManager]::CertificatePolicy = new-object IDontCarePolicy

	$BodyValue = "Username=$RestLogin&Password=$RestPassword"
	$url = "https://$SbcFqdn/rest/login"
	try
	{
		$Query = Invoke-RestMethod -Uri $url -Method Post -Body $BodyValue -SessionVariable SessionVar -verbose:$false
	}
	catch [System.Net.WebException]
	{
		$info = @{
			"Success" = $False;
			"Result" = $_.Exception; # Presumably "The remote name could not be resolved"
			"ErrorCode" = 404;
			"ErrorParam" = ""
		}
		$resultInfo = New-Object -TypeName PSObject -Property $info
		return $resultInfo
	}
	$Global:SessionVar = $SessionVar
	return (BasicHandler $Query)
}


#endregion Functions
#--------------------------------
# END  FUNCTIONS ---------------
#--------------------------------


#--------------------------------
# THE FUN STARTS HERE -----------
#--------------------------------

$ScriptVersion = "1.1"
$Error.Clear()
$Global:Debug = $psboundparameters.debug.ispresent
$Global:SessionVar = $null #This is the ID of the session we have open to the SBC

Write-Log -Level Info -Message "------------------------------------------------------" -NoConsole
Write-Log -Level Info -Message "Script launched " -NoConsole

If ($PsVersionTable.PsVersion.Major -lt 3)
{
	Write-Log -Level Error -Message "Sorry, your P$ version ($($PsVersionTable.PsVersion.ToString())) is too old: Invoke-RestMethod hasn't been invented yet"
	exit
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ($skipupdatecheck)
{
	Write-Log -Level Info -Message 'Skipping update check' -NoConsole
}
else
{
	write-progress -id 1 -Activity 'Initialising' -Status 'Performing update check' -PercentComplete (0)
	Get-UpdateInfo -title 'Update-RibbonADCache.ps1'
	write-progress -id 1 -Activity 'Initialising' -Status 'Back from performing update check' -Complete
}

$OutputValue = @()

try
{
	while (1)
	{
		Write-Log -Level Info -Message "About to login"
		if (($SbcFQDN -eq "") -or ($RestLogin -eq "") -or ($RestPassword -eq ""))
		{
			$SbcFQDN = read-UserInput "SBC FQDN" $SbcFQDN
			$RestLogin = read-UserInput "REST login name" $RestLogin
			$RestPassword = read-UserInput "REST password " $RestPassword

		}
		Write-Log -Level Info -Message "FQDN      = $($SbcFQDN)" -indent 1 -NoConsole
		Write-Log -Level Info -Message "RestLogin = $($RestLogin)" -indent 1 -NoConsole
		Write-Log -Level Info -Message "Password  = <Not Logged>" -indent 1 -NoConsole
		$result = Login $SbcFQDN $RestLogin $RestPassword
		if ($result.Success -eq $true)
		{
			Write-Log -Level Info -Message "Login successful" -indent 1
		}
		else
		{
			Write-Log -Level Error -Message ("Login failed. Error result = $($Result.Result)") -indent 1
			break
		}

		if ($QueryOnly)
		{
			Write-Log -Level Info -Message "Skipping clearing the Cache"
		}
		else
		{
			# Refresh the cache
			$url = "https://$SbcFQDN/rest/adconfig/?action=refreshadcache"
			$Query = Invoke-RestMethod -Uri $url -Method POST -WebSession $Global:SessionVar -verbose:$false
			$result = BasicHandler $Query
			if ($Result.Success -eq $true)
			{
				Write-Log -Level Info -Message "Refresh of the Cache requested"
			}
			else
			{
				Write-Log -Level Error -Message ("Refresh of the Cache failed. Error result = $($Result.Result)")
			}
		}

		# Query the DCs
		$url = "https://$SbcFQDN/rest/domaincontroller"
		$Query = Invoke-RestMethod -Uri $url -Method GET -WebSession $Global:SessionVar -verbose:$false
		$result = BasicHandler $Query
		if ($Result.Success -eq $true)
		{
			foreach ($DC in $Query.root.domaincontroller_list.domaincontroller_pk.href)
			{
				$DCQ = Invoke-RestMethod -Uri $DC -Method GET -WebSession $Global:SessionVar -verbose:$false
				$result = BasicHandler $DCQ
				if ($Result.Success -eq $true)
				{
					#YAY!
					$DcConfig = $DCQ.SelectNodes("/root/domaincontroller")
					$info = @{
					"ID" = $DcConfig.id;
					"DomainController" = $DcConfig.DomainController;
					"ADStatus" = $ADStatusLookup.Get_Item($DcConfig.rt_ADStatus);
					"CacheStatus"= $ADCacheStatusLookup.Get_Item($DcConfig.rt_CacheStatus);
					"BackupStatus" = $ADBackupStatusLookup.Get_Item($DcConfig.rt_BackupStatus)
					}
					$OutputValue += New-Object -TypeName PSObject -Property $info
				}
			}
		}
		else
		{
			Write-Log -Level Warning -Message "Query of existing AD values failed"
		}
		break
	}
}
catch
{
	if ($debug)
	{
		Write-Log -Level Error -Message "Unhandled crash. Error was $_ "
		$Global:error | Format-List -Property * -Force
	}
	else
	{
		Write-Log -Level Error -Message "Unhandled crash. Error was $_ "
	}
}
finally
{
	$OutputValue
}

Write-Log -Level Info -Message "Script exited" -NoConsole


# References
# Based on "Using REST to deploy an SBA on Sonus SBC1000/2000" by Adrien Plessis
# 	http://www.cusoon.fr/using-rest-to-deploy-an-sba-on-sonus-sbc10002000/#All_in_One
# REST "DomainController" resource: https://support.sonus.net/display/UXAPIDOC/Resource+-+domaincontroller
# Function return handling stolen with much gratitude from James Cussen: https://gallery.technet.microsoft.com/Skype-for-Business-Lync-04884260


#Code signing certificate kindly provided by Digicert:
# SIG # Begin signature block
# MIIZkAYJKoZIhvcNAQcCoIIZgTCCGX0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUyRnZOClui7ZaMye26T475SsZ
# l9+gghSeMIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMB4XDTIxMDEwMTAwMDAwMFoXDTMxMDEw
# NjAwMDAwMFowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAMLmYYRnxYr1DQikRcpja1HXOhFCvQp1dU2UtAxQ
# tSYQ/h3Ib5FrDJbnGlxI70Tlv5thzRWRYlq4/2cLnGP9NmqB+in43Stwhd4CGPN4
# bbx9+cdtCT2+anaH6Yq9+IRdHnbJ5MZ2djpT0dHTWjaPxqPhLxs6t2HWc+xObTOK
# fF1FLUuxUOZBOjdWhtyTI433UCXoZObd048vV7WHIOsOjizVI9r0TXhG4wODMSlK
# XAwxikqMiMX3MFr5FK8VX2xDSQn9JiNT9o1j6BqrW7EdMMKbaYK02/xWVLwfoYer
# vnpbCiAvSwnJlaeNsvrWY4tOpXIc7p96AXP4Gdb+DUmEvQECAwEAAaOCAbgwggG0
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEEGA1UdIAQ6MDgwNgYJYIZIAYb9bAcBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAfBgNVHSMEGDAWgBT0tuEgHf4prtLk
# YaWyoiWyyBc1bjAdBgNVHQ4EFgQUNkSGjqS6sGa+vCgtHUQ23eNqerwwcQYDVR0f
# BGowaDAyoDCgLoYsaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC10cy5jcmwwMqAwoC6GLGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtdHMuY3JsMIGFBggrBgEFBQcBAQR5MHcwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBPBggrBgEFBQcwAoZDaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRFRpbWVzdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEASBzctemaI7znGucgDo5nRv1CclF0CiNH
# o6uS0iXEcFm+FKDlJ4GlTRQVGQd58NEEw4bZO73+RAJmTe1ppA/2uHDPYuj1UUp4
# eTZ6J7fz51Kfk6ftQ55757TdQSKJ+4eiRgNO/PT+t2R3Y18jUmmDgvoaU+2QzI2h
# F3MN9PNlOXBL85zWenvaDLw9MtAby/Vh/HUIAHa8gQ74wOFcz8QRcucbZEnYIpp1
# FUL1LTI4gdr0YKK6tFL7XOBhJCVPst/JKahzQ1HavWPWH1ub9y4bTxMd90oNcX6X
# t/Q/hOvB46NJofrOp79Wz7pZdmGJX36ntI5nePk2mOHLKNpbh6aKLzCCBS8wggQX
# oAMCAQICEAqt2yhVXFSaEiY6y4bT9zkwDQYJKoZIhvcNAQELBQAwcjELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUg
# U2lnbmluZyBDQTAeFw0yMTA0MjMwMDAwMDBaFw0yMjA4MDQyMzU5NTlaMG0xCzAJ
# BgNVBAYTAkFVMRgwFgYDVQQIEw9OZXcgU291dGggV2FsZXMxEjAQBgNVBAcTCVBl
# dGVyc2hhbTEXMBUGA1UEChMOR3JlaWcgU2hlcmlkYW4xFzAVBgNVBAMTDkdyZWln
# IFNoZXJpZGFuMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAxrk1NuHH
# qyg9djhyuoE1UdImHdEItBzg/7zQ87RAQthP71A2GJ++zokQ6KfjbH5+UrEdODZN
# ibJF6/PnaVC1tUKPQHnauezk7ozu0JeUjLrxndxV8VEy3R/7wXp4hQ7XGaIehhhI
# u5+b6M0ZdTAmt93cT6AJYy8v/dPJr1DmZkj2KSbj10Ca9unAegKWsyDJmCQQ2EU5
# KxlRmPMwZK6as/SfAYVOxTnb5t7kO/F0HyKZJar5czLZn7CVWVke5QTqL6ZTnQg9
# 0u18c96gesFPAl247h+SgcLP4FOSzKVrF4NeMAyXlxettGiF2iei3r6zz8BEyhR0
# CXdbGzgmqDaU8QIDAQABo4IBxDCCAcAwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPA
# YPkt9mV1DlgwHQYDVR0OBBYEFDGB9TXcWUxGF52VHrnUqrZUeyXyMA4GA1UdDwEB
# /wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9o
# dHRwOi8vY3JsMy5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1
# oDOgMYYvaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1n
# MS5jcmwwSwYDVR0gBEQwQjA2BglghkgBhv1sAwEwKTAnBggrBgEFBQcCARYbaHR0
# cDovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYIKwYBBQUHAQEE
# eDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTgYIKwYB
# BQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNIQTJB
# c3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3
# DQEBCwUAA4IBAQDx1qRhZTX/nkQW4jCx2zWZsKJjMbeIUWMLi2dnuU9A9n1fIwwv
# +ab3jBKmoztY171Kxs0U97Tm/IzlwPeekIBKmTtThdBFmSqfU09eUPvtjLuI7H1j
# REAYH6MlzBIGRqbfaTSr7f+bSdSHsXZ68fB4zZyBg3s5N98yEFUe+978Of0hWRA5
# HlsNAdwjgih3dk9h1qBoqjVpt7VFLzpz7c99QBEND1zwn0VAwaxrFylraKjtnApK
# Gbu9Ow0YmL8kQ81B+pop8KzxQVEKA2A5wGpJciWgSSAatyEPZrPdcqIccktfV6gw
# pFZcN20IMqgQMv19mWLgywAJ2Er/ixi7G36qMIIFMDCCBBigAwIBAgIQBAkYG1/V
# u2Z1U0O1b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYD
# VQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAw
# WhcNMjgxMDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNl
# cnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdp
# Q2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG
# 9w0BAQEFAAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/
# 5aid2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH
# 03sjlOSRI5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxK
# hwjfDPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr
# /mzLfnQ5Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi
# 6CxR93O8vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCC
# AckwEgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAww
# CgYIKwYBBQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8v
# b2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6
# MHgwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3Vy
# ZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1s
# AAIEMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMw
# CgYIYIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1Ud
# IwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+
# 7A1aJLPzItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbR
# knUPUbRupY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7
# uq+1UcKNJK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7
# qPjFEmifz0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPa
# s7CM1ekN3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR
# 6mhsRDKyZqHnGKSaZFHvMIIFMTCCBBmgAwIBAgIQCqEl1tYyG35B5AXaNpfCFTAN
# BgkqhkiG9w0BAQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2Vy
# dCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMTYwMTA3MTIwMDAwWhcNMzEwMTA3MTIw
# MDAwWjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8A
# MIIBCgKCAQEAvdAy7kvNj3/dqbqCmcU5VChXtiNKxA4HRTNREH3Q+X1NaH7ntqD0
# jbOI5Je/YyGQmL8TvFfTw+F+CNZqFAA49y4eO+7MpvYyWf5fZT/gm+vjRkcGGlV+
# Cyd+wKL1oODeIj8O/36V+/OjuiI+GKwR5PCZA207hXwJ0+5dyJoLVOOoCXFr4M8i
# EA91z3FyTgqt30A6XLdR4aF5FMZNJCMwXbzsPGBqrC8HzP3w6kfZiFBe/WZuVmEn
# KYmEUeaC50ZQ/ZQqLKfkdT66mA+Ef58xFNat1fJky3seBdCEGXIX8RcG7z3N1k3v
# BkL9olMqT4UdxB08r8/arBD13ays6Vb/kwIDAQABo4IBzjCCAcowHQYDVR0OBBYE
# FPS24SAd/imu0uRhpbKiJbLIFzVuMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQDAgGGMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMIMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0
# cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNV
# HR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRB
# c3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMFAGA1UdIARJMEcwOAYKYIZI
# AYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20v
# Q1BTMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAQEAcZUS6VGHVmnN793a
# fKpjerN4zwY3QITvS4S/ys8DAv3Fp8MOIEIsr3fzKx8MIVoqtwU0HWqumfgnoma/
# Capg33akOpMP+LLR2HwZYuhegiUexLoceywh4tZbLBQ1QwRostt1AuByx5jWPGTl
# H0gQGF+JOGFNYkYkh2OMkVIsrymJ5Xgf1gsUpYDXEkdws3XVk4WTfraSZ/tTYYmo
# 9WuWwPRYaQ18yAGxuSh1t5ljhSKMYcp5lH5Z/IwP42+1ASa2bKXuh1Eh5Fhgm7oM
# LSttosR+u8QlK0cCCHxJrhO24XxCQijGGFbPQTS2Zl22dHv1VjMiLyI2skuiSpXY
# 9aaOUjGCBFwwggRYAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdp
# Q2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERp
# Z2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0ECEAqt2yhVXFSa
# EiY6y4bT9zkwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAw
# GQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisG
# AQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFFTOu6MqLSRT2SB3GZv/E5FmmM4PMA0G
# CSqGSIb3DQEBAQUABIIBADg0CJZFZHcx3j+axO54hQ4Kfi/uqoKefWT7YrxdnEYd
# uFPDb0u1n429OGB9XU+651FWtgXjuup2+MlkqMm9f7mqGrYIfHa2UuHwLLSBox0b
# iIV07VQ6hD8eHhz7co9Dl7HM0cJRtsxnLqItmD7AsjkGPUFHIHuqOnCktSNraMvp
# /Pe0i353JHeDKA+iCeRFsNOOsanxX2jeWOUS5pbkySbhfRQ85rYcMPkpL2/ivzKB
# M+I3rvl6Evf4xJjMB6JZBm6SBbRlLQy2P7KMkyu/QVBbpWhA59eHZ8MAYhhGLZTO
# U49U1fUoqStIAz755cr0rNGDIOsWTroOiS3rNykek1ShggIwMIICLAYJKoZIhvcN
# AQkGMYICHTCCAhkCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGln
# aUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQQIQDUJK4L46iP9g
# QCHOFADw3TANBglghkgBZQMEAgEFAKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0B
# BwEwHAYJKoZIhvcNAQkFMQ8XDTIxMTAyMzAwMTk0MlowLwYJKoZIhvcNAQkEMSIE
# IDzE36R6u1za6veLuvJn0VVlZsUPRWJiBDWjjtVqMr3hMA0GCSqGSIb3DQEBAQUA
# BIIBAJUAKzAok4IOKbPLuLJY7l6F0JW9RlQfUFkZ2MKP2n2KKrzGAPRywUsPIyu3
# TnyrU6CbI8PsAN9nkT9d2gfAWqz8SYOy+8ibflFEiPm1NBFGWLqRZKYmBZMESLLR
# jga8LAy0jpS9cVXsO9wywDnodOWvvU0cJS86RogQApVuzaLs2F7wTz77tLzSvqS7
# XUYeGv/FhC2F64jty2djF4mh1/TDddxaK9PIhAFyl6EgXSmPhUoiaYiJSGhDKtaE
# 6dhl2oZPkHfCjl7XDvI8QM39WUTwOv7ZWQgZYemVKmSwthv1kv7HDzDezgffZLzx
# F5c/jmp6GS9oBjg7979AOjLITyA=
# SIG # End signature block
