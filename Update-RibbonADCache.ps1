<#
.SYNOPSIS
This script will login to the REST account in your SBC and refresh the AD cache.

.DESCRIPTION
This script will login to the REST account in your SBC and refresh the AD cache. Add the "-QueryOnly" switch for it to look but not touch.
Outputs an object per configured Domain Controller with the relevant settings/results.


.NOTES
	Version				: 1.0
	Date				: 1st August 2018
	Author    			: Greig Sheridan

	Revision History 	:

			Wishlist / TODO:
				#?

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

	[xml]$XmlResult = $MyResult.Substring(5)
	$xmlresult
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

$ScriptVersion = "1.0"
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
		[xml]$XmlResult = $Query.Substring(5)
		if ($XmlResult.root.status.http_code.contains("200"))
		{
			foreach ($DC in $XmlResult.root.domaincontroller_list.domaincontroller_pk.href)
			{
				$DCQ = Invoke-RestMethod -Uri $DC -Method GET -WebSession $Global:SessionVar -verbose:$false
				[xml]$XmlResult = $DCQ.Substring(5)
				if ($XmlResult.root.status.http_code.contains("200"))
				{
					#YAY!
					$DcConfig = $XmlResult.SelectNodes("/root/domaincontroller")
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
# MIIceAYJKoZIhvcNAQcCoIIcaTCCHGUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUwHgGJkVeOVCpjYx1oUf5q7NI
# zB+gghenMIIFMDCCBBigAwIBAgIQA1GDBusaADXxu0naTkLwYTANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTIwMDQxNzAwMDAwMFoXDTIxMDcw
# MTEyMDAwMFowbTELMAkGA1UEBhMCQVUxGDAWBgNVBAgTD05ldyBTb3V0aCBXYWxl
# czESMBAGA1UEBxMJUGV0ZXJzaGFtMRcwFQYDVQQKEw5HcmVpZyBTaGVyaWRhbjEX
# MBUGA1UEAxMOR3JlaWcgU2hlcmlkYW4wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQC0PMhHbI+fkQcYFNzZHgVAuyE3BErOYAVBsCjZgWFMhqvhEq08El/W
# PNdtlcOaTPMdyEibyJY8ZZTOepPVjtHGFPI08z5F6BkAmyJ7eFpR9EyCd6JRJZ9R
# ibq3e2mfqnv2wB0rOmRjnIX6XW6dMdfs/iFaSK4pJAqejme5Lcboea4ZJDCoWOK7
# bUWkoqlY+CazC/Cb48ZguPzacF5qHoDjmpeVS4/mRB4frPj56OvKns4Nf7gOZpQS
# 956BgagHr92iy3GkExAdr9ys5cDsTA49GwSabwpwDcgobJ+cYeBc1tGElWHVOx0F
# 24wBBfcDG8KL78bpqOzXhlsyDkOXKM21AgMBAAGjggHFMIIBwTAfBgNVHSMEGDAW
# gBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAdBgNVHQ4EFgQUzBwyYxT+LFH+GuVtHo2S
# mSHS/N0wDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1Ud
# HwRwMG4wNaAzoDGGL2h0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3Vy
# ZWQtY3MtZzEuY3JsMDWgM6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hh
# Mi1hc3N1cmVkLWNzLWcxLmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgG
# CCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEE
# ATCBhAYIKwYBBQUHAQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wTgYIKwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFNIQTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMB
# Af8EAjAAMA0GCSqGSIb3DQEBCwUAA4IBAQCtV/Nu/2vgu+rHGFI6gssYWfYLEwXO
# eJqOYcYYjb7dk5sRTninaUpKt4WPuFo9OroNOrw6bhvPKdzYArXLCGbnvi40LaJI
# AOr9+V/+rmVrHXcYxQiWLwKI5NKnzxB2sJzM0vpSzlj1+fa5kCnpKY6qeuv7QUCZ
# 1+tHunxKW2oF+mBD1MV2S4+Qgl4pT9q2ygh9DO5TPxC91lbuT5p1/flI/3dHBJd+
# KZ9vYGdsJO5vS4MscsCYTrRXvgvj0wl+Nwumowu4O0ROqLRdxCZ+1X6a5zNdrk4w
# Dbdznv3E3s3My8Axuaea4WHulgAvPosFrB44e/VHDraIcNCx/GBKNYs8MIIFMDCC
# BBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0Ew
# HhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEV
# MBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29t
# MTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5n
# IENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfT
# CzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdgl
# rA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRn
# iolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7
# MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPr
# CGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z
# 3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8E
# BAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0g
# BEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nED
# wGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqG
# SIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9
# D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQG
# ivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEeh
# emhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJ
# RZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5
# gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIGajCCBVKgAwIBAgIQAwGa
# Ajr/WLFr1tXq5hfwZjANBgkqhkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMG
# A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEw
# HwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAw
# WhcNMjQxMDIyMDAwMDAwWjBHMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNl
# cnQxJTAjBgNVBAMTHERpZ2lDZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBT
# qZ8fZFnmfGt/a4ydVfiS457VWmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWR
# n8YUOawk6qhLLJGJzF4o9GS2ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRV
# fRiGBYxVh3lIRvfKDo2n3k5f4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3v
# J+P3mvBMMWSN4+v6GYeofs/sjAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA
# 8bLOcEaD6dpAoVk62RUJV5lWMJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGj
# ggM1MIIDMTAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8E
# DDAKBggrBgEFBQcDCDCCAb8GA1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIB
# kjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQG
# CCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMA
# IABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMA
# IABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMA
# ZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkA
# bgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgA
# IABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUA
# IABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAA
# cgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQAS
# KxOYspkH7R7for5XDStnAs0wHQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9
# MH0GA1UdHwR2MHQwOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRENBLTEuY3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkw
# JAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcw
# AoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElE
# Q0EtMS5jcnQwDQYJKoZIhvcNAQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI
# //+x1GosMe06FxlxF82pG7xaFjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7ea
# sGAm6mlXIV00Lx9xsIOUGQVrNZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8Oxw
# YtNiS7Dgc6aSwNOOMdgv420XEwbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQN
# JsQOfxu19aDxxncGKBXp2JPlVRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNt
# omHpigtt7BIYvfdVVEADkitrwlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwggbN
# MIIFtaADAgECAhAG/fkDlgOt6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJ
# BgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5k
# aWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBD
# QTAeFw0wNjExMTAwMDAwMDBaFw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5j
# b20xITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/J
# M/xNRZFcgZ/tLJz4FlnfnrUkFcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPs
# i3o2CAOrDDT+GEmC/sfHMUiAfB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ
# 8DIhFonGcIj5BZd9o8dD3QLoOz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNu
# gnM/JksUkK5ZZgrEjb7SzgaurYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJr
# GGWxwXOt1/HYzx4KdFxCuGh+t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3ow
# ggN2MA4GA1UdDwEB/wQEAwIBhjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUH
# AwIGCCsGAQUFBwMDBggrBgEFBQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIB
# xTCCAbQGCmCGSAGG/WwAAQQwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIw
# ggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQA
# aQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUA
# cAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMA
# UAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEA
# cgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkA
# dAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8A
# cgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIA
# ZQBuAGMAZQAuMAsGCWCGSAGG/WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsG
# AQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqg
# OKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURS
# b290Q0EuY3JsMB0GA1UdDgQWBBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSME
# GDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+
# ybcoJKc4HbZbKa9Sz1LpMUerVlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6
# hnKtOHisdV0XFzRyR4WUVtHruzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5P
# sQXSDj0aqRRbpoYxYqioM+SbOafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke
# /MV5vEwSV/5f4R68Al2o/vsHOE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qqu
# AHzunEIOz5HXJ7cW7g/DvXwKoO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQ
# nHcUwZ1PL1qVCCkQJjGCBDswggQ3AgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EANRgwbrGgA18btJ2k5C8GEwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFLiNcSJ8+/vsWRZZuB8t
# 5pSif78pMA0GCSqGSIb3DQEBAQUABIIBAE/W6IBkjzPUV9BtuTlNZLHBqXXTunDf
# vk3zZBFZM0OIHngl/SO2ry4hmJB/5OC7QQY3lCl/mGSkQgt+uDD5Of6r1GjZsW5i
# kaU1Dp2mRSGdjMl0zRrcGjKI2nGb5VJ8InL9D5E36tU6ssesKf9D44d60nqfzzCk
# 7gTqXKpN0UJYx35OBiNdWfxOy+4haQ+otFRsVib3K+gRduuiPb+fgSqR2sBemqlN
# 7S2LDjhrHCTr1MXZDLp8gTxW7QQFL6gM+VZGYLPGAq/w6jqADCwbWzmBnqXCebbe
# dGyc9oukc6gXWjp4daxYFhDex1MMTQOEzndk2LI0xsYN4jvCvXEPkV+hggIPMIIC
# CwYJKoZIhvcNAQkGMYIB/DCCAfgCAQEwdjBiMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYD
# VQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTECEAMBmgI6/1ixa9bV6uYX8GYw
# CQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcN
# AQkFMQ8XDTIwMDUwNTExMjUxMlowIwYJKoZIhvcNAQkEMRYEFAocfQboGG9vdeke
# /JDKA3i5rSGiMA0GCSqGSIb3DQEBAQUABIIBABbwH3MEV9DxQhajaVntzJ4EHbVz
# GiZ3L4JDW91890jIOMF2m+c6KWxzp1U/AoWpbJKBBlnB/F/Z+ftwp7dpMkc67pQm
# Ot0VbcQaBQ5O+DB7b4OJGQ2x7NU8P88dIXApy70eMNk1smJkeU7lIH/b8jOiVgGY
# 8oGrhYOmZTlxru5XeK63GeGsSOsvphuMnbRu6IBmlmP6bPbBGrv1oP/lmDN0qynA
# Wuh3XsuNyUF165NSJrG9TJ1VocqgZu593Wu3eBRz51SOe8gcqtFioY6JgYWeAEXX
# PUMsfy3xGXqua5SEi+06EsDvKekfy+nUcAxxnIm5sODDSMbGHulsE1kPkDY=
# SIG # End signature block
