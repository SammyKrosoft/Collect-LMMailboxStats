<#
.SYNOPSIS
	Script to collect Mailbox Statistics.
	Authors: Lewis Martin and Colin Pastuch, Senior IT Architects for the Government Of Canada (Shared Services Canada)
	Reviewer: Sam Drey, Microsoft Consultant

.DESCRIPTION
	Script to collect Mailbox Statistics, from Lewis Martin and Colin Pastuch, Senior IT Architects for the
	Government Of Canada (Shared Services Canada).

	This script is composed of the following sections:
		- a "Script Header" that initiates a Stopwatch to dump the time the script took to run
		- a Write-Log function to write a log file on the WINDOWS\TEMP directory of the machine it's executed from
		- a main section leveraging the following Exchange Cmdlets:
			> Get-MailboxDatabase
			> Get-Mailbox
			> Get-MailboxStatistics
			> Get-User
			> Get-MailboxJunkEmailConfiguration

	We are still optimizing this script to ensure it takes less than 10 hours for ~100K mailboxes.

.OUTPUTS
	Outputs are :
	- a Log file in the SYSTEMROOT\TEMP directory
	- a MailboxStats_dd-mm-yy_hh-mm-ss.csv file with all the mailbox statistics

.NOTES
	Next version will leverage Start-Job/Get-Job to speed up even more the script.

.LINK
    https://github.com/SammyKrosoft
#>

<# -------------------------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) -------------------------- #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch =  [system.diagnostics.stopwatch]::StartNew()
#Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorActionPreference = "SilentlyContinue"
#Script Version
$ScriptVersion = "1.0"
# Log or report file definition - dumping 2 examples, use both if you need to output a report AND a script execution Log
# or just use one (delete the unused)
$CSVReportFile = "$PSScriptRoot\MailboxStats_" + $((Get-Date).ToString('MM-dd-yyyy_hh-mm-ss')) + ".csv"
$ScriptExecutionLogReportFile = "$((Get-Location).Path)\Collect-LMMailboxStats-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# -------------------------- /SCRIPT_HEADER -------------------------- #>

<#Functions Section#>
function Write-Log
{
	<#
	.SYNOPSIS
		This function creates or appends a line to a log file.

	.PARAMETER  Message
		The message parameter is the log message you'd like to record to the log file.

	.EXAMPLE
		PS C:\> Write-Log -Message 'Value1'
		This example shows how to call the Write-Log function with named parameters.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)]
		[string]$Message,
		[Parameter(Mandatory=$false)]
		[string]$LogFileName = $ScriptExecutionLogReportFile
	)
	
	try
	{
		$DateTime = Get-Date -Format 'MM-dd-yy HH:mm:ss'
		$Invocation = "$($MyInvocation.MyCommand.Source | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)"
		Add-Content -Value "$DateTime - $Invocation - $Message" -Path $LogFileName
		Write-Host $Message
	}
	catch
	{
		Write-Error $_.Exception.Message
	}
}

<#END of Functions Section#>
Write-Log "************************************************************"
Write-Log "*              Starting New Script Session                 *"
Write-Log "************************************************************"

Write-Log "Retrieving databases..."
$AllDatabases = Get-MailboxDatabase | Select Identity,Name, Server
Write-Log "Found $($AllDatabases.count) databases."

$ObjectCollection = @()
$stopwatch_2 =  [system.diagnostics.stopwatch]::StartNew()

Write-Log "For each database, parsing all mailboxes to get MailboxStatistics, AD properties and Junk information... "
Foreach ($Database in $AllDatabases) {
    Write-Log "Retrieving mailboxes for database $($Database.name)"
    $MailboxList = @(Get-Mailbox -ResultSize Unlimited -Database $($Database.name)| Select PrimarySmtpAddress , ServerName, Alias , DisplayName , OrganizationalUnit , Database , WhenMailboxCreated , ProhibitSendReceiveQuota , UseDatabaseQuotaDefaults , HiddenFromAddressListsEnabled , SingleItemRecoveryEnabled , CustomAttribute14)
    Write-Log "Found $($MailboxList.count) mailboxes on database $($Database.name)"
    ForEach($Mailbox in $MailboxList){
		$stopwatch_2.Start()

	    $MailboxStats = Get-MailboxStatistics -Database $Database.Name | Select LastLogonTime, ItemCount,TotalItemSize
        $MailboxUserAD = Get-User $Mailbox.Alias | Select FirstName , LastName , Company , Department , WhenChanged
		$ErrorActionPreference = "Continue"
		Try {
			$Junk = Get-MailboxJunkEmailConfiguration -Id $Mailbox.Alias | Select Enabled
			$IsJunkEnabled = $Junk.Enabled
		}
		Catch {
			Write-Log "Unabled to get Junk info - marking as JunkInfoUnavailable in the report"
			$IsJunkEnabled = "JunkInfoUnavailable"
		}
                                                              
        $SingleObject = New-Object PSObject

		$SingleObject | Add-Member -MemberType NoteProperty -Value $Mailbox.PrimarySmtpAddress -Name "Email Address"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $Mailbox.Alias -Name "Alias"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $Mailbox.DisplayName -Name "Display Name"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $MailboxUserAD.FirstName -Name "First Name"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $MailboxUserAD.LastName -Name "Last Name"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $MailboxUserAD.Company -Name "Company"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $MailboxUserAD.Department -Name "Department"                
		$SingleObject | Add-Member -MemberType NoteProperty -Value $Mailbox.OrganizationalUnit  -Name "OU"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $MailboxStats.LastLogonTime  -Name "Last Logon Time"
		$SingleObject | Add-Member -MemberType NoteProperty -Value ($Mailbox.ServerName).ToUpper() -Name "Server Name"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $Mailbox.Database -Name "Database"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $Mailbox.WhenMailboxCreated -Name "When Mailbox Created"
		$SingleObject | Add-Member -MemberType NoteProperty -Value ($MailboxStats.TotalItemSize).Value.ToMB() -Name "Mailbox Size In MB"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $MailboxStats.ItemCount  -Name "Item Count"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $Mailbox.ProhibitSendReceiveQuota -Name "Prohibit Send Receive Quota"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $Mailbox.UseDatabaseQuotaDefaults -Name "UseDatabaseQuotaDefaults"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $MailboxUserAD.WhenChanged -Name "Last Modified"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $Mailbox.HiddenFromAddressListsEnabled -Name "Hidden From GAL"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $Mailbox.SingleItemRecoveryEnabled -Name "Single Item Recovery Enabled"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $IsJunkEnabled -Name "Junk Enabled"
		$SingleObject | Add-Member -MemberType NoteProperty -Value $Mailbox.CustomAttribute14 -Name "Owner"
		$ObjectCollection += $SingleObject
	}
	Write-Log "Processed $($MailboxList.count) mailboxes in $($stopwatch_2.elapsed.TotalSeconds) seconds..."
	$stopwatch_2.Reset()
}

# [Samdrey] Added hh-mm-ss into file name string to differentiate multiple runs the same day (testing purposes)
#$FileName = "MailboxStats_" + $((Get-Date).ToString('MM-dd-yyyy_hh-mm-ss')) + ".csv"
$FileName = $CSVReportFile
Write-Log "Done parsing all mailboxes, mailbox stats, and Junk info. Writing into file: $FileName"

$ObjectCollection | Export-Csv $FileName -NoTypeInformation -Encoding 'UTF8'

Notepad $FileName

# Original : Lewis Martin - [Samdrey] Commenting
# Copy-Item $FileName \\SD01CCVMM3100\c$\Scripts\MailboxStats

<# -------------------------- SCRIPT_FOOTER -------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Log "The script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# -------------------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) -------------------------- #>