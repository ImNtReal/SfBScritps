Param (
  $PoolFQDN = (Read-Host -Prompt "Please enter the Pool FQDN")
)

#### Script Information
#
# Originally written by Richard Brynteson
# http://masteringlync.com
####

# Convert UTC to Local timezone
function Convert-UTCtoLocal {
  param(
    [parameter(Mandatory=$true)]
    [String] $UTCTime
  )
  
  $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
  $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
  $LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
  
  Return $LocalTime
}

#Loop Through Front-End Pool
Foreach ($Computer in (Get-CsPool -Identity $PoolFQDN).Computers) {
  
  $Result = Invoke-SQLCmd -ServerInstance "$Computer\rtclocal" -Database rtcdyn -Query "SELECT ActiveConference.ConfId AS 'Conference ID', ActiveConference.Locked, Participant.UserAtHost AS  'Participant', Participant.JoinTime AS 'Join Time', Participant.EnterpriseId, ActiveConference.IsLargeMeeting AS 'Large Meeting' FROM   ActiveConference INNER JOIN Participant ON ActiveConference.ConfId = Participant.ConfId;"
  
  $Result | Add-Member -NotePropertyName 'Frontend' -NotePropertyValue $Computer
  
  $Result."Join Time" = Convert-UTCtoLocal -UTCTime $Result."Join Time"
  
  $Results += $Result
}

#$Results | ft 'Participant', 'Join Time', 'Large Meeting', 'Frontend' -GroupBy 'Conference ID'

$Results
