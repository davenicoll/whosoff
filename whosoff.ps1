<#
.SYNOPSIS
Retrieves and publishes information from Who's Off

.DESCRIPTION
This script requests the ical feed from Who's Off, transforms the data, and sends the information to Slack

.NOTES
None

.EXAMPLE
whosoff.ps1 -whosoffurl http://feeds.staff.whosoff.com/?u=0000000000 -slackchannel #yourchannel - slackapihookurl https://hooks.slack.com/services/A000000C/B000000U/SAKLJDKSALJDKLSAJD

.PARAMETER whosoffurl
The URL of the Who's Off feed to request data from. Who's Off provide feed data in .ics/iCal format

.PARAMETER proxy
If you need to use a proxy to POST a message to Slack, specify a URL to your proxy

.PARAMETER slackchannel
The Slack channel (or user) to post the update to

.PARAMETER slackapihookurl
Your Slack API hook URL to POST a message to. You can create this in the Slack administration area.

.PARAMETER useproxy
Whether or not to use a proxy to POST messages via.

.PARAMETER usemock
Whether to use the mock data file (for testing)

#>
#requires -version 4

param (
    [string]$whosoffurl = "http://feeds.staff.whosoff.com/?u=your-feed-id",
    [string]$proxy="http://YOURPROXY:8080",
    [string]$slackchannel="#your-channel",
    [string]$slackapihookurl="https://hooks.slack.com/services/your-slack-api-URL",
    [bool]$useproxy=$false,
    [bool]$runtests=$false,
    [bool]$usemock=$false
 )

[string]$datafile=[System.IO.Path]::GetTempFileName()
[string]$mockdatafile=".\whosoff-mock.ics"
$script:eventdata = ""
$events = ""
$today = [System.DateTime]::Now
$nextworkingdaysevents = ""
[string[]] $script:freedays = ""

function GetWhosOffData()
{
    if($usemock -eq $True)
    {
        if(Test-Path $mockdatafile)
        {
            $datafile = $mockdatafile
            write-output "Using mock Who's Off data from $datafile"
        }
        else
        {
            write-output "ERROR: $mockdatafile does not exist. Mock data not available"
            return
        }
    }
    else
    {
        try{
            write-output "Requesting Who's Off data from $whosoffurl"
            Invoke-WebRequest $whosoffurl -OutFile $datafile
            Write-Output "Saved data to $datafile"
        }
        catch [System.Net.WebException]
        {
            Write-Output "ERROR: Couldn't get data from Who's Off"
        }
    }

    $content = Get-Content $datafile -Raw

    $script:eventdata = $content -split 'BEGIN:VEVENT'
}

function ParseWhosOffData($day,[bool] $nextdaypass)
{
    <#
	When $nextdaypass is $False get data for anyone who is off on the specified $day
	When $nextdaypass is $True get data for anyone who commences a period of absence on the specified $day
	#>
	$parsedevents = ""

    foreach($element in $script:eventdata){

		$startdate =""
		$enddate =""

		if($element -match "DTSTART.+")
		{
			$d = $matches[0].split(":")
			$dy = $d[1].substring(0,8)
			$startdate = [datetime]::ParseExact($dy, 'yyyyMMdd', [CultureInfo]::InvariantCulture)
		}
		if($element -match "DTEND.+")
		{
			$d = $matches[0].split(":")
			$dy = $d[1].substring(0,8)
			$enddate = [datetime]::ParseExact($dy, 'yyyyMMdd', [CultureInfo]::InvariantCulture)
		}

		if(($startdate -ne "" -and $enddate -ne "") -or ($today -eq $startdate -and $day -eq $enddate))
		{
			if((($today -gt $startdate) -and ($day -lt $enddate) -and (-not $nextdaypass)) -Or (($day.ToString("yyyyMMdd") -eq $startdate.ToString("yyyyMMdd")) -and ($day -lt $enddate) -and $nextdaypass))
			{
                if($element -match "CATEGORIES.+")
                {
                    $status = $matches[0].replace("CATEGORIES:","")

					if($status -like "*sick leave*")
					{
						$status = "|Sickness|"
					}
					elseif($status -like "*maternity*")
					{
						$status = "|Maternity|"
					}
					elseif($status -like "*training*")
					{
						$status = "|Training|"
					}
					elseif($status -like "*out of office*")
					{
						$status = "|OoO|"
					}
					elseif($status -like "*working From home*")
					{
						$status = "|WFH|"
					}
					else
					{
						$status = "|Holiday|"
					}
                }

				if($element -match "SUMMARY.+")
				{
					$person = $matches[0].replace("SUMMARY:","").split("[")
                    $person = $person[0].trim()
				}

                $parsedevents += " • " + $person.trim() + $status + $enddate.ToString("dd/MM/yyyy") + [System.Environment]::NewLine
			}
		}

    }
    return $parsedevents
}

function GetFreeDays()
{

	$script:freedays = ""

    foreach($element in $script:eventdata){

		if($element -match "FREE DAY \(All Departments\)")
        {
            if($element -match "DTSTART.+")
            {
                $d = $matches[0].split(":")
				$script:freedays += $d[1].substring(0,8)
            }
		}

	 }
}

function SendSlackNotification($events, $nextworkingdaysevents)
{
    Write-Output("Posting to Slack ($slackchannel) using $slackapihookurl (Proxied: $useproxy)")

	$title1 = "<!here> The following people aren't in the office *today*..."
	$title2 = "From *" + $nextworkingday.ToString("dddd dd MMMM yyyy") + "* the following people will not be in the office..."

    if($events.length -gt 0)
    {
        $events = ConvertToSlackAttachment $events $title1
    }
    else
    {
        $events = ConvertToSlackAttachment $events "<!here> Everyone is in the office *today*."
    }

    if($nextworkingdaysevents.length -gt 0)
    {
        $nextworkingdaysevents = ConvertToSlackAttachment $nextworkingdaysevents $title2
    }
    else
    {
        $nextworkingdaysevents = ConvertToSlackAttachment $nextworkingdaysevents "There are no changes *tomorrow*."
    }

	$json = '{"attachments":[' + $events +  ',' + $nextworkingdaysevents + '], "channel":"' + $slackchannel + '", "username":"WhosOff", "icon_emoji":":calendar:" }'
    try{
    $ProgressPreference = 'silentlyContinue'
    if($useproxy -eq $TRUE)
    {
	    Invoke-WebRequest -Uri "$slackapihookurl" -Body $json -Method POST -Proxy "$proxy"
	}
    else
    {
        Invoke-WebRequest -Uri "$slackapihookurl" -Body $json -Method POST
    }
    }
    catch
    {
        Write-Output("ERROR: Couldn't post to Slack")
    }
}

function RemoveDataFile()
{
    if(test-path $datafile) {
        Remove-Item $datafile
    }
}

function ConvertToHtml($events)
{
    $events = $events.replace([System.Environment]::NewLine,"<br/>")
    $events = "<html xmlns='http://www.w3.org/1999/xhtml'><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'><style type='text/css'>html{background-color:#ffffff;}body{font-family:Calibri, Tahoma, Arial;font-size:1em;}table{height:100%;}td{padding:10px;font-family:calibri,Arial;font-size:12px;line-height:135%;}.email{width:100%;background-color:#00375B;}.email_container{background-color:#fff;border:2px solid #efefef;width:480pt;margin:0px auto;padding:20px;}.email_header{border-bottom:1px solid #efefef;vertical-align:bottom;height:80px;font-size:14px;font-weight:bold;}.email_header td{line-height:14px;font-size:14px;font-weight:bold;padding:10px;}.email_body{padding:10px;}.company_logo{http://staff.whosoff.com/Assets/Img/whosoff_email_logo.jpg}</style><title>Who's off today</title></head><body>The following staff will not be in the Stoke office today -<br/><br/>" + $events + "<br/></body></html>"

    return $events
}

function ConvertToSlackAttachment($events, $title)
{
	$data = $events -split ' • '
	$names = ""
	$reasonanddate = ""

	$data = $data|sort

    if($data.length -eq 0)
    {
        return '{"fallback":"","mrkdwn_in": ["pretext"],"pretext":"' + $title + '","color":"#8000ff"}'
    }

    foreach($element in $data){
		$subdata = $element.split([string[]]"|",[StringSplitOptions]"None")
		$names += $subdata[0].trim() + '\n'
		$reasonanddate += $subdata[2] + '    ' + $subdata[1] + '\n'
	}

	$names = ConvertToSlackField "Name" $names
	$reasonanddate = ConvertToSlackField "Return date    Reason" $reasonanddate
	$slack = '{"fallback":"","mrkdwn_in": ["pretext"],"pretext":"' + $title + '","color":"#8000ff","fields":[' + $names + ',' + $reasonanddate + ']}'

	return $slack
}

function ConvertToSlackField($title, $value)
{
	#Field format: {"title":"Notes","value":"This is much easier than I thought it would be.","short":true}
	$value = $value.replace([System.Environment]::NewLine,"")
	$value = $value.replace(" • ","")
	$value = $value.replace("'","''")
	$value = '{"title":"' + $title + '","value":"' + $value + '","short":true}'

	return $value
}

function GetNextWorkingDay($day)
{
	$day = $day +"$(1+$(@(1,2-eq7-($day).DayOfWeek)))"
	while (-Not (IsWorkingDay $day)){
		$day = $day.AddDays(1)
	}

	return $day
}

function IsWorkingDay([DateTime] $day)
{
	$strday = $day.ToString("yyyyMMdd")
	[bool] $returnvalue = -Not((($script:freedays -contains $strday)  -or ($day.DayOfWeek -eq 'Saturday') -Or ($day.DayOfWeek -eq 'Sunday')))

	return $returnvalue
}

# -- MAIN --

if($runtests -eq $TRUE)
{
    Write-Output("TODO: Create tests :)")
    return
}
else
{
    if($usemock -eq $False)
    {
        RemoveDataFile
        GetWhosOffData
        RemoveDataFile
    }
    else
    {
        GetWhosOffData
    }
    GetFreeDays

    if (IsWorkingDay $today)
    {
	    $events = ParseWhosOffData $today $FALSE
	    $nextworkingday = GetNextWorkingDay $today
	    $nextworkingdaysevents = ParseWhosOffData $nextworkingday $TRUE

	    SendSlackNotification $events $nextworkingdaysevents
    }
}
