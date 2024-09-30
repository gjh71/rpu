function Get-Calendar {
    [CmdletBinding()]
    param (
        [PSCustomObject[]]$events,
        $interest,
        $team=""
    )
    $vCalendar = @"
BEGIN:VCALENDAR
VERSION:2.0
METHOD:PUBLISH
X-WR-CALNAME:Sloeproeiwedstrijden-2024
PRODID:-//DevOps1//RPU-IcsGenerator-0.1//EN

"@
    foreach ($event in $events) {
        $evDateParts = $event.date.split("-")
        if ($interest){
            $evDescription = "?-M2-?-{0}-?" -f $event.description.ToUpper()
        }
        elseif ($team -ne ""){
            $evDescription = "{1}-{0}" -f $event.description, $team
        }
        else{
            $evDescription = $event.description
        }
        $parm = @{
            eventDate        = Get-Date -Year $evDateParts[2] -Month $evDateParts[1] -Day $evDateParts[0]
            eventDescription = $evDescription
            eventSummary     = $event.summary
            eventLocation    = $event.location
        }
        $vCalendar += Get-Event @parm
    }
    $vCalendar += @"
END:VCALENDAR
"@
    return $vCalendar
}

function Get-MD5Hash {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$input
    )
    
    $stringAsStream = [System.IO.MemoryStream]::new()
    $writer = [System.IO.StreamWriter]::new($stringAsStream)
    $writer.Write($input)
    $writer.Flush()
    $stringAsStream.Position = 0
    $md5Hash = Get-FileHash -InputStream $stringAsStream | Select-Object -ExpandProperty Hash
    return $md5Hash
}

function Get-Event {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [datetime]$eventDate,
        $eventDescription,
        $eventSummary,
        $eventLocation
    )
    # $eventGuid = [guid]::NewGuid()
    # make sure the same 'guid' is generated for the same event
    $md5Hash = Get-MD5Hash -input ("{0:yyyyMMdd}{1}{2}" -f $eventDate, $eventDescription, $eventLocation)
    $eventGuid = "{0}-{1}-{2}-{3}-{4}" -f $md5Hash.Substring(0,8), $md5Hash.Substring(8,4), $md5Hash.Substring(12,4), $md5Hash.Substring(16,4), $md5hash.Substring(20)
    $now = Get-Date
    $eventStart = Get-Date $eventDate -Hour 5 -Minute 30
    $eventEnd = Get-Date $eventDate -Hour 16 -Minute 30
    $vEvent = @"
BEGIN:VEVENT
UID:{0}
CREATED:{1:yyyyMMddTHHmmssZ}
DTSTAMP:{1:yyyyMMddTHHmmssZ}
LAST-MODIFIED:{1:yyyyMMddTHHmmssZ}
SEQUENCE:0
SUMMARY;LANGUAGE=nl-nl:{3}
DTSTART:{6:yyyyMMddTHHmmssZ}
DTEND:{7:yyyyMMddTHHmmssZ}
DESCRIPTION:{4}
LOCATION:{5}
TRANSP:OPAQUE
END:VEVENT

"@ -f $eventGuid, $now, $eventDate, $eventDescription, $eventSummary, $eventLocation, $eventStart, $eventEnd
    return $vEvent
}

$interest = $false
$team = "Mix2"
$file = Join-Path -Path $PSScriptRoot -ChildPath "2024-wedstrijden.csv"
$events = Get-Content -Path $file -raw | ConvertFrom-Csv
$events | Format-Table
$ics = Join-Path -Path $PSScriptRoot -ChildPath "calendar.ics"
Get-Calendar -events $events -interest $interest -team $team| out-file $ics
