$file = "./2024-wedstrijden.csv"
$events = Get-Content -Path $file -raw | ConvertFrom-Csv
$events | Format-Table


function CreateNewEvent {

    # Custom date formats that we want to use
    $longDateFormat = "yyyyMMddTHHmmssZ"
    $dateFormat = "yyyyMMdd"

    # Prompt the user for the start date in a specific format
    $startDate = Read-Host -Prompt ‘Enter the start date of the event in the format "yyyymmdd"’

    # Give the event a name specified by the user
    $eventSubject = Read-Host -Prompt ‘Enter the event subject’

    # This field is optional, but let’s ask for the details (description)
    $eventDesc = Read-Host -Prompt ‘Enter the event description summary (optional)’

    # Provide location information (also optional)
    $eventLocation = Read-Host -Prompt ‘Enter the event location (optional)’

    # Instantiate .NET StringBuilder
    $sb = [System.Text.StringBuilder]::new()

    # Fill in ICS/iCalendar properties based on RFC2445
    [void]$sb.AppendLine(‘BEGIN:VCALENDAR’)
    [void]$sb.AppendLine(‘VERSION:2.0’)
    [void]$sb.AppendLine(‘METHOD:PUBLISH’)
    [void]$sb.AppendLine(‘PRODID:-//Braunweb//PowerShell ICS Creator Sample//EN’)
    [void]$sb.AppendLine(‘BEGIN:VEVENT’)
    [void]$sb.AppendLine("UID:" + [guid]::NewGuid())
    [void]$sb.AppendLine("CREATED:" + [datetime]::Now.ToUniversalTime().ToString($longDateFormat))
    [void]$sb.AppendLine("DTSTAMP:" + [datetime]::Now.ToUniversalTime().ToString($longDateFormat))
    [void]$sb.AppendLine("LAST-MODIFIED:" + [datetime]::Now.ToUniversalTime().ToString($longDateFormat))
    [void]$sb.AppendLine("SEQUENCE:0")
    [void]$sb.AppendLine("DTSTART:" + $startDate)
    [void]$sb.AppendLine("RRULE:FREQ=YEARLY;INTERVAL=1")
    [void]$sb.AppendLine("DESCRIPTION:" + $eventDesc)
    [void]$sb.AppendLine("SUMMARY:" + $eventSubject)
    [void]$sb.AppendLine("LOCATION:" + $eventLocation)
    [void]$sb.AppendLine("TRANSP:TRANSPARENT")
    [void]$sb.AppendLine(‘END:VEVENT’)
    [void]$sb.AppendLine(‘END:VCALENDAR’)
}

function Get-Event{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [datetime]$eventDate,
        $eventDescription,
        $eventSummary,
        $eventLocation
    )
    $eventGuid = [guid]::NewGuid()
    $now = Get-Date
    $event = @"
BEGIN:VCALENDAR
VERSION:2.0
METHOD:PUBLISH
PRODID:-//Braunweb//PowerShell ICS Creator Sample//EN
BEGIN:VEVENT
UID:{0}
CREATED: {1:yyyyMMddHH:mm:ss}
DTSTAMP: {1:yyyyMMddHH:mm:ss}
LAST-MODIFIED: {1:yyyyMMddHH:mm:ss}
SEQUENCE:0
DTSTART:" + $startDat
RRULE:FREQ=YEARLY;INTERVAL=1
DESCRIPTION:" + $eventDes
SUMMARY:" + $eventSubjec
LOCATION:" + $eventLocatio
TRANSP:TRANSPARENT
END:VEVENT
END:VCALENDAR
"@ -f $eventGuid

}