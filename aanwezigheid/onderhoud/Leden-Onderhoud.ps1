$importFile = Join-Path -Path $PSScriptRoot -ChildPath "download.xlsx"
$attendance = Import-Excel -Path $importFile -WorksheetName "Sheet0"
$importCsv = Join-Path -Path $PSScriptRoot -ChildPath "download.csv"
# $attendance | Select-Object -Skip 1 | ConvertTo-Csv -NoTypeInformation | Set-Content -Path $importCsv -Force
$headers = $attendance[0]
$attendeeAttributes = @("Name", "Email", "Mobile")
foreach($header in $headers.PSObject.Properties.Name) {
    $parts = $header -split ","
    if ($parts.Length -eq 2) {
        $d = "{0:yyyy-MM-dd}" -f (get-date "1899-12-30").AddDays($parts[0])
        $eventtype = Invoke-Expression -Command ("`$attendance[0].'{0}'" -f $header)
        $d = "{0}-{1}" -f $eventtype, $d
        if ($attendeeAttributes -notcontains $eventtype) {
            $attendeeAttributes += $eventtype
        }
    }
    else{
        $d = $parts[0]
    }
    write-host($d)
}
$attendeeTemplate = New-Object PSObject
$templateValue = ""
foreach($attribute in $attendeeAttributes) {
    $attendeeTemplate | Add-Member -MemberType NoteProperty -Name $attribute -Value $templateValue
    if ($attribute -eq "Mobile") {
        $templateValue = 0
    }
}
$attendeeTemplate | Add-Member -MemberType NoteProperty -Name "Total" -Value 0

$attendees = @()
foreach($attendeeRecord in ($attendance | Select-Object -Skip 1 )){
    $attendee = $attendeeTemplate.PSObject.Copy()
    foreach($header in $headers.PSObject.Properties.Name) {
        if ($header -in $attendeeAttributes) {
                $attendee.$header = $attendeeRecord.$header
        }
        else {
            $customHeader = $attendance[0].$header
            if ($customHeader -in $attendeeAttributes) {
                $attendee.$customHeader += $attendeeRecord.$header
                $attendee.Total += $attendeeRecord.$header
            }
        }
    }
    $attendees += $attendee
}

$excelFilename = Join-Path -Path $PSScriptRoot -ChildPath "RPU-Onderhoud.xlsx"
$attendees | Export-Excel -Path $excelFilename -WorksheetName "Aanwezigheid" -AutoSize -FreezeTopRow -BoldTopRow

