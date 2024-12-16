$ledenAdministratieXlsFiLename = Join-Path -Path $PSScriptRoot -ChildPath "Leden-Sponsor-Contact.xlsx"
$ledenSpondFilename = Join-Path -Path $PSScriptRoot -ChildPath "Roeiploeg Utrecht-members.xlsx"
$onderhoudSpondFilename = Join-Path -Path $PSScriptRoot -ChildPath "Onderhoud Roeiploeg Utrecht-members.xlsx"

# Basis ledenadministratie
$ledenAdministratie = Import-Excel -Path $ledenAdministratieXlsFiLename -WorksheetName "Leden" | Where-Object { $_.Lid -eq "Actief lid" } `
    | Select-Object -Property "Voornaam", "Achternaam", "Team", `
        @{ Label = "Mobiel"; Expression = { $_.Mobiel.replace(" ", "").replace("-", "") } }, `
        @{ Label = "Email"; Expression = { $_."email adres" } }
$ledenSpond = Import-Excel -Path $ledenSpondFilename -WorksheetName "For import" `
    | Select-Object -Property "Name", "Cell", @{label = "Email"; expression = { $_.Email.Trim().ToLower() } }
$onderhoudSpond = Import-Excel -Path $onderhoudSpondFilename -WorksheetName "For import" `
    | Select-Object -Property "Name", "Cell", @{label = "Email"; expression = { $_.Email.Trim().ToLower() } }

$result = @()
# now try to match
foreach ($lid in $ledenAdministratie) {
    $lid.Voornaam = $lid.Voornaam.Trim()
    $lid.Achternaam = $lid.Achternaam.Trim()
    $lid.Team = $lid.Team.Trim()
    $matchObject = [PSCustomObject]@{
        Voornaam                    = $lid.Voornaam
        Achternaam                  = $lid.Achternaam
        Team                        = $lid.Team
        Mobiel                      = $lid.Mobiel
        Email                       = $lid.Email
        LedenAdministratie          = $true
        LedenSpond                  = $false
        OnderhoudSpond              = $false
        LedenAdministratieOpmerking = ""
        LedenSpondOpmerking         = ""
        OnderhoudSpondOpmerking     = ""
    }
    if ($matchObject.Mobiel.Length -gt 2 -and $matchObject.Mobiel.Substring(0, 2) -eq "06") {
        $matchObject.Mobiel = "+31" + $matchObject.Mobiel.Substring(1)
    }

    ### LedenSpond
    $match = $ledenSpond | Where-Object { ($_.Name -eq "$($lid.Voornaam) $($lid.Achternaam)") -and ($_.Cell -eq $lid.Mobiel) -and ($_.Email -eq $lid."email adres") }
    if ($null -eq $match) {
        $match = $ledenSpond | Where-Object { ($_.Name -eq "$($lid.Voornaam) $($lid.Achternaam)") }
        if ($match) {
            $matchObject.LedenSpond = $true
            $matchObject.LedenSpondOpmerking = "Naam"
        }
        $match = $ledenSpond | Where-Object { ($null -ne $_.Cell) -and ($_.Cell.replace(" ", "").replace("-", "") -eq $matchObject.Mobiel) }
        if ($match) {
            $matchObject.LedenSpond = $true
            $matchObject.LedenSpondOpmerking += ", mobiel"
        }
        $match = $ledenSpond | Where-Object { ($_.Email -eq $matchObject.Email) }
        if ($match) {
            $matchObject.LedenSpond = $true
            $matchObject.LedenSpondOpmerking += ", email"
        }
    }
    else {
        $matchObject.LedenSpond = $true
        $matchObject.LedenSpondOpmerking += "Naam, mobiel, email"
    }
    ### Onderhoud
    $match = $onderhoudSpond | Where-Object { ($_.Name -eq "$($lid.Voornaam) $($lid.Achternaam)") -and ($_.Cell -eq $lid.Mobiel) -and ($_.Email -eq $lid."email adres") }
    if ($null -eq $match) {
        $match = $onderhoudSpond | Where-Object { ($_.Name -eq "$($lid.Voornaam) $($lid.Achternaam)") }
        if ($match) {
            $matchObject.OnderhoudSpond = $true
            $matchObject.OnderhoudSpondOpmerking = "Naam"
        }
        $match = $onderhoudSpond | Where-Object { ($null -ne $_.Cell) -and ($_.Cell.replace(" ", "").replace("-", "") -eq $matchObject.Mobiel) }
        if ($match) {
            $matchObject.OnderhoudSpond = $true
            $matchObject.OnderhoudSpondOpmerking += ", mobiel"
        }
        $match = $onderhoudSpond | Where-Object { ($_.Email -eq $matchObject.Email) }
        if ($match) {
            $matchObject.OnderhoudSpond = $true
            $matchObject.OnderhoudSpondOpmerking += ", email"
        }
    }
    else {
        $matchObject.OnderhoudSpond = $true
        $matchObject.OnderhoudSpondOpmerking += "Naam, mobiel, email"
    }

    $result += $matchObject
}
$result | Where-Object { $_.LedenSpond -ne $true -or $_.OnderhoudSpond -ne $true } | ogv
$excelFilename = Join-Path -Path $PSScriptRoot -ChildPath "Compare-Ledenadministratie-Spond.xlsx"
$ledenAdministratie | Export-Excel -Path $excelFilename -WorksheetName "administratie" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
$ledenSpond | Export-Excel -Path $excelFilename -WorksheetName "spond" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
$onderhoudSpond | Export-Excel -Path $excelFilename -WorksheetName "onderhoud" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
$result | Export-Excel -Path $excelFilename -WorksheetName "Comparison" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -Show
