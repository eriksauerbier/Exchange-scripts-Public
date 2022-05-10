## PS-Skript zum löschen alter ActiveSync Geräte
## Stannek GmbH v.1.0 - 10.05.2022 - E.Sauerbier

# Exchange Exchange Management Shell laden
Add-PSSnapin *exchange*

# Parameter
$LastSuccessSyncBefore = 365 # Dieser Parameter kann nur alte Geräte anzeigen lassen

# Aktuelles Datum auslesen
$today = get-date
$lastsyncdate = $today.AddDays(-$LastSuccessSyncBefore)

# Alle Mailboxen auslesen
$mailboxes=Get-Mailbox 

# Objekt-Variable erzeugen
$RemovedDevices = [PSCustomObject]@()
$OldDevices  = [PSCustomObject]@()

# Alle Mobilen Geräte der ausgelesenen Mailboxen abfragen und wenn vorhanden in Variable schreiben

$devices = Get-mobiledevice # Abfrage der Mobilen Geräte für das Postfach

foreach ($device in $devices) {
   # Abfrage der letzten Syncronisierung
   $OldDevices += Get-MobileDeviceStatistics $device | where {$_.LastSuccessSync -le $lastsyncdate}
   
   }

# Ausgabe der zu löschenden Geräte

$break = $OldDevices | Select-Object DeviceModel, DeviceOS, LastSuccessSync, Identity | Out-GridView -Title "Folgende Geräte werden gelöscht" -PassThru

# Falls Abbrechen gedrückt wird, wird das Skript beendet
If ($break -eq $Null) {break}

#Alte Geräte löschen
foreach ($OldDevice in $OldDevices) {
        Remove-MobileDevice -Identity $OldDevice.Identity -Confirm:$false
        }
