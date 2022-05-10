## PS-Skript zum Abfragen aller ActiveSync Geräte
## Stannek GmbH v.1.1 - 10.05.2022 - E.Sauerbier

# Exchange Exchange Management Shell laden
Add-PSSnapin *exchange*

# Parameter
$LastSuccessSyncBefore = 0 # Dieser Parameter kann nur alte Geräte anzeigen lassen

# Aktuelles Datum auslesen
$today = get-date
$lastsyncdate = $today.AddDays(-$LastSuccessSyncBefore)


# Alle Mailboxen auslesen
$mailboxes=Get-Mailbox 

# Objekt-Variable erzeugen
$ASdevices = [PSCustomObject]@()

# Alle Mobilen Geräte der ausgelesenen Mailboxen abfragen und wenn vorhanden in Variable schreiben

foreach ($mailbox in $mailboxes) # Schleife zum Abfragen pro Mailbox
 	{ $devices = [PSCustomObject]@() # temporäre Objekt-Variable für Mobile Geräte erzeugen
      $devices = Get-mobiledevice -Mailbox $mailbox  # Abfrage der Mobilen Geräte für das Postfach
      # Wenn Mobile Geräte vorhanden dann wird dies in eine Objekt-Variable geschrieben
      if ($devices -ne $Null) {
            foreach ($device in $devices) {
                # Abfrage der letzten Syncronisierung
                $LastSuccesSync = Get-MobileDeviceStatistics $device | where {$_.LastSuccessSync -le $lastsyncdate}
                if ($LastSuccesSync) {
                $ASdevices += [PSCustomObject]@{
                Mailbox = $mailbox
                DeviceModel = $device.DeviceModel
                DeviceOS = $device.DeviceOS
                FirstSyncTime = $device.FirstSyncTime
                LasSyncTime = $LastSuccesSync | Select-Object -ExpandProperty LastSuccessSync
                }
                }
            }
      } 
}

# Ausgabe der Mobilen Geräte

$ASdevices | Out-GridView


