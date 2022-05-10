## PS-Skript zum Abfragen aller Postfächer die ActiveSync oder OWA aktiviert haben
## Stannek GmbH v.1.0 - 10.05.2022 - E.Sauerbier

# Exchange Exchange Management Shell laden
Add-PSSnapin *Exchange*

# Alle Mailboxen mit aktiven ActivySync oder OWA auslesen
$Mailboxes = Get-CASMailbox | where {($_.ActiveSyncEnabled -eq $True) -or ($_.OWAEnabled -eq $True)}

# Ausgabe der Mailboxen
$Mailboxes | Out-GridView
