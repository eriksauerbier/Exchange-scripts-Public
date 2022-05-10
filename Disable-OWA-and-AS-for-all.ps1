# Dieses Skript deaktiviert für alle Mailboxen Owa und ActivSync
# Stannek GmbH - E.Sauerbier - v.1.41 - 10.05.2022

# Exchange Snap-In laden
Add-PSSnapin *exchange*

# Mailboxen auslesen und Owa/ActivSync deaktivieren
Get-Mailbox | Set-CASMailbox -ActiveSyncEnabled $false -OWAEnabled $false -OWAforDevicesEnabled $false