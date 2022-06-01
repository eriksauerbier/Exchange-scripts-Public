# Dieses Skript löscht abgelaufene Exchange-Zertifikate eines bestimmten Subjects oder Generell (z.B. für Let's Encrypt)
# Stannek GmbH - v.1.1 - 01.06.2022 - E.Sauerbier

# Parameter
# Hier das Subject vom Zertifikat eingeben
$Subject = "CN=mail.domain.com"

# Exchange PSSnapin laden
Add-PSSnapin *Exchange*

# Abgelaufene Zertifikate auslesen
$RemoveCert = Get-ExchangeCertificate | Where {$_.Status -eq "DateInvalid" -and $_.Subject -eq $Subject}

# Alternativ kann mit folgenden Befehl auch alle abgelaufenen Zertifikate gelöscht werden
# Ist aber nicht Empfehlenswert, da es zu Problemen bei abgelaufenen Systemzertifikaten kommen kann
# $RemoveCert = Get-ExchangeCertificate | Where {$_.Status -eq "DateInvalid"}

# Prüft ob es zulöschende Zertifikate gibt, ansonsten Skript beenden
If ($RemoveCert -eq $Null) {break}

# Abgelaufene Zertifikate löschen
for ($i=0; $i -lt $RemoveCert.Length; $i++) {Remove-ExchangeCertificate -Thumbprint $RemoveCert[$i].Thumbprint -Confirm:$false}