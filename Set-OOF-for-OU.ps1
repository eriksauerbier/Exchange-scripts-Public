# Dieses Skript setzt den Abwesenheitsassistenten (OOF) für alle Benutzer oder eine bestimmte OU
# Stannek GmbH - v.1.1 - 02.06.2022 E.Sauerbier

#Parameter
$Starttime = "05.02.2023 18:00" # Format "dd.mm.jjjj HH:MM"
$Endtime = "10.02.2023 20:00" # Format "dd.mm.jjjj HH:MM"
$AutoReplyState = "Scheduled" # Alternative kann "Enabled" und "Disbaled" verwendet werden

# Dies ist der Text für die Abwesenheitsmeldung in HTML
$OOFText = @"
<html>
<head>
<title></title>
</head>
<body text="#000000" bgcolor="#FFFFFF" link="#FF0000" alink="#FF0000" vlink="#FF0000">
<font face="VERDANA,ARIAL,HELVETICA" size="-1">
<h3>Wichtiger Hinweis - Serverumstellung!</h3><br>

Sehr geehrte Damen und Herren,<br><br>
ab "Wochentag", xx.xx.xxxx, 18:00 Uhr bis einschlie&szlig;lich "Wochentag", xx.xx.xxxx, 20:00 Uhr findet bei uns eine gr&ouml;&szlig;ere Serverumstellung statt.<br>
In dieser zeit k&ouml;nnen wir keine E-Mails und Faxe empfangen und dementsprechend bearbeiten. Bei wichtigen Angelegenheiten nehmen Sie bitte telefonisch mit uns Kontakt auf.<br>
</font>
</body>
</html>
"@

# Exchange PS-Snapin laden
Add-PSSnapin *exchange*

# Organisationseinheit abfragen, wenn dies gewünscht ist
$OrgUnit = Get-OrganizationalUnit | Select-Object Name, DistinguishedName | Out-GridView -Title "Bei Bedarf Organisationseinheit wählen" -PassThru

# Abfrage ob OOF nur auf eine Organisationseinheit angewendet werden soll
If ($OrgUnit -eq $Null) {Get-Mailbox | Set-MailboxAutoReplyConfiguration -AutoReplyState $AutoReplyState -StartTime $Starttime -EndTime $Endtime -InternalMessage $OOFText -ExternalMessage $OOFText}
Else {Get-Mailbox -OrganizationalUnit $OrgUnit.DistinguishedName | Set-MailboxAutoReplyConfiguration -AutoReplyState $AutoReplyState -StartTime $Starttime -EndTime $Endtime -InternalMessage $OOFText -ExternalMessage $OOFText}