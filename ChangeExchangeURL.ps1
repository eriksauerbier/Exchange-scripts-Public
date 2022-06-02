# Skript zum Ändern der Exchange URLs
# Stannek GmbH - v.1.1 - 02.06.2022 - E.Sauerbier

#Exchange PSSnapin laden
Add-PSSnapin *exchange*

# Assembly für Inputbox laden 
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

# Abfrage der neuen ExchangeURL
$NewExchangeURL = [Microsoft.VisualBasic.Interaction]::InputBox('Exchange URL angeben (z.B. mail.domain.com)', 'Exchange URL')

# Skript abbrechen wenn nichts eingetragen wurde
If ($NewExchangeURL -eq "") {Break}

# Aktuelle ExchangeURL auslesen
$ExchangeURL = Get-OutlookAnywhere -Server $env:computername | Select-Object -ExpandProperty InternalHostname 

# Hinweis Meldung ausgeben
$Header = "Exchange URL"
$msg = "Die Exchange URL wird von $ExchangeURL in $NewExchangeURL geändert"
$Exclamation = [System.Windows.Forms.MessageBoxIcon]::Asterisk
$Noticebox = [System.Windows.Forms.Messagebox]::Show($msg,$Header,1,$Exclamation)

# Skript abbrechen wenn "Abbrechen" gedrückt wurde
If ($Noticebox -eq "Cancel") {Break}

# OWA URL ändern
$OWAurl = "https://" + "$NewExchangeURL" + "/owa"
write-host "OWA URL:" $OWAurl
Get-OwaVirtualDirectory -Server $env:computername | Set-OwaVirtualDirectory -internalurl $OWAurl -externalurl $OWAurl -WarningAction SilentlyContinue
 
# ECP URL ändern
$ECPurl = "https://" + "$NewExchangeURL" + "/ecp"
write-host "ECP URL:" $ECPurl
Get-EcpVirtualDirectory -server $env:computername| Set-EcpVirtualDirectory -internalurl $ECPurl -externalurl $ECPurl
 
# EWS URL ändern
$EWSurl = "https://" + "$NewExchangeURL" + "/EWS/Exchange.asmx"
write-host "EWS URL:" $EWSurl
Get-WebServicesVirtualDirectory -server $env:computername | Set-WebServicesVirtualDirectory -internalurl $EWSurl -externalurl $EWSurl -confirm:$false -force
 
# ActiveSync URL ändern
$EASurl = "https://" + "$NewExchangeURL" + "/Microsoft-Server-ActiveSync"
write-host "ActiveSync URL:" $EASurl
Get-ActiveSyncVirtualDirectory -Server $env:computername  | Set-ActiveSyncVirtualDirectory -internalurl $EASurl -externalurl $EASurl
 
# OfflineAdressbuch URL ändern
$OABurl = "https://" + "$NewExchangeURL" + "/OAB"
write-host "OAB URL:" $OABurl
Get-OabVirtualDirectory -Server $env:computername | Set-OabVirtualDirectory -internalurl $OABurl -externalurl $OABurl
 
# MAPIoverHTTP URL ändern
$MapiURL = "https://" + "$NewExchangeURL" + "/mapi"
write-host "MAPI URL:" $MapiURL
Get-MapiVirtualDirectory -Server $env:computername| Set-MapiVirtualDirectory -externalurl $MapiURL -internalurl $MapiURL
 
# Outlook Anywhere (RPCoverhTTP) #ndern
write-host "OA Hostname:" $NewExchangeURL
Get-OutlookAnywhere -Server $env:computername| Set-OutlookAnywhere -externalhostname $NewExchangeURL -internalhostname $NewExchangeURL -ExternalClientsRequireSsl:$true -InternalClientsRequireSsl:$true -ExternalClientAuthenticationMethod 'Negotiate' -wa 0
 
# Autodiscover URL ändern
$AutodiscoverURL = "https://" + "$NewExchangeURL" + "/Autodiscover/Autodiscover.xml"
write-host "Autodiscover URL:" $AutodiscoverURL
Get-ClientAccessServer $env:computername | Set-ClientAccessServer -AutoDiscoverServiceInternalUri $AutodiscoverURL
