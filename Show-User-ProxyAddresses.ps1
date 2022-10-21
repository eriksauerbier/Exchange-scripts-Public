## PS-Skript zum Anzeigen aller E-Mail-Adressen pro User
## Stannek GmbH v.1.0 - 20.10.2022 - E.Sauerbier

# Parameter

# ActiveDirectory PS-Modul laden
Import-Module ActiveDirectory

# Organisationseinheit erfragen
$OrgUnit = Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName | Out-GridView -Title "Bitte wünschte Organisationseinheit angeben" -OutPutMode Single

# Attribut "ProxyAddresses" auslesen

$ADObjects = Get-ADObject -SearchBase $OrgUnit.DistinguishedName -Filter * -Properties proxyAddresses

# Werte in CSV schreiben und ausgeben
ForEach ($Object In $ADObjects) {
  ForEach ($proxyAddress in $Object.proxyAddresses) {
    $Output = $Object.Name + "," + $proxyAddress
    # Zeile ausgeben
    Write-Host $Output
  }
}