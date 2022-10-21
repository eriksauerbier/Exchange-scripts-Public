## PS-Skript zum Exportieren aller E-Mail-Adressen (Attribut: proxyAddress) pro User
## Stannek GmbH v.1.0 - 20.10.2022 - E.Sauerbier

# Parameter
$FileOutputName = "ProxyAddresses-"+(Get-Date -Format dd-MM-yyyy)+".csv"

# ActiveDirectory PS-Modul laden
Import-Module ActiveDirectory

# Skriptpfad auslesen und Ausgabedatei erzeugen
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$FileOutput = $PSScriptRoot + "\" + $FileOutputName
"Name,proxyAddress" | Out-File $FileOutput

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
    # Wert in CSV schreiben
    $Output | Out-File $FileOutput -Append
  }
}
