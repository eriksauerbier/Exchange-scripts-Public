# Skript zum bereinigen der Exchange Log-Files im IIS und im Exchange
# Stannek GmbH - v.2.01 - 31.05.2022 - E.Sauerbier

# Parameter
$deleteafterdays = "14" # Alter in Tagen der Log-Files festlegen
$ScriptPath = "C:\Skripte\"
$ScriptName = $MyInvocation.MyCommand.Name
$TaskName = "Clean Exchange Logfiles"

# Assembly für Hinweisboxen laden
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Powershell-Module laden

import-module webadministration

# Pfade festlegen
$ExLogpath1 = $env:ExchangeInstallPath + "Logging"
$ExLogpath2 = $env:ExchangeInstallPath + "Bin\Search\Ceres\Diagnostics\Logs"

# Lösch Datum erzeugen
$deleteafter = (get-date).adddays(-$deleteafterdays)

# Task prüfen und ggf. anlegen

If (!(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)) {
    # Task nicht vorhanden, wird nun angelegt

    # Prüfen ob Skript auf im richtigen Pfad liegt
     if (!(Test-Path "$ScriptPath\$ScriptName")){
     $msg = "Das Skript für den Task liegt nicht unter $ScriptPath"
     $Header = "Das Skript für den Task fehlt"
     $Exclamation = [System.Windows.Forms.MessageBoxIcon]::Warning
     [System.Windows.Forms.Messagebox]::Show($msg,$header,1,$Exclamation)
     break
     }

    # Task anlegen
    $TaskArgument = "-File " + $Scriptpath + $ScriptName
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $TaskArgument
    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 7pm
    $User = New-ScheduledTaskPrincipal "NT AUTHORITY\SYSTEM"  -LogonType ServiceAccount
    $Tasksettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable -ExecutionTimeLimit 01:00:00
    $ST = New-ScheduledTask -Action $action -Principal $User -Trigger $trigger -Settings $Tasksettings 
    $task = Register-ScheduledTask $TaskName -InputObject $ST
    $task | Set-ScheduledTask
    }

# Exchange Log-Files bereinigen
Get-ChildItem -Path $ExLogpath1 -Recurse *.log -force  | Where-Object {$_.lastwritetime -lt $deleteafter} | Remove-Item -force -ErrorAction SilentlyContinue
Get-ChildItem -Path $ExLogpath2 -Recurse *.log -force  | Where-Object {$_.lastwritetime -lt $deleteafter} | Remove-Item -force -ErrorAction SilentlyContinue

# IIS Webseiten auslesen
$websites = get-website

# Schleife um zu löschende Logfiles auszulesen
foreach ($website in $websites)
    {# Logfile Verzeichnis auslesen
     $logfiledir = $website.logfile.directory
     # falls vorhanden Umgebungsvariable %Systemdrive% ersetzen 
     if ($LogFiledir -match "%SystemDrive%") {$logfiledir  = $logfiledir  -replace "%SystemDrive%",$env:SystemDrive}
     # Liste mit alle zu löschenden Logfiles erstellen 
     $logfilelist = Get-ChildItem $logfiledir -Recurse | Where-Object {! $_.PSIsContainer -and $_.lastwritetime -lt $deleteafter} | Select-Object fullname
     } 

# Logfiles in entsprechender Seite löschen
foreach ($logfile in $logfilelist)
     {remove-item $logfile.fullname}
