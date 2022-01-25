# Dieses Skript sucht nach einer bestimmte E-Mail-Adresse am Exchange-Server
# Stannek GmbH - v.2.01 - 25.01.2022 E.Sauerbier

# Paramter
$TextboxTitle = "E-Mail Adresse eingeben"
$TextBoxExplanation = "Bitte zu suchende E-Mail Adresse eingeben:"

# Exchange Snap-In laden
Add-PSSnapin *exchange* -ErrorAction Stop

# Assembly für Hinweisbox laden
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

### Funktionen laden 

# Funktion Eingabefenster
Function Start-TextBox ($Input_Type,$Title){

    ### Assemblys laden
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(310,160) # Größe des Fenster
    $form.StartPosition = 'CenterScreen'  # Fensterposition festlegen
    $form.Topmost = $true  # Fenster immer im Vordergrund öffnen

    ### OK-Button hinzufügen
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(70,85) # Position des Buttons
    $OKButton.Size = New-Object System.Drawing.Size(75,23) # Größe des Buttons
    $OKButton.Text = 'OK' # Text des Buttons
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    ### Cancel-Button hinzufügen
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(155,85) # Position des Button
    $CancelButton.Size = New-Object System.Drawing.Size(75,23) # Größe des Buttons
    $CancelButton.Text = 'Cancel' # Text des Buttons
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    ### Text hinzufügen
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,10) # Position des Textes
    $label.AutoSize = $True
    $Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Regular) # Formatierung des Textes
    $label.Font = $Font
    $label.Text = $Input_Type # Text der als Parameter angeben wurde
    $label.ForeColor = 'Blue' # Farbe des Textes 
    $form.Controls.Add($label)

    ### Eingabefeld hinzufügen
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,55) # Position des Eingabefelds
    $textBox.Size = New-Object System.Drawing.Size(275,100) # Größe des Eingabefelds
    $textBox.Multiline = $false # Erlaube keine mehrere Zeilen
    $form.Controls.Add($textBox)

    $form.Add_Shown({$textBox.Select()}) # Form aktivieren und anzeigen
    $result = $form.ShowDialog()

    # Wenn "OK" gedrückt wird das Ergebnis weiter gegeben
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $Return = $textBox.Lines
        Return $Return
    }

    # Wenn "Cancel" gedrückt, dann Skript beenden
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {Break}

}



# Textbox laden

$MailAdress = Start-TextBox $TextBoxExplanation $TextboxTitle

# nach E-Mail Adresse suchen
$SearchResult =  get-recipient -results unlimited | where {$_.emailaddresses -match $MailAdress} | select name,recipienttype

# Ergebnis ausgeben

$Header = "Suchergebnis"
$msg = "Die E-Mail Adresse wurde folgenden Objekt ("+ $SearchResult.recipienttype + ") zugeordnet: " + $SearchResult.Name
$Exclamation = [System.Windows.Forms.MessageBoxIcon]::Asterisk
[System.Windows.Forms.Messagebox]::Show($msg,$Header,1,$Exclamation)