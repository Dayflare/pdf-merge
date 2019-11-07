#PDF-Merge
#created by Florian Müller

#init Powershell GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#create form
#erstellt mit https://poshgui.com/
$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '750,349'
$Form.text                       = "PDF Merge"
$Form.TopMost                    = $false
$Form.FormBorderStyle            = 'FixedDialog'
$Form.MinimizeBox                = $false
$Form.MaximizeBox                = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Dieses Programm sucht im angegebenen Ordner alle .tif Dateien in Unterordnern und erzeugt eine PDF Datei daraus."
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 50
$Label1.location                 = New-Object System.Drawing.Point(18,22)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 712
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(18,80)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Pfad Hauptordner"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(18,60)
$Label2.Font                     = 'Microsoft Sans Serif,10'

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Start"
$Button1.width                   = 107
$Button1.height                  = 44
$Button1.Anchor                  = 'top'
$Button1.location                = New-Object System.Drawing.Point(18,206)
$Button1.Font                    = 'Microsoft Sans Serif,16,style=Bold'

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 711
$ProgressBar1.height             = 35
$ProgressBar1.location           = New-Object System.Drawing.Point(18,294)
$ProgressBar1.Style              = 'Continuous'

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.width                  = 236
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(18,150)
$TextBox2.Font                   = 'Microsoft Sans Serif,10'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "Name der PDF Datei"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(18,130)
$Label3.Font                     = 'Microsoft Sans Serif,10'

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Fortschritt: "
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(18,276)
$Label4.Font                     = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($Label1,$TextBox1,$Label2,$Button1,$ProgressBar1,$TextBox2,$Label3,$Label4))

 #Script 
  $Button1.Add_Click({
    $ProgressBar1.Value = 10
    $Label4.text = 'Fortschritt: Initialisierung...'
    #Sleep Timer werden nach Aktualisierung des Progress Bar gesetzt, weil sonst der Text nicht angezeigt wird
    Start-Sleep 1

    #Start Button wird deaktiviert für die Laufzeit damit das Skript nicht doppelt ausgeführt werden kann
    $Button1.Enabled = $false
    $workdir = $TextBox1.Text
    $pdfname = $TextBox2.Text

    #Fehlermeldung wenn Pfad oder Dateiname nicht ausgefüllt werden
    if (!$workdir -or !$pdfname) {
      [Windows.Forms.MessageBox]::Show('Pfad oder Name der PDF Datei wurde nicht angegeben!', 'PDFMerge', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning)
      $Button1.Enabled = $true
      return
    }
    else {
      #Debug
      #write-host "Path: $workdir"
      #write-host "PDF File Name: $pdfname"
    }

    Try {
      #Temp Ordner erstellen
      New-Item -ItemType Directory -path $workdir\pdfmerge -Force
    }
    Catch {
      Write-Host "Error: Can't create Temp Directory. No read/write access."
      Return
    }

    $ProgressBar1.Value = 30
    $Label4.text = 'Fortschritt: Suche und kopiere Dateien...'
    Start-Sleep 1

    Try {
      #Alle .tif und .pdf rekursiv in den temp ordner kopieren
      ForEach($File in (Get-ChildItem -Path $workdir -File -recurse -Exclude "pdfmerge" -Include '*.tif','*.pdf')){
        Copy-Item $File -Destination $workdir\pdfmerge -ErrorAction Stop
      }
      #Alte Routine, kopiert nicht Dateien die im Hauptordner liegen
      #ForEach($Folder in (Get-ChildItem -Directory $workdir -Exclude "pdfmerge")){
       # Get-ChildItem -Path $Folder -File -recurse -include '*.tif','*.pdf' | Copy-Item -Destination $workdir\pdfmerge -ErrorAction Stop
        #}
    }
    Catch {
      Write-Host "Error: Can't find or copy files."
      Return
    }
    
    $ProgressBar1.Value = 70
    $Label4.text = 'Fortschritt: Konvertiere und Verarbeite PDF Dateien...'
    Start-Sleep 1

    #Prüfung ob temp Ordner leer ist, also keine Dateien gefunden wurden
    #PDF24 wird dann nicht aufgerufen, Fehlermeldungen von PDF24 können nicht abgefangen werden
    if((Get-ChildItem $workdir\pdfmerge -force | Select-Object -First 1 | Measure-Object).Count -eq 0) {
      Write-Host "Keine Dateien zur Verarbeitung gefunden."
    }
    else {
      try {
        #Aufruf PDF24 und Übergabe der Dokumente aus temp Ordner
        & "C:\Program Files (x86)\PDF24\pdf24-DocTool.exe" -join -profile default/good -outputfile $workdir\$($pdfname).pdf -expanddirsrecursive $workdir\pdfmerge
      }
      catch {
        write-host "Error: Can't find PDF24 Program."
        return
      }
      #Temp Ordner wird gelöscht, sobald die fertig zusammengeführte Datei gespeichert wurde
      while (!(Test-Path "$workdir\$($pdfname).pdf")) { Start-Sleep 5 }

      $ProgressBar1.Value = 90
      $Label4.text = 'Fortschritt: Lösche temporäre Dateien'
      Start-Sleep 2

      try {
        Remove-Item -path $workdir\pdfmerge -recurse -Force
      }
      catch {
        write-host "Error: Can't delete temp files. No read/write access."
        return
      }
    }

    #Falls keine Dateien gefunden wurden und PDF24 nicht aufgerufen wurde, wird hier der temp Ordner gelöscht falls er existiert
    $ProgressBar1.Value = 90
    $Label4.text = 'Fortschritt: Lösche temporäre Dateien'
    Start-Sleep 2

      if (Test-Path "$workdir\pdfmerge"){
        Remove-Item -path $workdir\pdfmerge -recurse -Force
      }

    $ProgressBar1.Value = 100
    $Label4.text = 'Fortschritt: Fertig'
    Start-Sleep 1

    #Start Button wird wieder aktiviert
    $Button1.Enabled = $true
  })

#show window
[void]$form.ShowDialog()