#PDF-Merge
#created by Florian Müller

#cli params
param($Source, $output)

#init
#PDF Library
Add-Type -Path '.\PdfSharp.dll'
#Load WPF
Add-Type -AssemblyName PresentationFramework

#Main Function
Function Click_Start {
  Param ($workdir, $pdfname)

  $StatusBar.Value = 10
  $StatusText.Content = 'Initialisierung...'
  Start-Sleep 1

  #Start Button wird deaktiviert für die Laufzeit damit das Skript nicht doppelt ausgeführt werden kann
  $Button1.IsEnabled = $False

  #Fehlermeldung wenn Pfad oder Dateiname nicht ausgefüllt werden
  if (!$workdir -or !$pdfname) {
    New-WPFMessageBox -Content "Pfad oder Name der PDF Datei wurde nicht angegeben!"
    $Button1.IsEnabled = $true
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

  $StatusBar.Value = 30
  $StatusText.Content = 'Suche und kopiere Dateien...'
  Start-Sleep 1

  Try {
    If ($checkbox_PDF.IsChecked) {
      Write-Host "PDF File Extension checked."
      ForEach($File in (Get-ChildItem -Path $workdir -File -recurse -Exclude "pdfmerge" -Include '*.pdf')){
        Copy-Item $File -Destination $workdir\pdfmerge -ErrorAction Stop
      }
    }
    else {
      #Alle .tif und .pdf rekursiv in den temp ordner kopieren
      ForEach($File in (Get-ChildItem -Path $workdir -File -recurse -Exclude "pdfmerge" -Include '*.tif','*.pdf')){
        Copy-Item $File -Destination $workdir\pdfmerge -ErrorAction Stop
      }
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
  
  $StatusBar.Value = 70
  $StatusText.Content = 'Konvertiere und Verarbeite PDF Dateien...'
  Start-Sleep 1

  #Prüfung ob temp Ordner leer ist, also keine Dateien gefunden wurden
  #PDF24 wird dann nicht aufgerufen, Fehlermeldungen von PDF24 können nicht abgefangen werden
  if((Get-ChildItem $workdir\pdfmerge -force | Select-Object -First 1 | Measure-Object).Count -eq 0) {
    Write-Host "Keine Dateien zur Verarbeitung gefunden."
  }
  elseif ($checkbox_PDF.IsChecked -or $parameters) {
    Write-Host "Using PDF Sharp Library"
    Merge-PDF -path $workdir\pdfmerge -filename $workdir\$($pdfname).pdf
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

    $StatusBar.Value = 90
    $StatusText.Content = 'Lösche temporäre Dateien'
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
  $StatusBar.Value = 90
  $StatusText.Content = 'Lösche temporäre Dateien'
  Start-Sleep 2

    if (Test-Path "$workdir\pdfmerge"){
      Remove-Item -path $workdir\pdfmerge -recurse -Force
    }

  $StatusBar.Value = 100
  $StatusText.Content = 'Fertig'
  Start-Sleep 1

  #Start Button wird wieder aktiviert
  $Button1.IsEnabled = $true
}

#PDF Merge Funktion wenn nur PDF Dateien zusammengefügt werden
#Beispiel: Merge-PDF -path c:\pdf_files -filename c:\merged-files.pdf
Function Merge-PDF {
  Param ($path, $filename)

  $output = New-Object PdfSharp.Pdf.PdfDocument
  $PDFReader = [PdfSharp.Pdf.IO.PdfReader]
  $PdfDocumentOpenMode = [PdfSharp.Pdf.IO.PdfDocumentOpenMode]

  foreach ($i in (Get-ChildItem $path *.pdf -Recurse)) {
    $input = New-Object PdfSharp.Pdf.PdfDocument
    $input = $PdfReader::Open($i.fullname, $PdfDocumentOpenMode::Import)
    $input.Pages | ForEach-Object{$output.AddPage($_)}
  }

  $output.Save($filename)
}

#XAML GUI
[xml]$XAMLMain = @"
<Window x:Class="WpfApp1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
        mc:Ignorable="d"
        Title="PDF-Merge" Height="480" Width="800"
        ResizeMode="NoResize">
    <Grid>
        <GroupBox Header="PDF-Merge" HorizontalAlignment="Left" Height="150" Margin="37,21,0,0" VerticalAlignment="Top" Width="530">
            <Grid HorizontalAlignment="Left" Height="111" VerticalAlignment="Top" Width="503" Margin="10,10,0,0">
                <TextBox x:Name="SourcePath" HorizontalAlignment="Left" Height="23" Margin="10,31,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="480"/>
                <TextBox x:Name="DestinationPath" HorizontalAlignment="Left" Height="23" Margin="10,80,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="480"/>
                <Label x:Name="label_Source" Content="Quelle" HorizontalAlignment="Left" Margin="10,5,0,0" VerticalAlignment="Top"/>
                <Label x:Name="label_destination" Content="Ziel" HorizontalAlignment="Left" Margin="10,54,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.525,0.577"/>
            </Grid>
        </GroupBox>
        <Button x:Name="ButtonStart" Content="Start" HorizontalAlignment="Left" Margin="50,300,0,0" VerticalAlignment="Top" Width="150" Height="40" FontSize="24" FontWeight="Bold"/>
        <ProgressBar x:Name="StatusBar" HorizontalAlignment="Left" Height="30" Margin="50,379,0,0" VerticalAlignment="Top" Width="694"/>
        <GroupBox x:Name="groupBox" Header="Dateitypen" HorizontalAlignment="Left" Height="100" Margin="37,186,0,0" VerticalAlignment="Top" Width="530"/>
        <CheckBox x:Name="checkBox_PDF" Content="PDF" HorizontalAlignment="Left" Margin="50,210,0,0" VerticalAlignment="Top" />
        <CheckBox x:Name="checkBox_TIFF" Content="TIFF" HorizontalAlignment="Left" Margin="50,230,0,0" VerticalAlignment="Top"/>
        <Label x:Name="TextStatus" Content="" Margin="217,379,226.333,0" VerticalAlignment="Top" Width="350" Height="30" HorizontalContentAlignment="Center"/>
        <Label x:Name="label_status" Content="Fortschritt:" HorizontalAlignment="Left" Margin="53,348,0,0" VerticalAlignment="Top"/>
    </Grid>
</Window>
"@ -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'

#create GUI
$window=[Windows.Markup.XamlReader]::Load( (New-Object System.Xml.XmlNodeReader $XAMLMain))

#Deklariere Variablen für GUI Steuerelemente
$SourcePath = $window.FindName("SourcePath")
$DestinationPath = $window.FindName("DestinationPath")
$Button1 = $window.FindName("ButtonStart")
$StatusBar = $window.FindName("StatusBar")
$StatusText = $window.FindName("TextStatus")
$checkbox_PDF = $window.FindName("checkBox_PDF")
$checkbox_TIFF = $window.FindName("checkBox_TIFF")

 #Button Click Event
  $Button1.Add_Click({
    Click_Start -workdir $SourcePath.Text.ToString() -pdfname $DestinationPath.Text.ToString()
  })

#parameter check
If ($PSBoundParameters.ContainsKey('source') -and ($PSBoundParameters.ContainsKey('output'))) {
  Write-Host "Parameters provided"
  $parameters = $true
  Click_Start -workdir $source -pdfname $output
}
else {
  $parameters = $false
}

#show window
if ($parameters -eq $false) {
  $window.ShowDialog() | Out-Null
}