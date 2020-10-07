#PDF-Merge

#cli params
param($Source, $output)

#init
#PDF Library
Add-Type -Path '.\PdfSharp-gdi.dll'
#Load WPF
Add-Type -AssemblyName PresentationFramework

#Main Function
Function Click_Start {
  Param ($workdir, $pdfname)

  $StatusBar.Value = 10
  $StatusText.Content = 'Initialisierung...'
  Update-Gui
  Start-Sleep 1

  #Start Button wird deaktiviert für die Laufzeit damit das Skript nicht doppelt ausgeführt werden kann
  $Button1.IsEnabled = $False
  Update-Gui

  #Fehlermeldung wenn Pfad oder Dateiname nicht ausgefüllt werden
  if (!$workdir -or !$pdfname) {
    [System.Windows.MessageBox]::Show("Pfad oder Name der PDF Datei wurde nicht angegeben!")
    $Button1.IsEnabled = $true
    $StatusBar.Value = 0
    $StatusText.Content = ''
    Update-Gui
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
    [System.Windows.MessageBox]::Show("Error: Can't create Temp Directory.")
    Cleanup
    Return
  }

  $StatusBar.Value = 30
  $StatusText.Content = 'Suche und kopiere Dateien...'
  Update-Gui
  Start-Sleep 1

  If (!$checkbox_TIFF.IsChecked -AND !$checkbox_PDF.IsChecked){
    Write-Host "No checkbox checked. Abort."
    [System.Windows.MessageBox]::Show("Error: Es wurde kein Dateityp ausgewaehlt.")
    Cleanup
    return
  }

  Try {
    $excluded = "pdfmerge"

    If ($checkbox_PDF.IsChecked) {
      Write-Host "PDF File Extension checked."
      $included = @('*.pdf')
    }
    Else {
      #Alle .tif und .pdf rekursiv in den temp ordner kopieren
      $included = @('*.pdf','*.tif')
    }

    If ($checkbox_exclusion_Archiv.IsChecked) {
      Write-Host "Exclude Archiv Directory."
      $excluded = "pdfmerge|Archiv"
    }

    If ($checkbox_options_recursive.IsChecked)
    {
      $pathrecurse = $workdir
    }
    Else {
      $pathrecurse = "$workdir\*.*"
    }

    ForEach($File in (Get-ChildItem -Path $pathrecurse -File -recurse:$checkbox_options_recursive.IsChecked -Include $included | Where-Object {$_.PSParentPath -notmatch $excluded})){
      #Write-Host "$File"
      Copy-Item $File -Destination $workdir\pdfmerge -ErrorAction Stop
    }
  }

  Catch {
    Write-Host "Error: Can't find or copy files."
    [System.Windows.MessageBox]::Show("Error: Can't find or copy files.")
    Cleanup
    Return
  }
  
  $StatusBar.Value = 70
  $StatusText.Content = 'Konvertiere und Verarbeite PDF Dateien...'
  Start-Sleep 1
  Update-Gui

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
      if ($checkbox_TIFF.IsChecked) {
      #Aufruf PDF24 und Übergabe der Dokumente aus temp Ordner
      & "C:\Program Files (x86)\PDF24\pdf24-DocTool.exe" -join -profile default/good -outputfile $workdir\$($pdfname).pdf -expanddirsrecursive $workdir\pdfmerge
      }
    }
    catch {
      write-host "Error: Can't find PDF24 Program."
      Cleanup
      return
    }
    #Temp Ordner wird gelöscht, sobald die fertig zusammengeführte Datei gespeichert wurde
    while (!(Test-Path "$workdir\$($pdfname).pdf")) { Start-Sleep 5 }
    Cleanup
  }

  #Falls keine Dateien gefunden wurden und PDF24 nicht aufgerufen wurde, wird hier der temp Ordner gelöscht falls er existiert
  Cleanup
}

Function Cleanup {
  $StatusBar.Value = 90
  $StatusText.Content = 'Loesche temporaere Dateien'
  Update-Gui
  Start-Sleep 1

  if (Test-Path "$workdir\pdfmerge"){
    Remove-Item -path $workdir\pdfmerge -recurse -Force
  }

  $StatusBar.Value = 100
  $StatusText.Content = 'Fertig'
  Update-Gui
  Start-Sleep 1

  #Start Button wird wieder aktiviert
  $Button1.IsEnabled = $true
  Update-Gui

  if ($parameters) {
    Write-Host "Finished!"
  }
}

#PDF Merge Funktion wenn nur PDF Dateien zusammengefügt werden
#Beispiel: Merge-PDF -path c:\pdf_files -filename c:\merged-files.pdf
Function Merge-PDF {
  Param ($path, $filename)

  $pdf = New-Object PdfSharp.Pdf.PdfDocument

  foreach ($file in (Get-ChildItem $path *.pdf -Recurse)) {
    $document = [PdfSharp.Pdf.IO.PdfReader]::Open($file.FullName, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)

    for ($index = 0; $index -lt $document.PageCount; $index++){
      $page = $Document.Pages[$index]
      $null = $pdf.AddPage($page)
    }
  }
  $pdf.Save($filename)
}

#XAML GUI
#$App = New-Object -TypeName Windows.Application
#$App = New-Object Windows.Application
[xml]$XamlMain = @"
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
        <GroupBox x:Name="groupbox_Ordner" HorizontalAlignment="Left" Height="150" Margin="37,21,0,0" VerticalAlignment="Top" Width="550" Header="Ordner" AllowDrop="True">
            <StackPanel AllowDrop="True">
                <Label x:Name="label_Source" Content="Quelle" Margin="10,10,0,0"/>
                <TextBox x:Name="SourcePath" Height="23" Margin="15,0,0,0" TextWrapping="Wrap" Width="480" AllowDrop="True" ToolTip="Ordnerpfad mit PDF Dateien hier eingeben" HorizontalAlignment="Left"/>
                <Label x:Name="label_destination" Content="Dateiname" Margin="10,0,0,0"/>
                <TextBox x:Name="DestinationPath" Height="23" Margin="15,0,0,0" TextWrapping="Wrap" Width="480" ToolTip="Dateiname der zusammengefuegten PDF Datei, ohne Dateiendung, nur der Name. Die Datei wird im angegeben Pfad unter Quelle erstellt." AllowDrop="True" HorizontalAlignment="Left"/>
            </StackPanel>
        </GroupBox>
        <Button x:Name="ButtonStart" Content="Start" HorizontalAlignment="Left" Margin="37,303,0,0" VerticalAlignment="Top" Width="150" Height="40" FontSize="24" FontWeight="Bold"/>
        <ProgressBar x:Name="StatusBar" HorizontalAlignment="Left" Height="30" Margin="37,379,0,0" VerticalAlignment="Top" Width="707"/>
        <GroupBox x:Name="groupBox_Dateitypen" Header="Dateitypen" HorizontalAlignment="Left" Height="100" Margin="37,186,0,0" VerticalAlignment="Top" Width="180">
            <StackPanel>
                <CheckBox x:Name="checkBox_PDF" Content="PDF" Margin="5,5,0,0" VerticalAlignment="Top" IsChecked="True" />
                <CheckBox x:Name="checkBox_TIFF" Content="TIFF" Margin="5,5,0,0" VerticalAlignment="Top" HorizontalAlignment="Left" Width="40"/>
            </StackPanel>
        </GroupBox>
        <Label x:Name="TextStatus" Content="" Margin="217,379,226.333,0" VerticalAlignment="Top" Width="350" Height="30" HorizontalContentAlignment="Center"/>
        <Label x:Name="label_status" Content="Fortschritt:" HorizontalAlignment="Left" Margin="37,348,0,0" VerticalAlignment="Top"/>
        <GroupBox x:Name="groupBox_Ausnahmen" Header="Ausnahmen" HorizontalAlignment="Left" Height="100" Margin="222,186,0,0" VerticalAlignment="Top" Width="180">
            <CheckBox x:Name="checkbox_Ausnahme_Archiv" Content="Archiv Ordner" HorizontalAlignment="Left" Margin="3,8,0,0" VerticalAlignment="Top" IsChecked="True" />
        </GroupBox>
        <Rectangle x:Name="DragandDropBox" Fill="White" HorizontalAlignment="Left" Height="256" Margin="610,30,0,0" Stroke="Black" VerticalAlignment="Top" Width="134" Cursor="Arrow"/>
        <Label x:Name="DropText" Content="Ordner hier hin &#xD;&#xA;ziehen um den &#xD;&#xA;Pfad unter Quelle&#xD;&#xA;einzutragen!" HorizontalAlignment="Left" Margin="621,119,0,0" VerticalAlignment="Top" Height="81" Width="113" Cursor="Arrow"/>
        <GroupBox x:Name="groupBox_Optionen" Header="Optionen" HorizontalAlignment="Left" Height="100" Margin="407,186,0,0" VerticalAlignment="Top" Width="180">
            <CheckBox x:Name="checkbox_optionen_rekursiv" Content="Rekursiv" HorizontalAlignment="Left" Margin="4,8,0,0" VerticalAlignment="Top" ToolTip="Dateien aus allen Unterordnern zusammenfassen" />
        </GroupBox>
    </Grid>
</Window>
"@ -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'

#create GUI
$window=[Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XAMLMain))

function Update-Gui {
  $window.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Background, [action]{})
}

#close handle
$window.add_Closing({
  #$App.Shutdown()
  #Stop-Process $pid
})

#Deklariere Variablen für GUI Steuerelemente
$SourcePath = $window.FindName("SourcePath")
$DestinationPath = $window.FindName("DestinationPath")
$Button1 = $window.FindName("ButtonStart")
$StatusBar = $window.FindName("StatusBar")
$StatusText = $window.FindName("TextStatus")
$checkbox_PDF = $window.FindName("checkBox_PDF")
$checkbox_TIFF = $window.FindName("checkBox_TIFF")
$checkbox_exclusion_Archiv = $window.FindName("checkbox_Ausnahme_Archiv")
$checkbox_options_recursive = $window.FindName("checkbox_optionen_rekursiv")

$window.AllowDrop = $true
# $SourcePath.AllowDrop = $true

#Button Click Event
$Button1.Add_Click({
  Click_Start -workdir $SourcePath.Text.ToString() -pdfname $DestinationPath.Text.ToString()
})

#Drag and Drop Event
$window.Add_Drop({
    $content = [string]$_.Data.GetFileDropList()
    if ((Get-Item $content) -is [System.IO.DirectoryInfo]) {
      $SourcePath.Text = $content

      $DestinationName = Split-Path "$content" -Leaf
      $Date = Get-Date -Format "yyyy-MM-dd"
      Write-Host "suggested destination name will be: $($DestinationName)_$($Date)"
      $DestinationPath.Text = "$DestinationName"+"_"+"$Date"
    }
    else {
      write-host "dropped item not a folder, ignoring."
    }
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
  #$App.Run($window)
  $window.ShowDialog() | Out-Null
}
