# --- SCRIPT PER CONVERTIRE WORD IN PDF ---
# Non richiede installazioni esterne. Usa Microsoft Word installato sul PC.

# Carichiamo gli strumenti per mostrare le finestre di dialogo
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 1. Chiediamo all'utente di scegliere la cartella
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Seleziona la cartella contenente i file Word"
$folderBrowser.ShowNewFolderButton = $false

$result = $folderBrowser.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Operazione annullata dall'utente." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    exit
}

$startPath = $folderBrowser.SelectedPath

Write-Host "Avvio Word in background..." -ForegroundColor Cyan

# 2. Creiamo l'oggetto Word invisibile
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone
}
catch {
    Write-Host "Errore: Impossibile avviare Microsoft Word." -ForegroundColor Red
    Write-Host "Assicurati che Word sia installato."
    Read-Host "Premi Invio per uscire"
    exit
}

# Costante per il formato PDF (wdFormatPDF = 17)
$wdFormatPDF = 17
$count = 0
$errors = 0

# 3. Cerchiamo tutti i file .doc e .docx nelle sottocartelle
$files = Get-ChildItem -Path $startPath -Include *.doc, *.docx -Recurse -File | Where-Object { $_.Name -notlike "~$*" }

if ($files.Count -eq 0) {
    Write-Host "Nessun file Word trovato nella cartella selezionata." -ForegroundColor Yellow
}
else {
    Write-Host "Trovati $($files.Count) file. Inizio conversione..." -ForegroundColor Green
    
    foreach ($file in $files) {
        $pdfPath = [System.IO.Path]::ChangeExtension($file.FullName, ".pdf")
        
        # Saltiamo se il PDF esiste già
        if (Test-Path $pdfPath) {
            Write-Host "Saltato (esiste già): $($file.Name)" -ForegroundColor Gray
            continue
        }

        Write-Host "Convertendo: $($file.Name)..." -NoNewline

        try {
            # Apriamo il documento
            $doc = $word.Documents.Open($file.FullName, $false, $true) # ReadOnly=True
            
            # Salviamo come PDF
            $doc.SaveAs([ref]$pdfPath, [ref]$wdFormatPDF)
            
            # Chiudiamo il documento senza salvare modifiche
            $doc.Close([ref]$false)
            
            Write-Host " OK" -ForegroundColor Green
            $count++
        }
        catch {
            Write-Host " ERRORE" -ForegroundColor Red
            Write-Host "Dettagli: $($_.Exception.Message)" -ForegroundColor Red
            $errors++
        }
    }
}

# 4. Pulizia finale
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Host "`n--- Operazione Completata ---" -ForegroundColor Cyan
Write-Host "Convertiti: $count"
Write-Host "Errori: $errors"
Write-Host "Premi INVIO per chiudere questa finestra..."
Read-Host