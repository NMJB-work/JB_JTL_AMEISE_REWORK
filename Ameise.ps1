<#
.SYNOPSIS
    Führt SQL-Abfragen aus Dateien gegen einen SQL Server aus und exportiert die Ergebnisse als CSV oder Excel.

.DESCRIPTION
    Das Script liest eine oder mehrere SQL-Dateien ein, führt die Abfragen auf einem SQL Server aus
    und exportiert die Resultsets als CSV und/oder Excel-Dateien (XLSX).
    - Umlaut-kompatibler CSV-Export (Encoding: Default, z. B. Windows-1252).
    - Excel-Export über COM-Interop (ohne GUI-Prompts, mit AutoFilter und fixierter Kopfzeile).
    - Unterstützt eine einzelne SQL-Datei oder einen ganzen Ordner voller SQL-Dateien.

.PARAMETER SqlQueryFile
    Pfad zu einer einzelnen SQL-Datei. (ParameterSet: Single)

.PARAMETER SqlQueryFolder
    Ordner, aus dem alle .sql-Dateien verarbeitet werden. (ParameterSet: Folder)

.PARAMETER ServerName
    Name oder Instanz des SQL Servers (z. B. "SQL01" oder "SQL01\INSTANZ").

.PARAMETER DatabaseName
    Name der Datenbank, auf der die Abfrage ausgeführt wird.

.PARAMETER Username
    SQL-Login-Benutzername.

.PARAMETER Password
    SQL-Login-Passwort (im Klartext, z. B. für Automatisierung / geplante Tasks).

.PARAMETER OutputPath
    Zielordner für die Ausgabedateien. Standard ist das aktuelle Verzeichnis.

.PARAMETER OutputFile
    Basisname für die Ausgabedatei(en).
    - Single-Mode: direkt der Dateiname (ohne Erweiterung).
    - Folder-Mode: Basispräfix, an das der SQL-Dateiname angehängt wird.

.PARAMETER Format
    Ausgabformat: 'Csv' oder 'Excel'.

.PARAMETER AppendTimestamp
    Wenn gesetzt, wird an den Dateinamen ein Zeitstempel angehängt (z. B. _20251119_103000).

.PARAMETER KeepCsv
    Wenn gesetzt und Format = Excel, bleibt die CSV-Datei erhalten.
    Wenn nicht gesetzt, wird die CSV nach erfolgreicher Excel-Erzeugung gelöscht.

.PARAMETER SheetName
    Optionaler Name für das Excel-Blatt. Falls leer, bleibt der Standardname bestehen.

.EXAMPLE
    .\ExportReport.ps1 -SqlQueryFile ".\Query.sql" -ServerName "SQL01" -DatabaseName "ReportDB" `
        -Username "reportuser" -Password "geheim" -Format Excel

.EXAMPLE
    .\ExportReport.ps1 -SqlQueryFolder "C:\Sql\Reports" -ServerName "SQL01" -DatabaseName "ReportDB" `
        -Username "reportuser" -Password "geheim" -Format Csv -AppendTimestamp
#>

[CmdletBinding(DefaultParameterSetName = 'Single')]
param(
    # --- Eingabequelle: einzelne SQL-Datei ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Single')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$SqlQueryFile,

    # --- Eingabequelle: Ordner mit SQL-Dateien ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Folder')]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$SqlQueryFolder,

    # --- Verbindungsdaten ---
    [Parameter(Mandatory = $true)]
    [string]$ServerName,

    [Parameter(Mandatory = $true)]
    [string]$DatabaseName,

    [Parameter(Mandatory = $true)]
    [string]$Username,

    [Parameter(Mandatory = $true)]
    [string]$Password,

    # --- Ausgabe ---
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile = "QueryResults",

    [Parameter(Mandatory = $true)]
    [ValidateSet('Csv','Excel')]
    [string]$Format,

    [switch]$AppendTimestamp,
    [switch]$KeepCsv,

    [string]$SheetName
)

# --- Hilfsfunktion: eine einzelne SQL-Datei verarbeiten ---
function Invoke-DbQueryExport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SqlFile,

        [Parameter(Mandatory = $true)]
        [string]$BaseOutputName,

        [Parameter(Mandatory = $true)]
        [string]$ServerName,

        [Parameter(Mandatory = $true)]
        [string]$DatabaseName,

        [Parameter(Mandatory = $true)]
        [string]$Username,

        [Parameter(Mandatory = $true)]
        [string]$Password,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Csv','Excel')]
        [string]$Format,

        [switch]$KeepCsv,

        [string]$SheetName
    )

    if (-not (Test-Path $SqlFile -PathType Leaf)) {
        Write-Error "SQL-Datei nicht gefunden: '$SqlFile'"
        throw
    }

    # Ausgabe-Pfade definieren
    $BaseFilePath = Join-Path $OutputPath $BaseOutputName
    $CsvPath      = "$BaseFilePath.csv"
    $ExcelPath    = "$BaseFilePath.xlsx"

    Write-Host ""
    Write-Host "------------------------------------------------------------"
    Write-Host "Verarbeite SQL-Datei : $SqlFile"
    Write-Host "Zieldatei (Basis)    : $BaseFilePath"
    Write-Host "Format               : $Format"
    Write-Host "------------------------------------------------------------"

    # --- 1. SQL-Abfrage ausführen ---
    try {
        $Query = Get-Content $SqlFile -Raw

        Write-Host "Verbinde mit Server '$ServerName', DB '$DatabaseName' als '$Username'..."
        $Results = Invoke-Sqlcmd -ServerInstance $ServerName `
                                 -Database $DatabaseName `
                                 -Query $Query `
                                 -Username $Username `
                                 -Password $Password `
                                 -TrustServerCertificate `
                                 -ErrorAction Stop `
                                 -AbortOnError

        if (-not $Results) {
            Write-Warning "Abfrage lieferte keine Zeilen. CSV/Excel werden trotzdem (als leere Struktur) erzeugt."
        }

        # --- 2. CSV exportieren (Encoding: Default für Umlaute in Excel) ---
        $Results | Export-Csv -Path $CsvPath -NoTypeInformation -Delimiter ";" -Encoding Default
        Write-Host "✅ CSV exportiert: $CsvPath"
    }
    catch {
        Write-Error "🔴 Fehler bei SQL-Ausführung oder CSV-Export für '$SqlFile': $($_.Exception.Message)"
        throw
    }

    # --- 3. Optional: Konvertierung zu Excel ---
    if ($Format -eq 'Excel') {
        Write-Host "Starte Konvertierung von CSV zu Excel (XLSX)..."

        $Missing                     = [Type]::Missing
        $xlDelimited                 = 1
        $xlTextQualifierDoubleQuote  = 2
        $xlTextFormat                = 2
        $xlOpenXMLWorkbook           = 51

        $Excel      = $null
        $Workbook   = $null
        $Worksheet  = $null
        $Range      = $null

        try {
            $Excel = New-Object -ComObject Excel.Application -ErrorAction Stop
            $Excel.Visible = $false

            # CSV öffnen
            $Workbook = $Excel.Workbooks.Open($CsvPath)
            if (-not $Workbook) {
                throw "Workbook konnte nicht geöffnet werden."
            }

            $Worksheet = $Workbook.Sheets.Item(1)
            $Range     = $Worksheet.Range("A:A")

            # Spaltenformate vorbereiten (alle als Text),
            # Anzahl hier nach Bedarf anpassen (z. B. 50 Spalten)
            $NumColumns   = 50
            $ColumnFormats = @()
            for ($i = 1; $i -le $NumColumns; $i++) {
                # wichtig: ",@" für verschachtelte Arrays
                $ColumnFormats += ,@($i, $xlTextFormat)
            }

            # TextToColumns: Semikolon-getrennt, alles als Text, Locale = true
            $Range.TextToColumns(
                $Missing,                                # Destination
                $xlDelimited,                            # DataType
                $xlTextQualifierDoubleQuote,             # TextQualifier
                $false,                                  # ConsecutiveDelimiter
                $false,                                  # Tab
                $true,                                   # Semicolon
                $false,                                  # Comma
                $false,                                  # Space
                $false,                                  # Other
                $Missing,                                # OtherChar
                [System.Object[]]$ColumnFormats,         # FieldInfo
                $Missing,                                # DecimalSeparator
                $Missing,                                # ThousandsSeparator
                $Missing                                 # TrailingMinusNumbers
            )

            # Optional: SheetName setzen
            if ($SheetName -and $SheetName.Trim().Length -gt 0) {
                try {
                    $Worksheet.Name = $SheetName
                }
                catch {
                    Write-Warning "Konnte SheetName '$SheetName' nicht setzen: $($_.Exception.Message)"
                }
            }

            # AutoFilter + Kopfzeile fixieren
            $usedRange = $Worksheet.UsedRange
            $usedRange.AutoFilter() | Out-Null

            $Excel.ActiveWindow.SplitRow = 1
            $Excel.ActiveWindow.FreezePanes = $true

            # Falls Zieldatei existiert → löschen, um Excel-Dialoge zu vermeiden
            if (Test-Path $ExcelPath) {
                Remove-Item $ExcelPath -Force
            }

            # Speichern als XLSX
            $Workbook.SaveAs($ExcelPath, $xlOpenXMLWorkbook)
            Write-Host "✅ Excel-Datei erstellt: $ExcelPath"
        }
        catch {
            Write-Error "🔴 Fehler bei der Excel-Konvertierung für '$SqlFile': $($_.Exception.Message)"
            throw
        }
        finally {
            if ($Workbook)  { $Workbook.Close(0) | Out-Null }
            if ($Excel)     { $Excel.Quit()     | Out-Null }

            if ($Range)     { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Range)     | Out-Null }
            if ($Worksheet) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null }
            if ($Workbook)  { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook)  | Out-Null }
            if ($Excel)     { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)     | Out-Null }
        }

        # CSV löschen, sofern nicht explizit gewünscht
        if (-not $KeepCsv) {
            Remove-Item $CsvPath -Force
            Write-Host "🧹 Temporäre CSV-Datei gelöscht: $CsvPath"
        }
    }
    else {
        Write-Host "Format = CSV → Excel-Konvertierung wird übersprungen."
    }

    Write-Host "✔ Fertig: $SqlFile"
}

# --- Hauptlogik ---

# Zielordner sicherstellen
if (-not (Test-Path $OutputPath -PathType Container)) {
    Write-Host "Erstelle Ausgabeverzeichnis: $OutputPath"
    $null = New-Item -Path $OutputPath -ItemType Directory -Force
}

# Zeitstempel vorbereiten (falls benötigt)
$timestampSuffix = ""
if ($AppendTimestamp) {
    $timestampSuffix = "_" + (Get-Date -Format "yyyyMMdd_HHmmss")
}

try {
    switch ($PSCmdlet.ParameterSetName) {
        'Single' {
            # Single-Mode: ein SQL-File, direkter OutputFile-Name
            $baseName = "$OutputFile$timestampSuffix"
            Invoke-DbQueryExport -SqlFile $SqlQueryFile `
                                 -BaseOutputName $baseName `
                                 -ServerName $ServerName `
                                 -DatabaseName $DatabaseName `
                                 -Username $Username `
                                 -Password $Password `
                                 -OutputPath $OutputPath `
                                 -Format $Format `
                                 -KeepCsv:$KeepCsv `
                                 -SheetName $SheetName
        }

        'Folder' {
            # Folder-Mode: alle .sql im Ordner verarbeiten
            $sqlFiles = Get-ChildItem -Path $SqlQueryFolder -Filter *.sql -File
            if (-not $sqlFiles) {
                Write-Warning "Im Ordner '$SqlQueryFolder' wurden keine .sql-Dateien gefunden."
                exit 0
            }

            foreach ($sql in $sqlFiles) {
                $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($sql.Name)
                $baseName       = "${OutputFile}_${nameWithoutExt}$timestampSuffix"

                Invoke-DbQueryExport -SqlFile $sql.FullName `
                                     -BaseOutputName $baseName `
                                     -ServerName $ServerName `
                                     -DatabaseName $DatabaseName `
                                     -Username $Username `
                                     -Password $Password `
                                     -OutputPath $OutputPath `
                                     -Format $Format `
                                     -KeepCsv:$KeepCsv `
                                     -SheetName $SheetName
            }
        }
    }

    Write-Host ""
    Write-Host "🎉 Prozess abgeschlossen."
}
catch {
    Write-Error "Skript abgebrochen wegen eines Fehlers: $($_.Exception.Message)"
    exit 1
}
