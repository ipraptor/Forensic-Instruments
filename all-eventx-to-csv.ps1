# evtx-to-csv.ps1 — экспорт *.evtx → CSV с полным описанием событий
#==========================================
#  Autoexport all EVTX file in folder to CSV.
#  All saves in ./csv/* folder
#  Parsing process in columns:
#   Line, Timestamp,	Computer, Channel, Event ID, Record ID, Details
#  Anton Palamarchuk (info@expice.ru) 23-10-25v1
#==========================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptDir
$OutDir = Join-Path $ScriptDir 'csv'
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

Get-ChildItem -Path $ScriptDir -Filter *.evtx -File | ForEach-Object {
    $evtx = $_.FullName
    $base = $_.BaseName
    $csv  = Join-Path $OutDir ($base + '.csv')

    Write-Host "Обработка: $($_.Name)"
    $events = Get-WinEvent -Path $evtx -ErrorAction SilentlyContinue
    if (-not $events) { return }

    $i = 0
    $rows = foreach ($e in $events) {
        $i++
        $msg = try { $e.FormatDescription() } catch { $e.Message }
        if ($msg) { $msg = ($msg -replace "`r?`n", " | ") }

        [PSCustomObject][ordered]@{
            'Line'              = $i
            'Timestamp'         = $e.TimeCreated
            'Computer'          = $e.MachineName
            'Channel'           = $e.LogName
            'Event ID'          = $e.Id
            'Record ID'         = $e.RecordId
            'Description'       = $msg
        }
    }

    $rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
}

Write-Host "Готово. CSV сохранены в $OutDir"
