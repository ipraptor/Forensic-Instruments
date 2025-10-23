#==========================================
# The script searches for all *.csv files next to it
# and creates combined.csv in the same folder.
# Anton Palamarchuk (info@expice.ru) 23-10-25v1
#==========================================

# combine-csv.ps1 — объединяет все CSV в один файл с унифицированными колонками
# Антон, кладёшь рядом с CSV и запускаешь. Результат: .\combined.csv

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptDir
$OutFile = Join-Path $ScriptDir 'combined.csv'

# Определение разделителя по первой строке
function Get-Delimiter {
    param([string]$Path)
    $first = (Get-Content -Path $Path -TotalCount 1 -ErrorAction Stop)
    $c = ($first.ToCharArray() | Where-Object { $_ -eq ',' }).Count
    $s = ($first.ToCharArray() | Where-Object { $_ -eq ';' }).Count
    if ($s -gt $c) { return ';' } else { return ',' }
}

# Удаление BOM в имени колонки
function Strip-Bom {
    param([string]$s)
    if ($null -eq $s) { return $null }
    return ($s -replace "^\uFEFF","")
}

# Список файлов
$csvFiles = Get-ChildItem -Path $ScriptDir -Filter *.csv -File |
            Where-Object { $_.FullName -ne $OutFile }
if (-not $csvFiles) { Write-Host 'CSV не найдены'; exit }

# Объединённая схема колонок (в порядке обнаружения)
$schema = New-Object System.Collections.Generic.List[string]
# Для устранения дубликатов с разным BOM
$seen = New-Object 'System.Collections.Generic.HashSet[string]'

# Сначала пройдёмся по заголовкам всех файлов и соберём схему
foreach ($f in $csvFiles) {
    $delim = Get-Delimiter $f.FullName
    $rows = Import-Csv -Path $f.FullName -Delimiter $delim
    if (-not $rows) { continue }
    $hdr = $rows[0].PSObject.Properties.Name | ForEach-Object { Strip-Bom $_ }
    foreach ($h in $hdr) {
        if (-not $seen.Contains($h)) {
            [void]$seen.Add($h)
            [void]$schema.Add($h)
        }
    }
}

if ($schema.Count -eq 0) { Write-Host 'Пустые CSV'; exit }

# Теперь читаем и выравниваем строки под полную схему
$outRows = New-Object System.Collections.Generic.List[object]

foreach ($f in $csvFiles) {
    Write-Host "Добавление: $($f.Name)"
    $delim = Get-Delimiter $f.FullName
    $rows = Import-Csv -Path $f.FullName -Delimiter $delim
    if (-not $rows) { continue }

    foreach ($r in $rows) {
        # Карта: нормализованное имя -> значение
        $map = @{}
        foreach ($p in $r.PSObject.Properties) {
            $norm = Strip-Bom $p.Name
            $map[$norm] = $p.Value
        }

        # Формируем выровненную строку в порядке $schema
        $h = [ordered]@{}
        foreach ($col in $schema) {
            $h[$col] = if ($map.ContainsKey($col)) { $map[$col] } else { $null }
        }
        $outRows.Add([PSCustomObject]$h) | Out-Null
    }
}

$outRows | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
Write-Host "Готово: $OutFile"
