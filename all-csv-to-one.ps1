#==========================================
# The script searches for all *.csv files next to it
# and creates combined.csv in the same folder.
# Anton Palamarchuk (info@expice.ru) 23-10-25v1
#==========================================

# combine-csv.ps1 — объединяет все CSV в один файл с унифицированными колонками
# кладёте файл рядом с CSV и запускаете. Результат: .\combined.csv

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptDir
$OutFile = Join-Path $ScriptDir 'combined.csv'

# =========================================================================
# 💡 ВАЖНО: УКАЖИТЕ ИМЯ КОЛОНКИ С ДАТОЙ/ВРЕМЕНЕМ, КОТОРУЮ НУЖНО ФОРМАТИРОВАТЬ
# =========================================================================
$TimestampHeader = 'Timestamp' # <--- ИЗМЕНИТЕ ЭТО НА РЕАЛЬНОЕ ИМЯ КОЛОНКИ В ВАШИХ CSV!
# Если колонки нет, или она не найдена, скрипт продолжит работу, но форматировать будет нечего.
# Если в разных файлах колонка называется по-разному, вам нужно будет сначала
# стандартизировать имена колонок в функции Strip-Bom, либо добавить их в список.

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

# Функция для форматирования значения даты/времени
function Format-Timestamp {
    param([string]$Value)

    if ([string]::IsNullOrEmpty($Value)) { return $Value }

    # Пытаемся преобразовать строку в объект DateTimeOffset, чтобы учесть часовой пояс
    $dt = $null
    try {
        # Используем [System.DateTimeOffset]::Parse(), так как он более гибок в отношении
        # различных форматов дат, которые могут быть в исходных CSV.
        $dt = [System.DateTimeOffset]::Parse($Value)
        # Форматирование в нужный вид: YYYY-MM-DD hh:mm:ss.fff +zz:zz
        return $dt.ToString("yyyy-MM-dd HH:mm:ss.fff zzz")
    }
    catch {
        # Если преобразование не удалось, возвращаем исходное значение
        Write-Warning "Не удалось преобразовать значение '$Value' в формат даты/времени. Оставлено без изменений."
        return $Value
    }
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
            $value = if ($map.ContainsKey($col)) { $map[$col] } else { $null }

            # Проверяем, является ли текущая колонка той, что нужно форматировать
            if ($col -eq $TimestampHeader -and -not [string]::IsNullOrEmpty($value)) {
                # ПРИМЕНЯЕМ ФОРМАТИРОВАНИЕ
                $h[$col] = Format-Timestamp $value
            } else {
                $h[$col] = $value
            }
        }
        $outRows.Add([PSCustomObject]$h) | Out-Null
    }
}

$outRows | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
Write-Host "Готово: $OutFile"

