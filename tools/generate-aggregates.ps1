<#
.SYNOPSIS
    Строит запрос для снимка агрегатов (13.3, раздел 5): одна строка на метрику, «длинный» формат.

.DESCRIPTION
    Читает каталог колонок БД и печатает SQL, который для каждой пользовательской таблицы
    возвращает набор метрик: cnt, count.<колонка>, nulls.<колонка>, а для числовых колонок —
    sum/min/max, для дат — min/max в epoch-секундах.

    Особенности:
    - базовый тип колонки резолвится через system_type_id: пользовательские alias-типы
      (например MyMoney на базе money) корректно попадают в числовые/датские списки;
    - numeric/decimal шире 32 разрядов пропускаются для sum (сумма переполнит
      decimal(38,6), TRY_SUM на 2016–2019 нет) — такие колонки перечисляются в конце;
    - даты снимаются epoch-секундами (DATEDIFF_BIG от 1970-01-01, см. 13.3 §5 и 13.10 §4.2):
      строковые представления MSSQL и PostgreSQL не сравнимы;
    - time не имеет даты: min/max как строки, только для чтения человеком, не для diff;
    - datetimeoffset конвертируется к datetime2 — смещение теряется и проверяется отдельно;
    - NULL (пустая таблица) выводится как строка 'NULL'.

    Сгенерированный SQL — шаблон: просмотрите его перед запуском и убедитесь, что набор метрик
    соответствует разделу 5 инструкции. Для баз с сотнями таблиц запускайте по схемам (-Schema).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\generate-aggregates.ps1 `
      -Server srv01 -Database MyDb -OutFile <каталог-миграции>\MyDb\snapshot\T0\31-aggregates.sql

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\generate-aggregates.ps1 `
      -Server srv01 -Database MyDb -Schema dbo -OutFile <каталог-миграции>\MyDb\snapshot\T0\31-aggregates-dbo.sql

.NOTES
    Требуется модуль SqlServer (Invoke-Sqlcmd). Учётная запись — только чтение:
    db_datareader + VIEW DEFINITION.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Server,
    [Parameter(Mandatory = $true)][string]$Database,
    [Parameter(Mandatory = $true)][string]$OutFile,
    [string]$Schema,
    # SQL-аутентификация (без параметров — Windows-аутентификация, как раньше).
    # Пароль лучше не указывать в командной строке: он читается из env MSSQL_PASSWORD,
    # если -Password опущен.
    [string]$Username,
    [string]$Password
)

$ErrorActionPreference = 'Stop'

Import-Module SqlServer -ErrorAction Stop

if ($Schema) {
    $schemaFilter = "AND s.name = '$($Schema.Replace("'", "''"))'"
}
else {
    $schemaFilter = ''
}

if ($Username -and -not $Password) {
    $Password = $env:MSSQL_PASSWORD
    if (-not $Password) { throw "Указан -Username, но пароль не задан (-Password или env MSSQL_PASSWORD)." }
}

$metaQuery = @"
SELECT s.name AS sch, t.name AS tab, c.name AS col, st.name AS typ, c.precision AS prec
FROM sys.tables  t
JOIN sys.schemas s ON s.schema_id = t.schema_id
JOIN sys.columns c ON c.object_id = t.object_id
JOIN sys.types  st ON st.user_type_id = c.system_type_id
WHERE t.is_ms_shipped = 0 $schemaFilter
ORDER BY s.name, t.name, c.column_id;
"@

$sqlArgs = @{ ServerInstance = $Server; Database = $Database; Query = $metaQuery }
if ($Username) {
    $secure = ConvertTo-SecureString $Password -AsPlainText -Force
    $sqlArgs.Credential = New-Object System.Management.Automation.PSCredential($Username, $secure)
}
$meta = Invoke-Sqlcmd @sqlArgs
if (-not $meta) { throw "Не найдено ни одной пользовательской таблицы ($Server / $Database)." }

$numericTypes   = @('int', 'bigint', 'smallint', 'tinyint', 'decimal', 'numeric', 'money', 'smallmoney')
$epochDateTypes = @('date', 'datetime', 'datetime2', 'smalldatetime', 'datetimeoffset')

# Экранирование ] внутри идентификатора: [a]b] -> [a]]b]
function Quote-Name([string]$n) { return '[' + $n.Replace(']', ']]') + ']' }

$wideColumns = New-Object System.Collections.Generic.List[string]
$tableBlocks = New-Object System.Collections.Generic.List[string]

foreach ($table in ($meta | Group-Object sch, tab)) {

    $schName = $table.Group[0].sch
    $tabName = $table.Group[0].tab
    $target = (Quote-Name $schName) + '.' + (Quote-Name $tabName)
    $label = "N'" + ($schName + '.' + $tabName).Replace("'", "''") + "'"

    $metrics = New-Object System.Collections.Generic.List[string]
    $metrics.Add("SELECT $label AS object_name, N'cnt' AS metric, CAST(COUNT_BIG(*) AS nvarchar(64)) AS value FROM $target")

    foreach ($column in $table.Group) {

        $colRef  = Quote-Name $column.col
        $colName = $column.col.Replace("'", "''")

        $mCount = "N'count." + $colName + "'"
        $mNulls = "N'nulls." + $colName + "'"
        $mSum   = "N'sum."   + $colName + "'"
        $mMin   = "N'min."   + $colName + "'"
        $mMax   = "N'max."   + $colName + "'"

        $metrics.Add("SELECT $label AS object_name, $mCount AS metric, CAST(COUNT($colRef) AS nvarchar(64)) AS value FROM $target")
        $metrics.Add("SELECT $label, $mNulls, CAST(COUNT_BIG(*) - COUNT($colRef) AS nvarchar(64)) FROM $target")

        if ($numericTypes -contains $column.typ) {
            # numeric/decimal шире 32 разрядов: сумма переполнит decimal(38,6), TRY_SUM на 2016-2019 нет
            if (($column.typ -in @('decimal','numeric')) -and ([int]$column.prec -gt 32)) {
                $wideColumns.Add(("$schName.$tabName.$($column.col) ($($column.typ)($($column.prec)))"))
            }
            else {
                $metrics.Add("SELECT $label, $mSum, COALESCE(CAST(SUM(CAST($colRef AS decimal(38,6))) AS nvarchar(64)), N'NULL') FROM $target")
            }
            $metrics.Add("SELECT $label, $mMin, COALESCE(CAST(MIN($colRef) AS nvarchar(64)), N'NULL') FROM $target")
            $metrics.Add("SELECT $label, $mMax, COALESCE(CAST(MAX($colRef) AS nvarchar(64)), N'NULL') FROM $target")
        }
        elseif ($epochDateTypes -contains $column.typ) {
            # даты — только epoch-секунды: строковые форматы MSSQL и PostgreSQL не сравнимы
            $metrics.Add("SELECT $label, $mMin, COALESCE(CAST(DATEDIFF_BIG(SECOND, CAST('19700101' AS datetime2), MIN($colRef)) AS nvarchar(20)), N'NULL') FROM $target")
            $metrics.Add("SELECT $label, $mMax, COALESCE(CAST(DATEDIFF_BIG(SECOND, CAST('19700101' AS datetime2), MAX($colRef)) AS nvarchar(20)), N'NULL') FROM $target")
        }
        elseif ($column.typ -eq 'time') {
            # time не имеет даты: min/max как строки — только для чтения человеком, не для diff
            $metrics.Add("SELECT $label, $mMin, COALESCE(CAST(MIN($colRef) AS nvarchar(64)), N'NULL') FROM $target")
            $metrics.Add("SELECT $label, $mMax, COALESCE(CAST(MAX($colRef) AS nvarchar(64)), N'NULL') FROM $target")
        }
        # строки, binary, uniqueidentifier, float: только count/nulls — содержимое сверяется хешами
    }

    $tableBlocks.Add(($metrics -join "`r`nUNION ALL`r`n"))
}

$script = ($tableBlocks -join "`r`nUNION ALL`r`n") + "`r`nORDER BY object_name, metric;"

$dir = Split-Path -Parent $OutFile
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }

[IO.File]::WriteAllText($OutFile, $script, (New-Object System.Text.UTF8Encoding($false)))

$tableCount = $tableBlocks.Count
$metricCount = ([regex]::Matches($script, 'UNION ALL')).Count + 1

Write-Host ("Таблиц: {0}; метрик: {1}; файл: {2}" -f $tableCount, $metricCount, $OutFile)

if ($wideColumns.Count -gt 0) {
    Write-Warning ("Без sum (numeric шире 32 разрядов, переполнит decimal(38,6)): {0}" -f ($wideColumns -join '; '))
    Write-Warning "Сверяйте эти колонки хешами выгрузок или полным сравнением (13.3, раздел 7)."
}
