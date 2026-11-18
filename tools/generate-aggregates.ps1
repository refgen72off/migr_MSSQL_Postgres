<#
.SYNOPSIS
    Строит запрос для снимка агрегатов (13.3, раздел 5): одна строка на метрику, «длинный» формат.

.DESCRIPTION
    Читает каталог колонок БД и печатает SQL, который для каждой пользовательской таблицы
    возвращает набор метрик: cnt, count.<колонка>, nulls.<колонка>, а для числовых колонок —
    sum/min/max, для дат — min/max.

    Сгенерированный SQL — шаблон: просмотрите его перед запуском и убедитесь, что набор метрик
    соответствует разделу 5 инструкции. Для баз с сотнями таблиц запускайте по схемам (-Schema).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\generate-aggregates.ps1 `
      -Server srv01 -Database MyDb -OutFile C:\work\migr\MyDb\snapshot\T0\31-aggregates.sql

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\generate-aggregates.ps1 `
      -Server srv01 -Database MyDb -Schema dbo -OutFile C:\work\migr\MyDb\snapshot\T0\31-aggregates-dbo.sql

.NOTES
    Требуется модуль SqlServer (Invoke-Sqlcmd). Учётная запись — только чтение:
    db_datareader + VIEW DEFINITION.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Server,
    [Parameter(Mandatory = $true)][string]$Database,
    [Parameter(Mandatory = $true)][string]$OutFile,
    [string]$Schema
)

$ErrorActionPreference = 'Stop'

Import-Module SqlServer -ErrorAction Stop

if ($Schema) {
    $schemaFilter = "AND s.name = '$($Schema.Replace("'", "''"))'"
}
else {
    $schemaFilter = ''
}

$metaQuery = @"
SELECT s.name AS sch, t.name AS tab, c.name AS col, ty.name AS typ
FROM sys.tables  t
JOIN sys.schemas s ON s.schema_id = t.schema_id
JOIN sys.columns c ON c.object_id = t.object_id
JOIN sys.types  ty ON ty.user_type_id = c.user_type_id
WHERE t.is_ms_shipped = 0 $schemaFilter
ORDER BY s.name, t.name, c.column_id;
"@

$meta = Invoke-Sqlcmd -ServerInstance $Server -Database $Database -Query $metaQuery
if (-not $meta) { throw "Не найдено ни одной пользовательской таблицы ($Server / $Database)." }

$numericTypes = @('int', 'bigint', 'smallint', 'tinyint', 'decimal', 'numeric', 'money', 'smallmoney')
$dateTypes = @('date', 'datetime', 'datetime2', 'smalldatetime', 'datetimeoffset', 'time')

$tableBlocks = New-Object System.Collections.Generic.List[string]

foreach ($table in ($meta | Group-Object sch, tab)) {

    $schName = $table.Group[0].sch
    $tabName = $table.Group[0].tab
    $target = '[' + $schName + '].[' + $tabName + ']'
    $label = "N'" + ($schName + '.' + $tabName).Replace("'", "''") + "'"

    $metrics = New-Object System.Collections.Generic.List[string]
    $metrics.Add("SELECT $label AS object_name, N'cnt' AS metric, CAST(COUNT_BIG(*) AS nvarchar(64)) AS value FROM $target")

    foreach ($column in $table.Group) {

        $colRef = '[' + $column.col + ']'
        $colName = $column.col.Replace("'", "''")

        $mCount = "N'count." + $colName + "'"
        $mNulls = "N'nulls." + $colName + "'"
        $mSum = "N'sum." + $colName + "'"
        $mMin = "N'min." + $colName + "'"
        $mMax = "N'max." + $colName + "'"

        $metrics.Add("SELECT $label AS object_name, $mCount AS metric, CAST(COUNT($colRef) AS nvarchar(64)) AS value FROM $target")
        $metrics.Add("SELECT $label, $mNulls, CAST(COUNT_BIG(*) - COUNT($colRef) AS nvarchar(64)) FROM $target")

        if ($numericTypes -contains $column.typ) {
            $metrics.Add("SELECT $label, $mSum, CAST(SUM(CAST($colRef AS decimal(38,6))) AS nvarchar(64)) FROM $target")
            $metrics.Add("SELECT $label, $mMin, CAST(MIN($colRef) AS nvarchar(64)) FROM $target")
            $metrics.Add("SELECT $label, $mMax, CAST(MAX($colRef) AS nvarchar(64)) FROM $target")
        }
        elseif ($dateTypes -contains $column.typ) {
            $metrics.Add("SELECT $label, $mMin, CAST(MIN($colRef) AS nvarchar(64)) FROM $target")
            $metrics.Add("SELECT $label, $mMax, CAST(MAX($colRef) AS nvarchar(64)) FROM $target")
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
