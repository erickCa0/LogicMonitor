<# 
.SYNOPSIS
Exports LogicMonitor LM Uptime device configuration and monthly check counts.

.DESCRIPTION
Uses the official Logic.Monitor PowerShell module instead of manually signing
REST requests. The script uses Get-LMUptimeDevice as the source of truth for LM
Uptime resources, then exports backing device DataSources, instances, monthly
estimated counts, and optional actual datapoint counts via Get-LMDeviceInstanceData.
The default output format is CSV. Use -OutputFormat Json or Both if you also
want the richer nested JSON payloads.

Install the module first if needed:

  Install-Module Logic.Monitor -Scope CurrentUser

.EXAMPLE
.\Export-LMUptimeCounts.ps1 `
  -AccountName "your-portal" `
  -AccessId $env:LM_ACCESS_ID `
  -AccessKey $env:LM_ACCESS_KEY `
  -StartMonth 2026-01 `
  -EndMonth 2026-05 `
  -OutputDir .\exports\lm-uptime-2026

.EXAMPLE
Connect-LMAccount -AccountName "your-portal" -BearerToken $env:LM_BEARER_TOKEN
.\Export-LMUptimeCounts.ps1 -SkipConnect -Mode Both -StartMonth 2026-05 -EndMonth 2026-05
#>

[CmdletBinding()]
param(
    [string]$AccountName = $env:LM_COMPANY,
    [string]$AccessId = $env:LM_ACCESS_ID,
    [string]$AccessKey = $env:LM_ACCESS_KEY,
    [string]$BearerToken = $env:LM_BEARER_TOKEN,
    [switch]$UseCachedCredential,
    [string]$CachedAccountName,
    [switch]$SkipConnect,
    [switch]$DisconnectWhenDone,

    [ValidateSet('Estimated', 'Actual', 'Both')]
    [string]$Mode = 'Estimated',

    [string]$StartMonth,
    [string]$EndMonth,
    [int]$MonthsBack = 1,

    [string]$OutputDir,

    [ValidateSet('Csv', 'Json', 'Both')]
    [string]$OutputFormat = 'Csv',

    [ValidateSet('uptimewebcheck', 'uptimepingcheck')]
    [string]$Type,

    [Nullable[bool]]$IsInternal,
    [object]$Filter,
    [int]$BatchSize = 1000,

    [string[]]$CountDatasourceNames = @(
        'Web_Check_Individual',
        'Web_Check_Overall',
        'Ping_Check_Individual',
        'Ping_Check_Overall'
    ),

    [string[]]$PollingIntervalProperties = @(
        'pollingInterval',
        'polling_interval',
        'polling.interval',
        'uptime.pollingInterval',
        'uptime.polling.interval',
        'lm.uptime.pollingInterval',
        'lm.uptime.polling.interval',
        'checkInterval',
        'check.interval',
        'checkFrequency',
        'check.frequency',
        'frequency',
        'interval'
    ),

    [int]$DefaultIntervalMinutes = 5,

    [ValidateSet('Auto', 'Minutes', 'Seconds', 'Hours')]
    [string]$IntervalUnit = 'Auto',

    [switch]$IncludeDisabledInstances,
    [switch]$SkipDatasourceInventory,
    [switch]$SkipGroups,

    [ValidateSet('first', 'last', 'min', 'max', 'sum', 'average', 'none')]
    [string]$ActualAggregationType = 'none',
    [double]$ActualPeriod = 1,
    [int]$ActualWindowHours = 8
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

function Import-LogicMonitorModule {
    try {
        Import-Module Logic.Monitor -ErrorAction Stop
    }
    catch {
        throw "Unable to import the Logic.Monitor module. Install it with: Install-Module Logic.Monitor -Scope CurrentUser. Details: $($_.Exception.Message)"
    }
}

function Connect-LogicMonitorAccount {
    if ($SkipConnect) {
        return
    }

    $connectParams = @{
        DisableConsoleLogging = $true
        SkipVersionCheck      = $true
    }

    if ($UseCachedCredential -or $CachedAccountName) {
        $connectParams.UseCachedCredential = $true
        if ($CachedAccountName) {
            $connectParams.CachedAccountName = $CachedAccountName
        }
        Connect-LMAccount @connectParams
        return
    }

    if ($BearerToken) {
        if (-not $AccountName) {
            throw 'AccountName or LM_COMPANY is required when using BearerToken.'
        }
        $connectParams.AccountName = $AccountName
        $connectParams.BearerToken = $BearerToken
        Connect-LMAccount @connectParams
        return
    }

    if ($AccessId -and $AccessKey) {
        if (-not $AccountName) {
            throw 'AccountName or LM_COMPANY is required when using AccessId/AccessKey.'
        }
        $connectParams.AccountName = $AccountName
        $connectParams.AccessId = $AccessId
        $connectParams.AccessKey = $AccessKey
        Connect-LMAccount @connectParams
        return
    }

    throw 'Provide AccessId/AccessKey, BearerToken, cached credentials, or run Connect-LMAccount first and pass -SkipConnect.'
}

function Parse-Month {
    param([Parameter(Mandatory)][string]$Value)

    try {
        $culture = [Globalization.CultureInfo]::InvariantCulture
        $parsed = [DateTime]::ParseExact($Value, 'yyyy-MM', $culture)
        return [DateTime]::SpecifyKind($parsed, [DateTimeKind]::Utc)
    }
    catch {
        throw "Invalid month '$Value'. Use YYYY-MM, for example 2026-05."
    }
}

function Get-MonthWindows {
    if (($StartMonth -and -not $EndMonth) -or ($EndMonth -and -not $StartMonth)) {
        throw 'Provide both -StartMonth and -EndMonth, or neither.'
    }

    if ($MonthsBack -lt 1) {
        throw '-MonthsBack must be at least 1.'
    }

    if ($StartMonth) {
        $start = Parse-Month -Value $StartMonth
        $end = Parse-Month -Value $EndMonth
    }
    else {
        $now = [DateTime]::UtcNow
        $currentMonth = [DateTime]::SpecifyKind([DateTime]::new($now.Year, $now.Month, 1), [DateTimeKind]::Utc)
        $start = $currentMonth.AddMonths(-1 * $MonthsBack)
        $end = $currentMonth.AddMonths(-1)
    }

    if ($end -lt $start) {
        throw '-EndMonth must be the same as or after -StartMonth.'
    }

    $windows = @()
    $cursor = $start
    while ($cursor -le $end) {
        $windows += [pscustomobject]@{
            Month = $cursor.ToString('yyyy-MM')
            Start = $cursor
            End = $cursor.AddMonths(1)
        }
        $cursor = $cursor.AddMonths(1)
    }

    return $windows
}

function Get-OutputDirectory {
    if ($OutputDir) {
        $resolved = Resolve-Path -LiteralPath $OutputDir -ErrorAction SilentlyContinue
        if ($resolved) {
            return $resolved.Path
        }
        return $OutputDir
    }

    return "logicmonitor_uptime_powershell_export_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
}

function Export-JsonFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$InputObject
    )

    $InputObject | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function ConvertTo-CsvSafeValue {
    param($Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [string] -or
        $Value -is [int] -or
        $Value -is [long] -or
        $Value -is [double] -or
        $Value -is [decimal] -or
        $Value -is [bool] -or
        $Value -is [DateTime]) {
        return $Value
    }

    return ($Value | ConvertTo-Json -Depth 50 -Compress)
}

function ConvertTo-FlatCsvRows {
    param([array]$InputObjects)

    $rows = @()
    foreach ($item in @($InputObjects)) {
        $row = [ordered]@{}
        if ($null -eq $item) {
            continue
        }

        foreach ($property in $item.PSObject.Properties) {
            if ($property.Name -in @('properties', 'customProperties', 'systemProperties', 'inheritedProperties', 'autoProperties')) {
                if ($property.Value -is [System.Collections.IDictionary]) {
                    foreach ($key in $property.Value.Keys) {
                        $row["$($property.Name).$key"] = ConvertTo-CsvSafeValue -Value $property.Value[$key]
                    }
                }
                elseif ($property.Value) {
                    foreach ($nestedProperty in @($property.Value)) {
                        $name = Get-ObjectProperty -InputObject $nestedProperty -Name 'name'
                        if ($name) {
                            $row["$($property.Name).$name"] = ConvertTo-CsvSafeValue -Value (Get-ObjectProperty -InputObject $nestedProperty -Name 'value')
                        }
                    }
                }
                else {
                    $row[$property.Name] = $null
                }
                continue
            }

            $row[$property.Name] = ConvertTo-CsvSafeValue -Value $property.Value
        }
        $rows += [pscustomobject]$row
    }

    return $rows
}

function Export-FlatCsvFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [array]$Rows = @()
    )

    $flatRows = @(ConvertTo-FlatCsvRows -InputObjects $Rows)
    Write-CsvSafe -Path $Path -Rows $flatRows
}

function Export-MetadataCsv {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Metadata
    )

    $rows = @()
    foreach ($property in $Metadata.PSObject.Properties) {
        if ($property.Name -eq 'warnings') {
            continue
        }
        $rows += [pscustomobject]@{
            key = $property.Name
            value = ConvertTo-CsvSafeValue -Value $property.Value
        }
    }

    Write-CsvSafe -Path $Path -Rows $rows
}

function Export-WarningsCsv {
    param(
        [Parameter(Mandatory)][string]$Path,
        [array]$Warnings = @()
    )

    $rows = @()
    for ($index = 0; $index -lt $Warnings.Count; $index++) {
        $rows += [pscustomobject]@{
            warning_number = $index + 1
            warning = [string]$Warnings[$index]
        }
    }

    Write-CsvSafe -Path $Path -Rows $rows
}

function Get-ObjectProperty {
    param(
        $InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary] -and $InputObject.Contains($Name)) {
        return $InputObject[$Name]
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) {
        return $property.Value
    }

    return $null
}

function Get-FlatProperties {
    param([Parameter(Mandatory)]$InputObject)

    $flat = @{}

    foreach ($property in $InputObject.PSObject.Properties) {
        if ($property.Value -is [string] -or
            $property.Value -is [int] -or
            $property.Value -is [long] -or
            $property.Value -is [double] -or
            $property.Value -is [bool] -or
            $null -eq $property.Value) {
            $flat[$property.Name] = $property.Value
        }
    }

    foreach ($fieldName in @('properties', 'customProperties', 'systemProperties', 'inheritedProperties', 'autoProperties')) {
        $properties = Get-ObjectProperty -InputObject $InputObject -Name $fieldName
        if ($properties -is [System.Collections.IDictionary]) {
            foreach ($key in $properties.Keys) {
                $flat[[string]$key] = $properties[$key]
            }
        }
        elseif ($properties) {
            foreach ($property in @($properties)) {
                $name = Get-ObjectProperty -InputObject $property -Name 'name'
                if ($name) {
                    $flat[[string]$name] = Get-ObjectProperty -InputObject $property -Name 'value'
                }
            }
        }
    }

    return $flat
}

function ConvertTo-IntervalMinutes {
    param(
        $RawValue,
        [Parameter(Mandatory)][string]$PropertyName
    )

    if ($null -eq $RawValue) {
        return $null
    }

    $numeric = $null
    $suffix = ''

    if ($RawValue -is [int] -or $RawValue -is [long] -or $RawValue -is [double]) {
        $numeric = [double]$RawValue
    }
    else {
        $text = ([string]$RawValue).Trim().ToLowerInvariant()
        if (-not $text) {
            return $null
        }
        $match = [regex]::Match($text, '(\d+(?:\.\d+)?)')
        if (-not $match.Success) {
            return $null
        }
        $numeric = [double]::Parse($match.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture)
        $suffix = $text.Substring($match.Index + $match.Length).Trim()
    }

    if ($numeric -le 0) {
        return $null
    }

    $unit = $IntervalUnit
    if ($unit -eq 'Auto') {
        $lowerName = $PropertyName.ToLowerInvariant()
        if ($lowerName -like '*second*' -or $suffix -like 's*') {
            $unit = 'Seconds'
        }
        elseif ($suffix -like 'h*' -or $suffix -like '*hour*') {
            $unit = 'Hours'
        }
        else {
            $unit = 'Minutes'
        }
    }

    switch ($unit) {
        'Seconds' { return [Math]::Max(1, [int][Math]::Ceiling($numeric / 60)) }
        'Hours' { return [Math]::Max(1, [int][Math]::Ceiling($numeric * 60)) }
        default { return [Math]::Max(1, [int][Math]::Ceiling($numeric)) }
    }
}

function Get-DevicePollingInterval {
    param(
        [Parameter(Mandatory)]$Device,
        [Parameter(Mandatory)][System.Collections.Generic.List[string]]$Warnings
    )

    $flat = Get-FlatProperties -InputObject $Device

    foreach ($name in $PollingIntervalProperties) {
        if ($flat.ContainsKey($name)) {
            $parsed = ConvertTo-IntervalMinutes -RawValue $flat[$name] -PropertyName $name
            if ($null -ne $parsed) {
                return [pscustomobject]@{
                    Minutes = $parsed
                    Source = $name
                }
            }
        }
    }

    $lowerLookup = @{}
    foreach ($key in $flat.Keys) {
        $lowerLookup[$key.ToString().ToLowerInvariant()] = $key
    }

    foreach ($name in $PollingIntervalProperties) {
        $lowerName = $name.ToLowerInvariant()
        if ($lowerLookup.ContainsKey($lowerName)) {
            $originalName = $lowerLookup[$lowerName]
            $parsed = ConvertTo-IntervalMinutes -RawValue $flat[$originalName] -PropertyName $originalName
            if ($null -ne $parsed) {
                return [pscustomobject]@{
                    Minutes = $parsed
                    Source = $originalName
                }
            }
        }
    }

    $Warnings.Add(("Device {0} '{1}' did not expose a polling interval; used default {2} minutes." -f (Get-ObjectProperty $Device 'id'), (Get-DeviceDisplayName $Device), $DefaultIntervalMinutes))
    return [pscustomobject]@{
        Minutes = $DefaultIntervalMinutes
        Source = 'default'
    }
}

function Get-DeviceDisplayName {
    param([Parameter(Mandatory)]$Device)

    foreach ($name in @('displayName', 'name', 'hostName', 'id')) {
        $value = Get-ObjectProperty -InputObject $Device -Name $name
        if ($null -ne $value -and [string]$value -ne '') {
            return [string]$value
        }
    }

    return ''
}

function Get-DeviceCheckType {
    param([Parameter(Mandatory)]$Device)

    $deviceType = [string](Get-ObjectProperty -InputObject $Device -Name 'deviceType')
    switch ($deviceType) {
        '18' { return 'web' }
        '19' { return 'ping' }
        default {
            $type = [string](Get-ObjectProperty -InputObject $Device -Name 'type')
            if ($type -like '*ping*') { return 'ping' }
            return 'web'
        }
    }
}

function Get-PreferredDatasource {
    param(
        [Parameter(Mandatory)][array]$Datasources,
        [Parameter(Mandatory)][string]$CheckType
    )

    $preferredNames = if ($CheckType -eq 'ping') {
        @('Ping_Check_Individual', 'Ping_Check_Overall')
    }
    else {
        @('Web_Check_Individual', 'Web_Check_Overall')
    }

    foreach ($name in $preferredNames) {
        $match = @($Datasources | Where-Object { (Get-ObjectProperty -InputObject $_ -Name 'dataSourceName') -eq $name })
        if ($match.Count -gt 0) {
            return $match[0]
        }
    }

    $fallback = @($Datasources | Where-Object { $CountDatasourceNames -contains (Get-ObjectProperty -InputObject $_ -Name 'dataSourceName') })
    if ($fallback.Count -gt 0) {
        return $fallback[0]
    }

    return $null
}

function Test-InstanceEnabled {
    param($Instance)

    if ($IncludeDisabledInstances) {
        return $true
    }

    foreach ($name in @('stopMonitoring', 'disableMonitoring', 'disabled', 'isDisabled')) {
        $value = Get-ObjectProperty -InputObject $Instance -Name $name
        if ([bool]$value) {
            return $false
        }
    }

    return $true
}

function Get-InstanceDisplayName {
    param($Instance)

    foreach ($name in @('displayName', 'name', 'wildValue', 'id')) {
        $value = Get-ObjectProperty -InputObject $Instance -Name $name
        if ($null -ne $value -and [string]$value -ne '') {
            return [string]$value
        }
    }

    return ''
}

function Get-EstimatedRuns {
    param(
        [DateTime]$Start,
        [DateTime]$End,
        [int]$IntervalMinutes
    )

    if ($IntervalMinutes -lt 1 -or $End -le $Start) {
        return 0
    }

    $seconds = ($End - $Start).TotalSeconds
    return [int][Math]::Ceiling($seconds / ($IntervalMinutes * 60))
}

function Get-MonthChunks {
    param(
        [DateTime]$Start,
        [DateTime]$End
    )

    $chunks = @()
    $cursor = $Start
    while ($cursor -lt $End) {
        $chunkEnd = $cursor.AddHours($ActualWindowHours)
        if ($chunkEnd -gt $End) {
            $chunkEnd = $End
        }
        $chunks += [pscustomobject]@{
            Start = $cursor
            End = $chunkEnd
        }
        $cursor = $chunkEnd
    }

    return $chunks
}

function Measure-ItemDatapoints {
    param($Item)

    $maxCount = 0
    foreach ($propertyName in @('values', 'data', 'datapoints', 'timestamps', 'time', 'times')) {
        $value = Get-ObjectProperty -InputObject $Item -Name $propertyName
        if ($value -is [array]) {
            $maxCount = [Math]::Max($maxCount, @($value).Count)
        }
        elseif ($value -is [System.Collections.IDictionary]) {
            foreach ($nestedValue in $value.Values) {
                if ($nestedValue -is [array]) {
                    $maxCount = [Math]::Max($maxCount, @($nestedValue).Count)
                }
            }
        }
    }

    if ($maxCount -gt 0) {
        return $maxCount
    }

    return 1
}

function Measure-ReturnedDatapoints {
    param($Data)

    if ($null -eq $Data) {
        return 0
    }

    $items = @($Data)
    if ($items.Count -eq 0) {
        return 0
    }

    $total = 0
    foreach ($item in $items) {
        $total += Measure-ItemDatapoints -Item $item
    }

    return $total
}

function Write-CsvSafe {
    param(
        [Parameter(Mandatory)][string]$Path,
        [array]$Rows = @()
    )

    if ($Rows.Count -eq 0) {
        '' | Set-Content -LiteralPath $Path -Encoding UTF8
        return
    }

    $Rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

if ($BatchSize -lt 1 -or $BatchSize -gt 1000) {
    throw '-BatchSize must be between 1 and 1000.'
}
if ($DefaultIntervalMinutes -lt 1) {
    throw '-DefaultIntervalMinutes must be at least 1.'
}
if ($ActualWindowHours -lt 1 -or $ActualWindowHours -gt 24) {
    throw '-ActualWindowHours must be between 1 and 24 because Get-LMDeviceInstanceData supports a maximum 24-hour window.'
}

Import-LogicMonitorModule
Connect-LogicMonitorAccount

$warnings = [System.Collections.Generic.List[string]]::new()
$monthWindows = @(Get-MonthWindows)
$resolvedOutputDir = Get-OutputDirectory
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

try {
    $uptimeParams = @{
        BatchSize = $BatchSize
    }
    if ($Type) {
        $uptimeParams.Type = $Type
    }
    if ($PSBoundParameters.ContainsKey('IsInternal')) {
        $uptimeParams.IsInternal = $IsInternal
    }
    if ($null -ne $Filter) {
        $uptimeParams.Filter = $Filter
    }

    Write-Host 'Fetching LM Uptime devices...'
    $uptimeDevices = @(Get-LMUptimeDevice @uptimeParams)
    Write-Host ("Found {0} LM Uptime devices." -f $uptimeDevices.Count)

    $deviceGroups = @()
    if (-not $SkipGroups) {
        Write-Host 'Fetching device groups...'
        $deviceGroups = @(Get-LMDeviceGroup -BatchSize $BatchSize)
        Write-Host ("Found {0} device groups." -f $deviceGroups.Count)
    }

    $allDatasources = @()
    $allInstances = @()
    $logicalChecks = @()

    if ($SkipDatasourceInventory) {
        foreach ($device in $uptimeDevices) {
            $logicalChecks += [pscustomobject]@{
                Device = $device
                CheckType = Get-DeviceCheckType -Device $device
                CountDatasource = $null
                CountInstances = @()
                Datasources = @()
            }
        }
    }
    else {
        $totalDevices = $uptimeDevices.Count
        $index = 0
        foreach ($device in $uptimeDevices) {
            $index++
            $deviceId = Get-ObjectProperty -InputObject $device -Name 'id'
            if ($index -eq 1 -or $index % 25 -eq 0 -or $index -eq $totalDevices) {
                Write-Host ("Fetching DataSources and instances {0}/{1}..." -f $index, $totalDevices)
            }

            $datasources = @(Get-LMDeviceDatasourceList -Id $deviceId -BatchSize $BatchSize)
            $uptimeDatasources = @($datasources | Where-Object { $CountDatasourceNames -contains (Get-ObjectProperty -InputObject $_ -Name 'dataSourceName') })

            foreach ($datasource in $uptimeDatasources) {
                $allDatasources += [pscustomobject]@{
                    deviceId = $deviceId
                    deviceDisplayName = Get-DeviceDisplayName -Device $device
                    deviceDatasourceId = Get-ObjectProperty -InputObject $datasource -Name 'id'
                    dataSourceId = Get-ObjectProperty -InputObject $datasource -Name 'dataSourceId'
                    dataSourceName = Get-ObjectProperty -InputObject $datasource -Name 'dataSourceName'
                    instanceNumber = Get-ObjectProperty -InputObject $datasource -Name 'instanceNumber'
                    raw = $datasource
                }
            }

            $instancesByDatasourceId = @{}
            foreach ($datasource in $uptimeDatasources) {
                $dsId = Get-ObjectProperty -InputObject $datasource -Name 'id'
                $dsName = Get-ObjectProperty -InputObject $datasource -Name 'dataSourceName'
                $dataSourceId = Get-ObjectProperty -InputObject $datasource -Name 'dataSourceId'

                $instanceParams = @{
                    Id        = $deviceId
                    BatchSize = $BatchSize
                }
                if ($null -ne $dataSourceId -and [string]$dataSourceId -ne '') {
                    $instanceParams.DatasourceId = [int]$dataSourceId
                }
                else {
                    $instanceParams.DatasourceName = $dsName
                }

                $instances = @(Get-LMDeviceDatasourceInstance @instanceParams)
                $enrichedInstances = @()
                foreach ($instance in $instances) {
                    $enabledForCount = Test-InstanceEnabled -Instance $instance
                    $enriched = [pscustomobject]@{
                        deviceId = $deviceId
                        deviceDisplayName = Get-DeviceDisplayName -Device $device
                        deviceDatasourceId = $dsId
                        dataSourceId = $dataSourceId
                        dataSourceName = $dsName
                        instanceId = Get-ObjectProperty -InputObject $instance -Name 'id'
                        instanceName = Get-InstanceDisplayName -Instance $instance
                        wildValue = Get-ObjectProperty -InputObject $instance -Name 'wildValue'
                        countEnabled = $enabledForCount
                        raw = $instance
                    }
                    $enrichedInstances += $enriched
                    $allInstances += $enriched
                }
                $instancesByDatasourceId[[string]$dsId] = $enrichedInstances
            }

            $checkType = Get-DeviceCheckType -Device $device
            $countDatasource = Get-PreferredDatasource -Datasources $uptimeDatasources -CheckType $checkType
            if ($null -eq $countDatasource) {
                $warnings.Add(("Device {0} '{1}' did not have a matching count DataSource. Estimated executions are 0." -f $deviceId, (Get-DeviceDisplayName -Device $device)))
                $logicalChecks += [pscustomobject]@{
                    Device = $device
                    CheckType = $checkType
                    CountDatasource = $null
                    CountInstances = @()
                    Datasources = $uptimeDatasources
                }
                continue
            }

            $countDatasourceId = [string](Get-ObjectProperty -InputObject $countDatasource -Name 'id')
            $countInstances = @($instancesByDatasourceId[$countDatasourceId] | Where-Object { $_.countEnabled })
            $logicalChecks += [pscustomobject]@{
                Device = $device
                CheckType = $checkType
                CountDatasource = $countDatasource
                CountInstances = $countInstances
                Datasources = $uptimeDatasources
            }
        }
    }

    $monthlyRows = @()
    foreach ($logicalCheck in $logicalChecks) {
        $device = $logicalCheck.Device
        $deviceId = Get-ObjectProperty -InputObject $device -Name 'id'
        $interval = Get-DevicePollingInterval -Device $device -Warnings $warnings
        $countDatasource = $logicalCheck.CountDatasource
        $countDatasourceId = if ($null -ne $countDatasource) { Get-ObjectProperty -InputObject $countDatasource -Name 'id' } else { $null }
        $countDatasourceName = if ($null -ne $countDatasource) { Get-ObjectProperty -InputObject $countDatasource -Name 'dataSourceName' } else { $null }
        $countDatasourceDataSourceId = if ($null -ne $countDatasource) { Get-ObjectProperty -InputObject $countDatasource -Name 'dataSourceId' } else { $null }
        $countInstances = @($logicalCheck.CountInstances)
        $countInstanceNames = (($countInstances | ForEach-Object { $_.instanceName }) -join '; ')

        if ($countInstances.Count -eq 0 -and -not $SkipDatasourceInventory) {
            $warnings.Add(("Device {0} '{1}' has no enabled count instances for {2}; estimated executions are 0." -f $deviceId, (Get-DeviceDisplayName -Device $device), $countDatasourceName))
        }

        foreach ($window in $monthWindows) {
            $runsPerInstance = Get-EstimatedRuns -Start $window.Start -End $window.End -IntervalMinutes $interval.Minutes
            $estimatedExecutions = $runsPerInstance * $countInstances.Count
            $monthlyRows += [pscustomobject]@{
                month = $window.Month
                device_id = $deviceId
                device_display_name = Get-DeviceDisplayName -Device $device
                device_name = Get-ObjectProperty -InputObject $device -Name 'name'
                host_name = Get-ObjectProperty -InputObject $device -Name 'hostName'
                device_type = Get-ObjectProperty -InputObject $device -Name 'deviceType'
                is_internal = Get-ObjectProperty -InputObject $device -Name 'isInternal'
                check_type = $logicalCheck.CheckType
                count_datasource_name = $countDatasourceName
                device_datasource_id = $countDatasourceId
                datasource_id = $countDatasourceDataSourceId
                polling_interval_minutes = $interval.Minutes
                polling_interval_source = $interval.Source
                count_instance_count = $countInstances.Count
                count_instance_names = $countInstanceNames
                estimated_runs_per_instance = $runsPerInstance
                estimated_check_executions = $estimatedExecutions
                actual_check_executions = ''
                actual_count_status = 'not_requested'
            }
        }
    }

    if ($Mode -in @('Actual', 'Both')) {
        if ($SkipDatasourceInventory) {
            throw '-Mode Actual or Both requires DataSource inventory. Remove -SkipDatasourceInventory.'
        }

        $rowsByKey = @{}
        foreach ($row in $monthlyRows) {
            $key = '{0}|{1}|{2}|{3}' -f $row.device_id, $row.device_datasource_id, $row.check_type, $row.month
            $rowsByKey[$key] = $row
        }

        $totalWork = $logicalChecks.Count * $monthWindows.Count
        $workIndex = 0
        foreach ($logicalCheck in $logicalChecks) {
            $device = $logicalCheck.Device
            $deviceId = Get-ObjectProperty -InputObject $device -Name 'id'
            $countDatasource = $logicalCheck.CountDatasource
            if ($null -eq $countDatasource) {
                continue
            }
            $countDatasourceId = Get-ObjectProperty -InputObject $countDatasource -Name 'id'
            $instanceIds = @($logicalCheck.CountInstances | ForEach-Object { [string]$_.instanceId } | Where-Object { $_ })

            foreach ($window in $monthWindows) {
                $workIndex++
                if ($workIndex -eq 1 -or $workIndex % 20 -eq 0 -or $workIndex -eq $totalWork) {
                    Write-Host ("Counting actual uptime data {0}/{1}..." -f $workIndex, $totalWork)
                }

                $rowKey = '{0}|{1}|{2}|{3}' -f $deviceId, $countDatasourceId, $logicalCheck.CheckType, $window.Month
                if (-not $rowsByKey.ContainsKey($rowKey)) {
                    continue
                }

                $row = $rowsByKey[$rowKey]
                if ($instanceIds.Count -eq 0) {
                    $row.actual_check_executions = 0
                    $row.actual_count_status = 'no_enabled_instances'
                    continue
                }

                $actualTotal = 0
                $errors = @()
                foreach ($chunk in Get-MonthChunks -Start $window.Start -End $window.End) {
                    try {
                        $data = Get-LMDeviceInstanceData `
                            -Ids $instanceIds `
                            -StartDate $chunk.Start `
                            -EndDate $chunk.End `
                            -AggregationType $ActualAggregationType `
                            -Period $ActualPeriod
                        $actualTotal += Measure-ReturnedDatapoints -Data $data
                    }
                    catch {
                        $errors += $_.Exception.Message
                        break
                    }
                }

                $row.actual_check_executions = $actualTotal
                if ($errors.Count -gt 0) {
                    $row.actual_count_status = 'partial_error'
                    $warnings.Add(("Device {0} '{1}' actual-count errors: {2}" -f $deviceId, (Get-DeviceDisplayName -Device $device), ($errors[0..([Math]::Min(4, $errors.Count - 1))] -join ' | ')))
                }
                else {
                    $row.actual_count_status = 'ok'
                }
            }
        }
    }

    $totalsByKey = @{}
    foreach ($row in $monthlyRows) {
        $key = '{0}|{1}' -f $row.month, $row.check_type
        if (-not $totalsByKey.ContainsKey($key)) {
            $totalsByKey[$key] = [pscustomobject]@{
                month = $row.month
                check_type = $row.check_type
                resource_check_count = 0
                count_instance_count = 0
                estimated_check_executions = 0
                actual_check_executions = ''
            }
        }

        $bucket = $totalsByKey[$key]
        $bucket.resource_check_count++
        $bucket.count_instance_count += [int]$row.count_instance_count
        $bucket.estimated_check_executions += [int]$row.estimated_check_executions
        if ($row.actual_check_executions -ne '') {
            if ($bucket.actual_check_executions -eq '') {
                $bucket.actual_check_executions = 0
            }
            $bucket.actual_check_executions += [int]$row.actual_check_executions
        }
    }

    $combinedByMonth = @{}
    foreach ($row in $monthlyRows) {
        if (-not $combinedByMonth.ContainsKey($row.month)) {
            $combinedByMonth[$row.month] = [pscustomobject]@{
                month = $row.month
                check_type = 'combined'
                resource_check_count = 0
                count_instance_count = 0
                estimated_check_executions = 0
                actual_check_executions = ''
            }
        }

        $bucket = $combinedByMonth[$row.month]
        $bucket.resource_check_count++
        $bucket.count_instance_count += [int]$row.count_instance_count
        $bucket.estimated_check_executions += [int]$row.estimated_check_executions
        if ($row.actual_check_executions -ne '') {
            if ($bucket.actual_check_executions -eq '') {
                $bucket.actual_check_executions = 0
            }
            $bucket.actual_check_executions += [int]$row.actual_check_executions
        }
    }

    $monthlyTotals = @($totalsByKey.GetEnumerator() | Sort-Object Name | ForEach-Object { $_.Value })
    $monthlyTotals += @($combinedByMonth.GetEnumerator() | Sort-Object Name | ForEach-Object { $_.Value })

    $datasourceSummary = @($allDatasources | ForEach-Object {
        [pscustomobject]@{
            device_id = $_.deviceId
            device_display_name = $_.deviceDisplayName
            device_datasource_id = $_.deviceDatasourceId
            datasource_id = $_.dataSourceId
            datasource_name = $_.dataSourceName
            instance_number = $_.instanceNumber
        }
    })

    $metadata = [pscustomobject]@{
        generated_at_utc = ([DateTime]::UtcNow.ToString('o'))
        mode = $Mode
        output_format = $OutputFormat
        months = @($monthWindows | ForEach-Object { $_.Month })
        uptime_device_count = $uptimeDevices.Count
        logical_check_count = $logicalChecks.Count
        uptime_datasource_count = $allDatasources.Count
        uptime_instance_count = $allInstances.Count
        device_group_count = $deviceGroups.Count
        type = $Type
        is_internal = if ($PSBoundParameters.ContainsKey('IsInternal')) { $IsInternal } else { $null }
        filter = $Filter
        count_datasource_names = $CountDatasourceNames
        polling_interval_properties = $PollingIntervalProperties
        default_interval_minutes = $DefaultIntervalMinutes
        interval_unit = $IntervalUnit
        include_disabled_instances = [bool]$IncludeDisabledInstances
        skip_datasource_inventory = [bool]$SkipDatasourceInventory
        actual_window_hours = $ActualWindowHours
        actual_aggregation_type = $ActualAggregationType
        actual_period = $ActualPeriod
        warnings = @($warnings)
    }

    $exportPayload = [pscustomobject]@{
        metadata = $metadata
        uptime_devices = $uptimeDevices
        device_groups = $deviceGroups
        uptime_device_datasources = $allDatasources
        uptime_datasource_instances = $allInstances
        monthly_counts = $monthlyRows
        monthly_totals = $monthlyTotals
    }

    if ($OutputFormat -in @('Csv', 'Both')) {
        Export-FlatCsvFile -Path (Join-Path $resolvedOutputDir 'uptime_devices.csv') -Rows $uptimeDevices
        Export-FlatCsvFile -Path (Join-Path $resolvedOutputDir 'device_groups.csv') -Rows $deviceGroups
        Export-FlatCsvFile -Path (Join-Path $resolvedOutputDir 'uptime_device_datasources.csv') -Rows $allDatasources
        Export-FlatCsvFile -Path (Join-Path $resolvedOutputDir 'uptime_datasource_instances.csv') -Rows $allInstances
        Write-CsvSafe -Path (Join-Path $resolvedOutputDir 'monthly_uptime_test_counts.csv') -Rows $monthlyRows
        Write-CsvSafe -Path (Join-Path $resolvedOutputDir 'monthly_uptime_totals.csv') -Rows $monthlyTotals
        Write-CsvSafe -Path (Join-Path $resolvedOutputDir 'uptime_datasource_summary.csv') -Rows $datasourceSummary
        Export-MetadataCsv -Path (Join-Path $resolvedOutputDir 'uptime_run_metadata.csv') -Metadata $metadata
        Export-WarningsCsv -Path (Join-Path $resolvedOutputDir 'uptime_warnings.csv') -Warnings @($warnings)
    }

    if ($OutputFormat -in @('Json', 'Both')) {
        Export-JsonFile -Path (Join-Path $resolvedOutputDir 'uptime_powershell_export.json') -InputObject $exportPayload
        Export-JsonFile -Path (Join-Path $resolvedOutputDir 'uptime_devices.json') -InputObject $uptimeDevices
        Export-JsonFile -Path (Join-Path $resolvedOutputDir 'device_groups.json') -InputObject $deviceGroups
        Export-JsonFile -Path (Join-Path $resolvedOutputDir 'uptime_device_datasources.json') -InputObject $allDatasources
        Export-JsonFile -Path (Join-Path $resolvedOutputDir 'uptime_datasource_instances.json') -InputObject $allInstances
        Export-JsonFile -Path (Join-Path $resolvedOutputDir 'uptime_run_metadata.json') -InputObject $metadata
    }

    Write-Host ("Export complete: {0}" -f $resolvedOutputDir)
    if ($warnings.Count -gt 0) {
        $warningTarget = if ($OutputFormat -in @('Csv', 'Both')) { 'uptime_warnings.csv' } else { 'uptime_run_metadata.json' }
        Write-Warning ("Completed with {0} warning(s). See {1}." -f $warnings.Count, $warningTarget)
    }
}
finally {
    if ($DisconnectWhenDone -and -not $SkipConnect) {
        Disconnect-LMAccount -ErrorAction SilentlyContinue
    }
}
