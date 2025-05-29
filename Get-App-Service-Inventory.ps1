<#
  Export-AppServiceSizing.ps1  (v2025-05-29-p2)
  ▸ Collects App/Plan inventory, autoscale, stacks, networking,
    domain bindings — writes one XLSX with multiple sheets.
  ▸ Collects metrics data from Log Analytics (response time, CPU, memory).
  ▸ Handles >1 000 ARG rows via paging (Skip + First=1000).
#>

[CmdletBinding()]
param(
    [string]   $WorkspacePath = ".\AppServiceSizing_{0:yyyyMMdd}.xlsx" -f (Get-Date),
    [string[]] $Subscriptions,
    [string]   $AccountId,
    [string]   $TenantId,
    [string]   $LogAnalyticsWorkspaceId
)

Import-Module Az.Accounts           -EA Stop
Import-Module Az.ResourceGraph      -EA Stop
Import-Module Az.OperationalInsights -EA Stop
Import-Module ImportExcel           -EA Stop

#── Sign-in ──────────────────────────────────────────────────────────────
$connect = @{ ErrorAction = 'Stop' }
if ($AccountId) { $connect.AccountId = $AccountId }
if ($TenantId ) { $connect.TenantId  = $TenantId  }
if (-not (Get-AzContext)) { Connect-AzAccount @connect | Out-Null }

#── Helper: run ARG with paging, return one DataTable ────────────────────
function Invoke-ArgQuery {
    param([string]$Name,[string]$Query)

    Write-Host "   • $Name ..."
    $batch = 1000
    $skip  = 0
    $merged = $null

    while ($true) {
        $p = @{ Query = $Query; First = $batch }
        if ($skip -gt 0) { $p.Skip = $skip }
        if ($Subscriptions) { $p.Subscription = $Subscriptions }

        $page = Search-AzGraph @p
        if ($null -eq $page -or $null -eq $page.Data) {
            Write-Warning "No data returned for query: $Name"
            break
        }
        
        if (-not $merged) {
            # Handle different return types from Search-AzGraph
            if ($page.Data -is [System.Data.DataTable]) {
                # For DataTable, copy the first page completely
                $merged = $page.Data.Copy()
            } else {
                # For List or other collection types, convert to array
                $merged = @($page.Data)
            }
        } else {
            if ($merged -is [System.Data.DataTable] -and $page.Data -is [System.Data.DataTable]) {
                $merged.Merge($page.Data)
            } else {
                # For array/list types, combine using array addition
                $merged = @($merged) + @($page.Data)
            }
        }

        if ($merged -is [System.Data.DataTable]) {
            $fetched = $page.Data.Rows.Count
        } else {
            $fetched = @($page.Data).Count
        }
        if ($fetched -lt $batch) { break }
        $skip += $batch
    }

    return $merged
}

#── Helper: run Log Analytics query, return DataTable ────────────────────
function Invoke-MetricsQuery {
    param([string]$Name,[string]$Query)

    Write-Host "   📊 $Name ..."
    
    if (-not $LogAnalyticsWorkspaceId) {
        Write-Host "      ⚠️  Skipped (no workspace ID provided)"
        return $null
    }

    try {
        $result = Invoke-AzOperationalInsightsQuery -WorkspaceId $LogAnalyticsWorkspaceId -Query $Query
        if ($result -and $result.Results) {
            Write-Host "      ✓ Found $($result.Results.Count) records"
            return $result.Results
        } else {
            Write-Host "      ℹ️  No data returned"
            return $null
        }
    }
    catch {
        Write-Warning "Error executing metrics query '$Name': $($_.Exception.Message)"
        Write-Host "      💡 Tip: Ensure you have Log Analytics Reader permissions on the workspace"
        return $null
    }
}

#── 1  Apps ──────────────────────────────────────────────────────────────
$appInv = Invoke-ArgQuery Apps @'
Resources
| where type =~ "microsoft.web/sites"
| extend PlanId = tostring(properties.serverFarmId)
| extend OS = iff(tobool(properties.reserved), "Linux", "Windows")
| project subscriptionId, resourceGroup, name, OS, location, PlanId
'@

#── 2  Plans ─────────────────────────────────────────────────────────────
$planInv = Invoke-ArgQuery Plans @'
Resources
| where type =~ "microsoft.web/serverfarms"
| project subscriptionId, resourceGroup, Plan = name, SKU = sku.name,
         Region = location,
         NumberOfSites = toint(properties.numberOfSites),
         Workers       = toint(properties.numberOfWorkers),
         ZoneRedundant = properties.zoneRedundant
'@

#── 3  Autoscale ─────────────────────────────────────────────────────────
$auto = Invoke-ArgQuery Autoscale @'
Resources
| where type =~ "microsoft.insights/autoscalesettings"
| where properties.targetResourceUri has "/serverfarms/"
| extend TargetPlan = tostring(split(properties.targetResourceUri,"/")[8]),
         Min  = toint(properties.profiles[0].capacity.minimum),
         Max  = toint(properties.profiles[0].capacity.maximum),
         RuleCount = array_length(properties.profiles)
| project subscriptionId, TargetPlan, Min, Max, RuleCount
'@

#── 4  Stacks ────────────────────────────────────────────────────────────
$stack = Invoke-ArgQuery Stacks @'
Resources
| where type =~ "microsoft.web/sites"
| project subscriptionId, resourceGroup, name,
         Kind  = kind,
         Stack = properties.siteConfig.linuxFxVersion,
         NetFx = properties.siteConfig.netFrameworkVersion
'@

#── 5  Networking ────────────────────────────────────────────────────────
$net = Invoke-ArgQuery Networking @'
Resources
| where type =~ "microsoft.web/sites"
| project subscriptionId, resourceGroup, name,
         VNetSubnet   = properties.virtualNetworkSubnetId,
         PrivateEndpt = tostring(properties.privateEndpointConnections[0].id)
'@

#── 6  Domains ───────────────────────────────────────────────────────────
$domains = Invoke-ArgQuery Domains @'
Resources
| where type =~ "microsoft.web/sites/hostNameBindings"
| extend App = tostring(split(id,"/")[8])
| project subscriptionId, App, Host = name,
         SslState = properties.sslState, Thumbprint = properties.thumbprint
'@

#── Metrics queries ──────────────────────────────────────────────────────
Write-Host "`n🔍  Collecting metrics from Log Analytics..."

# Define metrics queries for easy maintenance
$metricsQueries = [ordered]@{
    'ResponseTime' = @{
        Name = 'Average Response Time by App per Day'
        Query = @'
AzureMetrics
| where TimeGenerated > ago(30d)
| where ResourceProvider == "MICROSOFT.WEB" and _ResourceId has "/sites/"
| where MetricName == "HttpResponseTime"
| extend parts = split(_ResourceId,'/')
| extend SubId = tostring(parts[2]),
         RG    = tostring(parts[4]),
         AppName = tostring(parts[8])
| summarize AvgRespSec = avg(Average) by bin(TimeGenerated,1d), SubId, RG, AppName
| extend AvgRespMs = AvgRespSec * 1000
| project TimeGenerated, SubId, RG, AppName, AvgRespMs
'@
    }
    'CpuSeconds' = @{
        Name = 'CPU Seconds per App per Day'
        Query = @'
AzureMetrics
| where TimeGenerated > ago(30d)
| where ResourceProvider == "MICROSOFT.WEB"
| where _ResourceId has "/sites/"
| where MetricName == "CpuTime"
| extend parts = split(_ResourceId,'/')
| extend SubId        = tostring(parts[2]),
         RG           = tostring(parts[4]),
         AppName      = tostring(parts[8])
| summarize CpuSeconds = sum(Total) by bin(TimeGenerated,1d), SubId, RG, AppName
| order by AppName, TimeGenerated
'@
    }
    'MemoryWorkingSet' = @{
        Name = 'Average Working Set per App per Day'
        Query = @'
AzureMetrics
| where TimeGenerated > ago(30d)
| where ResourceProvider == "MICROSOFT.WEB" and _ResourceId has "/sites/"
| where MetricName == "AverageMemoryWorkingSet"
| extend parts = split(_ResourceId,'/')
| extend SubId = tostring(parts[2]),
         RG    = tostring(parts[4]),
         AppName = tostring(parts[8])
| summarize AvgMemBytes = avg(Average) by bin(TimeGenerated,1d), SubId, RG, AppName
| extend AvgMemMiB = AvgMemBytes / 1024 / 1024
| project TimeGenerated, SubId, RG, AppName, AvgMemMiB
'@
    }
    'CpuMemoryPct' = @{
        Name = 'CPU and Memory % by App per Day'
        Query = @'
let base =
    AzureMetrics
    | where TimeGenerated > ago(30d)
    | where ResourceProvider == "MICROSOFT.WEB"
    | where _ResourceId has "/serverfarms/"
    | where MetricName in ("CpuPercentage","MemoryPercentage");
base
| extend parts = split(_ResourceId,'/')
| extend SubId     = tostring(parts[2]),
         RG        = tostring(parts[4]),
         PlanName  = tostring(parts[8])
| summarize AvgPct = avg(Average)
          by bin(TimeGenerated,1d), SubId, RG, PlanName, MetricName
| evaluate pivot(MetricName, any(AvgPct))
| project TimeGenerated, SubId, RG, PlanName,
          CpuPct = todouble(CpuPercentage),
          MemPct = todouble(MemoryPercentage)
'@
    }
}

# Execute metrics queries
$metricsResults = [ordered]@{}
foreach ($key in $metricsQueries.Keys) {
    $queryInfo = $metricsQueries[$key]
    $result = Invoke-MetricsQuery $queryInfo.Name $queryInfo.Query
    $metricsResults[$key] = $result
}

#── Excel output ─────────────────────────────────────────────────────────
Write-Host "💾  Writing $WorkspacePath ..."
Remove-Item $WorkspacePath -EA SilentlyContinue

$tables = [ordered]@{
    Apps       = $appInv
    Plans      = $planInv
    Autoscale  = $auto
    Stacks     = $stack
    Networking = $net
    Domains    = $domains
}

# Add metrics results to tables
foreach ($key in $metricsResults.Keys) {
    if ($metricsResults[$key]) {
        $tables[$key] = $metricsResults[$key]
    }
}

$first = $true
foreach ($sheet in $tables.Keys) {
    $dt = $tables[$sheet]
    if ($dt) {
        $rowCount = if ($dt -is [System.Data.DataTable]) { $dt.Rows.Count } else { @($dt).Count }
        if ($rowCount -gt 0) {
            $dt | Export-Excel $WorkspacePath `
                -WorksheetName $sheet `
                -TableName     $sheet `
                -AutoSize      `
                -FreezeTopRow  `
                -Append:(-not $first)
            $first = $false
            Write-Host ("      ✓ {0,-11} {1,6} rows" -f $sheet,$rowCount)
        } else {
            Write-Host ("      • Skipped {0} (empty)" -f $sheet)
        }
    } else {
        Write-Host ("      • Skipped {0} (empty)" -f $sheet)
    }
}

Write-Host "`n✅  Done — open '$WorkspacePath'"
