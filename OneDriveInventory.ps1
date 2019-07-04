. "$PSScriptRoot\Config.dev.ps1"
. "$PSScriptRoot\Common.ps1"

function GetInventoryTable()
{
    $table = New-Object -TypeName 'System.Data.DataTable' -ArgumentList @('Office365OneDriveInventory_stage')
    [void]$table.Columns.Add('id', 'int')
    [void]$table.Columns.Add('url', 'string')
    [void]$table.Columns.Add('owner', 'string')
    [void]$table.Columns.Add('storageMaximumLevel', 'long')
    [void]$table.Columns.Add('storageWarningLevel', 'long')
    [void]$table.Columns.Add('storageUsage', 'long')
    @(,$table)
}

function ExecuteStoredProcedure([string]$procedure)
{
    $conn = New-Object -TypeName 'System.Data.SqlClient.SqlConnection'
    $conn.ConnectionString = $Script:Config.ConnectionString
    $conn.Open()
    $cmd = New-Object -TypeName 'System.Data.SqlClient.SqlCommand'
    $cmd.Connection = $conn
    $cmd.CommandType = 'StoredProcedure'
    $cmd.CommandText = $procedure
    [void]$cmd.ExecuteNonQuery()
    $cmd.Dispose()
    $conn.Dispose()
}

function StoreInventory($table)
{
    if ($table.Rows.Count -eq 0) {
        return
    }
    $conn = New-Object -TypeName 'System.Data.SqlClient.SqlConnection'
    $conn.ConnectionString = $Script:Config.ConnectionString
    $conn.Open()
    $bulkCopy = New-Object -TypeName 'System.Data.SqlClient.SqlBulkCopy' -ArgumentList @($conn)
    $bulkCopy.DestinationTableName = 'Office365OneDriveInventory_stage'
    $bulkCopy.WriteToServer($table)
    $bulkCopy.Dispose()
    $conn.Dispose()
    $table.rows.Clear()
}

ConnectPnP
$allSites = Get-PnPTenantSite -IncludeOneDriveSites
ExecuteStoredProcedure 'dbo.spPrepareOneDriveInventory'
$inventoryTable = GetInventoryTable
foreach ($site in $allSites) {
    if ($site.Url -notlike $Script:Config.OneDriveUrlFilter) {
        continue
    }
    $row = $inventoryTable.NewRow()
    $row['url'] = $site.Url
    $row['owner'] = $site.Owner
    $row['storageMaximumLevel'] = $site.StorageMaximumLevel
    $row['storageWarningLevel'] = $site.StorageWarningLevel
    $row['storageUsage'] = $site.StorageUsage
    $inventoryTable.Rows.Add($row)
}
StoreInventory $inventoryTable
ExecuteStoredProcedure 'dbo.spCommitOneDriveInventory'
