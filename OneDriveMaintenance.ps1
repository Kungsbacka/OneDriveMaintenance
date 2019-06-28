. "$PSScriptRoot\Config.ps1"

function ConnectPnP() {
    $encryptedKey = $Script:Config.PEMPrivateKey
    $secureKey = ConvertTo-SecureString $encryptedKey
    $unsecureKey = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey)
    $PEMPrivateKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($unsecureKey)
    $params = @{
        Url = $Script:Config.SharePointAdminUrl
        Tenant = $Script:Config.TenantName
        ClientId = $Script:Config.AppRegistrationId
        PEMCertificate = $Script:Config.PEMCertificate
        PEMPrivateKey = $PEMPrivateKey
    }
    Connect-PnPOnline @params
    $Script:PnPContext = Get-PnPContext
}

function GetInventoryTable()
{
    $table = New-Object -TypeName 'System.Data.DataTable' -ArgumentList @('Office365OneDriveInventory_stage')
    [void]$table.Columns.Add('id', 'int')
    [void]$table.Columns.Add('url', 'string')
    [void]$table.Columns.Add('owner', 'string')
    [void]$table.Columns.Add('template', 'string')
    [void]$table.Columns.Add('status', 'string')
    [void]$table.Columns.Add('storageMaximumLevel', 'long')
    [void]$table.Columns.Add('storageWarningLevel', 'long')
    [void]$table.Columns.Add('storageUsage', 'long')
    [void]$table.Columns.Add('conditionalAccessPolicy', 'string')
    [void]$table.Columns.Add('limitedAccessFileType', 'string')
    [void]$table.Columns.Add('previousConditionalAccessPolicy', 'string')
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
$oneDriveSites = New-Object -TypeName 'System.Collections.Generic.Stack[string]'
$allSites | Where-Object Url -like $Script:Config.OneDriveUrlFilter | Foreach-Object {$oneDriveSites.Push($_.Url)}
ExecuteStoredProcedure 'dbo.spPrepareOneDriveInventory'
$inventoryTable = GetInventoryTable
while ($oneDriveSites.Count -gt 0) {
    $siteUrl = $oneDriveSites.Pop()
    try {
        $siteDetails = Get-PnPTenantSite -Url $siteUrl
    }
    catch {
        # This is a long running job and we might need to re-authenticate
        $oneDriveSites.Push($siteUrl)
        ConnectPnP
        continue
    }
    $previousConditionalAccessPolicy = $siteDetails.ConditionalAccessPolicy
    if ($siteDetails.ConditionalAccessPolicy -ne 'AllowLimitedAccess')
    {
        try {
            $siteDetails.ConditionalAccessPolicy = 'AllowLimitedAccess'
            $null = $siteDetails.Update()
            $Script:PnPContext.Load($siteDetails)
            Invoke-PnPQuery
        }
        catch {
            $oneDriveSites.Push($siteUrl)
            ConnectPnP
            continue
        }
    }
    $row = $inventoryTable.NewRow()
    $row['url'] = $siteDetails.Url
    $row['owner'] = $siteDetails.Owner
    $row['template'] = $siteDetails.Template
    $row['status'] = $siteDetails.Status
    $row['storageMaximumLevel'] = $siteDetails.StorageMaximumLevel
    $row['storageWarningLevel'] = $siteDetails.StorageWarningLevel
    $row['storageUsage'] = $siteDetails.StorageUsage
    $row['conditionalAccessPolicy'] = $siteDetails.ConditionalAccessPolicy
    $row['limitedAccessFileType'] = $siteDetails.LimitedAccessFileType
    $row['previousConditionalAccessPolicy'] = $previousConditionalAccessPolicy
    $inventoryTable.Rows.Add($row)
    if ($oneDriveSites.Count % 500 -eq 0) {
        StoreInventory $inventoryTable
    }
}
StoreInventory $inventoryTable
ExecuteStoredProcedure 'dbo.spCommitOneDriveInventory'
