. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\Common.ps1"

function RescheduleTask()
{
    $task = Get-ScheduledTask -TaskName 'SetOneDriveConditionalAccessPolicy'
    $task.Triggers | ForEach-Object {
        $_.StartBoundary = (Get-Date).AddMinutes(10).ToString('o')
    }
    $task | Set-ScheduledTask
}

function ReadFileIntoQueue([System.Collections.Generic.Queue[string]]$Queue)
{
    $queueFile = "$PSScriptRoot\onedrive.queue"
    if (Test-Path -Path $queueFile) {
        $Queue.Clear()
        Get-Content -Path $queueFile | ForEach-Object {$Queue.Enqueue($_)}
        Remove-Item -Path $queueFile -Force
    }
}

function WriteQueueToFile([System.Collections.Generic.Queue[string]]$Queue)
{
    if ($Queue.Count -gt 0) {
        $queueFile = "$PSScriptRoot\onedrive.queue"
        $Queue.GetEnumerator() | Set-Content -Path $queueFile -Force
    }
}

function WriteLogEntry([string]$Entry, [string]$Type)
{
    $logFile = "$PSScriptRoot\log.txt"
    $time = Get-Date -Format 'o'
    $padding = [string]::new(' ', 11 - $Type.Length)
    "$time $Type $padding $Entry" | Out-File -FilePath $logFile -Append
}

function HandleException($Queue, $Url, $ErrorRecord)
{
    WriteQueueToFile -Queue $Queue
    $logEntry = $Url + ' : ' + $ErrorRecord.ToString()
    WriteLogEntry -Entry $logEntry -Type 'Error'
    RescheduleTask
    exit
}

ConnectPnP
$queue = New-Object -TypeName 'System.Collections.Generic.Queue[string]'
ReadFileIntoQueue -Queue $queue
if ($queue.Count -eq 0) {
    Get-PnPTenantSite -IncludeOneDriveSites | Where-Object Url -like $Script:Config.OneDriveUrlFilter | Foreach-Object {
        $queue.Enqueue($_.Url)
    }
}
$count = 0
while ($count -lt 1000 -and $queue.Count -gt 0) {
    $siteUrl = $queue.Dequeue()
    $count++
    try {
        $siteDetails = Get-PnPTenantSite -Url $siteUrl
    }
    catch {
        HandleException -Queue $queue -Url $siteUrl -ErrorRecord $_
    }
    if ($siteDetails.ConditionalAccessPolicy -ne $Script:Config.ConditionalAccessPolicy)
    {
        try {
            $siteDetails.ConditionalAccessPolicy = $Script:Config.ConditionalAccessPolicy
            $null = $siteDetails.Update()
            $Script:PnPContext.Load($siteDetails)
            Invoke-PnPQuery
            WriteLogEntry -Entry ('Updated site ' + $siteDetails.Url) -Type 'Information'
        }
        catch {
            HandleException -Queue $queue -Url $siteUrl -ErrorRecord $_
        }
    }
}
WriteQueueToFile -Queue $queue
RescheduleTask
