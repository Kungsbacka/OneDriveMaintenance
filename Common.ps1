$ErrorActionPreference = 'Stop'

Import-Module -Name SharePointPnPPowerShellOnline -DisableNameChecking

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
