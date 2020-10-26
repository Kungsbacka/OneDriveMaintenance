$Script:Config = @{
    # SPO admin ULR
    SharePointAdminUrl = 'https://tenant-admin.sharepoint.com'
    # String to filter out OneDrive URLs from other sites
    OneDriveUrlFilter = 'https://tenat-my.sharepoint.com/personal/*'
    # Name of your tenant
    TenantName = 'tenant.onmicrosoft.com'
    # Application ID (Client ID)
    AppRegistrationId = '[App ID GUID]'
    # Certificate in PEM format
    PEMCertificate = '-----BEGIN CERTIFICATE-----...-----END CERTIFICATE-----'
    # Private key in PEM format and encrypted with DPAPI (ConvertTo-SecureString)
    PEMPrivateKey = ''
    # Connection string to inventory database
    ConnectionString = 'Server=dbserver.contoso.com;Database=Inventory;Integrated Security=True'
    # This is the Conditional Access Policy that is set on all OneDrives
    ConditionalAccessPolicy = 'AllowLimitedAccess'
}